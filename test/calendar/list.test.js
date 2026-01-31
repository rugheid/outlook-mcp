const handleListEvents = require('../../calendar/list');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');
const { resolveCalendarPath } = require('../../calendar/calendar-utils');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');
jest.mock('../../calendar/calendar-utils');

describe('handleListEvents', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
    resolveCalendarPath.mockClear();
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
  });

  describe('calendarView endpoint', () => {
    test('should use calendarView endpoint instead of events endpoint', async () => {
      resolveCalendarPath.mockResolvedValue('me/calendar');
      callGraphAPI.mockResolvedValue({ value: [] });

      await handleListEvents({ count: 10 });

      const [, , endpoint] = callGraphAPI.mock.calls[0];
      expect(endpoint).toBe('me/calendar/calendarView');
    });

    test('should include startDateTime and endDateTime parameters', async () => {
      resolveCalendarPath.mockResolvedValue('me/calendar');
      callGraphAPI.mockResolvedValue({ value: [] });

      const beforeCall = Date.now();
      await handleListEvents({ count: 10 });
      const afterCall = Date.now();

      const [, , , , queryParams] = callGraphAPI.mock.calls[0];

      expect(queryParams).toHaveProperty('startDateTime');
      expect(queryParams).toHaveProperty('endDateTime');

      // Verify startDateTime is approximately now
      const startDate = new Date(queryParams.startDateTime);
      expect(startDate.getTime()).toBeGreaterThanOrEqual(beforeCall);
      expect(startDate.getTime()).toBeLessThanOrEqual(afterCall);

      // Verify endDateTime is approximately 30 days from now
      const endDate = new Date(queryParams.endDateTime);
      const expectedEnd = new Date(startDate.getTime() + 30 * 24 * 60 * 60 * 1000);
      expect(Math.abs(endDate.getTime() - expectedEnd.getTime())).toBeLessThan(1000);
    });

    test('should not include $filter parameter', async () => {
      resolveCalendarPath.mockResolvedValue('me/calendar');
      callGraphAPI.mockResolvedValue({ value: [] });

      await handleListEvents({ count: 10 });

      const [, , , , queryParams] = callGraphAPI.mock.calls[0];
      expect(queryParams).not.toHaveProperty('$filter');
    });
  });

  describe('calendar parameter', () => {
    test('should use primary calendar when no calendar parameter is provided', async () => {
      resolveCalendarPath.mockResolvedValue('me/calendar');
      callGraphAPI.mockResolvedValue({ value: [] });

      await handleListEvents({ count: 10 });

      expect(resolveCalendarPath).toHaveBeenCalledWith(mockAccessToken, undefined);
    });

    test('should resolve calendar path when calendar parameter is provided', async () => {
      const calendarName = 'Work Calendar';
      resolveCalendarPath.mockResolvedValue('me/calendars/calendar-id-123');
      callGraphAPI.mockResolvedValue({ value: [] });

      await handleListEvents({ count: 10, calendar: calendarName });

      expect(resolveCalendarPath).toHaveBeenCalledWith(mockAccessToken, calendarName);

      const [, , endpoint] = callGraphAPI.mock.calls[0];
      expect(endpoint).toBe('me/calendars/calendar-id-123/calendarView');
    });
  });

  describe('event formatting', () => {
    test('should return formatted events when events are found', async () => {
      resolveCalendarPath.mockResolvedValue('me/calendar');

      const mockEvents = [
        {
          id: 'event1',
          subject: 'Test Meeting',
          start: { dateTime: '2024-03-10T10:00:00', timeZone: 'UTC' },
          end: { dateTime: '2024-03-10T11:00:00', timeZone: 'UTC' },
          location: { displayName: 'Conference Room' },
          bodyPreview: 'Discuss project updates'
        },
        {
          id: 'event2',
          subject: 'Weekly 1:1',
          start: { dateTime: '2024-03-11T14:00:00', timeZone: 'UTC' },
          end: { dateTime: '2024-03-11T15:00:00', timeZone: 'UTC' },
          location: { displayName: 'Teams' },
          bodyPreview: 'Regular sync'
        }
      ];

      callGraphAPI.mockResolvedValue({ value: mockEvents });

      const result = await handleListEvents({ count: 10 });

      expect(result.content[0].text).toContain('Found 2 events');
      expect(result.content[0].text).toContain('Test Meeting');
      expect(result.content[0].text).toContain('Weekly 1:1');
      expect(result.content[0].text).toContain('event1');
      expect(result.content[0].text).toContain('event2');
    });

    test('should return no events message when no events are found', async () => {
      resolveCalendarPath.mockResolvedValue('me/calendar');
      callGraphAPI.mockResolvedValue({ value: [] });

      const result = await handleListEvents({ count: 10 });

      expect(result.content[0].text).toBe('No calendar events found.');
    });
  });

  describe('error handling', () => {
    test('should handle authentication errors', async () => {
      ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

      const result = await handleListEvents({ count: 10 });

      expect(result.content[0].text).toContain('Authentication required');
    });

    test('should handle API errors', async () => {
      resolveCalendarPath.mockResolvedValue('me/calendar');
      callGraphAPI.mockRejectedValue(new Error('API Error'));

      const result = await handleListEvents({ count: 10 });

      expect(result.content[0].text).toContain('Error listing events');
    });
  });

  describe('count parameter', () => {
    test('should respect count parameter', async () => {
      resolveCalendarPath.mockResolvedValue('me/calendar');
      callGraphAPI.mockResolvedValue({ value: [] });

      await handleListEvents({ count: 25 });

      const [, , , , queryParams] = callGraphAPI.mock.calls[0];
      expect(queryParams.$top).toBe(25);
    });

    test('should use default count of 10 when not specified', async () => {
      resolveCalendarPath.mockResolvedValue('me/calendar');
      callGraphAPI.mockResolvedValue({ value: [] });

      await handleListEvents({});

      const [, , , , queryParams] = callGraphAPI.mock.calls[0];
      expect(queryParams.$top).toBe(10);
    });

    test('should enforce maximum count of 50', async () => {
      resolveCalendarPath.mockResolvedValue('me/calendar');
      callGraphAPI.mockResolvedValue({ value: [] });

      await handleListEvents({ count: 100 });

      const [, , , , queryParams] = callGraphAPI.mock.calls[0];
      expect(queryParams.$top).toBe(50);
    });
  });
});
