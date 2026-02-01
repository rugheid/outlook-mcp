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
    jest.spyOn(console, 'error').mockImplementation(() => {});
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    resolveCalendarPath.mockResolvedValue('me/calendar');
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('date range parameters', () => {
    test('should use default date range when not specified', async () => {
      callGraphAPI.mockResolvedValue({ value: [] });

      const beforeCall = Date.now();
      await handleListEvents({});
      const afterCall = Date.now();

      const [, , , , queryParams] = callGraphAPI.mock.calls[0];

      // Verify startDateTime is approximately now
      const startDate = new Date(queryParams.startDateTime);
      expect(startDate.getTime()).toBeGreaterThanOrEqual(beforeCall);
      expect(startDate.getTime()).toBeLessThanOrEqual(afterCall);

      // Verify endDateTime is approximately 30 days from now
      const endDate = new Date(queryParams.endDateTime);
      const expectedEnd = new Date(startDate.getTime() + 30 * 24 * 60 * 60 * 1000);
      expect(Math.abs(endDate.getTime() - expectedEnd.getTime())).toBeLessThan(1000);
    });

    test('should use custom startDateTime when provided', async () => {
      callGraphAPI.mockResolvedValue({ value: [] });

      const customStart = '2024-03-15T14:00:00Z';
      await handleListEvents({ startDateTime: customStart });

      const [, , , , queryParams] = callGraphAPI.mock.calls[0];
      expect(queryParams.startDateTime).toBe(new Date(customStart).toISOString());
    });

    test('should use custom endDateTime when provided', async () => {
      callGraphAPI.mockResolvedValue({ value: [] });

      const customEnd = '2024-03-15T15:00:00Z';
      await handleListEvents({ endDateTime: customEnd });

      const [, , , , queryParams] = callGraphAPI.mock.calls[0];
      expect(queryParams.endDateTime).toBe(new Date(customEnd).toISOString());
    });

    test('should use both custom dates for conflict checking', async () => {
      callGraphAPI.mockResolvedValue({ value: [] });

      const customStart = '2024-03-15T14:00:00Z';
      const customEnd = '2024-03-15T15:00:00Z';
      await handleListEvents({ startDateTime: customStart, endDateTime: customEnd });

      const [, , , , queryParams] = callGraphAPI.mock.calls[0];
      expect(queryParams.startDateTime).toBe(new Date(customStart).toISOString());
      expect(queryParams.endDateTime).toBe(new Date(customEnd).toISOString());
    });
  });

  describe('calendar parameter', () => {
    test('should use primary calendar when no calendar parameter is provided', async () => {
      callGraphAPI.mockResolvedValue({ value: [] });

      await handleListEvents({});

      expect(resolveCalendarPath).toHaveBeenCalledWith(mockAccessToken, undefined);
    });

    test('should resolve calendar path when calendar parameter is provided', async () => {
      const calendarName = 'Work Calendar';
      resolveCalendarPath.mockResolvedValue('me/calendars/calendar-id-123');
      callGraphAPI.mockResolvedValue({ value: [] });

      await handleListEvents({ calendar: calendarName });

      expect(resolveCalendarPath).toHaveBeenCalledWith(mockAccessToken, calendarName);

      const [, , endpoint] = callGraphAPI.mock.calls[0];
      expect(endpoint).toBe('me/calendars/calendar-id-123/calendarView');
    });
  });

  describe('calendarView endpoint', () => {
    test('should use calendarView endpoint', async () => {
      callGraphAPI.mockResolvedValue({ value: [] });

      await handleListEvents({});

      const [, , endpoint] = callGraphAPI.mock.calls[0];
      expect(endpoint).toBe('me/calendar/calendarView');
    });

    test('should include required query parameters', async () => {
      callGraphAPI.mockResolvedValue({ value: [] });

      await handleListEvents({});

      const [, , , , queryParams] = callGraphAPI.mock.calls[0];
      expect(queryParams).toHaveProperty('startDateTime');
      expect(queryParams).toHaveProperty('endDateTime');
      expect(queryParams).toHaveProperty('$top');
      expect(queryParams).toHaveProperty('$orderby', 'start/dateTime');
    });
  });

  describe('event formatting', () => {
    test('should return formatted events', async () => {
      const mockEvents = [
        {
          id: 'event1',
          subject: 'Team Meeting',
          start: { dateTime: '2024-03-15T14:00:00' },
          end: { dateTime: '2024-03-15T15:00:00' },
          location: { displayName: 'Room A' }
        },
        {
          id: 'event2',
          subject: 'Project Review',
          start: { dateTime: '2024-03-15T16:00:00' },
          end: { dateTime: '2024-03-15T17:00:00' },
          location: { displayName: 'Room B' }
        }
      ];

      callGraphAPI.mockResolvedValue({ value: mockEvents });

      const result = await handleListEvents({});

      expect(result.content[0].text).toContain('Found 2 event(s)');
      expect(result.content[0].text).toContain('Team Meeting');
      expect(result.content[0].text).toContain('Project Review');
      expect(result.content[0].text).toContain('event1');
      expect(result.content[0].text).toContain('event2');
    });

    test('should handle events without location', async () => {
      const mockEvents = [
        {
          id: 'event1',
          subject: 'Virtual Meeting',
          start: { dateTime: '2024-03-15T14:00:00' },
          end: { dateTime: '2024-03-15T15:00:00' },
          location: {}
        }
      ];

      callGraphAPI.mockResolvedValue({ value: mockEvents });

      const result = await handleListEvents({});

      expect(result.content[0].text).toContain('No location');
    });

    test('should return no events message with date range', async () => {
      callGraphAPI.mockResolvedValue({ value: [] });

      const result = await handleListEvents({
        startDateTime: '2024-03-15T00:00:00Z',
        endDateTime: '2024-03-15T23:59:59Z'
      });

      expect(result.content[0].text).toContain('No calendar events found');
    });
  });

  describe('count parameter', () => {
    test('should respect count parameter', async () => {
      callGraphAPI.mockResolvedValue({ value: [] });

      await handleListEvents({ count: 25 });

      const [, , , , queryParams] = callGraphAPI.mock.calls[0];
      expect(queryParams.$top).toBe(25);
    });

    test('should default to 10 events', async () => {
      callGraphAPI.mockResolvedValue({ value: [] });

      await handleListEvents({});

      const [, , , , queryParams] = callGraphAPI.mock.calls[0];
      expect(queryParams.$top).toBe(10);
    });

    test('should cap at 50 events', async () => {
      callGraphAPI.mockResolvedValue({ value: [] });

      await handleListEvents({ count: 100 });

      const [, , , , queryParams] = callGraphAPI.mock.calls[0];
      expect(queryParams.$top).toBe(50);
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

      const result = await handleListEvents({});

      expect(result.content[0].text).toContain('Authentication required');
    });

    test('should handle API error', async () => {
      callGraphAPI.mockRejectedValue(new Error('API Error'));

      const result = await handleListEvents({});

      expect(result.content[0].text).toContain('Error listing events');
    });
  });
});
