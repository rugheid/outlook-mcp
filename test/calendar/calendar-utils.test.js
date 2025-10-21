const {
  resolveCalendarPath,
  getCalendarIdByName,
  getAllCalendars,
  calendarCache
} = require('../../calendar/calendar-utils');
const { callGraphAPI } = require('../../utils/graph-api');

jest.mock('../../utils/graph-api');

describe('resolveCalendarPath', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    // Clear cache
    Object.keys(calendarCache).forEach(key => delete calendarCache[key]);
    // Mock console.error to avoid cluttering test output
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('default calendar', () => {
    test('should return primary calendar path when no calendar name is provided', async () => {
      const result = await resolveCalendarPath(mockAccessToken, null);
      expect(result).toBe('me/calendar');
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('should return primary calendar path when undefined is provided', async () => {
      const result = await resolveCalendarPath(mockAccessToken, undefined);
      expect(result).toBe('me/calendar');
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('should return primary calendar path when empty string is provided', async () => {
      const result = await resolveCalendarPath(mockAccessToken, '');
      expect(result).toBe('me/calendar');
      expect(callGraphAPI).not.toHaveBeenCalled();
    });
  });

  describe('calendar ID', () => {
    test('should use calendar ID directly when it looks like an ID (long string)', async () => {
      // Graph API calendar IDs are typically long base64-encoded strings
      const calendarId = 'AAMkAGI2TG93AAA=AAMkAGI2TG93AAA=AAMkAGI2TG93AAA=AAMkAGI2TG93AAA=';
      const result = await resolveCalendarPath(mockAccessToken, calendarId);

      expect(result).toBe(`me/calendars/${calendarId}`);
      expect(callGraphAPI).not.toHaveBeenCalled();
    });
  });

  describe('calendar name resolution', () => {
    test('should resolve calendar by name when found', async () => {
      const calendarId = 'calendar-id-123';
      const calendarName = 'Work Calendar';

      callGraphAPI.mockResolvedValueOnce({
        value: [
          { id: 'other-id', name: 'Personal Calendar' },
          { id: calendarId, name: calendarName }
        ]
      });

      const result = await resolveCalendarPath(mockAccessToken, calendarName);

      expect(result).toBe(`me/calendars/${calendarId}`);
      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'GET',
        'me/calendars',
        null,
        { $select: 'id,name' }
      );
    });

    test('should handle case-insensitive calendar name matching', async () => {
      const calendarId = 'calendar-id-456';
      const calendarName = 'WORK CALENDAR';

      callGraphAPI.mockResolvedValueOnce({
        value: [
          { id: calendarId, name: 'work calendar' }
        ]
      });

      const result = await resolveCalendarPath(mockAccessToken, calendarName);

      expect(result).toBe(`me/calendars/${calendarId}`);
    });

    test('should fall back to primary calendar when calendar name is not found', async () => {
      const calendarName = 'NonExistent Calendar';

      callGraphAPI.mockResolvedValueOnce({
        value: [
          { id: 'id1', name: 'Calendar 1' },
          { id: 'id2', name: 'Calendar 2' }
        ]
      });

      const result = await resolveCalendarPath(mockAccessToken, calendarName);

      expect(result).toBe('me/calendar');
    });

    test('should fall back to primary calendar when API call fails', async () => {
      const calendarName = 'Work Calendar';

      callGraphAPI.mockRejectedValueOnce(new Error('API Error'));

      const result = await resolveCalendarPath(mockAccessToken, calendarName);

      expect(result).toBe('me/calendar');
    });
  });
});

describe('getCalendarIdByName', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    Object.keys(calendarCache).forEach(key => delete calendarCache[key]);
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('should return calendar ID when exact match is found', async () => {
    const calendarId = 'calendar-id-123';
    const calendarName = 'Test Calendar';

    callGraphAPI.mockResolvedValueOnce({
      value: [{ id: calendarId, name: calendarName }]
    });

    const result = await getCalendarIdByName(mockAccessToken, calendarName);

    expect(result).toBe(calendarId);
    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'GET',
      'me/calendars',
      null,
      { $select: 'id,name' }
    );
  });

  test('should cache calendar lookups', async () => {
    const calendarId = 'calendar-id-789';
    const calendarName = 'Cached Calendar';

    callGraphAPI.mockResolvedValueOnce({
      value: [{ id: calendarId, name: calendarName }]
    });

    // First call should fetch from API
    const result1 = await getCalendarIdByName(mockAccessToken, calendarName);
    expect(result1).toBe(calendarId);
    expect(callGraphAPI).toHaveBeenCalledTimes(1);

    // Second call should use cache
    const result2 = await getCalendarIdByName(mockAccessToken, calendarName);
    expect(result2).toBe(calendarId);
    expect(callGraphAPI).toHaveBeenCalledTimes(1); // Still only 1 call
  });

  test('should return null when calendar is not found', async () => {
    const calendarName = 'NonExistent Calendar';

    callGraphAPI.mockResolvedValueOnce({
      value: [
        { id: 'id1', name: 'Other Calendar' }
      ]
    });

    const result = await getCalendarIdByName(mockAccessToken, calendarName);

    expect(result).toBeNull();
  });

  test('should return null when API call fails', async () => {
    const calendarName = 'Test Calendar';

    callGraphAPI.mockRejectedValueOnce(new Error('API Error'));

    const result = await getCalendarIdByName(mockAccessToken, calendarName);

    expect(result).toBeNull();
  });
});

describe('getAllCalendars', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('should return all calendars', async () => {
    const mockCalendars = [
      { id: 'cal1', name: 'Calendar 1', isDefaultCalendar: true },
      { id: 'cal2', name: 'Calendar 2', isDefaultCalendar: false }
    ];

    callGraphAPI.mockResolvedValueOnce({ value: mockCalendars });

    const result = await getAllCalendars(mockAccessToken);

    expect(result).toEqual(mockCalendars);
    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'GET',
      'me/calendars',
      null,
      {
        $select: 'id,name,canEdit,canShare,canViewPrivateItems,owner,isDefaultCalendar',
        $top: 50
      }
    );
  });

  test('should return empty array when no calendars found', async () => {
    callGraphAPI.mockResolvedValueOnce({ value: null });

    const result = await getAllCalendars(mockAccessToken);

    expect(result).toEqual([]);
  });

  test('should return empty array when API call fails', async () => {
    callGraphAPI.mockRejectedValueOnce(new Error('API Error'));

    const result = await getAllCalendars(mockAccessToken);

    expect(result).toEqual([]);
  });
});
