const handleRespondToEvent = require('../../calendar/respond');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('handleRespondToEvent', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('accepting events', () => {
    test('should accept an event successfully', async () => {
      callGraphAPI.mockResolvedValue({});

      const result = await handleRespondToEvent({
        eventId: 'event-123',
        response: 'accept'
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'me/events/event-123/accept',
        { sendResponse: true }
      );
      expect(result.content[0].text).toContain('accepted successfully');
    });

    test('should accept with comment', async () => {
      callGraphAPI.mockResolvedValue({});

      const result = await handleRespondToEvent({
        eventId: 'event-123',
        response: 'accept',
        comment: 'Looking forward to it!'
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'me/events/event-123/accept',
        { sendResponse: true, comment: 'Looking forward to it!' }
      );
      expect(result.content[0].text).toContain('Looking forward to it!');
    });

    test('should accept without notifying organizer', async () => {
      callGraphAPI.mockResolvedValue({});

      const result = await handleRespondToEvent({
        eventId: 'event-123',
        response: 'accept',
        sendResponse: false
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'me/events/event-123/accept',
        { sendResponse: false }
      );
      expect(result.content[0].text).toContain('Organizer was not notified');
    });
  });

  describe('declining events', () => {
    test('should decline an event successfully', async () => {
      callGraphAPI.mockResolvedValue({});

      const result = await handleRespondToEvent({
        eventId: 'event-456',
        response: 'decline'
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'me/events/event-456/decline',
        { sendResponse: true }
      );
      expect(result.content[0].text).toContain('declined successfully');
    });

    test('should decline with comment', async () => {
      callGraphAPI.mockResolvedValue({});

      const result = await handleRespondToEvent({
        eventId: 'event-456',
        response: 'decline',
        comment: 'Sorry, I have a conflict.'
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'me/events/event-456/decline',
        { sendResponse: true, comment: 'Sorry, I have a conflict.' }
      );
      expect(result.content[0].text).toContain('Sorry, I have a conflict.');
    });
  });

  describe('tentatively accepting events', () => {
    test('should tentatively accept an event', async () => {
      callGraphAPI.mockResolvedValue({});

      const result = await handleRespondToEvent({
        eventId: 'event-789',
        response: 'tentative'
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'me/events/event-789/tentativelyAccept',
        { sendResponse: true }
      );
      expect(result.content[0].text).toContain('tentatively accepted');
    });
  });

  describe('case insensitivity', () => {
    test('should handle uppercase response', async () => {
      callGraphAPI.mockResolvedValue({});

      await handleRespondToEvent({
        eventId: 'event-123',
        response: 'ACCEPT'
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'me/events/event-123/accept',
        expect.any(Object)
      );
    });

    test('should handle mixed case response', async () => {
      callGraphAPI.mockResolvedValue({});

      await handleRespondToEvent({
        eventId: 'event-123',
        response: 'Decline'
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'me/events/event-123/decline',
        expect.any(Object)
      );
    });
  });

  describe('validation', () => {
    test('should require eventId', async () => {
      const result = await handleRespondToEvent({
        response: 'accept'
      });

      expect(callGraphAPI).not.toHaveBeenCalled();
      expect(result.content[0].text).toContain('Event ID is required');
    });

    test('should require response', async () => {
      const result = await handleRespondToEvent({
        eventId: 'event-123'
      });

      expect(callGraphAPI).not.toHaveBeenCalled();
      expect(result.content[0].text).toContain('Response is required');
    });

    test('should reject invalid response type', async () => {
      const result = await handleRespondToEvent({
        eventId: 'event-123',
        response: 'maybe'
      });

      expect(callGraphAPI).not.toHaveBeenCalled();
      expect(result.content[0].text).toContain("Invalid response 'maybe'");
      expect(result.content[0].text).toContain('accept, decline, tentative');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

      const result = await handleRespondToEvent({
        eventId: 'event-123',
        response: 'accept'
      });

      expect(result.content[0].text).toContain('Authentication required');
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('should handle event not found error', async () => {
      callGraphAPI.mockRejectedValue(new Error('API call failed with status 404'));

      const result = await handleRespondToEvent({
        eventId: 'nonexistent-event',
        response: 'accept'
      });

      expect(result.content[0].text).toContain('Event not found');
    });

    test('should handle other API errors', async () => {
      callGraphAPI.mockRejectedValue(new Error('Network error'));

      const result = await handleRespondToEvent({
        eventId: 'event-123',
        response: 'accept'
      });

      expect(result.content[0].text).toContain('Error responding to event');
      expect(result.content[0].text).toContain('Network error');
    });
  });
});
