const handleSendDraft = require('../../email/send-draft');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('handleSendDraft', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('successful draft sending', () => {
    test('should send a draft successfully', async () => {
      const mockDraft = {
        id: 'draft-123',
        subject: 'Test Draft',
        toRecipients: [{ emailAddress: { address: 'test@example.com' } }],
        ccRecipients: [],
        bccRecipients: []
      };

      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI
        .mockResolvedValueOnce(mockDraft) // GET draft details
        .mockResolvedValueOnce({}); // POST send

      const result = await handleSendDraft({ id: 'draft-123' });

      expect(ensureAuthenticated).toHaveBeenCalledTimes(1);
      expect(callGraphAPI).toHaveBeenCalledTimes(2);
      expect(callGraphAPI).toHaveBeenNthCalledWith(
        1,
        mockAccessToken,
        'GET',
        'me/messages/draft-123'
      );
      expect(callGraphAPI).toHaveBeenNthCalledWith(
        2,
        mockAccessToken,
        'POST',
        'me/messages/draft-123/send',
        {}
      );
      expect(result.content[0].text).toContain('Draft sent successfully');
      expect(result.content[0].text).toContain('Test Draft');
      expect(result.content[0].text).toContain('Recipients: 1');
    });

    test('should display CC and BCC counts', async () => {
      const mockDraft = {
        id: 'draft-456',
        subject: 'Multi-recipient Draft',
        toRecipients: [
          { emailAddress: { address: 'one@example.com' } },
          { emailAddress: { address: 'two@example.com' } }
        ],
        ccRecipients: [{ emailAddress: { address: 'cc@example.com' } }],
        bccRecipients: [
          { emailAddress: { address: 'bcc1@example.com' } },
          { emailAddress: { address: 'bcc2@example.com' } }
        ]
      };

      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI
        .mockResolvedValueOnce(mockDraft)
        .mockResolvedValueOnce({});

      const result = await handleSendDraft({ id: 'draft-456' });

      expect(result.content[0].text).toContain('Recipients: 2 + 1 CC + 2 BCC');
    });

    test('should handle draft with no subject', async () => {
      const mockDraft = {
        id: 'draft-no-subject',
        toRecipients: [{ emailAddress: { address: 'test@example.com' } }],
        ccRecipients: [],
        bccRecipients: []
      };

      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI
        .mockResolvedValueOnce(mockDraft)
        .mockResolvedValueOnce({});

      const result = await handleSendDraft({ id: 'draft-no-subject' });

      expect(result.content[0].text).toContain('(no subject)');
    });
  });

  describe('validation', () => {
    test('should require draft ID', async () => {
      const result = await handleSendDraft({});

      expect(result.content[0].text).toContain('Draft ID is required');
      expect(callGraphAPI).not.toHaveBeenCalled();
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

      const result = await handleSendDraft({ id: 'draft-123' });

      expect(result.content[0].text).toBe(
        "Authentication required. Please use the 'authenticate' tool first."
      );
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('should handle draft not found error', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockRejectedValue(new Error('API call failed with status 404'));

      const result = await handleSendDraft({ id: 'nonexistent-draft' });

      expect(result.content[0].text).toContain('Draft not found');
      expect(result.content[0].text).toContain('may have been deleted or already sent');
    });

    test('should handle other Graph API errors', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockRejectedValue(new Error('Network error'));

      const result = await handleSendDraft({ id: 'draft-123' });

      expect(result.content[0].text).toBe('Error sending draft: Network error');
    });
  });
});
