const handleSaveDraft = require('../../email/save-draft');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('handleSaveDraft', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('creating new drafts', () => {
    test('should create a new draft with all fields', async () => {
      const mockDraft = {
        id: 'draft-123',
        subject: 'Test Subject',
        toRecipients: [{ emailAddress: { address: 'test@example.com' } }],
        ccRecipients: [],
        bccRecipients: []
      };

      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue(mockDraft);

      const result = await handleSaveDraft({
        to: 'test@example.com',
        subject: 'Test Subject',
        body: 'Test body content'
      });

      expect(ensureAuthenticated).toHaveBeenCalledTimes(1);
      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'me/messages',
        expect.objectContaining({
          subject: 'Test Subject',
          body: {
            contentType: 'text',
            content: 'Test body content'
          },
          toRecipients: [{ emailAddress: { address: 'test@example.com' } }]
        })
      );
      expect(result.content[0].text).toContain('Draft created successfully');
      expect(result.content[0].text).toContain('draft-123');
    });

    test('should use explicit HTML content type when specified', async () => {
      const mockDraft = { id: 'draft-456', subject: 'HTML Email' };

      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue(mockDraft);

      await handleSaveDraft({
        subject: 'HTML Email',
        body: '<p>Hello</p>',
        contentType: 'html'
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'me/messages',
        expect.objectContaining({
          body: {
            contentType: 'html',
            content: '<p>Hello</p>'
          }
        })
      );
    });

    test('should default to text content type', async () => {
      const mockDraft = { id: 'draft-789' };

      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue(mockDraft);

      await handleSaveDraft({
        body: 'Plain text content'
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'me/messages',
        expect.objectContaining({
          body: {
            contentType: 'text',
            content: 'Plain text content'
          }
        })
      );
    });

    test('should handle multiple recipients', async () => {
      const mockDraft = {
        id: 'draft-multi',
        toRecipients: [
          { emailAddress: { address: 'one@example.com' } },
          { emailAddress: { address: 'two@example.com' } }
        ],
        ccRecipients: [{ emailAddress: { address: 'cc@example.com' } }],
        bccRecipients: [{ emailAddress: { address: 'bcc@example.com' } }]
      };

      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue(mockDraft);

      const result = await handleSaveDraft({
        to: 'one@example.com, two@example.com',
        cc: 'cc@example.com',
        bcc: 'bcc@example.com',
        subject: 'Multi-recipient email'
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'me/messages',
        expect.objectContaining({
          toRecipients: [
            { emailAddress: { address: 'one@example.com' } },
            { emailAddress: { address: 'two@example.com' } }
          ],
          ccRecipients: [{ emailAddress: { address: 'cc@example.com' } }],
          bccRecipients: [{ emailAddress: { address: 'bcc@example.com' } }]
        })
      );
      expect(result.content[0].text).toContain('To: 2 recipient(s)');
      expect(result.content[0].text).toContain('CC: 1 recipient(s)');
      expect(result.content[0].text).toContain('BCC: 1 recipient(s)');
    });

    test('should create draft with only subject', async () => {
      const mockDraft = { id: 'draft-subject-only', subject: 'Just a subject' };

      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue(mockDraft);

      const result = await handleSaveDraft({
        subject: 'Just a subject'
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'me/messages',
        expect.objectContaining({
          subject: 'Just a subject'
        })
      );
      expect(result.content[0].text).toContain('Draft created successfully');
    });

    test('should set importance when provided', async () => {
      const mockDraft = { id: 'draft-important' };

      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue(mockDraft);

      await handleSaveDraft({
        subject: 'Important email',
        importance: 'high'
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'POST',
        'me/messages',
        expect.objectContaining({
          importance: 'high'
        })
      );
    });
  });

  describe('updating existing drafts', () => {
    test('should update an existing draft', async () => {
      const mockUpdatedDraft = {
        id: 'existing-draft-123',
        subject: 'Updated Subject'
      };

      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue(mockUpdatedDraft);

      const result = await handleSaveDraft({
        id: 'existing-draft-123',
        subject: 'Updated Subject'
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'PATCH',
        'me/messages/existing-draft-123',
        expect.objectContaining({
          subject: 'Updated Subject'
        })
      );
      expect(result.content[0].text).toContain('Draft updated successfully');
    });

    test('should require at least one field when updating', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);

      const result = await handleSaveDraft({
        id: 'draft-123'
      });

      expect(callGraphAPI).not.toHaveBeenCalled();
      expect(result.content[0].text).toContain('At least one field');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

      const result = await handleSaveDraft({
        subject: 'Test'
      });

      expect(result.content[0].text).toBe(
        "Authentication required. Please use the 'authenticate' tool first."
      );
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('should handle Graph API error', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleSaveDraft({
        subject: 'Test'
      });

      expect(result.content[0].text).toBe('Error saving draft: Graph API Error');
    });
  });
});
