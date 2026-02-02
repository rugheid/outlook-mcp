/**
 * Email module for Outlook MCP server
 */
const handleListEmails = require('./list');
const handleSearchEmails = require('./search');
const handleReadEmail = require('./read');
const handleSendEmail = require('./send');
const handleSaveDraft = require('./save-draft');
const handleSendDraft = require('./send-draft');
const handleMarkAsRead = require('./mark-as-read');

// Email tool definitions
const emailTools = [
  {
    name: "list-emails",
    description: "Lists recent emails from your inbox",
    inputSchema: {
      type: "object",
      properties: {
        folder: {
          type: "string",
          description: "Email folder to list (e.g., 'inbox', 'sent', 'drafts', default: 'inbox')"
        },
        count: {
          type: "number",
          description: "Number of emails to retrieve (default: 10, max: 50)"
        }
      },
      required: []
    },
    handler: handleListEmails
  },
  {
    name: "search-emails",
    description: "Search for emails using various criteria",
    inputSchema: {
      type: "object",
      properties: {
        query: {
          type: "string",
          description: "Search query text to find in emails"
        },
        folder: {
          type: "string",
          description: "Email folder to search in (default: 'inbox')"
        },
        from: {
          type: "string",
          description: "Filter by sender email address or name"
        },
        to: {
          type: "string",
          description: "Filter by recipient email address or name"
        },
        subject: {
          type: "string",
          description: "Filter by email subject"
        },
        hasAttachments: {
          type: "boolean",
          description: "Filter to only emails with attachments"
        },
        unreadOnly: {
          type: "boolean",
          description: "Filter to only unread emails"
        },
        count: {
          type: "number",
          description: "Number of results to return (default: 10, max: 50)"
        }
      },
      required: []
    },
    handler: handleSearchEmails
  },
  {
    name: "read-email",
    description: "Reads the content of a specific email",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to read"
        }
      },
      required: ["id"]
    },
    handler: handleReadEmail
  },
  {
    name: "send-email",
    description: "Composes and sends a new email",
    inputSchema: {
      type: "object",
      properties: {
        to: {
          type: "string",
          description: "Comma-separated list of recipient email addresses"
        },
        cc: {
          type: "string",
          description: "Comma-separated list of CC recipient email addresses"
        },
        bcc: {
          type: "string",
          description: "Comma-separated list of BCC recipient email addresses"
        },
        subject: {
          type: "string",
          description: "Email subject"
        },
        body: {
          type: "string",
          description: "Email body content"
        },
        contentType: {
          type: "string",
          description: "Body content type: 'text' for plain text, 'html' for HTML (default: 'text')",
          enum: ["text", "html"]
        },
        importance: {
          type: "string",
          description: "Email importance (normal, high, low)",
          enum: ["normal", "high", "low"]
        },
        saveToSentItems: {
          type: "boolean",
          description: "Whether to save the email to sent items"
        }
      },
      required: ["to", "subject", "body"]
    },
    handler: handleSendEmail
  },
  {
    name: "mark-as-read",
    description: "Marks an email as read or unread",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to mark as read/unread"
        },
        isRead: {
          type: "boolean",
          description: "Whether to mark as read (true) or unread (false). Default: true"
        }
      },
      required: ["id"]
    },
    handler: handleMarkAsRead
  },
  {
    name: "save-draft",
    description: "Creates a new email draft, updates an existing one, or creates a reply draft. Use replyToMessageId to create a properly threaded reply. Drafts are saved to the Drafts folder and can be reviewed in Outlook before sending.",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of an existing draft to update. Cannot be used with replyToMessageId."
        },
        replyToMessageId: {
          type: "string",
          description: "ID of a message to reply to. Creates a threaded reply draft with proper conversation linking. Cannot be used with id."
        },
        replyAll: {
          type: "boolean",
          description: "When true, reply to all recipients instead of just the sender. Only used with replyToMessageId. (default: false)"
        },
        to: {
          type: "string",
          description: "Comma-separated list of recipient email addresses. When replying, overrides the auto-filled recipients."
        },
        cc: {
          type: "string",
          description: "Comma-separated list of CC recipient email addresses. When replying, overrides the auto-filled CC recipients."
        },
        bcc: {
          type: "string",
          description: "Comma-separated list of BCC recipient email addresses"
        },
        subject: {
          type: "string",
          description: "Email subject. When replying, overrides the auto-generated 'RE: ...' subject."
        },
        body: {
          type: "string",
          description: "Email body content. When replying, replaces the auto-generated quoted content."
        },
        contentType: {
          type: "string",
          description: "Body content type: 'text' for plain text, 'html' for HTML (default: 'text')",
          enum: ["text", "html"]
        },
        importance: {
          type: "string",
          description: "Email importance (normal, high, low)",
          enum: ["normal", "high", "low"]
        }
      },
      required: []
    },
    handler: handleSaveDraft
  },
  {
    name: "send-draft",
    description: "Sends an existing email draft. Use list-emails with folder 'drafts' to find draft IDs.",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the draft to send"
        }
      },
      required: ["id"]
    },
    handler: handleSendDraft
  }
];

module.exports = {
  emailTools,
  handleListEmails,
  handleSearchEmails,
  handleReadEmail,
  handleSendEmail,
  handleSaveDraft,
  handleSendDraft,
  handleMarkAsRead
};
