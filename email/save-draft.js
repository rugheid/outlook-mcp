/**
 * Save draft email functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Parse comma-separated email addresses into Graph API format
 * @param {string} emails - Comma-separated email addresses
 * @returns {Array} - Array of recipient objects
 */
function parseRecipients(emails) {
  if (!emails) return [];
  return emails.split(',').map(email => ({
    emailAddress: {
      address: email.trim()
    }
  }));
}

/**
 * Save draft email handler - creates a new draft or updates an existing one
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSaveDraft(args) {
  const { id, to, cc, bcc, subject, body, contentType = 'text', importance = 'normal' } = args;

  // For new drafts, subject and body are recommended but not strictly required
  // For updates, at least one field should be provided
  if (id && !to && !cc && !bcc && !subject && !body && !args.contentType && !args.importance) {
    return {
      content: [{
        type: "text",
        text: "At least one field (to, cc, bcc, subject, body, contentType, importance) is required when updating a draft."
      }]
    };
  }

  try {
    // Get access token
    const accessToken = await ensureAuthenticated();

    // Build the message object with only provided fields
    const messageObject = {};

    if (subject !== undefined) {
      messageObject.subject = subject;
    }

    if (body !== undefined) {
      messageObject.body = {
        contentType,
        content: body
      };
    }

    const toRecipients = parseRecipients(to);
    if (toRecipients.length > 0) {
      messageObject.toRecipients = toRecipients;
    }

    const ccRecipients = parseRecipients(cc);
    if (ccRecipients.length > 0) {
      messageObject.ccRecipients = ccRecipients;
    }

    const bccRecipients = parseRecipients(bcc);
    if (bccRecipients.length > 0) {
      messageObject.bccRecipients = bccRecipients;
    }

    if (importance !== undefined) {
      messageObject.importance = importance;
    }

    let result;
    let action;

    if (id) {
      // Update existing draft
      result = await callGraphAPI(accessToken, 'PATCH', `me/messages/${id}`, messageObject);
      action = 'updated';
    } else {
      // Create new draft
      result = await callGraphAPI(accessToken, 'POST', 'me/messages', messageObject);
      action = 'created';
    }

    const responseText = [
      `Draft ${action} successfully!`,
      '',
      `Draft ID: ${result.id}`,
      result.subject ? `Subject: ${result.subject}` : null,
      toRecipients.length > 0 ? `To: ${toRecipients.length} recipient(s)` : null,
      ccRecipients.length > 0 ? `CC: ${ccRecipients.length} recipient(s)` : null,
      bccRecipients.length > 0 ? `BCC: ${bccRecipients.length} recipient(s)` : null,
      '',
      'Use send-draft with this ID to send the email, or find it in your Drafts folder in Outlook.'
    ].filter(Boolean).join('\n');

    return {
      content: [{
        type: "text",
        text: responseText
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error saving draft: ${error.message}`
      }]
    };
  }
}

module.exports = handleSaveDraft;
