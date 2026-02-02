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
 * Save draft email handler - creates a new draft, updates an existing one, or creates a reply draft
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSaveDraft(args) {
  const { id, replyToMessageId, replyAll = false, to, cc, bcc, subject, body, contentType = 'text', importance = 'normal' } = args;

  // Validate conflicting parameters
  if (id && replyToMessageId) {
    return {
      content: [{
        type: "text",
        text: "Cannot specify both 'id' (update existing draft) and 'replyToMessageId' (create reply). Use one or the other."
      }]
    };
  }

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

    let result;
    let action;

    if (replyToMessageId) {
      // Create a reply draft using the appropriate endpoint
      const replyEndpoint = replyAll
        ? `me/messages/${replyToMessageId}/createReplyAll`
        : `me/messages/${replyToMessageId}/createReply`;

      // createReply/createReplyAll returns a draft with threading headers set
      result = await callGraphAPI(accessToken, 'POST', replyEndpoint, {});
      action = 'created as reply';

      // If user provided any overrides, patch the draft
      const patchObject = buildMessageObject({ to, cc, bcc, subject, body, contentType, importance }, args);

      if (Object.keys(patchObject).length > 0) {
        result = await callGraphAPI(accessToken, 'PATCH', `me/messages/${result.id}`, patchObject);
      }
    } else if (id) {
      // Update existing draft
      const messageObject = buildMessageObject({ to, cc, bcc, subject, body, contentType, importance }, args);
      result = await callGraphAPI(accessToken, 'PATCH', `me/messages/${id}`, messageObject);
      action = 'updated';
    } else {
      // Create new standalone draft
      const messageObject = buildMessageObject({ to, cc, bcc, subject, body, contentType, importance }, args);
      result = await callGraphAPI(accessToken, 'POST', 'me/messages', messageObject);
      action = 'created';
    }

    // Count recipients for response
    const toCount = (result.toRecipients || []).length;
    const ccCount = (result.ccRecipients || []).length;
    const bccCount = (result.bccRecipients || []).length;

    const responseText = [
      `Draft ${action} successfully!`,
      '',
      `Draft ID: ${result.id}`,
      result.subject ? `Subject: ${result.subject}` : null,
      toCount > 0 ? `To: ${toCount} recipient(s)` : null,
      ccCount > 0 ? `CC: ${ccCount} recipient(s)` : null,
      bccCount > 0 ? `BCC: ${bccCount} recipient(s)` : null,
      replyToMessageId ? `Thread: linked to original message` : null,
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

/**
 * Build a message object from provided fields
 * Only includes fields that were explicitly provided
 * @param {object} fields - The field values
 * @param {object} args - Original args to check what was explicitly provided
 * @returns {object} - Message object for Graph API
 */
function buildMessageObject(fields, args) {
  const { to, cc, bcc, subject, body, contentType, importance } = fields;
  const messageObject = {};

  if (subject !== undefined && args.subject !== undefined) {
    messageObject.subject = subject;
  }

  if (body !== undefined && args.body !== undefined) {
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

  if (importance !== undefined && args.importance !== undefined) {
    messageObject.importance = importance;
  }

  return messageObject;
}

module.exports = handleSaveDraft;
