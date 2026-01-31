/**
 * Send draft email functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Send draft email handler - sends an existing draft
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSendDraft(args) {
  const { id } = args;

  if (!id) {
    return {
      content: [{
        type: "text",
        text: "Draft ID is required. Use list-emails with folder 'drafts' to find draft IDs."
      }]
    };
  }

  try {
    // Get access token
    const accessToken = await ensureAuthenticated();

    // First, get the draft details for the confirmation message
    const draft = await callGraphAPI(accessToken, 'GET', `me/messages/${id}`);

    // Send the draft - this endpoint takes no body
    await callGraphAPI(accessToken, 'POST', `me/messages/${id}/send`, {});

    const recipientCount = (draft.toRecipients || []).length;
    const ccCount = (draft.ccRecipients || []).length;
    const bccCount = (draft.bccRecipients || []).length;

    const responseText = [
      'Draft sent successfully!',
      '',
      `Subject: ${draft.subject || '(no subject)'}`,
      `Recipients: ${recipientCount}${ccCount > 0 ? ` + ${ccCount} CC` : ''}${bccCount > 0 ? ` + ${bccCount} BCC` : ''}`
    ].join('\n');

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

    // Check for common errors
    if (error.message.includes('404')) {
      return {
        content: [{
          type: "text",
          text: "Draft not found. The draft may have been deleted or already sent. Use list-emails with folder 'drafts' to see available drafts."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error sending draft: ${error.message}`
      }]
    };
  }
}

module.exports = handleSendDraft;
