/**
 * Respond to calendar event functionality
 * Handles accept, decline, and tentative responses to meeting invites
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

// Map response types to Graph API endpoints
const RESPONSE_ENDPOINTS = {
  accept: 'accept',
  decline: 'decline',
  tentative: 'tentativelyAccept'
};

const RESPONSE_LABELS = {
  accept: 'accepted',
  decline: 'declined',
  tentative: 'tentatively accepted'
};

/**
 * Respond to event handler - accept, decline, or tentatively accept a meeting invite
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleRespondToEvent(args) {
  const { eventId, response, comment, sendResponse = true } = args;

  // Validate required parameters
  if (!eventId) {
    return {
      content: [{
        type: "text",
        text: "Event ID is required. Use 'list-events' to find event IDs."
      }]
    };
  }

  if (!response) {
    return {
      content: [{
        type: "text",
        text: "Response is required. Must be one of: accept, decline, tentative"
      }]
    };
  }

  const normalizedResponse = response.toLowerCase();
  if (!RESPONSE_ENDPOINTS[normalizedResponse]) {
    return {
      content: [{
        type: "text",
        text: `Invalid response '${response}'. Must be one of: accept, decline, tentative`
      }]
    };
  }

  try {
    // Get access token
    const accessToken = await ensureAuthenticated();

    // Build API endpoint
    const endpoint = `me/events/${eventId}/${RESPONSE_ENDPOINTS[normalizedResponse]}`;

    // Request body
    const body = {
      sendResponse
    };

    // Only include comment if provided
    if (comment) {
      body.comment = comment;
    }

    // Make API call
    await callGraphAPI(accessToken, 'POST', endpoint, body);

    const responseLabel = RESPONSE_LABELS[normalizedResponse];
    const commentNote = comment ? `\nComment: "${comment}"` : '';
    const notifyNote = sendResponse ? '' : '\n(Organizer was not notified)';

    return {
      content: [{
        type: "text",
        text: `Event ${responseLabel} successfully.${commentNote}${notifyNote}`
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
          text: "Event not found. The event may have been deleted or the ID is incorrect. Use 'list-events' to find valid event IDs."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error responding to event: ${error.message}`
      }]
    };
  }
}

module.exports = handleRespondToEvent;
