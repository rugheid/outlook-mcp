/**
 * Authentication module for Outlook MCP server
 */
const tokenManager = require('./token-manager');
const TokenStorage = require('./token-storage');
const { authTools } = require('./tools');

// Create a singleton TokenStorage instance for auto-refresh capability
const tokenStorage = new TokenStorage();

/**
 * Ensures the user is authenticated and returns an access token
 * Uses TokenStorage which automatically refreshes expired tokens
 * @param {boolean} forceNew - Whether to force a new authentication
 * @returns {Promise<string>} - Access token
 * @throws {Error} - If authentication fails
 */
async function ensureAuthenticated(forceNew = false) {
  if (forceNew) {
    // Force re-authentication
    throw new Error('Authentication required');
  }

  // Use TokenStorage which auto-refreshes expired tokens using the refresh token
  const accessToken = await tokenStorage.getValidAccessToken();
  if (!accessToken) {
    throw new Error('Authentication required');
  }

  return accessToken;
}

module.exports = {
  tokenManager,
  tokenStorage,
  authTools,
  ensureAuthenticated
};
