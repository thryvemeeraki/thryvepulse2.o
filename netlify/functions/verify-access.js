const { json, parseBody, readToken } = require('./_shared');

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') return json(405, { error: 'Method not allowed.' });
  try {
    const { token } = parseBody(event);
    const result = readToken(token);
    if (result.expired) return json(410, { valid: false, expired: true });
    if (!result.valid) return json(401, { valid: false });
    return json(200, { valid: true, user: result.user });
  } catch (err) {
    console.error('verify-access error:', err && err.message);
    return json(500, { valid: false, error: 'Could not verify access.' });
  }
};
