const crypto = require('crypto');

const PRICE_PAISE = 9900;
const CURRENCY = 'INR';

function json(statusCode, body) {
  return {
    statusCode,
    headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' },
    body: JSON.stringify(body)
  };
}
function requiredEnv(name) {
  const value = process.env[name];

  if (!value || !value.trim()) {
    console.error(`Missing required environment variable: ${name}`);
    throw new Error(
      'Server configuration is incomplete. Please contact support.'
    );
  }

  return value;
}
function parseBody(event) {
  try { return JSON.parse(event.body || '{}'); }
  catch (_) { return {}; }
}
function authHeader() {
  const id = requiredEnv('RAZORPAY_KEY_ID');
  const secret = requiredEnv('RAZORPAY_KEY_SECRET');
  return 'Basic ' + Buffer.from(id + ':' + secret).toString('base64');
}
async function razorpay(path, options = {}) {
  const response = await fetch('https://api.razorpay.com/v1' + path, {
    ...options,
    headers: {
      Authorization: authHeader(),
      'Content-Type': 'application/json',
      ...(options.headers || {})
    }
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    const err = new Error((data.error && data.error.description) || 'Payment provider request failed.');
    err.status = response.status;
    throw err;
  }
  return data;
}
function makeToken(user) {
  const secret = requiredEnv('THRYVE_TOKEN_SECRET');
  const encoded = Buffer.from(JSON.stringify(user), 'utf8').toString('base64url');
  const signature = crypto.createHmac('sha256', secret).update(encoded).digest('base64url');
  return encoded + '.' + signature;
}
function readToken(token) {
  if (typeof token !== 'string' || token.length > 4096) return { valid: false };
  const parts = token.split('.');
  if (parts.length !== 2) return { valid: false };
  const [encoded, signature] = parts;
  const expected = crypto.createHmac('sha256', requiredEnv('THRYVE_TOKEN_SECRET')).update(encoded).digest();
  let supplied;
  try { supplied = Buffer.from(signature, 'base64url'); } catch (_) { return { valid: false }; }
  if (supplied.length !== expected.length || !crypto.timingSafeEqual(supplied, expected)) return { valid: false };
  let payload;
  try { payload = JSON.parse(Buffer.from(encoded, 'base64url').toString('utf8')); } catch (_) { return { valid: false }; }
  if (!payload || !payload.exp || !payload.paymentId || !payload.email || !['teen','adult'].includes(payload.cat)) return { valid: false };
  if (Date.now() >= payload.exp) return { valid: false, expired: true };
  return { valid: true, user: payload };
}
function safeError(err) {
  // Do not return secrets or raw provider payloads to the browser.
  if (err && err.status === 400) return 'The payment could not be confirmed. If money was deducted, please contact support before trying again.';
  return 'We could not complete the request. Please try again, or contact support if money was deducted.';
}
module.exports = { PRICE_PAISE, CURRENCY, json, requiredEnv, parseBody, razorpay, makeToken, readToken, safeError };
