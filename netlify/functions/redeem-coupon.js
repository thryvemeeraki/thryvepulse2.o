const crypto = require('crypto');
const { json, requiredEnv, parseBody, makeToken, safeError } = require('./_shared');

const GAS_URL = () => requiredEnv('APPS_SCRIPT_URL');
const ASSESSMENT_URL = 'https://thryvepulse2.netlify.app/thryve-assessment.html';

async function registerCashCoupon(payload) {
  const response = await fetch(GAS_URL(), {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' },
    body: new URLSearchParams({ payload: JSON.stringify(payload) }).toString(),
    redirect: 'follow'
  });
  const text = await response.text();
  let result;
  try { result = JSON.parse(text); } catch (_) { throw new Error('Registration service returned an unexpected response.'); }
  if (!response.ok || !result || result.ok !== true) {
    const err = new Error((result && result.error) || 'This coupon could not be redeemed. Please check the code or contact support.');
    err.status = 400;
    throw err;
  }
  return result;
}

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') return json(405, { error: 'Method not allowed.' });
  try {
    const data = parseBody(event);
    const name = String(data.name || '').trim().slice(0, 120);
    const email = String(data.email || '').trim().toLowerCase().slice(0, 254);
    const phone = String(data.phone || '').trim().slice(0, 30);
    const category = String(data.category || '');
    const couponCode = String(data.couponCode || '').trim().toUpperCase().slice(0, 40);
    if (!name || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || !phone || !['teen','adult'].includes(category) || !/^TP-CASH-[A-Z0-9]{12}$/.test(couponCode)) {
      return json(400, { error: 'Please check your registration details and cash coupon code.' });
    }
    const expiresAt = Date.now() + 60 * 60 * 1000;
    const user = { name, email, phone, cat: category, exp: expiresAt, couponCode, paymentMethod: 'Cash Coupon' };
    const token = makeToken(user);
    const link = ASSESSMENT_URL + '?token=' + encodeURIComponent(token);
    const expDate = new Date(expiresAt);
    const expTime = expDate.toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' }) +
      ' on ' + expDate.toLocaleDateString('en-IN', { day: 'numeric', month: 'long', year: 'numeric' });
    const catLabel = category === 'teen' ? 'Teen (Below 18)' : 'Adult (18 & Above)';
    const result = await registerCashCoupon({
      action: 'redeemCashCoupon', couponCode, name, email, phone,
      category: catLabel, link, expTime, amountInr: 99,
      paymentMethod: 'Cash Coupon'
    });
    return json(200, { ok: true, email: result.email || email });
  } catch (err) {
    console.error('redeem-coupon error:', err && err.message);
    const message = err && err.status === 400 ? err.message : safeError(err);
    return json(err && err.status === 400 ? 400 : 500, { error: message });
  }
};
