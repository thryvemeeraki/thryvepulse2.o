const crypto = require('crypto');
const { PRICE_PAISE, CURRENCY, json, requiredEnv, parseBody, razorpay, safeError } = require('./_shared');

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') return json(405, { error: 'Method not allowed.' });
  try {
    const data = parseBody(event);
    const name = String(data.name || '').trim().slice(0, 120);
    const email = String(data.email || '').trim().toLowerCase().slice(0, 254);
    const phone = String(data.phone || '').trim().slice(0, 30);
    const category = String(data.category || '');
    if (!name || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || !phone || !['teen','adult'].includes(category)) {
      return json(400, { error: 'Please check your name, email, phone number and age category.' });
    }
    const receipt = 'thryve_' + Date.now() + '_' + crypto.randomBytes(4).toString('hex');
    const order = await razorpay('/orders', {
      method: 'POST',
      body: JSON.stringify({
        amount: PRICE_PAISE,
        currency: CURRENCY,
        receipt,
        notes: { name, email, phone, category, product: 'Thryve Pulse Assessment' }
      })
    });
    return json(200, {
      orderId: order.id,
      amount: order.amount,
      currency: order.currency,
      keyId: requiredEnv('RAZORPAY_KEY_ID')
    });
  } catch (err) {
    console.error('create-order error:', err && err.message);
    return json(500, { error: safeError(err) });
  }
};
