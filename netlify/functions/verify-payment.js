const crypto = require('crypto');
const {
  PRICE_PAISE, CURRENCY, json, requiredEnv, parseBody, razorpay, makeToken, safeError
} = require('./_shared');

const GAS_URL = () => requiredEnv('APPS_SCRIPT_URL');

function signatureIsValid(orderId, paymentId, signature) {
  if (!orderId || !paymentId || typeof signature !== 'string' || !/^[a-f0-9]{64}$/i.test(signature)) return false;
  const expected = crypto.createHmac('sha256', requiredEnv('RAZORPAY_KEY_SECRET'))
    .update(orderId + '|' + paymentId).digest();
  const actual = Buffer.from(signature, 'hex');
  return actual.length === expected.length && crypto.timingSafeEqual(actual, expected);
}

async function saveRegistration(payload) {
  const response = await fetch(GAS_URL(), {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' },
    body: new URLSearchParams({ payload: JSON.stringify(payload) }).toString(),
    redirect: 'follow'
  });
  const text = await response.text();
  if (!response.ok || !/ok/i.test(text)) throw new Error('Registration service did not confirm the record.');
}

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') return json(405, { error: 'Method not allowed.' });
  try {
    const data = parseBody(event);
    const paymentId = String(data.razorpay_payment_id || '');
    const orderId = String(data.razorpay_order_id || '');
    const signature = String(data.razorpay_signature || '');
    if (!signatureIsValid(orderId, paymentId, signature)) return json(400, { error: 'Payment signature could not be verified.' });

    const order = await razorpay('/orders/' + encodeURIComponent(orderId), { method: 'GET' });
    if (order.amount !== PRICE_PAISE || order.currency !== CURRENCY || order.status === 'created' && !order.id) {
      return json(400, { error: 'This payment order is not valid for Thryve Pulse.' });
    }

    let payment = await razorpay('/payments/' + encodeURIComponent(paymentId), { method: 'GET' });
    if (payment.order_id !== orderId || payment.amount !== PRICE_PAISE || payment.currency !== CURRENCY) {
      return json(400, { error: 'Payment details do not match this assessment order.' });
    }
    if (payment.status === 'authorized') {
      try {
        await razorpay('/payments/' + encodeURIComponent(paymentId) + '/capture', {
          method: 'POST', body: JSON.stringify({ amount: PRICE_PAISE, currency: CURRENCY })
        });
      } catch (captureErr) {
        // The payment may already have been captured automatically; confirm below.
        console.warn('Capture request did not complete; rechecking payment status.');
      }
      payment = await razorpay('/payments/' + encodeURIComponent(paymentId), { method: 'GET' });
    }
    if (payment.status !== 'captured') {
      return json(402, { error: 'Payment is not captured yet. Please wait a moment and try again, or contact support if money was deducted.' });
    }

    const notes = order.notes || {};
    const name = String(notes.name || '').trim();
    const email = String(notes.email || '').trim().toLowerCase();
    const phone = String(notes.phone || '').trim();
    const category = String(notes.category || '');
    if (!name || !email || !phone || !['teen','adult'].includes(category)) {
      return json(400, { error: 'Registration details were missing from this order. Please contact support.' });
    }

    const expiresAt = Date.now() + 60 * 60 * 1000;
    const user = {
      name, email, phone, cat: category, exp: expiresAt,
      paymentId, orderId
    };
    const token = makeToken(user);
    const link = 'https://thryvepulse2.netlify.app/thryve-assessment.html?token=' + encodeURIComponent(token);
    const expDate = new Date(expiresAt);
    const expTime = expDate.toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' }) +
      ' on ' + expDate.toLocaleDateString('en-IN', { day: 'numeric', month: 'long', year: 'numeric' });
    const catLabel = category === 'teen' ? 'Teen (Below 18)' : 'Adult (18 & Above)';

    // Only after the provider confirms a captured ₹99 payment do we create the registration and email the link.
    await saveRegistration({
      action: 'register',
      name, email, phone,
      category: catLabel,
      link, expTime,
      paymentStatus: 'Paid',
      paymentMethod: 'Razorpay',
      couponCode: '',
      razorpayOrderId: orderId,
      razorpayPaymentId: paymentId,
      amountInr: 99
    });

    return json(200, { ok: true, email });
  } catch (err) {
    console.error('verify-payment error:', err && err.message);
    return json(500, { error: safeError(err) });
  }
};
