const https = require('https');

const SALES_EMAIL = 'Sales1@trustdi.com';

function sendEmail(apiKey, to, subject, html, fromName) {
  const from = fromName
    ? fromName + ' <Sales1@trustdi.com>'h
    : 'TrustDigital <Sales1@trustdi.com>';
  const data = JSON.stringify({ from, to: Array.isArray(to) ? to : [to], subject, html });
  return new Promise((resolve, reject) => {
    const options = {
      hostname: 'api.resend.com',
      path: '/emails',
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + apiKey, 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(data) }
    };
    const req = https.request(options, (res) => {
      let body = '';
      res.on('data', chunk => body += chunk);
      res.on('end', () => resolve({ status: res.statusCode, body }));
    });
    req.on('error', reject);
    req.write(data);
    req.end();
  });
}

module.exports = async function (context, req) {
  if (req.method === 'OPTIONS') {
    context.res = { status: 204, headers: { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'POST, OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type' } };
    return;
  }
  const payload = req.body;
  const { contact, emailHTML } = payload || {};
  if (!contact || !contact.email || !emailHTML) { context.res = { status: 400, body: JSON.stringify({ error: 'Missing required fields' }) }; return; }
  const apiKey = process.env.RESEND_API_KEY;
  if (!apiKey) { context.res = { status: 500, body: JSON.stringify({ error: 'Email service not configured' }) }; return; }
  const company = contact.company || 'Your Company';
  const name = contact.name || '';
  try {
    await sendEmail(apiKey, contact.email, 'Your Microsoft 365 Licensing Report — ' + company, emailHTML, 'TrustDigital');
    await sendEmail(apiKey, SALES_EMAIL, 'New M365 Report — ' + company + ' (' + name + ')', emailHTML, 'TrustDigital Calculator');
    context.res = { status: 200, headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' }, body: JSON.stringify({ success: true }) };
  } catch (err) {
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ success: false, error: err.message }) };
  }
};
