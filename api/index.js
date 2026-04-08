const { app } = require('@azure/functions');
const https = require('https');

const SALES_EMAIL = 'Sales1@trustdi.com';

// Sends one email via Resend API -- throws on non-2xx so errors surface
function sendEmail(apiKey, to, subject, html, fromName) {
  const from = fromName
    ? fromName + ' <noreply@trustdi.com>'
    : 'TrustDigital <noreply@trustdi.com>';

  const data = JSON.stringify({
    from,
    to: Array.isArray(to) ? to : [to],
    subject,
    html,
    reply_to: 'Sales1@trustdi.com'
  });

  return new Promise((resolve, reject) => {
    const options = {
      hostname: 'api.resend.com',
      path: '/emails',
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + apiKey,
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(data)
      }
    };

    const req = https.request(options, (res) => {
      let body = '';
      res.on('data', chunk => body += chunk);
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          resolve({ status: res.statusCode, body });
        } else {
          reject(new Error('Resend API ' + res.statusCode + ': ' + body));
        }
      });
    });

    // 30-second timeout to prevent hanging indefinitely
    req.setTimeout(30000, () => {
      req.destroy(new Error('Request timed out after 30s'));
    });

    req.on('error', reject);
    req.write(data);
    req.end();
  });
}

app.http('send-report', {
  methods: ['GET', 'POST'],
  authLevel: 'anonymous',
  route: 'send-report',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') {
      return {
        status: 204,
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'POST, OPTIONS',
          'Access-Control-Allow-Headers': 'Content-Type'
        }
      };
    }

    let payload;
    try {
      payload = await request.json();
    } catch (e) {
      return {
        status: 400,
        headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' },
        body: JSON.stringify({ error: 'Invalid JSON body' })
      };
    }

    const { contact, emailHTML, wantsHRForm, hrTemplate } = payload || {};

    if (!contact || !contact.email || !emailHTML) {
      return {
        status: 400,
        headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' },
        body: JSON.stringify({ error: 'Missing required fields: contact.email and emailHTML required' })
      };
    }

    const apiKey = process.env.RESEND_API_KEY;
    if (!apiKey) {
      return {
        status: 500,
        headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' },
        body: JSON.stringify({ error: 'Email service not configured -- RESEND_API_KEY missing' })
      };
    }

    const company = contact.company || 'Your Company';
    const name    = contact.name    || '';

    const customerSubject = 'Your Microsoft 365 Licensing Report — ' + company;
    const salesSubject    = 'New M365 Report — ' + company + ' (' + name + ')';
    const hrSubject       = 'HR Data Request — ' + company;

    try {
      // Email 1: customer gets their full licensing report
      await sendEmail(apiKey, contact.email, customerSubject, emailHTML, 'TrustDigital');

      // Email 2: Sales1 gets the same report with a sales-friendly subject
      await sendEmail(apiKey, SALES_EMAIL, salesSubject, emailHTML, 'TrustDigital Calculator');

      // Email 3 (optional): HR data request template sent to customer
      if (wantsHRForm && hrTemplate) {
        const hrHTML = '<html><body style="font-family:Arial,sans-serif;max-width:680px;margin:0 auto;padding:24px">'
          + String(hrTemplate).replace(/\n/g, '<br>')
          + '</body></html>';
        await sendEmail(apiKey, contact.email, hrSubject, hrHTML, 'TrustDigital');
      }

      return {
        status: 200,
        headers: {
          'Content-Type': 'application/json',
          'Access-Control-Allow-Origin': '*'
        },
        body: JSON.stringify({ success: true })
      };
    } catch (err) {
      return {
        status: 500,
        headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' },
        body: JSON.stringify({ success: false, error: err.message })
      };
    }
  }
});
