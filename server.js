require('dotenv').config();
const express = require('express');
const fetch = require('node-fetch');
const qs = require('querystring');
const cookieSession = require('cookie-session');
const bodyParser = require('body-parser');

const app = express();
app.use(bodyParser.json());
app.use(cookieSession({ name: 'session', keys: [process.env.SESSION_SECRET || 'devkey'], maxAge: 24*60*60*1000 }));

const { CLIENT_ID, CLIENT_SECRET, REDIRECT_URI, TENANT = 'common', PORT = 3000 } = process.env;
const AUTH_BASE = `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0`;
const AUTHORIZE_URL = `${AUTH_BASE}/authorize`;
const TOKEN_URL = `${AUTH_BASE}/token`;
const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

function buildAuthUrl(state) {
  const params = {
    client_id: CLIENT_ID,
    response_type: 'code',
    redirect_uri: REDIRECT_URI,
    response_mode: 'query',
    scope: 'openid profile email offline_access Mail.Read',
    state
  };
  return AUTHORIZE_URL + '?' + qs.stringify(params);
}

app.get('/login', (req, res) => {
  const state = Math.random().toString(36).substring(2,15);
  req.session.oauth_state = state;
  res.redirect(buildAuthUrl(state));
});

app.get('/auth/callback', async (req, res) => {
  const { code, state, error } = req.query;
  if (error) return res.status(400).send('Auth error: ' + JSON.stringify(req.query));
  if (!code) return res.status(400).send('No code');
  if (!state || state !== req.session.oauth_state) return res.status(400).send('Invalid state');

  const tokenResp = await fetch(TOKEN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: qs.stringify({ client_id: CLIENT_ID, scope: 'Mail.Read offline_access openid profile email', code, redirect_uri: REDIRECT_URI, grant_type: 'authorization_code', client_secret: CLIENT_SECRET })
  });
  const tokenJson = await tokenResp.json();
  if (tokenJson.error) return res.status(500).send(tokenJson);

  req.session.access_token = tokenJson.access_token;
  req.session.refresh_token = tokenJson.refresh_token;

  res.redirect('/');
});

async function graphFetch(path, token, opts = {}){
  const res = await fetch(GRAPH_BASE + path, { headers: { Authorization: `Bearer ${token}` }, ...opts });
  return res;
}

app.get('/api/me', async (req, res) => {
  const token = req.session.access_token;
  if (!token) return res.status(401).json({ error: 'not_authenticated' });
  const r = await graphFetch('/me', token);
  if (!r.ok) return res.status(r.status).send(await r.text());
  res.json(await r.json());
});

app.get('/api/messages', async (req, res) => {
  const token = req.session.access_token;
  if (!token) return res.status(401).json({ error: 'not_authenticated' });
  const page = parseInt(req.query.page || '1');
  const top = parseInt(req.query.top || '12');
  const skip = (page - 1) * top;
  const search = req.query.search;

  let q = `/me/messages?$top=${top}&$skip=${skip}&$select=subject,from,receivedDateTime,bodyPreview`;
  if (search) {
    q = `/me/messages?$top=50&$select=subject,from,receivedDateTime,bodyPreview`;
  }

  const r = await graphFetch(q, token);
  if (!r.ok) return res.status(r.status).send(await r.text());
  const j = await r.json();
  let items = j.value || [];
  if (search) {
    const s = search.toLowerCase();
    items = items.filter(it => ((it.subject || '') + ' ' + (it.bodyPreview || '') + ' ' + (it.from?.emailAddress?.address||'')).toLowerCase().includes(s));
    const total = items.length;
    items = items.slice(0, top);
    return res.json({ value: items, total });
  }
  res.json({ value: items, total: items.length + skip });
});

app.get('/api/messages/:id', async (req, res) => {
  const token = req.session.access_token;
  if (!token) return res.status(401).json({ error: 'not_authenticated' });
  const id = req.params.id;
  const r = await graphFetch(`/me/messages/${id}`, token);
  if (!r.ok) return res.status(r.status).send(await r.text());
  res.json(await r.json());
});

app.get('/api/messages/:id/raw', async (req, res) => {
  const token = req.session.access_token;
  if (!token) return res.status(401).json({ error: 'not_authenticated' });
  const id = req.params.id;
  const r = await graphFetch(`/me/messages/${id}/$value`, token);
  if (!r.ok) return res.status(r.status).send(await r.text());
  const raw = await r.text();
  res.setHeader('Content-disposition', `attachment; filename="message_${id}.eml"`);
  res.setHeader('Content-Type', 'message/rfc822');
  res.send(raw);
});

app.get('/logout', (req, res) => { req.session = null; res.redirect('/'); });

app.get('/', (req, res) => {
  res.send('Hotmail Reader Backend. For frontend, run the React app.');
});

app.listen(PORT, ()=> console.log(`Backend listening on ${PORT}`));
