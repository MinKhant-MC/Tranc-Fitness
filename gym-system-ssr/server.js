const fs = require('fs');
const http = require('http');
const path = require('path');
const { URL } = require('url');

const ROOT = __dirname;
const PORT = Number(process.env.PORT || 3000);

const ROUTES = {
  '/': 'index.html',
  '/login': 'index.html',
  '/register': 'register.html',
  '/member': 'member.html',
  '/admin': 'admin.html'
};

const TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.webmanifest': 'application/manifest+json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.ico': 'image/x-icon',
  '.svg': 'image/svg+xml',
  '.txt': 'text/plain; charset=utf-8'
};

function send(res, status, body, type) {
  res.writeHead(status, {
    'Content-Type': type || 'text/plain; charset=utf-8',
    'X-Content-Type-Options': 'nosniff',
    'Referrer-Policy': 'strict-origin-when-cross-origin',
    'Cache-Control': status === 200 ? 'public, max-age=60' : 'no-store'
  });
  res.end(body);
}

function isInsideRoot(filePath) {
  const relative = path.relative(ROOT, filePath);
  return relative && !relative.startsWith('..') && !path.isAbsolute(relative);
}

function renderPage(fileName, route) {
  const filePath = path.join(ROOT, fileName);
  let html = fs.readFileSync(filePath, 'utf8');
  const data = {
    appName: 'Tranc Gym',
    route,
    renderedAt: new Date().toISOString()
  };
  const marker = `<script>window.__TRANC_GYM_SSR__=${JSON.stringify(data)};</script>`;
  if (html.includes('</head>')) {
    html = html.replace('</head>', `  ${marker}\n</head>`);
  }
  return html;
}

function serveStatic(res, pathname) {
  const safePath = decodeURIComponent(pathname).replace(/^\/+/, '');
  const filePath = path.join(ROOT, safePath);
  if (!isInsideRoot(filePath) || !fs.existsSync(filePath) || !fs.statSync(filePath).isFile()) {
    send(res, 404, 'Not found');
    return;
  }
  const ext = path.extname(filePath).toLowerCase();
  send(res, 200, fs.readFileSync(filePath), TYPES[ext] || 'application/octet-stream');
}

const server = http.createServer((req, res) => {
  try {
    const parsed = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
    const pathname = parsed.pathname.replace(/\/+$/, '') || '/';
    if (ROUTES[pathname]) {
      send(res, 200, renderPage(ROUTES[pathname], pathname), TYPES['.html']);
      return;
    }
    serveStatic(res, parsed.pathname);
  } catch (error) {
    send(res, 500, `Server error: ${error.message}`);
  }
});

server.listen(PORT, () => {
  console.log(`Tranc Gym SSR is running at http://localhost:${PORT}`);
});
