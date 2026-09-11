// Tiny static file server for site/ (local testing and Playwright). Usage: node tools/serve.mjs [PORT]
import http from 'node:http';
import { createReadStream, statSync } from 'node:fs';
import { extname, join, normalize, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = normalize(fileURLToPath(new URL('../site/', import.meta.url)));
const port = Number(process.argv[2] || process.env.PORT || 8787);
const types = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8', '.webmanifest': 'application/manifest+json',
  '.svg': 'image/svg+xml', '.png': 'image/png', '.ico': 'image/x-icon', '.woff2': 'font/woff2', '.txt': 'text/plain; charset=utf-8',
};

http.createServer((req, res) => {
  let p = decodeURIComponent(new URL(req.url, 'http://localhost').pathname);
  if (p.endsWith('/')) p += 'index.html';
  const file = normalize(join(root, p));
  if (!file.startsWith(root.endsWith(sep) ? root : root + sep) && file !== root) { res.writeHead(403); return res.end('forbidden'); }
  let st;
  try { st = statSync(file); } catch { res.writeHead(404, { 'Content-Type': 'text/plain' }); return res.end('not found: ' + p); }
  if (st.isDirectory()) { res.writeHead(302, { Location: p + '/' }); return res.end(); }
  res.writeHead(200, { 'Content-Type': types[extname(file).toLowerCase()] || 'application/octet-stream', 'Cache-Control': 'no-store', 'Content-Length': st.size });
  createReadStream(file).pipe(res);
}).listen(port, () => console.log(`serving ${root} at http://localhost:${port}/`));
