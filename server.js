// Tiny local server for the Meesho Keychain Catalog Builder.
//
// Why this exists: the app's AI features (Gemini photo analysis / image
// generation) can't be called directly from a browser page — Google's
// image-generation API rejects the browser's CORS preflight. This server
// makes those same requests from Node instead (no CORS involved there) and
// hands the result back to the page. Everything else in the app (the form,
// Excel export, ImgBB image hosting) works the same with or without this
// server — only the two AI buttons need it running.
//
// Usage: node server.js   (or double-click start.bat on Windows)
// Then open the printed http://localhost:PORT URL in your browser.

const http = require('http');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const JSZip = require('jszip');

const PORT = process.env.PORT || 5177;
const INDEX_HTML_PATH = path.join(__dirname, 'index.html');
const TEMPLATE_XLSX_PATH = path.join(__dirname, 'template.xlsx');
const TEMPLATE_SHEET_NAME = 'Keychains-Fill this';
const TEMPLATE_FIRST_DATA_ROW = 5;
const GEMINI_ENDPOINT = 'https://generativelanguage.googleapis.com/v1beta/interactions';

// Maps our internal field keys to their column number (D=4 .. AP=42) on the real
// "Keychains-Fill this" sheet. Must stay in sync with the FIELDS array in index.html.
const COLUMN_MAP = {
  title: 4, size: 5, price: 6, return_price: 7, mrp: 8, gst: 9, hsn: 10,
  weight: 11, inventory: 12, country_of_origin: 13, manufacturer: 14,
  manufacturer_address: 15, manufacturer_pincode: 16, packer: 17,
  packer_address: 18, packer_pincode: 19, importer: 20, importer_address: 21,
  importer_pincode: 22, color: 23, generic_name: 24, material: 25,
  net_quantity: 26, dimension_unit: 27, height: 28, length: 29, width: 30,
  shape: 31, type: 32, main_image: 33, other_image_1: 34, other_image_2: 35,
  other_image_3: 36, style_id: 37, sku: 38, brand_name: 39, group_id: 40,
  description: 41, brand: 42,
};
const REVERSE_COLUMN_MAP = {};
for (const [key, col] of Object.entries(COLUMN_MAP)) REVERSE_COLUMN_MAP[col] = key;

// Clones template.xlsx and writes one data row per product into "Keychains-Fill this",
// starting at row 5. Shared by both the combined export and the per-product ZIP export.
async function buildCatalogWorkbookBuffer(products) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(TEMPLATE_XLSX_PATH);
  const sheet = workbook.getWorksheet(TEMPLATE_SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${TEMPLATE_SHEET_NAME}" not found in template.xlsx`);

  products.forEach((product, i) => {
    const row = sheet.getRow(TEMPLATE_FIRST_DATA_ROW + i);
    for (const [key, col] of Object.entries(COLUMN_MAP)) {
      const value = product[key];
      if (value !== undefined && value !== null && value !== '') {
        row.getCell(col).value = value;
      }
    }
    row.commit();
  });

  return workbook.xlsx.writeBuffer();
}

// Filesystem-safe filename built from a product's SKU (falls back to a generic name).
function filenameForProduct(product, fallback) {
  const base = (product && product.sku ? product.sku : fallback).toString().replace(/[^A-Za-z0-9_-]+/g, '_');
  return `${base || fallback}.xlsx`;
}

function sendJson(res, status, obj) {
  const body = JSON.stringify(obj);
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Content-Length': Buffer.byteLength(body),
  });
  res.end(body);
}

function readBody(req, maxBytes) {
  return new Promise((resolve, reject) => {
    let size = 0;
    const chunks = [];
    req.on('data', (chunk) => {
      size += chunk.length;
      if (size > maxBytes) {
        reject(new Error('Request body too large'));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
    req.on('error', reject);
  });
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://localhost:${PORT}`);

  if (req.method === 'GET' && (url.pathname === '/' || url.pathname === '/index.html')) {
    fs.readFile(INDEX_HTML_PATH, (err, data) => {
      if (err) {
        res.writeHead(500, { 'Content-Type': 'text/plain' });
        res.end('Could not read index.html: ' + err.message);
        return;
      }
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(data);
    });
    return;
  }

  // Proxies a Gemini "interactions" request. The browser never talks to
  // Google directly, which is what avoids the CORS failure.
  if (req.method === 'POST' && url.pathname === '/api/gemini') {
    try {
      const body = await readBody(req, 25 * 1024 * 1024);
      const { apiKey, model, input } = JSON.parse(body || '{}');
      if (!apiKey || !model || !input) {
        return sendJson(res, 400, { error: 'Missing apiKey, model, or input' });
      }
      const upstream = await fetch(GEMINI_ENDPOINT, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'x-goog-api-key': apiKey },
        body: JSON.stringify({ model, input }),
      });
      const text = await upstream.text();
      res.writeHead(upstream.status, { 'Content-Type': 'application/json; charset=utf-8' });
      res.end(text);
    } catch (err) {
      sendJson(res, 500, { error: String((err && err.message) || err) });
    }
    return;
  }

  // Fetches an arbitrary public image URL server-side and returns it as
  // base64, so the "Analyze photo" / "Auto-fill images" AI buttons work
  // even when the Main Image is a pasted URL rather than a freshly-Browsed
  // file (a plain browser fetch() could itself be blocked by CORS on
  // whatever host is serving that image).
  if (req.method === 'GET' && url.pathname === '/api/fetch-image') {
    try {
      const target = url.searchParams.get('url');
      if (!target || !/^https?:\/\//i.test(target)) {
        return sendJson(res, 400, { error: 'Invalid or missing url parameter' });
      }
      const upstream = await fetch(target);
      if (!upstream.ok) {
        return sendJson(res, 502, { error: `Could not fetch that image (upstream returned ${upstream.status})` });
      }
      const buf = Buffer.from(await upstream.arrayBuffer());
      const mimeType = upstream.headers.get('content-type') || 'image/jpeg';
      sendJson(res, 200, { mimeType, data: buf.toString('base64') });
    } catch (err) {
      sendJson(res, 500, { error: String((err && err.message) || err) });
    }
    return;
  }

  // Clones the real Meesho template (all sheets, dropdowns, formatting intact) and
  // writes one data row per product into "Keychains-Fill this", starting at row 5.
  if (req.method === 'POST' && url.pathname === '/api/export-xlsx') {
    try {
      const body = await readBody(req, 25 * 1024 * 1024);
      const { products, filename } = JSON.parse(body || '{}');
      if (!Array.isArray(products) || !products.length) {
        return sendJson(res, 400, { error: 'No products to export' });
      }
      if (!fs.existsSync(TEMPLATE_XLSX_PATH)) {
        return sendJson(res, 500, { error: 'template.xlsx is missing from the app folder' });
      }

      const buffer = await buildCatalogWorkbookBuffer(products);
      const safeName = (filename || 'Keychains-Fill-this.xlsx').toString().replace(/[^A-Za-z0-9_.-]+/g, '_');
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${safeName}"`,
        'Content-Length': buffer.length,
      });
      res.end(buffer);
    } catch (err) {
      sendJson(res, 500, { error: String((err && err.message) || err) });
    }
    return;
  }

  // Exports every product as its own single-row Excel file (same template clone as above),
  // bundled into one ZIP so a whole catalog can be downloaded as individual per-SKU files.
  if (req.method === 'POST' && url.pathname === '/api/export-xlsx-zip') {
    try {
      const body = await readBody(req, 25 * 1024 * 1024);
      const { products } = JSON.parse(body || '{}');
      if (!Array.isArray(products) || !products.length) {
        return sendJson(res, 400, { error: 'No products to export' });
      }
      if (!fs.existsSync(TEMPLATE_XLSX_PATH)) {
        return sendJson(res, 500, { error: 'template.xlsx is missing from the app folder' });
      }

      const zip = new JSZip();
      const usedNames = new Set();
      for (let i = 0; i < products.length; i++) {
        const buffer = await buildCatalogWorkbookBuffer([products[i]]);
        let name = filenameForProduct(products[i], `product_${i + 1}`);
        while (usedNames.has(name)) name = name.replace(/\.xlsx$/, `_${i + 1}.xlsx`);
        usedNames.add(name);
        zip.file(name, buffer);
      }
      const zipBuffer = await zip.generateAsync({ type: 'nodebuffer' });
      res.writeHead(200, {
        'Content-Type': 'application/zip',
        'Content-Disposition': 'attachment; filename="Meesho-Catalog-Individual.zip"',
        'Content-Length': zipBuffer.length,
      });
      res.end(zipBuffer);
    } catch (err) {
      sendJson(res, 500, { error: String((err && err.message) || err) });
    }
    return;
  }

  // Reads a Meesho error-report xlsx (the file Meesho sends back after a failed bulk
  // upload — same template, with columns B/C filled in per row: ERROR STATUS / ERROR
  // MESSAGE) and returns the rows that have a problem, reverse-mapped back into our
  // internal field keys so the client can load them straight into the Add/Edit form.
  if (req.method === 'POST' && url.pathname === '/api/import-error-report') {
    try {
      const body = await readBody(req, 25 * 1024 * 1024);
      const { fileBase64 } = JSON.parse(body || '{}');
      if (!fileBase64) return sendJson(res, 400, { error: 'Missing file' });
      const buffer = Buffer.from(fileBase64, 'base64');
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      const sheet = workbook.worksheets.find(ws => /fill this/i.test(ws.name)) || workbook.getWorksheet(TEMPLATE_SHEET_NAME);
      if (!sheet) {
        return sendJson(res, 400, { error: 'Could not find a "...Fill this" data sheet in this file — is this the right file?' });
      }

      const errorRows = [];
      sheet.eachRow((row, rowNumber) => {
        if (rowNumber < TEMPLATE_FIRST_DATA_ROW) return;
        const hasAnyData = row.getCell(4).value; // Product Name — presence marks a real data row
        if (!hasAnyData) return;
        const errorStatus = (row.getCell(2).value || '').toString().trim();
        const errorMessage = (row.getCell(3).value || '').toString().trim();
        if (!errorMessage && errorStatus.toUpperCase() !== 'INVALID') return; // only rows with a flagged problem

        const product = {};
        for (const [col, key] of Object.entries(REVERSE_COLUMN_MAP)) {
          let v = row.getCell(Number(col)).value;
          if (v && typeof v === 'object') {
            // Defensively unwrap rich-text/formula-result cell objects; anything unrecognized
            // is dropped rather than risk leaking "[object Object]" into the form.
            if (Array.isArray(v.richText)) v = v.richText.map(rt => rt.text).join('');
            else if (v.text !== undefined) v = v.text;
            else if (v.result !== undefined) v = v.result;
            else v = null;
          }
          if (v !== null && v !== undefined && v !== '') product[key] = v.toString();
        }
        errorRows.push({ row: rowNumber, errorStatus, errorMessage, product });
      });

      sendJson(res, 200, { rows: errorRows });
    } catch (err) {
      sendJson(res, 500, { error: String((err && err.message) || err) });
    }
    return;
  }

  res.writeHead(404, { 'Content-Type': 'text/plain' });
  res.end('Not found');
});

server.listen(PORT, () => {
  console.log(`Meesho Keychain Catalog Builder running at http://localhost:${PORT}`);
  console.log('Press Ctrl+C to stop.');
});
