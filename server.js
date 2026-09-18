// Local server for the Meesho Catalog Builder.
//
// Why this exists: the app's AI features (Gemini photo analysis / image
// generation) can't be called directly from a browser page — Google's
// API rejects the browser's CORS preflight. This server makes those same
// requests from Node instead (no CORS involved there) and hands the result
// back to the page. It also does the Excel work (cloning Meesho's real
// template, reading error reports) and — the newer part — parses an
// uploaded Meesho category template into a field schema so the app's form
// can be generated for ANY category, not just Keychains.
//
// Usage: node server.js   (or double-click start.bat on Windows)
// Then open the printed http://localhost:PORT URL in your browser.

const http = require('http');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const JSZip = require('jszip');
const { GoogleGenAI } = require('@google/genai');

const PORT = 3000;
const INDEX_HTML_PATH = path.join(__dirname, 'index.html');
const LEGACY_TEMPLATE_XLSX_PATH = path.join(__dirname, 'template.xlsx'); // the original bundled Keychains template
const CATEGORIES_DIR = path.join(__dirname, 'categories');

/* ============================================================================
 * Generic Meesho template parsing
 *
 * Every Meesho category template we've seen follows the same layout on its
 * "<Category>-Fill this" sheet:
 *   row 1: a merged category title, repeated across every column
 *   row 2: "* Compulsory Field" / "Optional Field" / "Do not fill ..." flags
 *   row 3: "Field Name\n\nDescription..." (first line is the field's label)
 *   row 4: "Watch Explainer Video" / blank
 *   row 5+: actual data rows, each with an Excel data validation that tells
 *           us the field's type (list/whole/decimal/textLength) and, for
 *           dropdowns, either an inline list or a range on "Validation Sheet".
 * That's enough to reconstruct the whole form without hardcoding a column
 * layout per category — this function is the only place that needs to know
 * the layout convention; everything else works off its output.
 * ========================================================================= */

function cellText(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') {
    if (Array.isArray(v.richText)) return v.richText.map((rt) => rt.text).join('');
    if (v.text !== undefined) return cellText(v.text);
    if (v.result !== undefined) return cellText(v.result);
    if (v.hyperlink !== undefined) return String(v.hyperlink);
    return '';
  }
  return String(v);
}

function colLetterToNum(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

// Resolves an Excel "list" data validation's formulae into actual option strings —
// either an inline `"A,B,C"` list or a `'Sheet Name'!$X$3:$X$40` range reference.
function resolveListOptions(formulae, workbook) {
  if (!formulae || !formulae.length) return [];
  const f = formulae[0];
  if (f.startsWith('"') && f.endsWith('"')) {
    return f.slice(1, -1).split(',').map((s) => s.trim()).filter(Boolean);
  }
  const m = f.match(/^'?([^'!]+)'?!\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)$/);
  if (m) {
    const [, sheetName, col1, row1, , row2] = m;
    const ws = workbook.getWorksheet(sheetName);
    if (!ws) return [];
    const colNum = colLetterToNum(col1);
    const r1 = parseInt(row1, 10);
    const r2 = parseInt(row2, 10);
    const vals = [];
    for (let r = r1; r <= r2; r++) {
      const v = cellText(ws.getRow(r).getCell(colNum).value).trim();
      if (v) vals.push(v);
    }
    return vals;
  }
  return [];
}

function slugify(name) {
  return (name || 'category').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '') || 'category';
}

// Known Meesho field labels map onto the same internal keys the app has always used
// (form.title, form.sku, form.main_image, ...) so the rest of the app can special-case
// them (AI prompts, image Browse+analyze, SKU tracking, variation grouping) regardless
// of which category's template they came from — these labels are Meesho's own
// boilerplate and are expected to read identically across categories. Anything that
// doesn't match becomes a category-specific field, keyed by a slug of its own label.
const LABEL_ALIASES = [
  [/^product name$/i, 'title'],
  [/^variation$/i, 'size'],
  [/^meesho price$/i, 'price'],
  [/wrong\s*\/\s*defective returns price/i, 'return_price'],
  [/^mrp$/i, 'mrp'],
  [/^gst\s*%?$/i, 'gst'],
  [/^hsn id$/i, 'hsn'],
  [/net weight/i, 'weight'],
  [/^inventory$/i, 'inventory'],
  [/country of origin/i, 'country_of_origin'],
  [/manufacturer name/i, 'manufacturer'],
  [/manufacturer address/i, 'manufacturer_address'],
  [/manufacturer pincode/i, 'manufacturer_pincode'],
  [/^packer name$/i, 'packer'],
  [/packer address/i, 'packer_address'],
  [/packer pincode/i, 'packer_pincode'],
  [/importer name/i, 'importer'],
  [/importer address/i, 'importer_address'],
  [/importer pincode/i, 'importer_pincode'],
  [/^color$/i, 'color'],
  [/generic name/i, 'generic_name'],
  [/net quantity/i, 'net_quantity'],
  [/product dimension unit/i, 'dimension_unit'],
  [/product height/i, 'height'],
  [/product length/i, 'length'],
  [/product width/i, 'width'],
  [/^image\s*1\b/i, 'main_image'],
  [/^image\s*2\b/i, 'other_image_1'],
  [/^image\s*3\b/i, 'other_image_2'],
  [/^image\s*4\b/i, 'other_image_3'],
  [/product id\s*\/\s*style id/i, 'style_id'],
  [/^sku id$/i, 'sku'],
  [/^brand name$/i, 'brand_name'],
  [/^group id$/i, 'group_id'],
  [/product description/i, 'description'],
  [/^brand$/i, 'brand'],
];

function keyForLabel(label, usedKeys) {
  for (const [pattern, key] of LABEL_ALIASES) {
    if (pattern.test(label)) return key;
  }
  let base = label.toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '') || 'field';
  let key = base;
  let n = 2;
  while (usedKeys.has(key)) key = `${base}_${n++}`;
  return key;
}

// Same grouping the app has always shown for Keychains, generalized by field-name
// pattern — Meesho's legal/pricing/image field names are boilerplate across
// categories, so this keeps working the same way for any template.
function groupForField(key, label) {
  if (['title', 'generic_name', 'group_id', 'brand_name', 'brand'].includes(key)) return 'identity';
  if (['price', 'return_price', 'mrp', 'gst', 'hsn'].includes(key)) return 'pricing';
  if (['main_image', 'other_image_1', 'other_image_2', 'other_image_3'].includes(key)) return 'images';
  if (['country_of_origin', 'manufacturer', 'manufacturer_address', 'manufacturer_pincode',
    'packer', 'packer_address', 'packer_pincode', 'importer', 'importer_address', 'importer_pincode'].includes(key)) return 'legal';
  return 'details';
}

const GROUP_TITLES = {
  identity: 'Listing Identity',
  pricing: 'Pricing & Tax',
  images: 'Images',
  details: 'Product Details',
  legal: 'Manufacturer, Packer & Importer',
};

function parseTemplateSchema(workbook, categoryName) {
  const sheet = workbook.worksheets.find((ws) => /fill this/i.test(ws.name));
  if (!sheet) throw new Error('Could not find a "...Fill this" data sheet in this template — is this a real Meesho category template?');

  let flagRow = null;
  for (let r = 1; r <= 10 && !flagRow; r++) {
    let found = false;
    sheet.getRow(r).eachCell((cell) => { if (/compulsory field/i.test(cellText(cell.value))) found = true; });
    if (found) flagRow = r;
  }
  if (!flagRow) throw new Error('Could not find the "* Compulsory Field" row — this doesn\'t look like a Meesho template.');
  const descRow = flagRow + 1;
  const firstDataRow = descRow + 2; // skips the "Watch Explainer Video" row, matches every template we've seen

  const usedKeys = new Set();
  const fields = [];
  const colCount = Math.max(sheet.actualColumnCount || 0, sheet.columnCount || 0, 60);
  for (let c = 2; c <= colCount; c++) {
    const flag = cellText(sheet.getRow(flagRow).getCell(c).value).trim();
    if (!flag || /do not fill/i.test(flag)) continue; // system-only columns (ERROR STATUS/MESSAGE, etc.)
    const descLines = cellText(sheet.getRow(descRow).getCell(c).value).split(/\n+/).map((s) => s.trim()).filter(Boolean);
    const label = descLines[0];
    if (!label) continue;
    const help = descLines.slice(1).join(' ');
    const required = /compulsory/i.test(flag);
    const dv = sheet.getRow(firstDataRow).getCell(c).dataValidation;
    let type = 'text';
    let options = null;
    if (dv && dv.type === 'list') {
      type = 'select';
      options = resolveListOptions(dv.formulae, workbook);
    } else if (dv && (dv.type === 'whole' || dv.type === 'decimal' || dv.type === 'custom')) {
      type = 'number';
    } else if (/description/i.test(label)) {
      type = 'textarea';
    } else if (/^product (height|length|width)$/i.test(label)) {
      // Meesho's own sheet allows free text here (no strict numeric validation), but a
      // number input is friendlier and Meesho accepts a plain number either way.
      type = 'number';
    }
    const key = keyForLabel(label, usedKeys);
    usedKeys.add(key);
    fields.push({ key, label, help, required, type, options, col: c, group: groupForField(key, label) });
  }

  const groupsPresent = [...new Set(fields.map((f) => f.group))];
  const groupOrder = ['identity', 'pricing', 'images', 'details', 'legal'];
  const groups = groupOrder.filter((g) => groupsPresent.includes(g)).map((g) => ({ key: g, title: GROUP_TITLES[g] }));

  const sizeField = fields.find((f) => f.key === 'size');

  return {
    name: categoryName || sheet.name.replace(/-?fill this/i, '').trim() || 'Category',
    sheetName: sheet.name,
    firstDataRow,
    fields,
    groups,
    sizeOptions: sizeField ? sizeField.options || [] : [],
  };
}

function columnMapFromSchema(schema) {
  const map = {};
  schema.fields.forEach((f) => { map[f.key] = f.col; });
  return map;
}

/* ============================================================================
 * Per-category storage — each category lives in categories/<slug>/ as
 * template.xlsx (the real file, cloned on export) + schema.json (parsed once
 * on upload, re-read on every request rather than re-parsed).
 * ========================================================================= */

function categoryDir(slug) { return path.join(CATEGORIES_DIR, slug); }

function loadCategory(slug) {
  const dir = categoryDir(slug);
  const schemaPath = path.join(dir, 'schema.json');
  const templatePath = path.join(dir, 'template.xlsx');
  if (!fs.existsSync(schemaPath) || !fs.existsSync(templatePath)) {
    throw new Error(`Unknown category "${slug}" — upload its template first.`);
  }
  const schema = JSON.parse(fs.readFileSync(schemaPath, 'utf8'));
  return { slug, schema, templatePath };
}

function saveCategory(slug, schema, fileBuffer) {
  const dir = categoryDir(slug);
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(path.join(dir, 'template.xlsx'), fileBuffer);
  fs.writeFileSync(path.join(dir, 'schema.json'), JSON.stringify(schema, null, 2));
}

function listCategories() {
  if (!fs.existsSync(CATEGORIES_DIR)) return [];
  return fs.readdirSync(CATEGORIES_DIR)
    .filter((slug) => fs.existsSync(path.join(categoryDir(slug), 'schema.json')))
    .map((slug) => {
      const schema = JSON.parse(fs.readFileSync(path.join(categoryDir(slug), 'schema.json'), 'utf8'));
      return { slug, name: schema.name };
    });
}

// Bootstraps the original bundled Keychains template as the default "keychains"
// category the first time the server runs, using the exact same parser as any
// category a user uploads later — one code path, no special-cased Keychains logic.
function bootstrapKeychainsCategory() {
  const slug = 'keychains';
  if (fs.existsSync(path.join(categoryDir(slug), 'schema.json'))) return;
  if (!fs.existsSync(LEGACY_TEMPLATE_XLSX_PATH)) return;
  (async () => {
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(LEGACY_TEMPLATE_XLSX_PATH);
      const schema = parseTemplateSchema(workbook, 'Keychains');
      saveCategory(slug, schema, fs.readFileSync(LEGACY_TEMPLATE_XLSX_PATH));
      console.log('Bootstrapped the "keychains" category from the bundled template.xlsx');
    } catch (err) {
      console.error('Could not bootstrap the Keychains category:', err.message);
    }
  })();
}

// Clones a category's template.xlsx and writes one data row per product into its
// "...Fill this" sheet, starting at its detected first data row. Shared by the
// combined export and the per-product ZIP export.
async function buildCatalogWorkbookBuffer(category, products) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(category.templatePath);
  const sheet = workbook.getWorksheet(category.schema.sheetName);
  if (!sheet) throw new Error(`Sheet "${category.schema.sheetName}" not found in this category's template.xlsx`);
  const columnMap = columnMapFromSchema(category.schema);

  products.forEach((product, i) => {
    const row = sheet.getRow(category.schema.firstDataRow + i);
    for (const [key, col] of Object.entries(columnMap)) {
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

bootstrapKeychainsCategory();

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

  // Health check endpoint
  if (req.method === 'GET' && url.pathname === '/api/health') {
    return sendJson(res, 200, { status: 'ok' });
  }

  // Runtime config endpoint
  if (req.method === 'GET' && url.pathname === '/api/config') {
    const key = (process.env.GEMINI_API_KEY || '').trim();
    // In sandbox environments, dummy or expired tokens starting with AQ.Ab8RN6I are non-functional
    const isFunctionalEnvKey = Boolean(key && !key.startsWith('AQ.Ab8RN6I'));
    return sendJson(res, 200, { hasServerApiKey: isFunctionalEnvKey });
  }

  // Proxies a Gemini request using the official @google/genai SDK.
  if (req.method === 'POST' && url.pathname === '/api/gemini') {
    try {
      const body = await readBody(req, 25 * 1024 * 1024);
      const { apiKey, model, input } = JSON.parse(body || '{}');
      const providedKey = (apiKey && apiKey.trim()) || '';
      const envKey = (process.env.GEMINI_API_KEY || '').trim();
      const keyToUse = providedKey || (!envKey.startsWith('AQ.Ab8RN6I') ? envKey : '');

      if (!keyToUse) {
        return sendJson(res, 401, {
          error: 'Gemini API key is required. Please provide a Gemini API Key in ⚙️ Settings (get a free key at https://aistudio.google.com/app/apikey).',
          isAuthError: true,
          code: 'UNAUTHENTICATED'
        });
      }

      // Map models to valid current Gemini models
      let targetModel = (model || 'gemini-3.8-flash').trim();
      if (targetModel.includes('3.5-flash') || targetModel.includes('1.5') || targetModel.includes('2.0')) {
        targetModel = 'gemini-3.8-flash';
      }
      if (targetModel.includes('image')) {
        targetModel = 'gemini-3.1-flash-image';
      }

      // Convert input payload to @google/genai format
      const parts = [];
      if (Array.isArray(input)) {
        for (const item of input) {
          if (item.type === 'text' && item.text) {
            parts.push({ text: item.text });
          } else if (item.type === 'image' && item.data) {
            parts.push({
              inlineData: {
                mimeType: item.mime_type || item.mimeType || 'image/png',
                data: item.data,
              },
            });
          }
        }
      } else if (typeof input === 'string') {
        parts.push({ text: input });
      }

      if (!parts.length) {
        return sendJson(res, 400, { error: 'No prompt or content provided to Gemini.' });
      }

      const ai = new GoogleGenAI({
        apiKey: keyToUse,
        httpOptions: {
          headers: {
            'User-Agent': 'aistudio-build',
          },
        },
      });

      const response = await ai.models.generateContent({
        model: targetModel,
        contents: parts.length === 1 && parts[0].text ? parts[0].text : { parts },
      });

      const responseText = response.text || '';

      // Check for generated image if any
      let generatedImage = null;
      if (response.candidates && response.candidates[0] && response.candidates[0].content && response.candidates[0].content.parts) {
        for (const part of response.candidates[0].content.parts) {
          if (part.inlineData && part.inlineData.data) {
            generatedImage = {
              data: part.inlineData.data,
              mime_type: part.inlineData.mimeType || 'image/png',
            };
            break;
          }
        }
      }

      const contentParts = [];
      if (responseText) {
        contentParts.push({ type: 'text', text: responseText });
      }
      if (generatedImage) {
        contentParts.push({ type: 'image', data: generatedImage.data, mime_type: generatedImage.mime_type });
      }

      return sendJson(res, 200, {
        text: responseText,
        steps: [
          {
            content: contentParts,
          },
        ],
      });
    } catch (err) {
      let errMsg = String((err && err.message) || err);
      const isAuthProblem = (
        errMsg.includes('401') ||
        errMsg.includes('UNAUTHENTICATED') ||
        errMsg.includes('ACCESS_TOKEN_TYPE_UNSUPPORTED') ||
        errMsg.includes('API_KEY_SERVICE_BLOCKED') ||
        errMsg.includes('invalid authentication credentials')
      );
      if (isAuthProblem) {
        console.warn('[Gemini Notice] Authentication required or key invalid:', errMsg.slice(0, 100));
        return sendJson(res, 401, {
          error: 'Gemini API authentication error: The provided Gemini API Key is invalid or not authorized. Please enter a valid Gemini API Key in ⚙️ Settings (get a free key at https://aistudio.google.com/app/apikey).',
          isAuthError: true,
          code: 'UNAUTHENTICATED'
        });
      }
      console.warn('[Gemini Notice] Request error:', errMsg.slice(0, 150));
      return sendJson(res, 500, { error: errMsg });
    }
    return;
  }

  // Fetches an arbitrary public image URL server-side and returns it as
  // base64, so the AI photo tools work even when the Main Image is a pasted
  // URL rather than a freshly-Browsed file (a plain browser fetch() could
  // itself be blocked by CORS on whatever host is serving that image).
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

  // Parses an uploaded Meesho category template into a field schema and stores it
  // (template.xlsx + schema.json) under categories/<slug>/, so the client can build
  // a form for any category, not just the bundled Keychains one.
  if (req.method === 'POST' && url.pathname === '/api/categories') {
    try {
      const body = await readBody(req, 25 * 1024 * 1024);
      const { fileBase64, name } = JSON.parse(body || '{}');
      if (!fileBase64) return sendJson(res, 400, { error: 'Missing file' });
      const buffer = Buffer.from(fileBase64, 'base64');
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      const schema = parseTemplateSchema(workbook, name);
      const slug = slugify(name || schema.name);
      saveCategory(slug, schema, buffer);
      sendJson(res, 200, { slug, schema });
    } catch (err) {
      sendJson(res, 400, { error: String((err && err.message) || err) });
    }
    return;
  }

  // Lists every category the server knows about (bundled Keychains + anything uploaded).
  if (req.method === 'GET' && url.pathname === '/api/categories') {
    try {
      sendJson(res, 200, { categories: listCategories() });
    } catch (err) {
      sendJson(res, 500, { error: String((err && err.message) || err) });
    }
    return;
  }

  // Returns one category's stored schema (so the client can restore it after a
  // refresh without re-uploading the template).
  if (req.method === 'GET' && url.pathname.startsWith('/api/categories/')) {
    try {
      const slug = decodeURIComponent(url.pathname.slice('/api/categories/'.length));
      const category = loadCategory(slug);
      sendJson(res, 200, { slug, schema: category.schema });
    } catch (err) {
      sendJson(res, 404, { error: String((err && err.message) || err) });
    }
    return;
  }

  // Clones the category's real Meesho template (all sheets, dropdowns, formatting
  // intact) and writes one data row per product starting at its first data row.
  if (req.method === 'POST' && url.pathname === '/api/export-xlsx') {
    try {
      const body = await readBody(req, 25 * 1024 * 1024);
      const { products, filename, category: categorySlug } = JSON.parse(body || '{}');
      if (!Array.isArray(products) || !products.length) {
        return sendJson(res, 400, { error: 'No products to export' });
      }
      const category = loadCategory(categorySlug || 'keychains');

      const buffer = await buildCatalogWorkbookBuffer(category, products);
      const safeName = (filename || 'Meesho-Catalog.xlsx').toString().replace(/[^A-Za-z0-9_.-]+/g, '_');
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
      const { products, category: categorySlug } = JSON.parse(body || '{}');
      if (!Array.isArray(products) || !products.length) {
        return sendJson(res, 400, { error: 'No products to export' });
      }
      const category = loadCategory(categorySlug || 'keychains');

      const zip = new JSZip();
      const usedNames = new Set();
      for (let i = 0; i < products.length; i++) {
        const buffer = await buildCatalogWorkbookBuffer(category, [products[i]]);
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
      const { fileBase64, category: categorySlug } = JSON.parse(body || '{}');
      if (!fileBase64) return sendJson(res, 400, { error: 'Missing file' });
      const category = loadCategory(categorySlug || 'keychains');
      const columnMap = columnMapFromSchema(category.schema);
      const reverseColumnMap = {};
      for (const [key, col] of Object.entries(columnMap)) reverseColumnMap[col] = key;

      const buffer = Buffer.from(fileBase64, 'base64');
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      const sheet = workbook.worksheets.find((ws) => /fill this/i.test(ws.name));
      if (!sheet) {
        return sendJson(res, 400, { error: 'Could not find a "...Fill this" data sheet in this file — is this the right file?' });
      }

      const titleCol = columnMap.title || 4;
      const errorRows = [];
      sheet.eachRow((row, rowNumber) => {
        if (rowNumber < category.schema.firstDataRow) return;
        const hasAnyData = row.getCell(titleCol).value; // Product Name — presence marks a real data row
        if (!hasAnyData) return;
        const errorStatus = cellText(row.getCell(2).value).trim();
        const errorMessage = cellText(row.getCell(3).value).trim();
        if (!errorMessage && errorStatus.toUpperCase() !== 'INVALID') return; // only rows with a flagged problem

        const product = {};
        for (const [col, key] of Object.entries(reverseColumnMap)) {
          const v = cellText(row.getCell(Number(col)).value).trim();
          if (v) product[key] = v;
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

server.listen(PORT, '0.0.0.0', () => {
  console.log(`Meesho Catalog Builder running at http://0.0.0.0:${PORT}`);
  console.log('Press Ctrl+C to stop.');
});
