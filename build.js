#!/usr/bin/env node
/* Builds the self-contained index.html:
 *   template.html + vendor libs + settings.json (baked defaults) + app.js
 * Run: node build.js
 */
'use strict';
const fs = require('fs');
const path = require('path');

const root = __dirname;
const read = (p) => fs.readFileSync(path.join(root, p), 'utf8');

// `</script` inside inlined JS would end the <script> tag early. Escaping the
// slash is a no-op inside JS strings/regexes, which is the only place it can
// legally appear in valid code.
const safeInline = (js) => js.replace(/<\/script/gi, '<\\/script');

const settings = JSON.parse(read('settings.json')); // validates JSON as a side effect

let html = read('src/template.html');
const inject = (token, content) => {
  if (!html.includes(token)) throw new Error('Missing token in template: ' + token);
  html = html.replace(token, () => content);
};

inject('<!--INJECT:XLSX-->', safeInline(read('vendor/xlsx.full.min.js')));
inject('<!--INJECT:JSPDF-->', safeInline(read('vendor/jspdf.umd.min.js')));
inject('/*INJECT:DEFAULTS*/null', safeInline(JSON.stringify(settings)));
inject('<!--INJECT:APP-->', safeInline(read('src/app.js')));

fs.writeFileSync(path.join(root, 'index.html'), html);
const kb = Math.round(fs.statSync(path.join(root, 'index.html')).size / 1024);
console.log(`index.html built (${kb} KB), defaults version ${settings.version}, ${settings.teachers.length} teachers`);
