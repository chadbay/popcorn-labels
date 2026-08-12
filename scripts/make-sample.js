#!/usr/bin/env node
/* Generates sample_orders.xlsx shaped like the weekly export the original
 * script expects: two header rows, then data rows with teacher in column A,
 * item code in column C, and student count in column G.
 * Run: node scripts/make-sample.js
 */
'use strict';
const path = require('path');
const XLSX = require(path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js'));

const rows = [
  ['Teacher', 'Student', 'Item', 'SKU?', 'Price', 'Qty?', 'Count', 'H', 'I', 'J'],
  ['', '', '', '', '', '', '', '', '', ''],
  // teacher, -, item code, -, -, -, count
  ['Aune', '', 'SE-POPCORN', '', '', '', 12, '', '', ''],
  ['Berry', '', 'SE-POPCORN', '', '', '', 15, '', '', ''],
  ['Davis', '', 'SE-POPCORN', '', '', '', 10, '', '', ''],
  ['Davis', '', 'SE-POPCORN-SPRING-ONLY', '', '', '', 3, '', '', ''],  // same teacher, second code -> sums to 13
  ['Judy Hurst', '', 'SE-POPCORN', '', '', '', 9, '', '', ''],
  ['Martinez', '', 'SE-POPCORN', '', '', '', 14, '', '', ''],
  ['Bernacki', '', 'SE-POPCORN', '', '', '', 16, '', '', ''],
  ['Brightwell', '', 'SE-POPCORN', '', '', '', 11, '', '', ''],
  ['Dunagan', '', 'SE-POPCORN', '', '', '', 13, '', '', ''],
  ['Martin', '', 'SE-POPCORN', '', '', '', 17, '', '', ''],
  ['Shipley', '', 'SE-POPCORN', '', '', '', 12, '', '', ''],
  ['Espich', '', 'SE-POPCORN', '', '', '', 18, '', '', ''],
  ['Falsone', '', 'SE-POPCORN', '', '', '', 14, '', '', ''],
  ['Kissinger', '', 'SE-POPCORN', '', '', '', 10, '', '', ''],
  ['Richardson', '', 'SE-POPCORN', '', '', '', 15, '', '', ''],
  ['Slawson', '', 'SE-POPCORN', '', '', '', 13, '', '', ''],
  ['Chamness', '', 'SE-POPCORN', '', '', '', 20, '', '', ''],
  ['Grant', '', 'SE-POPCORN', '', '', '', 11, '', '', ''],
  ['Sarah Hurst', '', 'SE-POPCORN', '', '', '', 12, '', '', ''],
  ['Medina', '', 'SE-POPCORN', '', '', '', 14, '', '', ''],
  ['Oden', '', 'SE-POPCORN', '', '', '', 9, '', '', ''],
  ['Rivas', '', 'SE-POPCORN', '', '', '', 16, '', '', ''],
  ['Brookshire', '', 'SE-POPCORN', '', '', '', 13, '', '', ''],
  ['Hemphill', '', 'SE-POPCORN', '', '', '', 15, '', '', ''],
  ['Hernandez', '', 'SE-POPCORN', '', '', '', 12, '', '', ''],
  ['Ellis', '', 'SE-POPCORN', '', '', '', 14, '', '', ''],
  ['Roe', '', 'SE-POPCORN', '', '', '', 10, '', '', ''],
  ['Rosier', '', 'SE-POPCORN', '', '', '', 17, '', '', ''],
  ['Fields', '', 'SE-POPCORN', '', '', '', 19, '', '', ''],
  ['Melanson', '', 'SE-POPCORN', '', '', '', 12, '', '', ''],
  ['Park', '', 'SE-POPCORN', '', '', '', 16, '', '', ''],
  ['Platt', '', 'SE-POPCORN', '', '', '', 11, '', '', ''],
  ['Polivka', '', 'SE-POPCORN', '', '', '', 14, '', '', ''],
  // Edge cases:
  ['Nguyen', '', 'SE-POPCORN', '', '', '', 8, '', '', ''],        // new teacher, not in roster
  ['Unknown', '', 'SE-POPCORN', '', '', '', 4, '', '', ''],       // teacher not specified
  [null, '', 'SE-POPCORN', '', '', '', 2, '', '', ''],            // empty teacher cell
  ['Aune', '', 'SE-TSHIRT', '', '', '', 99, '', '', ''],          // non-popcorn item, ignored
];

const ws = XLSX.utils.aoa_to_sheet(rows);
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'Orders');
const out = path.join(__dirname, '..', 'sample_orders.xlsx');
// XLSX.writeFile's fs autodetection fails under newer Node; write a buffer instead.
require('fs').writeFileSync(out, XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
console.log('Wrote', out);
