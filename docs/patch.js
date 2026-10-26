/* SPI Generator — workbook writer.
 *
 * Instead of re-saving the workbook through a spreadsheet library (which drops
 * VBA, styling and anything else it doesn't model), this edits the sheet XML in
 * the original .xlsm and re-zips it. Everything not touched — macros, styles,
 * drawings, print settings — comes out byte-for-byte identical.
 *
 * Two things must happen for Excel to show the right numbers afterwards:
 *   1. the target cells get their new value and their row is unhidden;
 *   2. fullCalcOnLoad is set so Excel recalculates the totals that depend on
 *      those cells (the Python tool got this for free by dropping calcChain).
 */
var SPIPatch = (function () {
  'use strict';

  var enc = new TextEncoder();

  function numStr(n) {
    if (!isFinite(n)) throw new Error('refusing to write a non-finite number');
    if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
    return String(Number(n.toPrecision(15)));
  }

  function setAttr(attrs, name, value) {
    var re = new RegExp('\\b' + name + '="[^"]*"');
    if (re.test(attrs)) return attrs.replace(re, name + '="' + value + '"');
    return attrs + ' ' + name + '="' + value + '"';
  }

  function dropAttr(attrs, name) {
    return attrs.replace(new RegExp('\\s*\\b' + name + '="[^"]*"', 'g'), '');
  }

  function colIndex(letters) {
    var n = 0;
    for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n;
  }

  function colName(n) {
    var s = '';
    while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - r - 1) / 26; }
    return s;
  }

  /* New cell markup for a numeric value, keeping the existing style. */
  function numericCell(address, attrs, value) {
    var keep = dropAttr(dropAttr(attrs || '', 'r'), 't');   // address is added back; numbers have no t
    return '<c r="' + address + '"' + keep + '><v>' + numStr(value) + '</v></c>';
  }

  function rebuildRow(attrs, body, changes) {
    if (body === undefined || body === null) body = '';

    // Unhide the row: Excel treats a missing hidden attribute as visible.
    if (/\bhidden="(1|true)"/.test(attrs)) attrs = dropAttr(attrs, 'hidden');

    var cells = [];
    var cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, c;
    while ((c = cellRe.exec(body))) {
      var ref = /\br="([A-Z]+)(\d+)"/.exec(c[1]);
      cells.push({
        start: c.index,
        end: cellRe.lastIndex,
        raw: c[0],
        address: ref ? ref[1] + ref[2] : null,
        col: ref ? colIndex(ref[1]) : 0,
        attrs: c[1],
        inner: c[2] === undefined ? '' : c[2]
      });
    }

    var sorted = changes.slice().sort(function (a, b) { return a.col - b.col; });
    var inserts = [];

    sorted.forEach(function (ch) {
      var address = colName(ch.col) + ch.row;
      var hit = cells.filter(function (x) { return x.address === address; })[0];

      if (hit) {
        if (/<f\b/.test(hit.inner)) {
          throw new Error('Cell ' + address + ' (' + ch.task + ', ' + ch.payItem + ') holds a formula, ' +
            'so adding units there would destroy it. Unhide the pay item row in the SPI and try again.');
        }
        hit.raw = numericCell(address, hit.attrs, ch.newValue);
        return;
      }

      // No cell there yet: build one and remember where to splice it in.
      var before = cells.filter(function (x) { return x.col > ch.col; })[0];
      inserts.push({
        at: before ? before.start : body.length,
        markup: numericCell(address, '', ch.newValue),
        col: ch.col
      });
    });

    var out = body;
    cells.slice().reverse().forEach(function (cell) {
      out = out.slice(0, cell.start) + cell.raw + out.slice(cell.end);
    });
    inserts.sort(function (a, b) { return b.at - a.at; }).forEach(function (ins) {
      out = out.slice(0, ins.at) + ins.markup + out.slice(ins.at);
    });

    // Keep the row's spans hint honest if we added cells outside it.
    var span = /\bspans="(\d+):(\d+)"/.exec(attrs);
    if (span && inserts.length) {
      var lo = parseInt(span[1], 10), hi = parseInt(span[2], 10);
      inserts.forEach(function (ins) { lo = Math.min(lo, ins.col); hi = Math.max(hi, ins.col); });
      attrs = setAttr(attrs, 'spans', lo + ':' + hi);
    }

    return '<row' + attrs + '>' + out + '</row>';
  }

  /* Patch just the rows we care about, leaving the rest of the sheet string alone. */
  function patchSheet(xml, changes) {
    var byRow = {};
    changes.forEach(function (ch) {
      (byRow[ch.row] = byRow[ch.row] || []).push(ch);
    });

    var out = '';
    var last = 0;
    var rowRe = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g, m;
    while ((m = rowRe.exec(xml))) {
      var r = /\br="(\d+)"/.exec(m[1]);
      if (!r || !byRow[parseInt(r[1], 10)]) continue;
      out += xml.slice(last, m.index);
      out += rebuildRow(m[1], m[2], byRow[parseInt(r[1], 10)]);
      last = rowRe.lastIndex;
    }
    out += xml.slice(last);
    return out;
  }

  function forceRecalc(workbookXml) {
    var m = /<calcPr\b[^>]*?\/>/.exec(workbookXml);
    if (m) {
      var tag = setAttr(m[0].replace(/\s*\/>$/, ''), 'fullCalcOnLoad', '1') + '/>';
      return workbookXml.slice(0, m.index) + tag + workbookXml.slice(m.index + m[0].length);
    }
    return workbookXml.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
  }

  /* Rebuild the .xlsm: same parts, one sheet patched.
   * Unchanged parts keep their original compressed bytes, so this stays fast
   * and the download is the same size as the file you started with. */
  function build(originalBytes, wb, sheetName, changes) {
    if (!(originalBytes instanceof Uint8Array)) originalBytes = new Uint8Array(originalBytes);

    var path = wb.sheetPath(sheetName);
    var parts = fflate.unzipSync(originalBytes, {
      filter: function (f) { return f.name === path || f.name === 'xl/workbook.xml'; }
    });

    var replacements = {};
    var dec = new TextDecoder('utf-8');
    replacements[path] = enc.encode(patchSheet(dec.decode(parts[path]), changes));
    replacements['xl/workbook.xml'] = enc.encode(forceRecalc(dec.decode(parts['xl/workbook.xml'])));

    try {
      return ZipRepack.repack(originalBytes, replacements);
    } catch (e) {
      // Anything exotic about the zip: fall back to a full re-zip.
      var files = fflate.unzipSync(originalBytes);
      for (var name in replacements) files[name] = replacements[name];
      return fflate.zipSync(files, { level: 3 });
    }
  }

  return { build: build, patchSheet: patchSheet, forceRecalc: forceRecalc };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = SPIPatch;
