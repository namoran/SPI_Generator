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

      // No cell there yet: build its markup. It is placed by column order
      // during assembly below, not by splicing at a position in the
      // original body text.
      inserts.push({
        markup: numericCell(address, '', ch.newValue),
        col: ch.col
      });
    });

    // Assemble the row body in column order: walk the existing cells,
    // emit each insert before the first existing cell with a higher
    // column, then emit the cell itself. Text between cells is carried
    // through untouched. No offset into the original body is applied to
    // the modified text -- a replacement whose markup is shorter or
    // longer than the original shifts every later offset, which is how
    // an insert could previously land inside a neighbouring cell's tag
    // and produce a workbook that would not open.
    var out = '';
    var pending = inserts.slice().sort(function (a, b) { return a.col - b.col; });
    var cursor = 0;
    cells.forEach(function (cell) {
      while (pending.length && pending[0].col < cell.col) out += pending.shift().markup;
      out += body.slice(cursor, cell.start);
      out += cell.raw;
      cursor = cell.end;
    });
    pending.forEach(function (ins) { out += ins.markup; });
    out += body.slice(cursor);

    // Keep the row's spans hint honest if we added cells outside it.
    var span = /\bspans="(\d+):(\d+)"/.exec(attrs);
    if (span && inserts.length) {
      var lo = parseInt(span[1], 10), hi = parseInt(span[2], 10);
      inserts.forEach(function (ins) { lo = Math.min(lo, ins.col); hi = Math.max(hi, ins.col); });
      attrs = setAttr(attrs, 'spans', lo + ':' + hi);
    }

    return '<row' + attrs + '>' + out + '</row>';
  }

  /* The full markup of one row in a sheet string, or null if the row is
   * absent (a pay item with no cells anywhere in the workbook). */
  function extractRow(xmlText, rn) {
    var rowRe = /<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g, m;
    while ((m = rowRe.exec(xmlText))) {
      var r = /\br="(\d+)"/.exec(m[0]);
      if (r && parseInt(r[1], 10) === rn) return m[0];
    }
    return null;
  }

  /* Map of cell address -> cell markup for one row's markup. A malformed
   * row (a cell tag nested inside another cell's attributes, say) parses
   * to a map that is missing the damaged cells, which is exactly what the
   * changed-row verification in patchSheet compares against. */
  function cellMap(rowXml) {
    var map = {};
    if (!rowXml) return map;
    var cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, c;
    while ((c = cellRe.exec(rowXml))) {
      var ref = /\br="([A-Z]+\d+)"/.exec(c[1]);
      if (ref) map[ref[1]] = c[0];
    }
    return map;
  }

  /* Patch the rows we care about, leaving the rest of the sheet string alone.
   * Rows that are missing from the sheet XML entirely (a pay item with no
   * existing cells anywhere in the workbook) get a new <row> element inserted
   * in row order — without this, changes landing on empty rows were silently
   * dropped from the output file. */
  function patchSheet(xml, changes) {
    var byRow = {};
    changes.forEach(function (ch) {
      (byRow[ch.row] = byRow[ch.row] || []).push(ch);
    });

    var missing = Object.keys(byRow).map(Number).sort(function (a, b) { return a - b; });
    function insertMissing(upToRow) {
      var s = '';
      while (missing.length && missing[0] < upToRow) {
        var rn = missing.shift();
        s += rebuildRow(' r="' + rn + '"', '', byRow[rn]);
      }
      return s;
    }

    var out = '';
    var last = 0;
    var rowRe = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g, m;
    while ((m = rowRe.exec(xml))) {
      var r = /\br="(\d+)"/.exec(m[1]);
      if (!r) continue;
      var rn = parseInt(r[1], 10);
      out += xml.slice(last, m.index);
      out += insertMissing(rn);
      if (byRow[rn]) {
        out += rebuildRow(m[1], m[2], byRow[rn]);
        missing = missing.filter(function (x) { return x !== rn; });
      } else {
        out += m[0];
      }
      last = rowRe.lastIndex;
    }
    out += xml.slice(last);
    if (missing.length) {
      var rest = insertMissing(Infinity);
      var close = out.indexOf('</sheetData>');
      if (close === -1) throw new Error('Could not find the end of the sheet data to add new pay item rows.');
      out = out.slice(0, close) + rest + out.slice(close);
    }

    verifyPatched(xml, out, changes);
    return out;
  }

  /* Fail loudly rather than hand back a file missing changes -- and check
   * the written value, not just that the cell exists, so a change that was
   * overwritten or written incorrectly cannot slip through silently. */
  function verifyPatched(xml, out, changes) {
    var byRow = {};
    changes.forEach(function (ch) {
      (byRow[ch.row] = byRow[ch.row] || []).push(ch);
    });

    changes.forEach(function (ch) {
      var address = colName(ch.col) + ch.row;
      var at = out.indexOf('<c r="' + address + '"');
      if (at === -1) {
        throw new Error('Internal error: change to cell ' + address + ' (' + ch.task + ', ' + ch.payItem + ') was not written.');
      }
      var end = out.indexOf('</c>', at);
      var cellXml = end === -1 ? '' : out.slice(at, end);
      var v = /<v>([^<]*)<\/v>/.exec(cellXml);
      if (!v || v[1] !== numStr(ch.newValue)) {
        throw new Error('Internal error: change to cell ' + address + ' (' + ch.task + ', ' + ch.payItem +
          ') was written with the wrong value (expected ' + numStr(ch.newValue) + ').');
      }
    });

    // Also verify the cells the plan did NOT change, in every row it
    // touched: each one must survive with its original markup, and no
    // unplanned cell may appear. The value check above cannot see damage
    // to an unchanged neighbour -- on PO C0C875 an insert landed inside
    // the following cell's tag, both changed cells read back correctly,
    // and the broken file was reported as success.
    Object.keys(byRow).forEach(function (rnKey) {
      var rn = parseInt(rnKey, 10);
      var planned = {};
      byRow[rn].forEach(function (ch) { planned[colName(ch.col) + rn] = true; });
      var origMap = cellMap(extractRow(xml, rn));
      var newMap = cellMap(extractRow(out, rn));
      Object.keys(origMap).forEach(function (addr) {
        if (planned[addr]) return;
        if (newMap[addr] !== origMap[addr]) {
          throw new Error('Internal error: cell ' + addr + ' in row ' + rn +
            ' was not part of this change but was altered or lost while writing the row.');
        }
      });
      Object.keys(newMap).forEach(function (addr) {
        if (!origMap[addr] && !planned[addr]) {
          throw new Error('Internal error: cell ' + addr + ' in row ' + rn +
            ' appeared in the output but was not part of this change.');
        }
      });
    });
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

  return { build: build, patchSheet: patchSheet, forceRecalc: forceRecalc, verifyPatched: verifyPatched };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = SPIPatch;
