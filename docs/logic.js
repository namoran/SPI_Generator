/* SPI Generator — logic.
 *
 * Port of SPI-Generator.py. Reads the RFC, works out how many units to add to
 * each task/pay-item cell of the SPI, and reports the changes. The actual
 * workbook edit happens in patch.js, which modifies the original file in place
 * so the VBA, styles and everything else survive untouched.
 *
 * Depends on PPRXlsx, PAY_ITEMS and TASKS. No DOM access.
 */
var SPI = (function () {
  'use strict';

  var RFC_SHEET = 'RFC';
  var SOW_SHEET = 'SOW Units';
  var RFC_FIRST_ROW = 22;   // Python: min_row=22 (skips the surcharge rows)
  var RFC_LAST_ROW = 111;   // Python: max_row=111
  var SOW_FIRST_ROW = 11;   // Python: min_row=11
  var SOW_FIRST_COL = 7;    // Python: min_col=7; TASK 1 lands on column I

  function payIndex(pi) {
    var i = PAY_ITEMS.indexOf(pi);
    if (i === -1) {
      throw new Error('Pay item "' + pi + '" is not in the pay item list. ' +
        'If FDEP added a new pay item, regenerate data.js (tools/gen-data.py).');
    }
    return i;
  }

  function taskIndex(task) {
    var i = TASKS.indexOf(task);
    if (i === -1) throw new Error('Unknown task "' + task + '"');
    return i;
  }

  /* Python get_spi_row / get_spi_column. */
  function cellFor(pi, task) {
    return { row: SOW_FIRST_ROW + payIndex(pi), col: SOW_FIRST_COL + taskIndex(task) };
  }

  /* Python load_rfc: rows 22..111, columns A..I, rows with no task skipped. */
  function loadRfc(sheet) {
    var m = sheet.range(RFC_FIRST_ROW, RFC_LAST_ROW, 1, 9);
    var out = [];
    for (var i = 0; i < m.length; i++) {
      if (m[i][0] === null) continue;
      out.push({
        task: 'TASK ' + String(m[i][0]),
        payItem: String(m[i][1]),
        uom: m[i][5],
        units: m[i][7],
        total: m[i][8],
        sourceRow: i + RFC_FIRST_ROW
      });
    }
    return out;
  }

  /* Work out every cell to touch, with its current and new value.
   * An RFC can list the same pay item more than once in a task (e.g. a
   * negative line removing a placeholder and a positive line entering the
   * actual amount), so lines for the same cell are folded together in RFC
   * order -- the same accumulation the Python version performs when it
   * adds each line into the live cell. One planned change per cell. */
  function plan(rfcSheet, sowSheet) {
    var items = loadRfc(rfcSheet);
    var byCell = {};
    var changes = [];
    items.forEach(function (it) {
      var at = cellFor(it.payItem, it.task);
      var units = it.units === null ? 0 : it.units;
      if (typeof units !== 'number') {
        throw new Error('Row ' + it.sourceRow + ' of the RFC has a non-numeric quantity for pay item ' +
          it.payItem + ' (' + it.task + ').');
      }
      var key = at.row + '|' + at.col;
      var ch = byCell[key];
      if (!ch) {
        var current = sowSheet.value(at.row, at.col);
        var oldVal = current === null ? 0 : current;
        if (typeof oldVal !== 'number') {
          throw new Error('Row ' + it.sourceRow + ' of the RFC has a non-numeric quantity for pay item ' +
            it.payItem + ' (' + it.task + ').');
        }
        ch = {
          task: it.task,
          payItem: it.payItem,
          row: at.row,
          col: at.col,
          address: colName(at.col) + at.row,
          oldValue: oldVal,
          units: 0,
          newValue: oldVal,
          lines: 0,
          parts: []
        };
        byCell[key] = ch;
        changes.push(ch);
      }
      ch.units += units;
      ch.newValue += units;
      ch.lines += 1;
      ch.parts.push(units);
    });
    return changes;
  }

  function colName(n) {
    var s = '';
    while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - r - 1) / 26; }
    return s;
  }

  function report(changes) {
    var out = [];
    changes.forEach(function (c) {
      out.push(c.units + ' unit(s) added to ' + c.oldValue + ' in ' + c.task + ', ' + c.payItem +
               '. The new value is ' + c.newValue + '.');
      if (c.lines > 1) {
        out.push('Note: this combines ' + c.lines + ' RFC lines (' + c.parts.join(', ') +
                 ') for ' + c.task + ', ' + c.payItem + '.');
      }
    });
    return out;
  }

  function open(buffer, label) {
    try {
      return PPRXlsx.open(buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer));
    } catch (e) {
      throw new Error('Could not read ' + label + ': ' + e.message);
    }
  }

  function sheetOrThrow(wb, name, label) {
    var ws = wb.sheet(name);
    if (!ws) throw new Error(label + ' has no "' + name + '" sheet. Sheets found: ' + wb.names.join(', '));
    return ws;
  }

  return {
    loadRfc: loadRfc,
    plan: plan,
    report: report,
    cellFor: cellFor,
    open: open,
    sheetOrThrow: sheetOrThrow,
    RFC_SHEET: RFC_SHEET,
    SOW_SHEET: SOW_SHEET
  };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = SPI;
