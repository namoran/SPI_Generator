/* SPI Generator — page wiring. */
(function () {
  'use strict';

  var files = { rfc: null, spi: null };
  var runBtn = document.getElementById('run');
  var result = document.getElementById('result');
  var banner = document.getElementById('banner');
  var changes = document.getElementById('changes');
  var rows = document.getElementById('rows');
  var download = document.getElementById('download');
  var url = null;

  function slot(name) { return document.getElementById('slot-' + name); }

  function setFile(name, file) {
    if (!file) return;
    files[name] = file;
    var el = slot(name);
    el.classList.add('has');
    var label = el.querySelector('.file');
    label.textContent = file.name + ' · ' + (file.size / 1048576).toFixed(2) + ' MB';
    label.classList.remove('hidden');
    runBtn.disabled = !(files.rfc && files.spi);
    result.classList.add('hidden');
  }

  document.querySelectorAll('input[type=file]').forEach(function (input) {
    input.addEventListener('change', function () { setFile(input.dataset.slot, input.files[0]); });
  });

  ['rfc', 'spi'].forEach(function (name) {
    var el = slot(name);
    ['dragenter', 'dragover'].forEach(function (ev) {
      el.addEventListener(ev, function (e) { e.preventDefault(); el.classList.add('over'); });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      el.addEventListener(ev, function (e) { e.preventDefault(); el.classList.remove('over'); });
    });
    el.addEventListener('drop', function (e) {
      var f = e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) setFile(name, f);
    });
  });

  function readBuffer(file) {
    if (file.arrayBuffer) return file.arrayBuffer();
    return new Promise(function (resolve, reject) {
      var r = new FileReader();
      r.onload = function () { resolve(r.result); };
      r.onerror = function () { reject(new Error('could not read ' + file.name)); };
      r.readAsArrayBuffer(file);
    });
  }

  function esc(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  function show(kind, text) {
    result.classList.remove('hidden');
    banner.className = 'banner ' + kind;
    banner.textContent = text;
  }

  runBtn.addEventListener('click', function () {
    runBtn.disabled = true;
    changes.classList.add('hidden');
    download.classList.add('hidden');
    show('ok', 'Processing…');

    setTimeout(function () {
      Promise.all([readBuffer(files.rfc), readBuffer(files.spi)])
        .then(function (bufs) {
          var rfcWb = SPI.open(bufs[0], 'the RFC file');
          var spiWb = SPI.open(bufs[1], 'the old SPI');
          var plan = SPI.plan(
            SPI.sheetOrThrow(rfcWb, SPI.RFC_SHEET, 'the RFC file'),
            SPI.sheetOrThrow(spiWb, SPI.SOW_SHEET, 'the old SPI')
          );

          var out = SPIPatch.build(new Uint8Array(bufs[1]), spiWb, SPI.SOW_SHEET, plan);

          if (url) URL.revokeObjectURL(url);
          url = URL.createObjectURL(new Blob([out], {
            type: 'application/vnd.ms-excel.sheet.macroEnabled.12'
          }));
          download.href = url;

          rows.innerHTML = plan.map(function (c) {
            return '<tr class="' + (c.units ? 'added' : '') + '">' +
              '<td>' + esc(c.task) + '</td>' +
              '<td>' + esc(c.payItem) + '</td>' +
              '<td>' + esc(c.address) + '</td>' +
              '<td class="num">' + esc(c.oldValue) + '</td>' +
              '<td class="num">' + esc(c.units) + '</td>' +
              '<td class="num">' + esc(c.newValue) + '</td></tr>';
          }).join('');

          show('ok', plan.length + ' pay item' + (plan.length === 1 ? '' : 's') +
            ' transferred. Download the updated SPI below — it opens with the macros intact.');
          changes.classList.remove('hidden');
          download.classList.remove('hidden');
        })
        .catch(function (err) {
          show('error', err && err.message ? err.message : String(err));
        })
        .then(function () { runBtn.disabled = false; });
    }, 20);
  });
})();
