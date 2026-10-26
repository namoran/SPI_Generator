/* zip-repack.js — rewrite a zip without recompressing untouched entries.
 *
 * An .xlsm is a zip, and re-deflating every part costs ~600 ms on a 1.4 MB
 * workbook. Only the parts we changed need compressing; the rest can be copied
 * across as the original compressed bytes. Result: ~40 ms and an output file
 * the same size as the input.
 *
 * Falls back (throws) on anything unusual — zip64, encryption, comments — so
 * the caller can use a plain re-zip instead.
 */
var ZipRepack = (function () {
  'use strict';

  var EOCD = 0x06054b50, CEN = 0x02014b50, LOC = 0x04034b50;

  var crate = (function () {
    var t = new Uint32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = (c & 1) ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(buf) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < buf.length; i++) c = crate[(c ^ buf[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  function rd32(b, i) { return (b[i] | (b[i + 1] << 8) | (b[i + 2] << 16) | (b[i + 3] << 24)) >>> 0; }
  function rd16(b, i) { return b[i] | (b[i + 1] << 8); }

  function findEocd(b) {
    for (var i = b.length - 22; i >= Math.max(0, b.length - 22 - 65535); i--) {
      if (rd32(b, i) === EOCD) return i;
    }
    throw new Error('not a zip file');
  }

  function repack(bytes, replacements) {
    if (!(bytes instanceof Uint8Array)) bytes = new Uint8Array(bytes);
    var eocd = findEocd(bytes);
    var count = rd16(bytes, eocd + 10);
    var cdOffset = rd32(bytes, eocd + 16);
    var cdSize = rd32(bytes, eocd + 12);
    if (cdOffset === 0xFFFFFFFF || count === 0xFFFF) throw new Error('zip64 not supported here');

    var out = [];
    var central = [];
    var written = 0;
    var p = cdOffset;

    for (var i = 0; i < count; i++) {
      if (rd32(bytes, p) !== CEN) throw new Error('bad central directory at ' + p);
      var flags = rd16(bytes, p + 8);
      var method = rd16(bytes, p + 10);
      var time = rd16(bytes, p + 12);
      var date = rd16(bytes, p + 14);
      var crc = rd32(bytes, p + 16);
      var csize = rd32(bytes, p + 20);
      var usize = rd32(bytes, p + 24);
      var nameLen = rd16(bytes, p + 28);
      var extraLen = rd16(bytes, p + 30);
      var commentLen = rd16(bytes, p + 32);
      var localOff = rd32(bytes, p + 42);
      var name = new TextDecoder().decode(bytes.subarray(p + 46, p + 46 + nameLen));

      if (csize === 0xFFFFFFFF || usize === 0xFFFFFFFF || localOff === 0xFFFFFFFF) {
        throw new Error('zip64 entry ' + name);
      }
      if (flags & 0x1) throw new Error('encrypted entry ' + name);
      if (method !== 0 && method !== 8) throw new Error('unsupported compression method ' + method);

      // Where the entry's bytes really start (local header has its own extras).
      if (rd32(bytes, localOff) !== LOC) throw new Error('bad local header for ' + name);
      var lNameLen = rd16(bytes, localOff + 26);
      var lExtraLen = rd16(bytes, localOff + 28);
      var dataStart = localOff + 30 + lNameLen + lExtraLen;

      var nameBytes = new TextEncoder().encode(name);
      var payload, outMethod = method, outCrc = crc, outCsize = csize, outUsize = usize;

      if (replacements[name]) {
        var raw = replacements[name];
        var deflated = fflate.deflateSync(raw, { level: 6 });
        if (deflated.length < raw.length) {
          payload = deflated; outMethod = 8;
        } else {
          payload = raw; outMethod = 0;
        }
        outCrc = crc32(raw);
        outCsize = payload.length;
        outUsize = raw.length;
      } else {
        payload = bytes.subarray(dataStart, dataStart + csize);
      }

      var newFlags = flags & ~0x08;   // sizes are written up front, no data descriptor

      var local = new Uint8Array(30 + nameBytes.length);
      var dv = new DataView(local.buffer);
      dv.setUint32(0, LOC, true);
      dv.setUint16(4, 20, true);
      dv.setUint16(6, newFlags, true);
      dv.setUint16(8, outMethod, true);
      dv.setUint16(10, time, true);
      dv.setUint16(12, date, true);
      dv.setUint32(14, outCrc, true);
      dv.setUint32(18, outCsize, true);
      dv.setUint32(22, outUsize, true);
      dv.setUint16(26, nameBytes.length, true);
      dv.setUint16(28, 0, true);
      local.set(nameBytes, 30);

      out.push(local, payload);

      var cen = new Uint8Array(46 + nameBytes.length);
      var cv = new DataView(cen.buffer);
      cv.setUint32(0, CEN, true);
      cv.setUint16(4, 20, true);
      cv.setUint16(6, 20, true);
      cv.setUint16(8, newFlags, true);
      cv.setUint16(10, outMethod, true);
      cv.setUint16(12, time, true);
      cv.setUint16(14, date, true);
      cv.setUint32(16, outCrc, true);
      cv.setUint32(20, outCsize, true);
      cv.setUint32(24, outUsize, true);
      cv.setUint16(28, nameBytes.length, true);
      cv.setUint16(30, 0, true);
      cv.setUint16(32, 0, true);
      cv.setUint16(34, 0, true);
      cv.setUint16(36, 0, true);
      cv.setUint32(38, 0, true);
      cv.setUint32(42, written, true);
      cen.set(nameBytes, 46);
      central.push(cen);

      written += local.length + payload.length;
      p += 46 + nameLen + extraLen + commentLen;
    }

    var cdStart = written;
    var cdLen = 0;
    central.forEach(function (c) { out.push(c); cdLen += c.length; });

    var end = new Uint8Array(22);
    var ev = new DataView(end.buffer);
    ev.setUint32(0, EOCD, true);
    ev.setUint16(8, count, true);
    ev.setUint16(10, count, true);
    ev.setUint32(12, cdLen, true);
    ev.setUint32(16, cdStart, true);
    ev.setUint16(20, 0, true);
    out.push(end);

    var total = 0;
    out.forEach(function (chunk) { total += chunk.length; });
    var result = new Uint8Array(total);
    var at = 0;
    out.forEach(function (chunk) { result.set(chunk, at); at += chunk.length; });
    return result;
  }

  return { repack: repack, crc32: crc32 };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = ZipRepack;
