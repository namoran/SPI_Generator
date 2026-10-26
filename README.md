# SPI Generator

Transfers the line items from a PRP change order (RFC) into an SPI and gives you
the updated workbook back.

**Use it at <https://namoran.github.io/SPI_Generator/>** — the old Streamlit
version is still at <https://spigenerator.streamlit.app/>.

Drop in the RFC and the SPI, press **Generate SPI**, download the result.
Nothing is uploaded: the workbook is edited in the browser tab.

## What comes back

The output is the original hand-maintained SPI with only the changed cells
touched. That matters, because the SPI is a macro-enabled workbook that people
also use for invoicing:

- the VBA project is copied through untouched (`vbaProject.bin` is byte-for-byte
  identical),
- styles, data validation, drawings, print settings and every other sheet are
  left alone,
- pay item rows that were hidden become visible, so the added quantities show,
- `fullCalcOnLoad` is set so Excel recalculates the row totals on open.

Re-saving through a spreadsheet library would quietly drop most of that, which
is why `docs/patch.js` edits the sheet XML inside the original zip instead, and
`docs/zip-repack.js` copies the compressed bytes of every unchanged part. That
also makes it quick: about 100–250 ms for a 1.4 MB workbook, versus ~3 s to
re-compress all 128 parts.

## How it works

`SPI-Generator.py` is the original Streamlit app and remains the reference
implementation. `docs/` is a port of its logic to plain JavaScript:

| file | job |
| --- | --- |
| `docs/index.html` | the page |
| `docs/ui.js` | drag/drop, planning, download |
| `docs/logic.js` | port of the Python comparison/planning logic |
| `docs/patch.js` | XML surgery on the SOW sheet + zip rebuild |
| `docs/zip-repack.js` | rewrite a zip without recompressing unchanged parts |
| `docs/ppr-xlsx.js` | minimal `.xlsm` reader |
| `docs/fflate.js` | zip decompression, vendored |
| `docs/data.js` | pay item / task lists, generated from the Python source |

## Tests

```sh
PYTHON=.venv/bin/python node tests/generate.test.mjs
```

The test builds a workbook with the browser code, then runs `tests/verify.py`
with openpyxl to confirm:

- every expected cell holds the right arithmetic,
- no unrelated cell in the SOW sheet changed,
- the rows that received a value are no longer hidden,
- `vbaProject.bin` is byte-identical and all 128 parts are present,
- `fullCalcOnLoad` is set.

Fixtures are expected in `tests/fixtures/` and are not committed, since they are
real client workbooks:

```
5632 RFC1 VERSION 4.xlsm                            (the RFC)
PART2-AttachmentB-SPI-138506266-SA - Copy.xlsm      (the SPI to update)
```

A useful extra check is LibreOffice, which recalculates the whole workbook:

```sh
soffice --headless --convert-to xlsx --outdir /tmp/lo tests/tmp/actual.xlsm
```

Every affected row should then satisfy `H == SUM(I:R)`.

If FDEP adds a pay item, update `SPI-Generator.py` and regenerate the browser
copies of the lists:

```sh
python3 tools/gen-data.py          # rewrite docs/data.js
python3 tools/gen-data.py --check  # fail if docs/data.js is stale
```

## Deploying

GitHub Pages serves `docs/` from `main`. There is no build step — commit and
push.
