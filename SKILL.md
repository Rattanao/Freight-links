---
name: manifest-enter-edi-check
description: >
  Reconciles an ocean-freight carrier's MANIFEST (.xls/.xlsx, free-form cargo report)
  against a customer's ENTER document (.pdf, per-B/L confirmation/amendment sheet,
  often with hand-typed red annotations for STATUS/DG/REEFER TEMP/TRANSIT — including
  notes that cover a whole range of sub-B/Ls at once) before filing a Thailand customs
  import declaration (ใบขนสินค้าขาเข้า). Auto-detects which of two known ENTER PDF
  layouts is in use. Produces a single color-coded EDI.xlsx report comparing every
  sub-B/L on CONSIGNEE, CONTAINER NO., STATUS, TOTAL PACKAGE, PACKAGING type, GROSS
  WEIGHT, MEASUREMENT, MARKS, DESCRIPTION, REEFER TEMP, DG (CLASS/UN), and
  TRANSIT/TRANSHIPMENT, plus a TOTAL row cross-checking summed values against the
  totals each document states for itself, and a critical row for any B/L that exists
  in ENTER but is entirely missing from MANIFEST (cargo at risk of being left off the
  declaration).
  Use this whenever the user uploads or references files named/shaped like
  MANIFEST.xls, MANIFEST_*.xls, or ENTER.pdf, or asks in Thai or English to
  "ตรวจสอบเอกสารขาเข้าเรือ", "เทียบ MANIFEST กับ ENTER", "ทำ EDI", "ตรวจ B/L",
  or similar import-document reconciliation — even if only one of the two files is
  attached yet, or the wording is loose. This is Rattana's own confirmed report format
  (built and iterated across real jobs for multiple customers/carriers) — trust it
  over improvising a new layout from scratch.
---

# MANIFEST × ENTER → EDI.xlsx reconciliation

This packages a working, already-validated script (`scripts/build_edi.py`) plus the
exact report conventions the user settled on after many rounds of feedback. Don't
re-derive the format — run the script and follow the conventions below.

## When the user gives you files

1. Save the MANIFEST file and the ENTER PDF somewhere sensible (their existing
   `input/` folder if the project already has one, otherwise the folder they
   uploaded to). Any filename containing "MANIFEST"/"ENTER" is fine — the script
   also just picks up the only `.xls`/`.xlsx` and only `.pdf` in the folder if the
   names don't match.
2. **If `EDI.xlsx` already exists at the destination and is open in Excel**, saving
   will raise `PermissionError: [Errno 13]`. Build to a scratch path first, verify,
   then copy over the real path (retry the copy once — Excel sometimes releases the
   lock a moment later); if it's still locked, hand the user the scratch copy and ask
   them to close Excel before you copy it into place.
3. Run:
   ```bash
   python scripts/build_edi.py --manifest <path> --enter <path> --outdir <job folder>/output
   ```
   **Always write the finished `EDI.xlsx` into the job's `output/` folder** (sibling of
   `input/`) — the user asked for this on every run (2026-10-26). Without `--outdir` the
   script now defaults to that `output/` folder too; never leave the report in the job root
   or a scratch folder as the final copy.
   (or with no args at all if both files sit in the current folder / an `input/`
   subfolder — it auto-detects them). It prints a summary: total B/Ls, how many rows
   need review, and a list of exactly what's wrong per B/L — relay that to the user
   before they open the spreadsheet, leading with anything DG/REEFER/TRANSIT-related
   since those carry legal weight on the declaration.
4. Send the resulting `EDI.xlsx` to the user.

## If the MANIFEST layout doesn't match

`parse_manifest()` in the script assumes one specific free-form carrier layout
(column 0 = B/L number / S:/C:/N: lines, column 5 **or 6** = marks + container number
lines, column 10 = package count + STATUS + description [+ DG/REEFER/TRANSIT tokens],
column 16 = gross weight then measurement — see the header comment in the script and
`references/format-notes.md`). If a new MANIFEST file's columns are laid out
differently, read a raw dump of it first (`xlrd`, print every row) before assuming
the parser will work — adjust the column indices in `parse_manifest()` to match
rather than guessing blind. The B/L prefix itself (e.g. "HASL" for Heung-A) is
auto-detected from the file, so a different carrier with the same layout shape
usually needs no code changes at all.

The marks/container column specifically is **auto-detected** too
(`detect_marks_col()`), not hardcoded to 5 — some MANIFEST exports (confirmed on a real
Heung-A KMTC file) put the actual data one column to the right of where the header
cell's merge range visually suggests (header "MARKS AND NUMBERS" merges starting at
column 7, but the data underneath starts at column 6). Trust the auto-detected column
over the header position if they ever disagree — it's picked from where the container-
number pattern (`\d+\.\s*[A-Z]{4}\d{6,7}`, e.g. `1. HALU5202915`) actually appears most
often in the data rows, not from where a header label happens to sit.

Also watch for the whole-file grand-total footer (`PORT TOTAL:` / `TOTAL MEASUREMENT:`
/ `TOTAL GROSSWEIGHT:` / `TOTAL PACK:`) leaking into the **last** B/L block's MARKS —
the per-page `PortOfDischarge: ...` footer that normally ends a block sits in column 0
and is caught by `END.match(_cell(rows[j], 0))`, but the true end-of-file grand total on
a real file sat in **column 1** instead (`PORT TOTAL:` at col 1, not col 0). The end-of-
block check scans every column in the row for an `END` match now, not just column 0 —
don't narrow that back down to column 0 only.

## If the ENTER layout doesn't match

`parse_enter()` auto-detects and handles **four** known block layouts (it tries the
literal anchor `B/L NO. :` — with a space before the colon — first, then `B/L NO.:`
without the space, then falls back to the column-position-based layouts C and D if no
anchor is found anywhere in the document). See `references/format-notes.md` for a
full worked example of each:

- **Layout A**: separate `MARK&NO.` / `NO. OF PKGS` / `DESCRIPTION OF GOODS` columns,
  block ends on the anchor `CNTR :` then `B/L NO. :` (space before the colon) then the
  consignee name.
- **Layout B**: no separate MARKS column — one free-text paragraph covers marks and
  description together — block ends on `CONSIGNEE :` / `CONT.NO.:` / `B/L NO.:` (no
  space before this colon) in that order. Because marks and description aren't really
  separable in this layout, `ENT[bl]['combined_text'] = True` is set and the report
  compares the combined text as a **word bag** (`same_bag()`, order-independent)
  instead of the strict character-for-character check used for layout A — flag it
  orange (needs human confirmation) rather than red (confident mismatch) when the
  word sets don't match, since the discrepancy may just be how the two documents
  happened to wrap the same content.
- **Layout C** (`parse_enter_layout_c()`): a real ruled table with no `B/L NO.` anchor
  text anywhere — seen on Cargoport (Thailand) Ltd.'s "AMENDMENT" letter form. Columns
  are identified by x0 position bands instead of anchor text, and a sub-B/L's number
  is split across two lines (base number, then its suffix letter directly below in the
  same column) rather than printed as one string — this is also why the generic
  fallback parser misses 8 of 9 sub-B/Ls on this layout (it can't reassemble a key
  that's never printed contiguously). STATUS and CONTAINER NO. are usually stated once
  for the whole document rather than per block. See `references/format-notes.md` for
  the full column-band map and the reasoning behind each design choice.
- **Layout D** (`parse_enter_layout_d()`): a real ruled table like layout C, seen on AGN
  International Logistics (Thailand) Co., Ltd.'s "AMENDMENT" form — anchor is the
  literal header text `B/L CHANGE NO.` in the leftmost column. Unlike layout C, each
  sub-B/L's full number is printed on **one line** (never split base+suffix), and the
  CONSIGNEE column has a `NOTIFY:` sub-line stacked directly under the consignee name
  in the *same* x0 band — split the column's lines at that label to get `cons`/`notify`
  separately (`rec['notify']` — a new ENT field only this layout populates). The
  `CARGO MOVEMENT n` line sits inside the DESCRIPTION column's x0 band but is **not**
  goods description — filter it out of `desc` (same exclusion category as MANIFEST's
  STATUS token/container-type code) and let the separate CARGO MOVEMENT anchor-scan (see
  below) pick it up on its own. This form's FreeText annotations aren't limited to
  STATUS/DG/REEFER/TRANSIT corrections the way other carriers' are — it also uses one
  for `NOTIFY:` (duplicating the printed NOTIFY line) — `_apply_annot()` must route a
  `NOTIFY:`-prefixed annotation to its own bucket, not the STATUS catch-all, or it
  silently corrupts `status_raw` for that B/L.
  Some AGN files draw the **whole page sideways**: the content is landscape but
  `page.rotation == 0`, and every line has `dir == (0,-1)`. In that case the raw x/y are
  swapped and nothing lands in the right column band. `_page_lines()` detects the majority
  line direction and de-rotates coordinates before layout D reads them. Use it, not raw
  `bbox`, for any new layout-D code. On these files, QUANTITY can also wrap
  (`203` / `CARTONS`), so join all lines in the QTY band. STATUS is printed once in the
  header (`STATUS : *** เร่งเปิดตู้เข้าโกดังด่วน ***`) and applies to every B/L unless a
  FreeText annotation overrides it. Its Thai comes out garbled by the font
  (`เร=งเป@ดตCDเขDาโกดIงด=วน`), and `_LD_TH_FIX` maps it back.
  **H.I.T. INTERCON (Handle Inter)** uses this same `B/L CHANGE NO.` form (added 2026-10-27,
  SAWASDEE SPICA 2609S job), but its columns sit further left: MARKS ~166 | QUANTITY ~308 |
  DESCRIPTION ~383 | CONSIGNEE ~535 | GROSS WEIGHT ~704. With AGN's edges the description fell
  into QUANTITY and the consignee into DESC. Pages are identified **per page** by the footer code
  `[FM-CS-01-05]` (the company name is only on page 1, and the hahm bundle has one Handle page
  among AGN pages). They use their own edges `_LD_HANDLE_EDGES`, shifted by the page's `QUANTITY`
  header x, since one hahm page sits ~6pt further left. The block stops 6pt above `TOTAL` because
  the `12,437.740 KGS.` total sits just above it. Result: 8/8 B/Ls passed. This also fixed a
  false CONSIGNEE/DESC flag on hahm B/L 801096. The other 17 jobs were identical cell for cell.
- **Layout F** (`parse_enter_layout_f()`): CMS Logistics "INBOUND CARGO MANIFEST" —
  page 1 is a cover letter (totals + "AS PER ATTACHED SHEET"), then per block: goods
  rows (marks x0<155 / pkgs <255 / desc <480 / KGS.+CBM ≥480), then `CONSIGNEE : <name>`
  with `D/O No. <B/L>` on the same row, then `CNTR NO: <cntr>/<size>`. The B/L is only
  printed after the `D/O No.` label, so annotation anchoring strips that label first
  (otherwise the page has no anchors and FreeText STATUS notes are silently dropped).
  Blocks **split across pages** (goods at the bottom of one page, CONSIGNEE at the top of
  the next) are handled by carrying leftover rows forward. The totals line is
  `TOTAL QUANTITY n PACKAGES G.W. n KGS. MEAS. n CBM.`; a whole-document STATUS note
  like `เปิดตู้เข้าโกดัง-ทุกBL` applies to every B/L — so does any plain FreeText STATUS
  note placed **below** the `TOTAL QUANTITY` line (THE M CORP puts `เปิดตู้เข้าโกดัง` there
  with no "ทุกBL" wording). THE M CORP GLOBAL uses this same layout, printing the label as
  `D / O No .` with spaces, and some of its PDFs use control characters (`\x04`/`\x05`/
  `\x07`/`\x08`) instead of spaces between words. `pymupdf.Page.get_text` is wrapped at
  import time (`_clean_get_text`) to turn those into spaces everywhere — keep it, or
  anchors fail and openpyxl raises `IllegalCharacterError`.
  **Never copy a job folder's local `build_edi.py` over this skill's script** — that has
  already wiped out layout F twice. Merge the changes into this file instead.
  A browser version of this layout (for colleagues, no install) lives at
  https://github.com/Rattanao/M-CORP (`docs/core.js` = JS port, page at
  https://rattanao.github.io/M-CORP/). If layout F parsing changes here, port the same
  change to that repo's `docs/core.js` too.
  **Penanshin uses this same layout F parser.** Its variant of the form differs in a few ways:
  - The consignee label is `CONSIGNEE [n]`, and the name sits in its own column (x≈136). It
    can wrap over several lines (name + address) down to the container row, and all of those
    lines are joined.
  - The container label is `CTRN NO` (not `CNTR NO:`).
  - `D/O No.` and the B/L number are separate text items on the **CTRN NO row**, not the
    CONSIGNEE row. The parser looks for the B/L anywhere from the CONSIGNEE row through the
    container row.
  - The Weight and Meas column starts at x≈460, not ≈488. The column edge is read from each
    page's `Weight…` header instead of `_LF_DESC_MAX`.
  - The footer is `Print Date/Time:` with the date as a separate item. Everything from the
    footer down is dropped, along with stray `:` items.
  - The totals line is `TOTAL PKGS: n PACKAGES` / `TOTAL G.W.: n KGS.` / `TOTAL MEAS: n CBMS.`.
  - The STATUS note is a FreeText `เปิดตู้เข้าโกดัง ทุก BL` at the top of page 1.

  Verified 2026-10-13: after these changes, the CMA, Cargoport, DSV, Gateway, hahm, Quality,
  Shipco and THE M CORP jobs produced reports identical to the previous script. The same
  generalization was **not** ported to the M-CORP page; it doesn't need it.
  Its browser version lives at https://github.com/Rattanao/Penanshin (page at
  https://rattanao.github.io/Penanshin/; see "Web version" below). Any layout F change must
  keep both CMS/THE M CORP and Penanshin working, so rerun one job of each.

  **EASTERN AIR LOGISTICS** (added 2026-10-26) uses this same CMS-style form. Two fixes came
  from it and from the CMS SKY ORION 2612S job:
  - Some marks contain their own `CONSIGNEE:` line (`SHIPPER : ESA` / `CONSIGNEE:` / `J & U …`).
    A CONSIGNEE row only counts as a block header if its own `CNTR NO:` row comes before the
    next CONSIGNEE row. Otherwise that B/L's whole block was swallowed (B/L A came out empty,
    5 pkgs / 824 kg short of ENTER's own total).
  - `CARGO MOVEMENT: L` / `CARGO MOVEMENT: 7` is printed inside each B/L's goods block, above its
    `D/O No.` row. Layout F reads it from the block itself (`_mv_local`), because the general
    scanner ties a movement line to the B/L *above* it, which is the wrong B/L on this form.
  - Every CARGO MOVEMENT reader (general scanner, DESC comparison, layout D filter) now accepts
    `CARGO MOVEMENT 7`, `CARGO MOVEMENT: 7` and `CARGO MOVEMENT = 7`.
- **Layout G** (`parse_enter_layout_g()`): DSV "AMENDMENT OF MANIFEST" — **one page per
  B/L**. Each page has the B/L after `B/L CHANGED:`, the consignee below `Consignee` (up to
  `Notify party`), and a printed `STATUS: "ลากตู้ด่วน" CY **HOT BOX**` line. Under
  `Mark & Numbers` the columns are: marks x<175 | `40   Carton` 175–255 | desc 255–495 |
  KGS/CBM ≥495. The container line `SKHU1857320 / 40'HC` sits at the bottom of the desc
  column; it is pulled out as `cont`, not desc. Everything from `DSV Standard Terms` down is
  dropped. Added 2026-10-15 (Heung-A HOCHIMINH 2610S job); the Cargoport, Gateway,
  Penanshin, Quality and THE M CORP reports were unchanged.
  The same job exposed a MANIFEST bug: the container-type-code filter (`45G0X1`) also
  swallowed invoice lines like `V266411`. The filter now allows at most 1–3 more characters
  after `[A-Z]\d`.

- **Layout H** (`parse_enter_layout_h()`): MAC-NELS SHIPPING "ATTACHED SHEETS". The header is
  `MARKS NO.` / `PACKAGES` / `DESCRIPTION OF GOODS`, and column edges are read from it. Each
  block has goods rows, then `CONSIGNEE [n]` : name, then `CTRN NO` : container with
  `G.W.: n KGS.` on the right, then `B/L NO` : B/L with `MEAS: n CBMS.` on the right. The
  `B/L NO` row ends the block. Totals are `TOTAL PKGS/G.W./MEAS` on the last page. Without this
  parser the generic fallback shifts every block by one B/L. Packaging `PLTS` maps to PALLET.
  Verified 2026-10-15: 14/14 B/Ls passed, and Penanshin output was unchanged.
  Its browser version lives at https://github.com/Rattanao/MAC-NELS (local folder
  `Downloads/check edi-MAC-NELS`, page at https://rattanao.github.io/MAC-NELS/). It is the
  Penanshin page with `parseEnter` swapped for a JS port of layout H (`isMacNelsSheet()`
  rejects other forms) and `PLTS` → PALLET in `canonPkg`. Parity check 2026-10-15: 420/420
  cells matched `build_edi.py`. Port any layout H change there too.

- **Layout I** (`parse_enter_layout_i()`): UNION WORLD SHIPPING "AMENDMENT" (added 2026-10-19,
  SAWASDEE MIMOSA 2610S job). It has the same `B/L CHANGE NO.` header as AGN layout D, but the
  columns sit elsewhere, so layout D misreads it. It is identified **per page** by three things:
  `B/L CHANGE NO.` at the far left, `HB/L NO` at x>650, and a `STATUS = ..` line in the HB/L
  column. The ASIA ATLANTIC page in the hahm ENTER has the first two but not the third.
  Column bands: B/L <100 | marks <255 | qty <340 | desc <485 | cons <610 | KGS/M3 <688 |
  HB/L ≥688.
  - The B/L is split across two lines: base `HASLC0126343042` with `8A` below it. Layout C and
    the fallback parser turn that base into a fake B/L, so any key that is a strict prefix of a
    layout-I B/L is dropped.
  - The HB/L column holds the HB/L number, then the containers (comma-separated, sometimes
    wrapped mid-number), then `STATUS = CY`. A Thai STATUS goes on the next line, and its
    garbled `ҟ` is mapped back to `้`.
  - The totals line is `TOTAL n PACKAGE(S) G.W n KGS. n M3`.
  - Container check: when ENTER lists more than one container for a B/L, the whole set must
    match MANIFEST.
  - MANIFEST STATUS lines like `LCL CONSIGNEE'S LABOUR` now count as STATUS, not DESC.

  Regression check: the Able, CMA, Cargoport, DSV, Gateway, MAC-NELS, Penanshin, Quality and
  hahm reports are identical to the previous script, cell for cell including fills.
  Browser version: repo https://github.com/Rattanao/Union-world (local folder
  `Downloads/check edi-union world`), page at https://rattanao.github.io/Union-world/ once Pages
  is enabled. See "Check EDI Union World" under "Web version" below. Port any layout I change
  there too.

- **Layout J** (`parse_enter_layout_j()`): LX PANTOS "AMENDMENT OF MANIFEST OF" (added 2026-10-19,
  SAWASDEE ALTAIR 2609S job). It has **one B/L per page**, and the customer usually sends **one PDF
  per B/L** (`ENTER <BL>.pdf`, `ENTER <BL>A.pdf`, …). The script takes one `--enter` file, so merge
  them into a single PDF first (pymupdf `insert_pdf`, sorted by name length then name) and pass that.
  It is gated on the text `LX PANTOS` + `AMENDMENT OF MANIFEST OF`. Per page: `B/L NO.` then
  `: <BL>`; goods rows under `MARK` are marks x<195 | `10 PACKAGES` 195–335 | desc ≥335; the `(W:`
  row has G.W. before `KGS./M :` and MEAS before `CBM.)`. The values for `CONSIGNEE` /
  `NOTIFY PARTY` / `CONTAINER NO` / `STATUS` are at x≥130. `SAME AS ABOVE` notify = consignee.
  STATUS reads like `***CY ลากตู้*** D/O รับแหลมฉบัง` (real Thai; it only looks garbled in the
  cp874 console). There is no document-level total. Without this parser the generic fallback
  missed the base B/L and read STATUS as `:`. Result: 7/7 B/Ls passed.

- **Layout K** (`parse_enter_layout_k()`): UNITEX / Planet "AMENDMENT" form `FM-CSI-01` (added
  2026-10-19, SAWASDEE MIMOSA 2610S job). It has the same `B/L CHANGE NO.` + `HB/L NO` header as
  layout I, but it's identified per page by `CNTR NO.:` sitting in the MARKS column. Column bands:
  B/L <100 | marks <255 | qty <340 | desc <485 | cons <600 | KGS/M3 <688 | HB/L ≥688.
  - The B/L is split across two lines (`HASLC0126900` / `9665A`), so the tail can be up to 6
    characters. Fake prefix keys from layout C or the fallback parser are dropped.
  - The MARKS column holds the marks, then `CNTR NO.:`, then `SKLU6020052/20'GP`, then `STATUS: CY`.
  - The last block stops 6pt above `TOTAL`, because `G.W n KGS.` sits a little above it and
    otherwise leaks into CONSIGNEE.
  - Result: 4/4 B/Ls passed. Layout K returns 0 B/Ls on the AGN, hahm, Union World and Cargoport
    ENTERs, and the Union World report was identical cell for cell.

- **Layout L** (`parse_enter_layout_l()`): SINOTRANS amendment letter (added 2026-10-27, SAWASDEE
  SPICA 2609S job). It has **one B/L per page**, with a blank page after each one. It is identified
  by `MB/L NO` + `ISSUED D/O AS FOLLOWING` + `CONTR NO`. Per page:
  - `MB/L NO. : <BL>`. On some pages the vessel's `V.2609S` wraps onto the front of this line, so
    the label is searched mid-line. One page printed `HHASLC…`; leading letters are stripped until
    the number matches `bl_re`.
  - `CNEE : <name>`.
  - Goods rows under `MARKS & NOS.`: marks x<165 | `6  PALLETS` 165–265 | desc ≥265, down to the
    `CONTR NO.` row.
  - Container `SKHU… / 40' HQ` sits on the CONTR NO. row. G.W. is `G.W. : n KGS`, MEAS is the
    `n CBM` item, and STATUS is `CY/CY`.
  - The FreeText `CS IMPORT` stamp at the foot is ignored.
  - Result: 26/26 B/Ls passed. Regression on 17 other jobs: 0 cell differences.
  - Browser version: repo https://github.com/Rattanao/Sinotrans (see "Check EDI Sinotrans" under
    "Web version" below). Port any layout L change there too.

- **Layout M** (`parse_enter_layout_m()`): FREIGHT LINKS EXPRESS "IMPORT CARGO MANIFEST" (added
  2026-10-27, SAWASDEE SPICA 2609S job). It is gated on `IMPORT CARGO MANIFEST` + `CUSTOMS B/L NO` +
  `MARKINGS & NOS`.
  - Page 1 is a cover letter (`OBL NO : <master B/L>`) with `[ X ]` checkboxes (CY / UNSTUFFING INTO
    GODOWN / …) and FreeText notes (`**เร่งเปิดตู้เข้าโกดัง`, `จ่ายเงินกทม. Do รับแหลม`). That page has no B/L
    anchor, so the general annotation scan drops them. Layout M reads the ticked label and every
    FreeText note longer than one character, and uses them as STATUS for every B/L.
  - The container comes from the page header `Container No : SKHU… - 40 FT`.
  - Column edges are read from the `NO OF PACKAGES` / `DESCRIPTION OF GOODS` header. Each block is the
    goods rows, then `CUSTOMS B/L NO <B/L>` / `CONSIGNEE :` name / `GROSS WEIGHT : n KG` +
    `MEASUREMENT : n M3`, then a `____` line.
  - In `2 BAGS (1 PALLET)`, the bracket part moves to the start of desc, as MANIFEST writes it.
  - Totals are `Total Packages : n` / `Total Weight Kg : n` / `Total Volume m3 : n`.
  - Without this parser the generic fallback read the container as `ASLC0126613` and merged marks with
    desc. Result: 7/7 B/Ls parsed, and the other jobs were identical cell for cell.
  - FreeText notes on the table pages go to the first `CUSTOMS B/L NO` row below them.
  - Browser version: repo https://github.com/Rattanao/Freight-links (see "Check EDI Freight Links"
    under "Web version" below). Port any layout M change there too.

- **Able (layout B variant) + MANIFEST as PDF** (added 2026-10-15). MANIFEST can now be a PDF
  (`MANIFEST*.pdf`, a landscape "CARGO MANIFEST" printout). `_pdf_manifest_rows()` turns it into the
  same row/column grid as the .xls (column chosen by each line's left edge on the rotated page). That
  PDF has no `1)…4)` rule lines, so a transit B/L's ประเทศปลายทาง / CARGO MOVEMENT shows **orange**
  "ไม่มีใน MANIFEST PDF" (check in the system), not red. The Able ENTER uses the layout B anchors
  `CONSIGNEE :` / `CONT NO.  :` / `B/L  NO.  :`, but with these differences:
  - The B/L suffix letter is a separate text item (`HASLK…293` + `A`). The parser joins them.
  - Units sit on their own lines (`3,700.000` / `KGS.` / `9.740` / `CBM`).
  - A `Notify Party :` section can appear between the consignee and CONT NO. It goes to `rec['notify']`.
  - The typed notes `" CARGO MOVEMENT = 7 "` and `"... CODE 7 "` sit at the end of the page's text
    stream. They are moved back to their real y position and become `movement_raw`.
  - Some MARKS lines contain a description span at x≈217. These are split into two lines.

  Moving lines back to their y position is limited to these MOVEMENT notes. Moving every line
  broke Gateway and hahm. Regression check: all 11 other jobs gave identical results to the
  previous script.

**`--only-enter`:** when the MANIFEST covers the whole vessel and the ENTER covers only one
customer (as with MAC-NELS, using the vessel MANIFEST from the hahm job), pass `--only-enter`.
It drops other customers' B/Ls and skips the vessel GRAND TOTAL comparison, so the TOTAL row
compares this consignment only.

Both layouts also read **FreeText PDF annotations** (text the customer typed on top
of the PDF, usually in red) and assign each one to whichever B/L block sits below it
on the page by y-coordinate — this is how hand-added STATUS corrections, DG CLASS/UN,
REEFER temperatures, and INTRANSIT/TRANSHIPMENT notes that aren't in the document's
original printed text get picked up at all. **Never skip these** — per the user's
SOP, annotated text is the customer's real instruction and overrides the printed
form. Watch for **range annotations**: some customers write one note covering a whole
span of sub-B/Ls instead of repeating it per block, e.g.
`HASLK01262506564-T    STATUS :เปิดตู้เข้าโกดัง` meaning the base B/L through its
`-T`-suffixed sibling (682, 682A, 682B, … 682T) all share that STATUS —
`parse_enter()` detects the `<base>-<letter> <instruction>` shape and expands it to
every sub-B/L in the range rather than misfiling the whole note onto whichever single
block happens to sit nearest by position.

If a new ENTER PDF matches neither known layout, dump `page.get_text()` and
`page.annots()` for a page or two first to see the actual anchors and annotation
positions before touching the parser — don't guess blind.

## Report conventions — apply these by default, don't ask

These were fixed through direct user feedback; treat them as settled defaults for
this report type, not something to re-decide each time.

- **Font:** Aptos Narrow, size 9, throughout the whole sheet.
- **Background:** white unless there's a specific reason to color a cell — color
  carries meaning here, so don't add it decoratively. Red (`FDD7D7`) = mismatch
  (prefix the cell text with `⚠`), orange (`FCE4D6`) = needs human confirmation,
  yellow (`FFF2CC`) = no ENTER doc to compare against for that B/L, green (`C6EFCE`)
  = pass (only used in the summary column), navy (`1F4E78`) fill + white bold = header.
- **Column order:** an urgent-summary column comes **first** (column A), then B/L
  NO., then CNEE.(ENTER)/CNEE.(MANIFEST), then **TAX ID vs NOTIFY** right after those
  two (user-requested position — it's a MANIFEST-internal check, not an ENTER/MANIFEST
  pair, so it doesn't fit the alternating pattern but belongs visually next to the
  consignee columns it's most related to), then the remaining ENTER/MANIFEST pairs for
  container, status, package count+type, weight, measurement, marks, description, then
  reefer/DG/transit, then SHED NO./ประเทศปลายทาง/CARGO MOVEMENT, then a detail-notes
  column **last**. See `HEAD` in the script for the exact list — don't reorder it
  without the user asking. The four check functions and all the per-B/L comparison
  logic in `build_report()` are written against a fixed *internal* column numbering
  (`mk()` calls, `cs[...]`, `COL_SHED`/`COL_DESTCODE`/`COL_MOVEMENT`/`COL_TAXID`/
  `COL_NOTE`) that does **not** match `HEAD`'s visual order — `_OLD_TO_NEW` /
  `_reorder()` / `_remap_cols()` translate from internal numbering to the real sheet
  position exactly once, right before each of the four places a row gets written (the
  per-B/L row, the "no ENTER" row, the "B/L missing from MANIFEST" row, and the TOTAL
  row). If the user asks to move a column again, change `HEAD` and `_OLD_TO_NEW` only —
  don't renumber the `mk()` calls scattered through the comparison logic.
- **CNEE. and NOTIFY should be the same company** (per the user's SOP) — TAX ID vs
  NOTIFY (`check_taxid()`) enforces this against MANIFEST's own N:/NOTIFY PARTY field
  by default. The one exception is a B/L with a TRANSIT/TRANSHIPMENT sentence: there,
  the correct notify party may be one ENTER states separately (a forwarder handling the
  in-transit leg) rather than MANIFEST's own N: — `compute_extra_checks()` passes
  ENTER's own `rec['notify']` (only layout D populates this field so far) as a second
  acceptable match, but **only** when the B/L has transit text; a non-transit B/L must
  still match MANIFEST's N: exactly.
- **STATUS canon:** CY, CY/CY, FCL, ลากตู้ → `CY`; LCL, ขน(ส่ง) → `LCL`; CFS,
  LCL/CFS, เปิดตู้ → `LCL/CFS`. Show the raw text plus `(= <canonical>)`.
  If the customer's ENTER document uses different colloquial terms for the same
  three buckets, extend `canon_status()` rather than leaving them unclassified.
- **PACKAGING:** reduce to the bare noun (CASE/PALLET/CARTON/BOX/PACKAGE/SET/…) —
  never leave a raw manifest code, a package count, or a `(s)` suffix in this column.
- **MARKS / DESCRIPTION:** compare strictly — normalize by stripping everything but
  letters and digits, then require exact equality (`same_detail()`). This tolerates
  only the line-wrapping/spacing differences that free-form MANIFESTs introduce (a
  stray mid-word space from PDF/Excel extraction never counts as a mismatch); any
  real wording difference (an extra/missing word or sentence) must still be flagged,
  even a small one. Compare MARKS(ENTER) vs MARKS(MANIFEST) and DESC(ENTER) vs
  DESC(MANIFEST) **as separate pairs**, not as one combined blob — flagging a column
  that is byte-identical to its counterpart just because the *other* column in the
  same row disagrees is confusing to the reviewer. For layout-B ENTER files (real
  x0-split MARKS/DESCRIPTION columns, see format-notes.md) flag orange, not red,
  since the MANIFEST's own split point comes from Excel line-wrapping and can
  legitimately land the same content on the other side of the line from ENTER's real
  column boundary — flag **both documents'** cells for whichever field actually
  differs, not just the MANIFEST side, so a reviewer scanning only the ENTER column
  still sees that something needs checking. **Never truncate or drop content** — the DESCRIPTION (MANIFEST) cell must
  read exactly like the source file's "NO. OF PKGS / DESCRIPTION OF GOODS" column,
  including any DG/REEFER/TRANSIT sentence embedded in it (those get *also* pulled
  into their own dedicated columns for quick scanning, but must stay in DESCRIPTION
  too — extracting them is additive, never destructive). The only content that's
  legitimately excluded from DESCRIPTION is pure structural noise that isn't
  "description of goods" in any reading and already has its own column elsewhere:
  the STATUS token and the container-type code (e.g. `45G0X1 PART CNTR`).
- **DG / REEFER TEMP / TRANSIT-TRANSHIPMENT:** flag red on any B/L where these show
  up **at all**, in either document — not just when the two documents disagree with
  each other. These always need customs follow-up regardless of agreement. The
  TRANSIT/TRANSHIPMENT **column** shows only the distilled phrase via
  `extract_transit_phrase()` — `"INTRANSIT TO <destination>"` or
  `"TRANSHIPMENT TO <destination>"` — never the raw sentence or an `ENTER:`/
  `MANIFEST:` source label; some MANIFESTs write a full travel narrative
  (`"IN TRANSIT CARGO FROM HOCHIMINH PORT, VIETNAM VIA LAEM CHABANG PORT BY SEA TO
  MYWADDY, MYANMAR BY TRUCK"`) and only the actual destination after the last `TO`
  (before a `VIA`/`BY`/`FROM` clause or the sentence end) belongs on-screen — the full
  sentence still belongs in DESCRIPTION per the rule above, just not repeated in this
  column. When neither document mentions transit/transhipment for a B/L, show `-`
  in this column (not `LOCAL`).
- **B/L in ENTER but missing from MANIFEST entirely:** its own critical red row
  (`⚠⚠ ... วิกฤต`), not just a console message — this is cargo that risks being left
  off the customs declaration altogether, distinct from (and worse than) a B/L that's
  in MANIFEST with no ENTER to compare against (which gets a yellow row instead).
- **TOTAL row:** sums every numeric column for both ENTER and MANIFEST (right-align
  those columns, comma-separated thousands), a packaging-kind breakdown, counts of
  STATUS/DG/REEFER/TRANSIT flags, and cross-checks against whatever total each
  source document states about itself (MANIFEST's own "GRAND TOTAL" footer, ENTER's
  own "PACKAGES/KGS/CBM" summary) — not just against the other document. Distinct
  containers count once each, not once per B/L sharing them.
- Row 2 = `Vessel <name> | Master B/L <number>` in plain black bold, one line — both
  values come from the parsed data, don't hardcode a shipment's vessel or B/L number.
  Row 3 = the color/legend key, bold red font.
- Row heights are fixed (not auto-fit) and header/data columns are sized to their
  own header text by default, except MARKS/DESC. which get extra width for full text.
- **Column alignment:** every data column is left-aligned, except G.W. KGM (ENTER),
  G.W. KGM (MANIFEST), MEAS. MTQ (ENTER), and MEAS. MTQ (MANIFEST), which are
  right-aligned. Header row (row 4) stays centered. Don't right-align TOTAL PKG/PKG.
  or any other column just because it holds a number.

If the user asks for a formatting tweak that isn't covered above (a column width, a
color, a reordering), make the change and treat it as durable for *this* report going
forward, the same way the conventions above were built up — update this SKILL.md's
list if the change seems likely to matter for the next job too, not just this one.

## Regression check across every forwarder — run after ANY change to build_edi.py

The user wants this skill to work for every forwarder's ENTER, so a fix for one form must not
break another. Before changing `build_edi.py`, copy it (`cp scripts/build_edi.py /tmp/base.py`),
make the change, then run:

```bash
python scripts/regress_all.py --base /tmp/base.py
```

It runs the old and new script on every `Downloads/check edi-*` job (input/ or job root; LX Pantos'
per-B/L PDFs are merged first; jobs without both files are listed as SKIP) and prints every cell
whose text or fill changed. Every difference must be explained by the intended change; anything
else is a regression to fix before handing results to the user.

Last full run 2026-10-26, 15 jobs (AGN-LAO, Able, CMA/CMS, Cargoport, DSV, Gateway, LX Pantos,
MAC-NELS, Penanshin, Planet, Quality, Shipco, THE M CORP, hahm, union world): no errors, identical
B/L and STATUS/DG/REEFER/TRANSIT counts; the only changes were the row-3 legend, the TOTAL label
wording, and the intended CMS fixes. Eastern Air has an ENTER only (checked by parsing: 10/10 B/Ls,
sums equal ENTER's own total). Macpro has no sample file.

**Rule changes must reach every website too.** Each forwarder site (rattanao.github.io/<name>:
Cargoport, check-edi-AGN/Able/DSV/Quality/Shipco, Gateway, LX-Pantos, M-CORP, MAC-NELS,
Penanshin, Planet, Union-world, CMS, Eastern-Air, Sinotrans, Freight-links) has its own `core.js` JS port, and the hub
**Check-MNF-Forwarder** carries a further copy of every tool under `tools/<name>/core.js`. A rule
or reading fix in `build_edi.py` is not done until it is ported to all of them (they share the same
function shapes: `checkMovement`, `extractTransitPhrase`, `parseManifest`'s `END` check,
`buildReport`'s TOTAL row, `index.html`'s chip and Excel legend), then each site's old and new
`core.js` are run on that forwarder's real job and every report cell compared, as for the Python
script. Done 2026-10-26 for the 3-TRANSHIPMENT rule: 28 copies, only intended differences.

TOTAL row label: when the three sums match but other flags exist, it reads
`⚠ ยอดรวมตรงกัน แต่มีจุดต้องตรวจ: TRANSIT n B/L, …` (still red), not "ยอดรวมไม่ตรง". "ยอดรวมไม่ตรง"
is only for a real PACKAGE/G.W./MEAS. mismatch.

MANIFEST block end: in the MARKS and DESCRIPTION columns only real footer text ends a B/L block
(`PORT TOTAL`, `GRAND TOTAL`, `TOTAL PACK/GROSSWEIGHT/MEASUREMENT`, `PortOfDischarge`). Marks like
`TOTAL 4 CTNS ONLY` used to end the block early and drop the container line below it.

## 4 extra fixed-rule checks (SHED NO. / ประเทศปลายทาง / CARGO MOVEMENT / TAX ID)

**Display rule (user, 2026-10-09):** in all 4 columns, a value that passes its locked rule
shows `-`, not the value itself. The actual MANIFEST value is shown only when it's wrong
(red) or needs a person to check it (orange). `-` therefore means either "correct" or "no rule
applies".

Added per the user's fixed customs rules (not a MANIFEST↔ENTER comparison — these check
the MANIFEST's own printed values against fixed reference tables). They live in their
own 4 columns at the end of the report (before the notes column) and are computed by
`compute_extra_checks()` / `check_shed()` / `check_dest_code()` / `check_movement()` /
`check_taxid()` in `build_edi.py`. All 4 values are printed on **one column (col 15) of
MANIFEST.xls**, as 4 consecutive lines per B/L block: `1) SHED NO. ####` /
`2) ประเทศปลายทาง (XX)` / `3) CARGO MOVEMENT (X)` / `4) TAX ID: (ชื่อ)` — parsed into
`rec['shed_no']` / `rec['dest_code_raw']` / `rec['movement_raw']` / `rec['tax_id_raw']`
in `parse_manifest()`.

- **1) SHED NO. — user-specified SHED (takes priority):** the user states the job's SHED
  number when asking for the check (e.g. "ตรวจสอบเอกสาร SHED NO. 0141"). Pass it as
  `--shed 0141`. Then **every** B/L must match that number: a match shows `-`, a mismatch
  is red. The one exception is LAOS, which keeps the old rule (0124 when discharge =
  BANGKOK). If the user didn't state a SHED, ask for it, or run without `--shed` to fall
  back to the 7-category rules below.
- **1) SHED NO. (fallback without --shed)** — only 7 categories have a required SHED NO.; everything else shows
  `-` (no rule). The category is decided by the **PORT OF DISCHARGE** field (tracked
  per-page in `parse_manifest()` as `rec['port_discharge']`) plus, for two of the
  categories, other MANIFEST content:
  - `USED ENGINE` keyword anywhere in DESCRIPTION, **and** discharge = BANGKOK → `0126`
  - destination country = LAOS (from the TRANSIT sentence), **and** discharge = BANGKOK
    → `0124`
  - DG detected on the B/L, **and** discharge = LAEM CHABANG → `2826`
  - discharge field itself reads `THLKR` → `0332` ; `BMT` → `0110` ; `SCT` → `0302` ;
    `UNITHAI` → `0113`. **No real example of these 4 has been seen yet** (confirmed by
    the user: they show up as the PORT OF DISCHARGE value itself, not a separate
    "Place of Delivery" field) — the code is ready for them but unverified against a
    real file. If one of these ever fires, double-check the match against the actual
    cell text before trusting it blindly.
- **2) ประเทศปลายทาง (XX)** — only checked when the B/L has a TRANSIT/TRANSSHIPMENT
  sentence with a recognized destination country (`DEST_CODE` table: LAOS→LA (user-confirmed 2026-10-12; was LL),
  MYANMAR→MM, CHINA→CN, MALAYSIA→MY, PHILIPPINES→PH, CAMBODIA→KH, MARSHALL ISLAND→MH,
  INDIA→IN). No transit sentence, or an unrecognized country → `-`, not checked.
- **3) CARGO MOVEMENT** — **TRANSHIPMENT always reads `(3-TRANSHIPMENT)`** (user rule,
  2026-10-26), to any destination including Laos, whether the TRANSHIPMENT wording is in
  MANIFEST or only in ENTER's transit text; this is checked before the INTRANSIT rules below.
  TRANSIT to anywhere except Laos must read `(7-TRANSIT)`.
  TRANSIT to Laos is the one exception: MANIFEST's code (`7-TRANSIT` or
  `L-INTRANSIT TO LAOS`) must match whatever ENTER's own `CARGO MOVEMENT <n>` line says
  for that same B/L (plain printed text on the ENTER PDF, not a FreeText annotation —
  parsed in `parse_enter()` into `ent_movement`/`rec['movement_raw']`, matched to the
  nearest B/L anchor **above** it — the opposite direction from FreeText annotations,
  since this text sits below its own block's B/L-number line, not above the next
  block's). If ENTER has no movement value for that B/L, flag **orange** ("ต้องเทียบ
  ENTER ด้วยคน"), not red — a real gap, not a confident mismatch. No transit sentence
  at all → `-` (STATUS `O-OTHER`/`5-CONSOLIDATE` cases are correct by definition).
- **4) TAX ID: (ชื่อ)** — must company-match (`comatch()`, same fuzzy logic as
  CONSIGNEE) against the B/L's own `N :`/NOTIFY PARTY name. Empty `TAX ID: ()` → `-`,
  not checked (common in real files).

## Web version: Check EDI Quality (GitHub Pages)

The Quality ENTER form (layout B — `MARKS & NOS | DESCRIPTIONS OF GOODS | GROSS WT.& MEAS`,
anchors `CONSIGNEE :` / `CONT.NO.:` / `B/L NO.:`) also exists as an in-browser page so
colleagues can use it without installing Python: repo `Rattanao/check-edi-Quality`,
live at https://rattanao.github.io/check-edi-Quality/ (Settings → Pages → Deploy from a
branch → `main` / root; the root `index.html` redirects to `docs/`).

- **Limitation (state it on the page and to the user): the web page supports only
  Quality's AMENDMENT form.** `isQualityAmendment()` rejects anything else (including the
  AGN form, which has `B/L CHANGE NO.`) with a message instead of guessing. `CNTRS*`
  files are ignored (not compared). For other customers' ENTER forms use the Python
  script / this skill, or build a sibling page the same way (`check-edi-AGN` is the
  template: `docs/index.html` + `docs/core.js`, same UI, only the ENTER parser differs).
- `docs/core.js` is a JS port of `build_edi.py` (parseManifest + report logic shared with
  the AGN page; `parseEnter` rewritten for the Quality form). pdf.js gives word-level
  items `{x,y,w,t}`, so the port groups items into rows, splits each row into segments at
  gaps > 20pt, and classifies by `rel = x − page label x` (the x of the `CONSIGNEE`/
  `CONT.NO.`/`B/L NO.` labels on that page — page margins differ between files, e.g.
  20.6pt vs 56.6pt, so never hard-code absolute x): `rel < 100` → MARKS, `rel ≥ 100` →
  DESCRIPTION/values, `rel ≥ 300` → KGS/CBM column. Classify by **segment**, not by single
  word — long MARKS lines legitimately contain words far to the right.
- **pdf.js splits words differently per file** (one Quality PDF gave `B/L NO.:` as one
  item, another gave `B/L` + `NO.:` as two). Detect labels on the merged **row segments**
  (`buildRows`), never on single items — a first version tested item-level and the live
  page wrongly said "not a Quality form" for a valid file. Keep `parseManifest` in sync
  with Python too (the container-type filter must stay `^\d{0,2}[A-Z]\d[A-Z0-9]{1,3}(?:\s|$)`,
  otherwise real DESC lines like `F413GGNAF03P` / `C5102P:CU95/SN5%` get dropped).
- Same report rules as the Python script: MARKS/DESC compared per pair with
  `sameDetail()` (spaces ignored), mismatch → orange on both documents' cells; TRANSIT
  column `-` when absent; consignee joined across wrapped lines; blocks may straddle a
  page break (rows are indexed globally across pages).
- **Verify any change to the port against Python:** run `build_edi.py`, run the JS port in
  node (`pdfjs-dist@3.11.174`, `xlsx@0.18.5`, same versions as the CDN tags), and diff
  every cell — the 2026-10-09 Quality job (9 B/L × 28 columns) matched with 0 differences.
- Never commit customer files: `.gitignore` excludes `input/`, `output/`, `*.xls*`, `*.pdf`.

**Check EDI Penanshin** is a sibling page built the same way: repo `Rattanao/Penanshin`
(local folder `Downloads/check edi-Penanshin`), live at https://rattanao.github.io/Penanshin/
once Pages is enabled.
- `docs/core.js` is the Quality port with its ENTER parser swapped for a JS version of
  layout F. Column edges come from the header items `No. of Pkgs` / `Description` /
  `Weight`. `isPenanshinManifest()` rejects any other form.
- It also includes the AGN fixes: Laos = `(LA)`, the 4 rule columns show `-` when correct,
  and `CARGO MOVEMENT n` is ignored when comparing MARKS/DESC.
- MARKS/DESC mismatches are **red**, matching Python layout F (not `combined_text`).
- 2026-10-13 parity check against `build_edi.py`: 224 of 224 cells matched, values and
  fill colours.
- Before publishing, replace real B/L numbers in the copied `build_edi.py` comments with
  made-up ones (e.g. `HASLK01260000123`), as the Quality repo did.
- On this PC, bash heredocs strip backslashes, so they corrupt regex code (`\b` became a
  backspace character). Make code edits with the Edit tool or a script file, never a heredoc.

**Check EDI DSV** is another sibling page: repo `Rattanao/check-edi-DSV` (local folder
`Downloads/check edi-DSV`), live at https://rattanao.github.io/check-edi-DSV/ (Pages: `main` /
root, enabled 2026-10-15).
- `docs/core.js` is the Penanshin port with its ENTER parser swapped for a JS version of
  layout G. `isDsvAmendment()` accepts a file containing `AMENDMENT OF MANIFEST` and
  `B/L CHANGED`. In a PDF that bundles several forms (e.g. the hahm job), only the DSV pages are
  read. FreeText annotations apply to the B/L on the same page.
- It also includes the MANIFEST type-code regex fix (`V266411` is no longer dropped).
- 2026-10-15 parity check against `build_edi.py`: 168 of 168 report cells matched on the DSV
  job, and 72 of 72 ENTER fields matched on the 6 DSV pages of the hahm job.
- Don't name scratch scripts after stdlib modules (a scratch `html.py` broke `import pymupdf`).

**Check EDI Able** is another sibling page: repo `Rattanao/check-edi-Able` (local folder
`Downloads/check edi-Able`, pushed 2026-10-15), page at https://rattanao.github.io/check-edi-Able/
once Pages is enabled (Settings → Pages → `main` / root).
- `docs/core.js` is the MAC-NELS port with two changes:
  - Its ENTER parser is replaced by a JS version of the Able layout-B variant, `isAbleForm()`.
    pdf.js returns single words already sorted by position, so the parser groups them into rows,
    then into segments (gap ≤ 20pt, and never across the MARKS|DESC line at 155pt right of the
    `CONSIGNEE` label). The late-typed MOVEMENT notes need no reordering here.
  - It also accepts a MANIFEST PDF: `isManifestPdf()` detects it, and `manifestRowsFromPdf()`
    uses the same column edges as `_pdf_manifest_rows()`.
- It includes the `no_rule_lines` orange review, the transit `-` fix and the type-code regex fix.
- 2026-10-15 parity check against `build_edi.py` on the Able job: 588 of 588 cells matched
  (values and fills). The page was also tested end to end in the browser, including the Excel
  export.
- On this PC, bash heredocs strip `\` from regexes. Edit `index.html`/`core.js` with the Edit tool.

**Check EDI Union World** is another sibling page: repo `Rattanao/Union-world` (local folder
`Downloads/check edi-union world`, pushed 2026-10-19), page at https://rattanao.github.io/Union-world/
once Pages is enabled (Settings → Pages → `main` / root).
- `docs/core.js` is the Able port. Its ENTER parser is replaced by a JS version of layout I
  (`isUnionForm()` / `isUnionPage()` use the same per-page check as Python), and the global is
  `window.CheckUnion`. `pageLines()` must **not** merge pdf.js items: the B/L number sits only ~5pt
  left of the MARKS column, and merging glued them together (0 B/Ls found). `pdfPages()` already
  joins touching pieces.
- pdf.js places the `G.W … KGS.` total line slightly **above** `TOTAL`, so the declared total is read
  from lines within ±12pt of `TOTAL`, sorted left to right.
- It also has the multi-container set compare, the `LCL CONSIGNEE'S LABOUR` STATUS fix, and the
  TOTAL-row note without a ⚠ prefix (matching Python). The Able, MAC-NELS, Penanshin, Quality and
  DSV pages still add ⚠ to that note, a cosmetic mismatch only visible when totals differ.
- **Header has the Heung A Line logo, same as the DO page** (https://rattanao.github.io/DO/): the logo
  is `docs/assets/hal-logo.png` (747×129, copied from the DO site) in
  `<span class="brand"><img …></span>` before the `<h1>`. `.brand` is a white rounded box (padding
  5px 10px, img height 30px), `header .wrap` uses `align-items: center`, and `.by` is bold 15px. It
  is also the favicon. Use the same header for any new sibling page.
- 2026-10-19 parity check against `build_edi.py`: 364 of 364 cells matched (values and fills),
  both without a SHED and with `--shed 0121`. The page was also tested end to end in the browser,
  including the Excel export.
- Real B/L and container numbers in the repo's `build_edi.py` comments were replaced with made-up
  ones (`HASLC0126000012`, `HALU1234567`, …). The skill's own copy keeps the real numbers.

**Check EDI Planet** is another sibling page: repo `Rattanao/Planet` (local folder
`Downloads/check edi-Planet`, pushed 2026-10-19), live at https://rattanao.github.io/Planet/
(Pages: `main` / root, enabled 2026-10-19).
- `docs/core.js` is the Union World port. Its ENTER parser is replaced by a JS version of layout K
  (`isPlanetForm()` / `isPlanetPage()` gate on `CNTR NO.:` in the MARKS band, the same per-page
  check as Python), and the global is `window.CheckPlanet`. The header uses the same Heung A Line
  logo as the DO page and Union World.
- 2026-10-19 parity check against `build_edi.py`: 140 of 140 cells matched (values and fills), both
  without a SHED and with `--shed 0122`. The page was also tested in the browser, including the Excel
  export. It rejects the Union World, AGN and hahm ENTERs.
- The repo's `build_edi.py` is the skill copy with real numbers scrubbed. A small mapping script
  replaced them with made-up ones, and its output was confirmed identical to the skill copy's.
- The preview server in `.claude/launch.json` must pass the same port in `runtimeArgs` as in `port`.
  If the app reassigns the port, the page fails to load (`chrome-error`). Pick a free port instead.

**Check EDI LX Pantos** is another sibling page: repo `Rattanao/LX-Pantos` (local folder
`Downloads/check edi-LX Pantos`, pushed 2026-10-19), live at https://rattanao.github.io/LX-Pantos/
(Pages: `main` / root, enabled 2026-10-19). It accepts **only the LX Pantos ENTER form** (user-confirmed).
- `docs/core.js` is the Union World port. Its ENTER parser is replaced by a JS version of layout J
  (`isLxForm()` / `isLxPage()`: `LX PANTOS` + `AMENDMENT OF MANIFEST OF` + a left `B/L NO.` + `MARK` + a
  `(W:` row on the same page), and the global is `window.CheckLx`. The page accepts **many ENTER PDFs
  at once** (one per B/L), so the user never has to merge them. pdf.js splits `PKG :10(GT)` into
  `PKG` + `:10(GT)` and the STATUS into 3 pieces; `rowOrder()` puts pieces back into rows (|dy|<3) and
  then left to right. A FreeText note applies to the B/L of its own page. The header uses the same
  Heung A Line logo as the DO page.
- 2026-10-19 parity check against `build_edi.py`: 224 of 224 cells matched (values and fills), without
  a SHED, with `--shed 2813` and with `--shed 0141`. The page was tested in the browser with the 8 real
  files, including the Excel export. It rejects the Union World ENTER.
- The repo's `build_edi.py`, `core.js` comments and README example use made-up B/L, container and
  consignee values.

**Check EDI Eastern Air** is another sibling page: repo `Rattanao/Eastern-Air` (local folder
`Downloads/check edi-Eastern Air`, which is also the job folder; pushed 2026-10-26), page at
https://rattanao.github.io/Eastern-Air/ (Pages: `main` / root, enabled 2026-10-26).
- EASTERN AIR LOGISTICS uses the CMS "INBOUND CARGO MANIFEST" form (layout F), so the site is a copy
  of the CMS repo (`/tmp/claude/s/CMS`) renamed: global `window.CheckEastern` (the gate is still
  `isCmsInbound()`, so it also accepts CMS files), title/h1 `CHECK EDI EASTERN AIR`, download
  `EDI Eastern Air.xlsx`, and the same Heung A Line logo header (`docs/hal-logo.png`).
- `scripts/build_edi.py`, `SKILL.md` and `references/format-notes.md` in the repo are the **current
  skill copies** (not the older CMS repo copy), with real B/L and container numbers replaced by a
  mapping script. The scrubbed script gave a report identical to the skill copy's (0 cell diffs).
- 2026-10-26 parity check against `build_edi.py` on the SAWASDEE VEGA 2613S job: 308 of 308 cells
  matched (values and fills), both without a SHED and with `--shed 2835`. The page was tested in the
  browser, including the Excel export.
- Any layout F change must now be ported to CMS, M-CORP, Penanshin **and** Eastern-Air `core.js`.
- `.gitignore` excludes `input/*`, `output/`, `EDI.xlsx` and `.claude/` (local preview config).

**Check EDI Sinotrans** is another sibling page: repo `Rattanao/Sinotrans` (local folder
`Downloads/check edi-Sinotrans`, which is also the job folder; pushed 2026-10-27), page at
https://rattanao.github.io/Sinotrans/ once Pages is enabled (Settings → Pages → `main` / root).
- `docs/core.js` is the LX Pantos port (taken after its 2026-10-26 TRANSHIPMENT commit). Its ENTER
  parser is replaced by a JS version of layout L. `isSinoForm()` / `isSinoPage()` look for
  `MB/L NO` + `ISSUED D/O AS FOLLOWING` + `MARKS & NOS.` + `CONTR NO.`, and the global is
  `window.CheckSino`. The header uses the same Heung A Line logo as the DO page
  (`docs/assets/hal-logo.png`).
- pdf.js splits this form into single words (`MARKS` `&` `NOS.`, `4` `PALLETS`, `G.W.` `:` `670.00`
  `KGS`), so the parser groups items into rows (|dy|<3) and reads row text. A value and its unit
  can come as one item (`62.50 KGS`, `0.243 CBM`) or as two, so both are accepted.
- FreeText notes apply only above the `CONTR NO.` row. The `CS IMPORT` stamp at the foot is ignored.
- 2026-10-27 parity check against `build_edi.py`: 756 of 756 cells matched (values and fills),
  without a SHED, with `--shed 2813` and with `--shed 0141`. The page rejects the ENTERs of all 18
  other forwarder jobs. It was tested in the browser, including the Excel export.
- The repo's `build_edi.py` is the skill copy, with real B/L and container numbers (and two customer
  names in comments) replaced by made-up ones. It gave a report identical to the skill copy's.
- The page is **not** in the Check-MNF-Forwarder hub yet.

**Check EDI Freight Links** is another sibling page: repo `Rattanao/Freight-links` (local folder
`Downloads/check edi-Freight links`, which is also the job folder; pushed 2026-10-27), page at
https://rattanao.github.io/Freight-links/ once Pages is enabled (Settings → Pages → `main` / root).
- `docs/core.js` is the CMS port (taken after its 2026-10-26 TRANSHIPMENT commit). Its ENTER parser
  is replaced by a JS version of layout M. `isFreightLinks()` looks for `IMPORT CARGO MANIFEST` +
  `CUSTOMS B/L NO` + `MARKINGS & NOS`, and the global is `window.CheckFreightLinks`. The header uses
  the same Heung A Line logo (`docs/assets/hal-logo.png`). Download name: `EDI Freight Links.xlsx`.
- The parser groups pdf.js items into rows (|dy|<3) and reads row text. `pdfPages()` also adds
  FreeText text as items, so those are filtered out before the rows are built. The ticked checkbox
  on the cover letter is the FreeText `X` (its y + 10 matches the label row's baseline).
- FreeText notes on the table pages go to the first `CUSTOMS B/L NO` row below them, in both Python
  and JS.
- 2026-10-27 parity check against `build_edi.py`: 224 of 224 cells matched (values and fills),
  without a SHED and with `--shed 2813`. The page rejects the ENTERs of all 20 other forwarder jobs.
  It was tested in the browser, including the Excel export.
- The repo's `build_edi.py` is the skill copy, with real B/L and container numbers replaced by
  made-up ones. Across all 19 jobs, its reports were identical to the skill copy's.
- The page is **not** in the Check-MNF-Forwarder hub yet.

## Reference

- `scripts/build_edi.py` — the full implementation; read its docstring and inline
  comments before modifying anything, they document every column-index assumption.
- `references/format-notes.md` — worked examples of the exact MANIFEST row shape and
  **both** known ENTER layouts (A and B) this parser expects, plus the B/L-range
  annotation pattern, taken from real validated jobs across different customers.
