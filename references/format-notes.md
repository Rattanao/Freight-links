# Worked example: MANIFEST.xls and ENTER.pdf row shapes

Taken from two real validated jobs — a Heung-A Line consolidation, Master B/L
HASLK01262506564 (21 sub-B/Ls, vessel HEUNG-A HOCHIMINH 2610S), and a second Heung-A
job, Master B/L HASLK01262946186 (19+ sub-B/Ls, vessel SAWASDEE SUNRISE 2610S). The
MANIFEST shape below is common to both; the two jobs' ENTER PDFs use different block
layouts (A and B respectively), both handled by `parse_enter()`. Use this to sanity
check a new file against what `parse_manifest()` / `parse_enter()` expect, and to see
where a different carrier or customer format is likely to diverge.

## MANIFEST.xls (read with `xlrd`, sheet 0, `ragged_rows=True`)

One block per B/L. The header row of each block carries the B/L number in column 0;
subsequent rows carry shipper/consignee/notify, marks, container, package/status/
description, and weight/measurement, until the next B/L header row or an END marker
(`PORT TOTAL`, `GRAND TOTAL`, `PortOfDischarge`).

```
col0                          col5                              col10                         col16
HASLK01262506564              SAMYANG VALVE MADE IN KOREA BOX    5 PX (PALLET(s))              0.000 KGM
S : HEUNG A LINE CO.,LTD.                                        LCL/CFS                       0.170 MTQ
C : TR CO.,LTD                1. SKHU1204423                     45G0X1 PART CNTR
N : TR GOAL SUPPLY CO.,LTD                                       (5 WOODEN CASE) SAMYANG VALVE FACTORY
```

- Column 0: `S :` / `C :` / `N :` prefix = Shipper / Consignee / Notify Party. A
  name that wraps onto a second physical row (no prefix, not starting with a
  container-list "1."/"2." marker) gets appended to whichever of consignee/notify
  was most recently seen.
- Column 5: marks text on most rows; a line matching `^\d+\.\s*[A-Z]{4}\d{6,7}` is a
  container number instead (`1. SKHU1204423` → `SKHU1204423`), not marks text.
- Column 10 on the header row = package count + packaging code, e.g. `5 PX
  (PALLET(s))` — `PX` is the carrier's short code, `PALLET(s)` is the human-readable
  form; both mean packaging = PALLET after `canon_pkg()`. On later rows it holds
  STATUS (`LCL/CFS`, `CY`, …), a container-type code to ignore (`45G0X1 PART CNTR`),
  or free-text description — and, when present, DG (`CLASS 9 UN 3077`), reefer
  temperature (`TEMP: -15'C`), or a transit/transhipment sentence. These three get
  **also** copied into their own dedicated fields (`rec['dg']`, `rec['reefer']`,
  `rec['transit']`) for quick scanning/flagging, but must stay appended to
  `rec['desc']` too — DESCRIPTION (MANIFEST) has to read exactly like the source
  column, word for word. Only STATUS tokens and container-type codes are true
  structural noise and get dropped from description entirely (they already have
  their own columns and carry no shipment-content information).
  A transit/transhipment sentence often wraps across several consecutive column-10
  rows before the block ends (e.g. `"IN TRANSIT CARGO FROM HOCHIMINH PORT,` /
  `VIETNAM VIA LAEM CHABANG PORT BY SEA TO` / `MYWADDY, MYANMAR BY TRUCK"` across 3
  rows) — once the transit keyword is seen, every remaining column-10 line in that
  block is treated as a continuation of the same sentence, not a fresh description.
- Column 16: gross weight on the header row (`... KGM`), measurement (`... MTQ`) on
  the next row down.
- The declared-totals footer looks like (note the section label sits in **column 1**,
  not column 0 — a common trap):
  ```
  col1              col6                    col7
  GRAND TOTAL:      TOTAL MEASUREMENT:      91.255 MTQ
                     TOTAL GROSSWEIGHT:      36,190.150 KGM
                     TOTAL PACK:             101 PACKAGES
  ```

## ENTER.pdf — layout A (separate MARKS / DESCRIPTION columns)

One page typically holds two B/L blocks (top half / bottom half), each shaped like:

```
<marks text, one or more lines>
<package count>
<description text>
<packaging word: PALLET / CARTON / CASE / BOX / SET / PACKAGE / WOODEN CASE / …>
G.W.: 1360.000 KGS.
MEAS.: 4.170 CBM.
SKHU1204423-40
CNTR :
HASLK01262506564
B/L NO. :
TR GOAL SUPPLY CO.,LTD
```

- The block is parsed **bottom-up** from the literal anchor line `B/L NO. :` (space
  before the colon) — the line above it is the B/L number, the line after it is the
  consignee name typed into a box on the form.
- `CNTR :` sits right above the B/L number; the line above `CNTR :` holds the
  container number plus a `-40`/`-20` suffix (container size) that gets stripped.
- `parse_enter()` tries this layout first (by searching for the exact `B/L NO. :`
  anchor); if none are found it falls back to layout B below.

## ENTER.pdf — layout B (marks and description in a different block shape)

Seen on a different customer's ENTER PDF for the same carrier. Anchor is
`B/L NO.:` (**no** space before the colon — this is the detail that tells layout A
and layout B apart) instead of layout A's `B/L NO. :`. One page holds several
blocks shaped like:

```
MARKS & NOS          DESCRIPTIONS OF GOODS         GROSS WT.& MEAS      (header, once per page)
                     8 METAL BOXES                 11,136.000 KGS
                                                    16.486 CBM
VITTORIA TYRES-BKK   SOLUTION STYRENE BUTADIENE RUBBER (SSBR)
                     SSBR F1038
CONSIGNEE : 
VITTORIA TYRES (THAILAND) CO., LTD.
CONT.NO.:
HLHU2780542/40'HC
B/L NO.:
HASLK01262946186
```

- **This layout still has real, separate MARKS & NOS / DESCRIPTIONS OF GOODS
  columns** — it is not one combined free-text paragraph. Don't rely on line order
  to tell them apart (a "first short line = marks, rest = description" heuristic
  gets it backwards on plenty of blocks); instead read each line's **x0 bbox**
  (`page.get_text('dict')`, not the plain `get_text()` string) — the left column
  (MARKS, same x0 as the `CONSIGNEE :`/`CONT.NO.:`/`B/L NO.:` labels, ~20–25pt from
  the page edge) vs. the right column (DESCRIPTION, ~185pt+). `parse_enter()`
  collects `(line_text, x0)` pairs for exactly this reason — don't strip the x0 back
  out when touching this code.
- Still set `combined_text: True` and compare MARKS(ENTER) vs MARKS(MANIFEST) and
  DESC(ENTER) vs DESC(MANIFEST) **separately** (each with `same_detail()`, not
  `same_bag()` — letters/digits only, ignoring all whitespace, so a stray mid-word
  space from PDF extraction like `"C O., LTD"` doesn't false-flag) and flag orange
  (not red) only the pair that actually differs — the MANIFEST's own split point
  between MARKS and DESC comes from Excel line-wrap position, not real column
  geometry, so it can legitimately disagree with ENTER's split even when the total
  content matches. **Never flag one side of a pair (say, MARKS(ENTER)) unless the
  matching field on the other document differs too** — flagging a column whose
  own text is byte-identical to its counterpart just because a *different* column
  in the same row disagrees is confusing to the reviewer ("why is this marked wrong,
  it says the same thing?"). And flag **both documents' cells for that field**, not
  just the MANIFEST side — a reviewer scanning only the ENTER column should still be
  able to tell something needs checking.
- `CONSIGNEE :` values can wrap onto a **second physical line** in the PDF (long
  legal-entity suffixes like `"PUBLIC COMPANY LIMITED."` land on their own line) —
  join every line between the `CONSIGNEE :` label and the next `CONT.NO.:` label,
  not just the first one, or the consignee name silently loses its second half.
- The pkg-count line can read `"1 PACKAGES"`, `"1 PKGS"`, `"1 PLT"`/`"6 PLT"`,
  `"39 CARTONS"`, `"3 BOXES"`, etc. — normalize `PLTS?`→`PALLET` and `PKGS?`→`PACKAGE`
  before passing to `canon_pkg()`.
- This file's own declared totals showed up as a single `GRAND TOTAL` line with all
  three numbers *before* their own labels, in one run: `GRAND TOTAL 102 PACKAGES
  34,949.260 KGS (2x40'HC) 87.870 CBM` — handled by trying this pattern first, before
  falling back to the layout-A-style patterns below.

## ENTER.pdf — layout C (Cargoport "AMENDMENT" letter, real table, no B/L-NO anchor)

Seen from forwarder Cargoport (Thailand) Ltd. (form code `FM-IMP-02`), Master B/L
HASLS21266308271, vessel HEUNG-A HOCHIMINH V. 2610S — an LCL console shipment with
8 sub-B/Ls. This is a genuine ruled table repeated across pages, headed once per page
by `D/O NO. | CONSIGNEE | MARKS & NUMBERS | QUANTITY | DESCRIPTION | GROSS WEIGHT |
HB/L NO`, but **the header label x-positions do not line up with their own column's
data x-positions** (a template quirk — do not try to match data to the nearest header
label's x0; match by content shape instead, see below). Neither `B/L NO. :` nor
`B/L NO.:` appears anywhere, so layouts A and B both miss every block; the generic
fallback also missed 8 of 9 sub-B/Ls here (see below), so this needed a dedicated
`parse_enter_layout_c()`.

```
y≈278  x≈47  HASLS21266308271        x≈143 LEATHER CHEMICALS         x≈277 1 PALLET(S)   x≈363 HM-3116
                                            MADE IN SPAIN                                        MILANO ABRASIVE WAX
                                                                                                   ATENAS POLISHING WAX
                                     x≈505 CK SHOES (THAILAND)        x≈629 G.W.200.500 KGS.  x≈666 1.000 M3   x≈709 CS146003001
                                            CO.,LTD.
y≈333  x≈47  HASLS21266308271
y≈345  x≈47  A                       x≈143 PROJECT NAME : ...         x≈277 1 CASE(S)        x≈363 ELEVATOR SPARE PARTS ...
```

- **The HB/L NO column is the leftmost column of the table (x0 ≈ 40–95)**, found by
  content, not by the "HB/L NO" header label's own x0 (which sits at the far right of
  the header row — a mismatch that's easy to trust wrongly). Every block re-prints the
  **full base B/L number** here; a sub-B/L's single-letter suffix (`A`, `B`, …) prints
  as its own separate text line directly below the base number, same x-column. A block
  with no suffix (the bare master B/L) has no letter line following it — the very next
  thing in that column is simply the next block's base number. Detect blocks by
  scanning this column top-to-bottom: a `bl_re` match starts a block; if the *next*
  entry in the same column is a lone `[A-Z]` letter that (base+letter) still satisfies
  `bl_re`, consume it as this block's suffix, otherwise the block has no suffix.
- The other six fields are separated by clean, consistent x0 bands (with margin for
  jitter across rows): MARKS & NUMBERS ~95–240, QUANTITY ~240–320, DESCRIPTION
  ~320–470, CONSIGNEE ~470–590, GROSS WEIGHT ~590–650, MEASUREMENT ~650–690, D/O NO.
  ≥690. QUANTITY is a single `"<count> <TYPE>(S)"` line (e.g. `1 PALLET(S)`,
  `29 CARTON(S)`) — reuse the same `^([\d,]+)\s+([A-Za-z].*)$` pattern as layout B.
- **No per-block CONTAINER column** — this is one LCL console container for the whole
  shipment, stated exactly once as `CONT NO.` in the document header (before the
  table), e.g. `CONT NO.` ... `BMOU8864248(40DCH)   NO DG` on the same visual row.
  Read that row by y-proximity to the `CONT NO.` label and apply the container number
  to **every** sub-B/L record — there is nothing else to compare per row.
- **STATUS is usually stated once, for the whole document**, on the `TOTAL` summary
  row at the end of the table (e.g. `STATUS : CFS/CFE  <Thai clause>`), not repeated
  per block and not a FreeText annotation. Apply it as every block's `status_raw`
  unless a block already got a more specific one from an actual annotation. The Thai
  clause after `STATUS :` may come through as unrecoverable replacement-character
  garbage if the PDF embeds a symbol font with a broken cmap — that's fine, the ASCII
  token before it (`CFS/CFE`) is what `canon_status()` needs and survives intact.
- The `TOTAL` row itself must be **excluded from the last block's content** (there is
  no next B/L anchor to bound it otherwise, so without an explicit cut the whole
  `TOTAL <n> PACKAGE(S) ... G.W <n> KGS. <n> M3.` line plus the STATUS/signature block
  below it would silently get appended into the last sub-B/L's fields). Stop
  collecting table rows the moment a line's stripped text is exactly `TOTAL`, and read
  the declared totals from what follows instead — note this format's own total line
  uses `PACKAGE(S)` / `G.W ... KGS.` / `... M3.` wording, not the `PACKAGES`/`CBM`
  wording the other two layouts' totals use, so it needs its own regex.
- MARKS and DESCRIPTION are genuinely separate columns here (same as layout A) and
  compared strictly (`same_detail()`, red on mismatch) — not treated as `combined_text`
  — because the two documents' content lined up character-for-character once split by
  the correct x-bands in a real validated job (marks metadata blocks like `PROJECT
  NAME : ... PURCHASE ORDER : 4551221020` matched MANIFEST's column-5 continuation
  lines exactly). If a future Cargoport file's split disagrees with MANIFEST's own
  wrap point, downgrade this to orange/`combined_text` for that file rather than
  assuming the mismatch is real.
- Why the **generic fallback** (`parse_enter_generic`) isn't enough for this layout:
  it finds B/L numbers by scanning every line for a `bl_re` match, but here the suffix
  letter is a separate text line from the base number, so it only ever recovers the
  bare base B/L (all 8 suffixed blocks collapse onto that one key and get overwritten
  by whichever block's fields were parsed last). Any new layout where a sub-B/L's
  identifier is split across more than one text line needs this kind of dedicated,
  content-aware column parser — the generic fallback can't reconstruct a key that's
  never printed as one contiguous string.

## Both layouts: annotations and declared totals

- Customer-typed **annotations** (PDF FreeText objects, usually red, large font)
  float independently of the printed text and can say almost anything: a STATUS
  override (`STATUS CY`, `เปิดตู้เข้าโกดัง`, `ลากตู้`, `ขนส่งแรงงาน-เรือ`), a DG
  class/UN (`CLASS 9 UN 3077`), a reefer temp (`TEMP: -15'C`), or a transit/
  transhipment note (`INTRANSIT TO LAOS`, `TRANSHIPMENT TO MYANMAR`). These do
  **not** appear in `page.get_text()` reading order anywhere near their visual
  position — pymupdf's plain-text extraction dumps annotation text at the very end
  of the page's text, after the last printed block. That's why the parser reads
  `page.annots()` directly and assigns each one to a B/L by **y-coordinate**
  (whichever B/L anchor sits below the annotation on the page), not by text order.
- **Range annotations**: one FreeText note can cover several sub-B/Ls at once instead
  of repeating per block, e.g. `HASLK01262506564-T    STATUS :เปิดตู้เข้าโกดัง`
  positioned at the very top of the page (above every real B/L anchor on it) — this
  means "B/L 682 through 682T (682, 682A, 682B, … 682T) all have this STATUS", not
  "the nearest block has this exact text". `parse_enter()` matches the
  `^([A-Z0-9]{6,}?)-([A-Z])\s+(.+)$` shape and applies the instruction (`rest`, the
  part after the range prefix) to every letter in `A..<end_letter>` plus the bare
  base. It can also be written **the other way round** — instruction first, then the
  `<base>-<letter>` marker at the very end, e.g.
  `เปิดตู้เข้าโกดัง HASLK01262946186-S` (meaning B/L 400 through 400S) — match this
  shape too (`^(.+?)\s+([A-Z0-9]{6,}?)-([A-Z])$`), and confirm the trailing token
  against `bl_re` before treating it as a range marker rather than assuming direction
  from position alone.
  **Important**: only the instruction portion (`rest`) is added to the raw-text
  filter list used to strip annotation debris out of the block-parsing token stream —
  the full original text (which contains a real B/L number, at the start *or* the
  end depending on which way round it was written) must *not* be used for that
  filter, or it will swallow the legitimate bare-B/L-number line for the range's base
  B/L as a false-positive substring match. As a second line of defense, `parse_enter()`
  never lets a candidate line that itself matches `bl_re` be dropped by the
  annotation-debris filter, no matter what any annotation's text contains — a real
  B/L-number line is a block-boundary anchor and losing even one silently drops that
  entire block (and corrupts the next block's parsing, since the block slice
  boundaries shift by one line). This filter is also scoped **per page** — text is
  only stripped from the same page the annotation sits on, never document-wide —
  because a short word that's part of one page's annotation sentence (e.g. a country
  name like `"THAILAND"` inside a `"...VIA BANGKOK,THAILAND"` transit note) can
  coincidentally also be a legitimate printed line for an unrelated B/L on a
  different page, and a document-wide filter would wrongly delete it there too.
- The document's own stated totals usually sit on the last page, but watch the
  label/value order — it isn't always "label then value":
  ```
  PACKAGES :
  101.00
  KGS :
  36,190.15
  91.255
  CBM :
  ```
  Here PACKAGES and KGS are label-then-value, but the CBM figure (91.255) appears
  **before** its own "CBM :" label. `parse_enter()`'s declared-totals regex handles
  this specific ordering; if a new ENTER format puts all three in a different order
  again, adjust that regex rather than assuming the simple per-field version will
  catch it.

## Known limits worth telling the user about

- The B/L-prefix auto-detection (`detect_bl_prefix`) picks the most common
  leading-letters run in column 0 that's followed by digits — works well for a
  single-carrier file, but a MANIFEST mixing multiple carriers' B/L prefixes in one
  sheet will only pick up the majority one. Pass `--manifest` unaffected; if this
  matters, split the file per carrier first.
- Container-grouping annotations drawn as boxes/circles (vector shapes, not
  FreeText) on the ENTER PDF are **not** detected — only typed FreeText annotations
  are. If a customer relies on drawn shapes to group containers, that grouping won't
  surface in the report; mention it to the user rather than silently missing it.
