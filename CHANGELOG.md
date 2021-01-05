This changelog includes changes across our components.  Not all changes affect
all builds.  Some builds are immediately pushed to address customer issues in
specific builds, so listed version numbers may not be available for this module.

1.20201120.1

- more style features in `apply_style_delta` 


1.20201106.1

- extract custom table style metadata
- first iteration of `apply_style_delta` and `get_computed_style` 
- (edit) append new style metadata during `book_append_sheet`

1.20201022.1

- FODS skip master styles text (GH2150)
- empty invalid cell (GH2136)

1.20201004.1

- documented background `fgColor` / `bgColor` / `patternType`
- ODS blank sheet creates a workbook with range A1:A1
- XLSX formula newline character normalized to `\n`
- (chart) cleaned up `axes` types

1.20200908.1

- CSV / HTML / DOM parsing now set formats for negative numbers and percentages
- HTML inserts newline character in forced breaks (`<p>...</p><p>...</p>` in TD)
- (image) "base64" / "binary" types now specify `string` instead of `"string"`

1.20200828.1

- (ssf) updated to 1.20200828.1 (zh-CN zh-TW ja-JP ko-KR th-TH localized tables)

1.20200821.1

- ignore gradients in differential styles

1.20200817.1

- (image) escape xml entities in shapes and chart text boxes

1.20200811.1

- Comments workaround properly handles hidden rows and columns

1.20200808.1

- Type correction for data validation error / input message suppression
- `cell_set_hyperlink` now sets text color to theme "hyperlink" color
- Comments workaround for bug in Excel for Windows

1.20200807.1

- XLML create stub cells to force border generation

1.20200801.1

- (image) Shape rich text properties
- International "This Workbook" codename support

1.20200728.1

- Comment `hidden` attribute correctly parsed from VML
- Custom XML types and documentation

1.20200720.1

- Default row height and column width through `!sheetFormat`
- `dist` folder now includes types

1.20200715.1

- updated type definitions to include PPI parameter
- multi-range conditional format bug fix
- cell selection sets active pane based on freeze settings

1.20200713.1

- (edit) properly update cells which happen to be table headers
- (chart) XLS shapes

1.20200710.1

- Read/write currently-selected cell for each worksheet

1.20200709.1

- (crypto) XLSX Try default password "VelvetSweatshop" if unspecified
- (edit) Update single-cell array formulae

1.20200707.2

- XLSX DXF properly parse mapped themes and styles

1.20200707.1

- XLSX shared formula shift row ranges like `2:4`
- (crypto) XLS RC4 correctly rotate key after buffer ending on boundary

1.20200705.1

- XLSB formula parsing fixes

1.20200703.1

- x14 Conditional Formatting support
- HTML support IE border specifications

1.20200702.1

- Types include Column Styles and Number Format
- HTML TD background style used when inner SPANs are transparent

1.20200626.1

- `decode_range` for single cell ensures s and e are separate objects GH2026
- (edit) `template_set_aoa` supports cell objects

1.20200623.1

- HTML Export ignore empty prefix columns

1.20200622.1

- HTML Export forced width (always write a table width)
- HTML Export Default width when sheet has no default width and no custom width
- HTML Export text is wrapped when neighboring cell has content
- HTML Export better treatment of borders in merge cells
- Types include `xlfn` (read/readFile) and `gridcolor` (html export)
- (pivot) Type definitions added

1.20200617.1

- x14 Data Validations read and write multiple block ranges

1.20200604.1

- Disabled `PRN` parsing by default (aligned with OSS 0.16.2)
- Correctly distinguish between styles with omitted `sz`

1.20200526.1

- (crypto) XLS correctly handle large continued records
- (ssf) updated to 1.20200527.1 (zh-CN format code 31)

1.20200525.1

- XLSX correct for LibreOffice invalid Data Validations
- Improved performance of style population

1.20200521.1

- DOM parse `origin` option correctly handles merge cells

1.20200520.1

- XLSX Conditional formatting multiple region rule parsed as one rule
- Column width `auto` skips merge cells when calculating widths

1.20200519.1

- XLSX Conditional formatting "text" rule now writes formula data
- HTML Export: `overflow: hidden` and table styles have `width` set
- SSF update to 1.20200519.1 (fixed IE11 regression)

1.20200514.1

- (crypto) standalone build proper treatment of byte arrays
- CSV handle `\r\n` with separators (see GH1943)
- External workbook properly parse sheet names with encoded characters

1.20200510.1

- Diagonal Border differential style XML fixed
- Clarify how to set "Mark as Final" custom property
- (crypto) Try default password "VelvetSweatshop" if unspecified

1.20200507.1

- Updated SSF (better treatment of parens in negative values and phone numbers)

1.20200506.1

- `sheet_set_range_style` support dense mode

1.20200505.1

- `set_date_style` to control date processing style
- XLSB external 3d missing ranges now return `#REF!` when missing data
- (edit) use column and row styles when applying styles to missing cells

1.20200427.1

- (crypto) reworked package.json so bundlers play nice

1.20200425.1

- Updated SSF (more precise support for over 11 decimal places)

1.20200424.1

- Integration with NetSuite SuiteScript -- proper global detection
- XLSX Conditional formatting "blank" rule now writes formula data
- XLSX read `nodim` option to skip dimension

1.20200417.1

- regex vulnerability cleanup (mirrors fix in 0.16.0 in open source)
- (pivot) control row and column header captions
- (pivot) fixed address calculation when using multiple pivot filters

1.20200415.1

- XLML styled write
- XLML row heights use pt instead of px
- (pivot) Replicate Excel bug in rounding dates down to seconds
- (pivot) Column sort order and collapse

1.20200410.1

- XLSX/XLSB/XLS/XLSB read/write "shrink to fit" alignment property
- `xlfn` read / readFile option for raw formulae
- HTML writer applies row hidden property to TD element

1.20200406.1

- XLSX skip parsing empty extended props, omit blanks on write

1.20200401.1

- @sheet/ssf updated (more intl support)

1.20200329.1

- AMD change `define` call to better support `require.js`
- (chart): more type definitions

1.20200323.1

- `sheet_add_json` safeguards

1.20200319.1

- XLSX custom properties properly escape double quotes in values

1.20200318.1

- XLML support improper lowercase tags
- XLML embedded HTML run support
- XLML style interpret font tag with empty name as Arial
- DOM parser skips rows of nested tables (generates one cell)
- `sheet_add_dom` utility function

1.20200316.1

- fixed uses of `hasOwnProperty`
- proper encoding of ampersand in XLSX properties
- updated CFB to 1.1.4

1.20200309.1

- tables `header` option to disable header row
- tables throw an error when headers collide
- proper encoding of chinese characters in:
  + (core): print header/footer
  + (image/chart): shape text
  + (template editor) defined names and VML


1.20200219.1

- faster {encode,decode}\_cell
- correct row number regexes

1.20200212.1

- Print "Fit All {Rows,Columns} on One Page"

1.20200210.1

- XLSB read Rich Text runs
- XLS / XLSB write Rich Text runs (requires bookSST: true)

