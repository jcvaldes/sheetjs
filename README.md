# [SheetJS Pro](http://sheetjs.com)

The [community edition documentation](https://docs.sheetjs.com/) covers basic
API and functionality.  This build includes the following features:

(note: the easiest way to explore this document is to search for keywords and
read the relevant sections.  For general inspiration, skim the headers!  Large
code samples are hidden behind details-summary blocks and can be revealed by
clicking on "Click to show")

## Importing and Exporting with Styles

The `cellStyles` option should be passed to the `read` and `write` functions:

```js
var wb = XLSX.readFile("test.xlsx", {cellStyles: true});
XLSX.writeFile(wb, "out.xlsx", {cellStyles: true});
```

For XLS and XLSB, `bookSST` must also be specified to support Rich Text runs
(this is a technical file format limitation, it does not affect XLSX):

```js
XLSX.writeFile(wb, "out.xls", {cellStyles: true, bookSST: true});
XLSX.writeFile(wb, "out.xlsb", {cellStyles: true, bookSST: true});
```


### Color Objects

The most consistent way to specify a color is with a RGB string or hex code:

```js
ws.A1.s = { color: { rgb: "FF0000" }}; // red
```

<details>
  <summary><b>Other color specifications</b> (click to show)</summary>

For round-tripping worksheet styles, it is also possible to use a color from the
default system theme:

```js
ws.A2.s = { color: { theme: 1, tint: 0.4 }};
```

There is also a basic palette accessible as "indexed" colors:

```js
ws.A2.s = { color: { index: 12 }};
```

</details>


### Text Styling

Each cell has a style object accessible at the `.s` key of the cell, following
the schema:

| key         | description                                                 |
|:------------|:------------------------------------------------------------|
| `bold`      | Set to `true` to bold the cell text                         |
| `italic`    | Set to `true` to italicize the cell text                    |
| `underline` | 1 = single, 2 = double                                      |
| `sz`        | Font size in pts (e.g. 10, 12, 14)                          |
| `strike`    | Set to `true` for strike-through effect                     |
| `name`      | Font name                                                   |
| `color`     | Color object                                                |
| `valign`    | Vertical alignment: `sub` (subscript) `super` (superscript) |

For example, the following worksheet tests bold and italic:

```js
var ws = XLSX.utils.aoa_to_sheet([
  ["Normal"],
  ["Bold"],
  ["Italic"],
  ["B+I"],
]);

ws["A2"].s = { bold: true };
ws["A3"].s = { italic: true };
ws["A4"].s = { bold: true, italic: true };
```


#### Rich Text

For more advanced text styling, the `.R` key of a cell can be set to an array of
cell objects.  The writer will apply the relevant style to each individual block
and concatenate the result.

For example, the following cell will contain the text "<b>bold</b><i>italic</i>"
where `bold` is bold and `italic` is italicized:

```js
ws["A1"].R = [
  {t:'s', v:'bold', s:{bold:true}},
  {t:'s', v:'italic', s:{italic:true}},
];
```

If the rich text array is specified, it will be used in lieu of the `.v` text.


#### Hyperlink Styling

By default, a hyperlink attached to a cell does not modify the text style.  The
Excel behavior (styling a link in blue and converting to a purple style when you
click the link) is a special theme text color `color: { theme: 10 }`:

```js
ws["B2"].l = { Target: "https://sheetjs.com" }; // link doesn't modify style

if(!ws["B2"].s) ws["B2"].s = {};
ws["B2"].s.color = { theme: 10 }
```

The helper function `cell_set_hyperlink` will apply the hyperlink text style in
addition to assigning a hyperlink:

```js
XLSX.utils.cell_set_hyperlink(ws["B2"], "https://sheetjs.com");
```


### Cell Alignment and Wrapping

Cell Alignment properties are controlled in the `alignment` key of the style
object.  The supported `alignment` features are explained in the table below:

| key            | description                                                 |
|:---------------|:------------------------------------------------------------|
| `horizontal`   | Horizontal alignment: `"left" "center" "right" "justify"`   |
| `vertical`     | Vertical alignment: `"top" "center" "bottom"`               |
| `indent`       | Indent level (zero-indexed, 0 is default, 7 is maximum)     |
| `wrapText`     | If `true`, text is wrapped: `\n` chars appear as new lines  |
| `shrinkToFit`  | Excel Text control :: "Shrink to fit"                       |
| `textRotation` | Text rotation in degrees                                    |

For example, the following cell will contain the wrapped text `"a\nb\nc"`, using
top vertical alignment and left horizontal alignment:

```js
ws["A1"] = {
  t: "s", /* string cell */
  v: "a\nb\nc", /* use standard \n newline.  do not include \r */
  s: {
    alignment: {
      horizontal: "left",
      vertical: "top",
      wrapText: true
    }
  }
};
```


The following sample stresses the 9 alignment pairs and text wrapping:

<details>
  <summary><b>Alignment Example</b> (click to show)</summary>

```js
var ws = XLSX.utils.aoa_to_sheet([
  ["TL", "TC", "TR", "a\nb\nc", "a\nb\nc"],
  ["CL", "CC", "CR", "d\ne\nf"],
  ["BL", "BC", "BR", "g\nh\ni"],
]);

/* --- Horizontal Alignment --- */
/* Left-align   A1:A3 */
XLSX.utils.sheet_set_range_style(ws, "A1:A3", {alignment: { horizontal: "left" }});
/* Center-align B1:B3 */
XLSX.utils.sheet_set_range_style(ws, "B1:B3", {alignment: { horizontal: "center" }});
/* Right-align  C1:C3 */
XLSX.utils.sheet_set_range_style(ws, "C1:C3", {alignment: { horizontal: "right" }});

/* --- Vertical Alignment --- */
/* Top-align    A1:C1 */
XLSX.utils.sheet_set_range_style(ws, "A1:C1", {alignment: { vertical: "top" }});
/* Center-align A2:C2 */
XLSX.utils.sheet_set_range_style(ws, "A2:C2", {alignment: { vertical: "center" }});
/* Bottom-align A3:C3 */
XLSX.utils.sheet_set_range_style(ws, "A3:C3", {alignment: { vertical: "bottom" }});

/* --- Wrap D1:D4 (leave E1 unwrapped) --- */
XLSX.utils.sheet_set_range_style(ws, "D1:D3", {alignment: { wrapText: true }});
```

</details>

### Cell Background

The cell style object can contain information about the background:

| key           | excel name         | description                 |
|:--------------|:-------------------|:----------------------------|
| `patternType` | "Pattern Style"    | type of pattern (see below) |
| `fgColor`     | "Background Color" | primary background color    |
| `bgColor`     | "Pattern Color"    | secondary background color  |

If omitted, `patternType` is assumed to be "solid", corresponding to a single
color as indicated by the `fgColor` field (the `bgColor` secondary color is
disregarded).  Other patterns are listed below:

<details>
  <summary><b>Pattern Styles</b> (click to show)</summary>

The description can be found in the tooltips that appear when hovering over a
pattern in the Pattern Style picker.  The table entries are listed in row-major
order based on the appearances in the Pattern Style picker:

| pattern style     | Excel 2019 Description       |
|:------------------|:-----------------------------|
| `solid`           | Solid                        |
| `darkGray`        | 75% Gray                     |
| `mediumGray`      | 50% Gray                     |
| `lightGray`       | 25% Gray                     |
| `gray125`         | 12.5% Gray                   |
| `gray0625`        | 6.25% Gray                   |
| `darkHorizontal`  | Horizontal Stripe            |
| `darkVertical`    | Vertical Stripe              |
| `darkDown`        | Reverse Diagonal Stripe      |
| `darkUp`          | Diagonal Stripe              |
| `darkGrid`        | Diagonal Crosshatch          |
| `darkTrellis`     | Thick Diagonal Crosshatch    |
| `lightHorizontal` | Thin Horizontal Stripe       |
| `lightVertical`   | Thin Vertical Stripe         |
| `lightDown`       | Thin Reverse Diagonal Stripe |
| `lightUp`         | Thin Diagonal Stripe         |
| `lightGrid`       | Thin Horizontal Crosshatch   |
| `lightTrellis`    | Thin Diagonal Crosshatch     |

</details>

For example, to set a solid cell background, assign to the `.s.fgColor` property:

```js
ws.A3.s = { fgColor: { rgb: 0xFF0000 } }; // green background
```


### Cell Borders

The `s.top`, `s.bottom`, `s.left` and `s.right` properties control cell borders.
They are shaped as follows:

| key     | description             |
|:--------|:------------------------|
| `style` | border type (see below) |
| `color` | color object            |

The following table shows the supported border styles in Excel (first column) as
well as the styles when a worksheet is written to HTML (second column) and the
conditions under which the cell border will map to the Excel border.  Since
Excel supports different border styles from HTML, some styles will never be
matched but all Excel border styles will be written to HTML.

| Excel border value | generated `border-style` | matching `border-style`  |
|:-------------------|:-------------------------|--------------------------|
| (default)          | `none`                   | `none`                   |
| `thin`             | `solid 1px`              | `solid`  (width    1px)  |
| `hair`             | `solid 1px`              |                          |
| `medium`           | `solid 2px`              | `solid`  (width    2px)  |
| `thick`            | `solid 3px`              | `solid`  (width >= 3px)  |
| `double`           | `double 3px`             | `double` (width >= 1px)  |
| `dotted`           | `dotted 1px`             | `dotted` (width    1px)  |
| `dashDotDot`       | `dotted 1px`             |                          |
| `mediumDashDotDot` | `dotted 2px`             | `dotted` (width >= 2px)  |
| `dashed`           | `dashed 1px`             | `dashed` (width    1px)  |
| `dashDot`          | `dashed 1px`             |                          |
| `slantDashDot`     | `dashed 1px`             |                          |
| `mediumDashed`     | `dashed 2px`             | `dashed` (width >= 2px)  |
| `mediumDashDot`    | `dashed 2px`             |                          |

The `thin` border style will be used for all other HTML border styles (groove,
ridge, inset, outset) as they have no Excel equivalent.

For example:

```js
ws["A4"].s = {
  top: { style: "thin" }, // thin black border on top
  bottom: { style: "thick", color: { rgb: 0xFF0000 } }, // red thick border
  left: { style: "dashed", color: { rgb: 0x00FF00 } } // green dashed border
}
```


### Default Row Heights and Column Widths

The `!sheetFormat` key of the worksheet specifies default cell dimensions:

| key   | description         |
|:------|:--------------------|
| `row` | Row height object   |
| `col` | Column width object |

Row height can be specified with `hpx` or `hpt`.  Row visibility is controlled
by the `hidden` key.

Column width can be specified with `wch` or `wpx` or `width`.  To hide all of
the columns, set `wpx` to zero.

NOTE: the default column width cannot be automatic!  Setting to `auto: 1` will
have surprising results.

```js
ws['!sheetFormat'] = {
  row: {
    hpx: 36 // default row height 36 px
  },
  col: {
    wpx: 100 // default column width 100px
  }
};
```



### Cell Widths

Excel calculates column widths on save.  For new files, use the column object
`.auto` property to trigger a calculation when the file is saved:


```js
if(!ws['!cols']) ws['!cols'] = [];
ws['!cols'][1] = { auto: 1 }; // set the second column to auto-fit width
```


### Default Cell Styles for Columns

By default, Excel will use the "Normal" style when filling out cells that are
not within the range of the worksheet.  Assigning a number format to the `z`
field and a cell style object to the `s` field will override the column default:

```js
var column = XLSX.utils.decode_col("B");
var number_format = "#,##0.00"; // thousands separator + 2 decimal places
var style = { bold: true }; // always bold new cells in the column

if(!ws['!cols']) ws['!cols'] = [];
if(!ws['!cols'][column]) = ws['!cols'][column] = { auto: 1 }; // default col
ws["!cols"][column].z = number_format;
ws["!cols"][column].s = style;
```


### Naming Styles

Styles are not named by default.  To force a name, set the `style` property of
the `.s` style object:

```js
ws["A4"].s.style = "Test Name";
```


### Cell-Level Protection

By default, worksheets are not protected.  To enable any sort of protection,
the `"!protect"` key of the worksheet object must be set.  The `password` key
of the protection object specifies the password to unlock the sheet.  If no
password is set, any user can click "Unprotect Sheet" to unprotect.  More
details are available at <https://docs.sheetjs.com/#worksheet-object>

When a worksheet is locked, by default all cells are locked but the formulae are
visible.  These options are adjustable on a cell level.  The cell `.s` style
object additionally supports the following properties:

| key         | description                                                 |
|:------------|:------------------------------------------------------------|
| `hidden`    | Set to `true` to hide the formula if sheet is protected     |
| `editable`  | Set to `true` to enable edits (by default cells are locked) |

<details>
  <summary><b>Code Example</b> (click to show)</summary>

```js
var ws = XLSX.utils.aoa_to_sheet([[1,2],[3,4]]);
ws["A1"].f = "2-1";  // for illustrative purposes, assign formulae to A1:A2
ws["A2"].f = "A1+2";

ws["!protect"] = {}; /* enable worksheet protection */

/* hide formulae for cells A1:B1 */
XLSX.utils.sheet_set_range_style(ws, "A1:B1", { hidden: true });
/* enable editing for cells A1:A2 */
XLSX.utils.sheet_set_range_style(ws, "A1:A2", { editable: true });
```

</details>

Since cell-level protection is associated with the cell style, the option
`cellStyles: true` must be set when writing with `write` or `writeFile`.


### Active selection

The worksheet `!sel` key controls the actively selected cells in a worksheet.
The parsers will return an object with the following keys:

| key     | description                                   |
|:--------|:----------------------------------------------|
| `cell`  | actively selected cell                        |
| `range` | current selected cells (can be a split range) |

For example, the following specifies that the range A1:B3 should be selected
with A2 as the active cell:

```js
ws["!sel"] = { cell: "A2", range: "A1:B3" };
```


### API

#### Set Style to a Range of Cells

`XLSX.utils.sheet_set_range_style` applies a style to a range of cells.  Text
and background styles are applied to every cell, while `top` / `bottom` / `left`
/ `right` borders are applied to the exterior sides of the range.

In addition to the standard options from the style object, the following keys
are supported:

| key     | description                                                        |
|:--------|:-------------------------------------------------------------------|
| `z`     | Number format (function will set cell `.z` property of each cell)  |
| `incol` | Interior Vertical border (applied to every non-exterior border)    |
| `inrow` | Interior Horizontal border (applied to every non-exterior border)  |

The interior borders are applied to every applicable cell:

- `incol` sets the `left` border for every cell not in the first column and sets
  the `right` border for every cell not in the last column.
- `inrow` sets the `top` border for every cell not in the first row and sets the
  `bottom` border for every cell not in the last row.

For example, given the sheet

```
XXX| A | B | C | D |
---+---+---+---+---+
 1 | 1 | 2 | 3 | 4 |
 2 | 5 | 6 | 7 | 8 |
 3 | 9 | A | B | C |
 4 | D | E | F | 0 |
```

This code will set:
- background and text color for every cell in `B2:C3` and
- left border of cells `B2:B3`
- right border of cells `C2:C3`
- top border of cells `B2:C2`
- bottom border of cells `B3:C3`

```js
XLSX.utils.sheet_set_range_style(
  ws, // worksheet
  "B2:C3", // range
  { // style object
    fgColor: { rgb: 0x0000FF }, // blue solid background
    color: { rgb: 0xFFFFFF }, // white text
    top: { style: "thick", color: { rgb: 0xFFFF00 } }, // thick yellow border
    bottom: { style: "thick", color: { rgb: 0xFF0000 } }, // red thick border
    left: { style: "dashed", color: { rgb: 0x00FF00 } } // green dashed border
  }
);
```

Set a style to `false` to remove (e.g. `bold: false` will remove bold cells)


#### Determining Final Style

`XLSX.utils.get_computed_style(ws, address)` applies conditional formatting,
table, and other styles that may affect the computed cell style.


## Cell Comments

Cell comments are stored in the `c` array of cell objects.

Each comment object supports the following properties:

| key | description           |
|:----|:----------------------|
| `a` | Author of the comment |
| `t` | Plaintext             |
| `R` | Rich text array       |

Additional named properties can be attached directly to the array:

| key      | description                                        |
|:---------|:---------------------------------------------------|
| `hidden` | Hide comment (displays when mouse hovers over cell |
| `!pos`   | Position of comment                                |
| `s`      | Style (see below for limitations)                  |

<details>
  <summary><b>How to specify comment position</b> (click to show)</summary>

Note: for users of the Pro Image build or any build that includes image or chart
support, this is the same content displayed in "Image Dimensions and Location".

Comments can start anywhere in the worksheet.  There are three position styles:

**Absolute Position with Size**

| key | interpretation                             |
|:----|:-------------------------------------------|
| `x` | X-coordinate of upper-left corner (pixels) |
| `y` | Y-coordinate of upper-left corner (pixels) |
| `w` | width (pixels)                             |
| `h` | height (pixels)                            |

The following example specifies an absolute size of 456 x 123 pixels starting
from the pixel position (0, 0):

```js
{
  /* location of upper-left corner (in pixels) */
  x: 0, y: 0,
  /* object size (in pixels) -- Excel will scale to fit dimensions */
  w: 456, h: 123,
};
```

**Relative Position with Size**

This style lets you anchor to a given starting cell.

| key | interpretation                                          |
|:----|:--------------------------------------------------------|
| `r` | row of upper-left corner (0-indexed)                    |
| `c` | col of upper-left corner (0-indexed)                    |
| `x` | X distance from upper-left corner of base cell (pixels) |
| `y` | Y distance from upper-left corner of base cell (pixels) |
| `w` | width (pixels)                                          |
| `h` | height (pixels)                                         |

The most common use case is to set a comment to show up to the right of a given
cell.  To do that, set `r` to the cell row, `c` to cell column + 1. For example:

```js
var addr = XLSX.utils.encode_cell({r:3, c:4});
ws[addr].c["!pos"] = {
  r: 3, // same row as cell
  c: 4+1, // next column
  x: 0, y:0, // start at the topleft corner
  w: 456, h: 123
}
```

**Relative Position of Corners**

This lets you create comments that appear to span across a cell or range.

| key | interpretation                                           |
|:----|:---------------------------------------------------------|
| `r` | row of upper-left corner (0-indexed)                     |
| `c` | col of upper-left corner (0-indexed)                     |
| `x` | X distance from upper-left corner of base cell (pixels)  |
| `y` | Y distance from upper-left corner of base cell (pixels)  |
| `R` | row of lower-right corner (0-indexed)                    |
| `C` | col of lower-right corner (0-indexed)                    |
| `X` | X distance from lower-right corner of base cell (pixels) |
| `Y` | Y distance from lower-right corner of base cell (pixels) |

The following example specifies an comment covering over the cell B5.

- The address of the upper-left corner is `r = 4, c = 1`
- The upper-left pixel offset will be 0 in both directions.
- The address of the lower-right corner is `R = r + 1 = 5, C = c + 1 = 2`
- The upper-left pixel offset will be 0 in both directions.


```js
var addr = { r: 4, c: 1 };
var rowspan = 1, colspan = 1;
ws["C3"].c["!pos"] = {
  /* upper-left corner cell address*/
  c: addr.c, r: addr.r,
  /* lower-right corner cell address*/
  C: addr.c + colspan, R: addr.r + rowspan,
  /* zero pixel offset for both points */
  x: 0, y: 0, X: 0, Y: 0
};
```


</details>

The `s` key supports the following styles:

| key       | description        |
|:----------|:-------------------|
| `fgColor` | Primary fill color |

The following example:

- adds a cell "C3" to the worksheet with value "abc"
- adds a comment with a single author "SheetJS" and rich text
- sets comment positioning to start at cell D5 and span 100 px width / height
- makes comment fill in yellow

```js
var comment_part = {
  a:"SheetJS",
  R: [
    {t: "s", v: "plain text "},
    {t: "s", v: "bold text ", s: { bold: true } }
  ]
};
var comment = [];
comment.push(comment_part);
comment["!pos"] = {c: 3 /* "D" */, r: 4 /* "5" */, w: 100, h: 100};
comment.s = { fgColor: { rgb: "FFFF00" } };

XLSX.utils.sheet_add_aoa(ws, [["abc"]], {origin: "C3"});
ws["C3"].c = comment;
```



## Print Settings

Print settings are generally stored in the `"!print"` key of the worksheet.  For
backwards compatibility, legacy key names are supported but not documented here.

<details>
  <summary><b>Example Print Settings</b> (click to show)</summary>

```js
ws["!print"] = {

  /* Print area A1:E50 */
  area: { s:{r:0, c:0}, e:{c:4,r:49} },

  /* Forced row break at row 3-4 */
  rowBreaks: [{R:3}],

  /* Forced column break at column C-D */
  colBreaks: [{C:3}],

  /* margins */
  margins: {
    left: 0.7,
    right: 0.6,
    top: 0.5,
    bottom: 0.4,
    header: 0.3,
    footer: 0.2
  },

  /* raw header string */
  header: "&&SheetJS",

  /* parameterized footer settings */
  footer: {
    odd: {
      center: { R: [
        {w: "&A", s: { bold: true } },
        {w: "&D", s: { italic: true } },
        {w: "&F", s: { underline: true } },
        {w: "red", s: { color: {rgb: 0xFF0000 } } },
        {w: "&Z", s: { underline: 2, strike: true}}
      ] }
    },
    even: {
      right: { w: "&I&D&I" },
      left: { w: "&A", s: { bold: true } }
    }
  },

  /* other properties */
  props: {
    orientation: "landscape",
    paper: "Legal",
    dpi: 1200,
    first: 3,
    comments: "displayed",
    errors: "n/a",
    gridlines: true,
    bw: true,
    draft: true,
    headings: true,
    order: "over",
    centerX: false,
    centerY: true,
    scale: 50 // this sets 50% scale
    //fit: { width: 2, height: 3 } // this force fits to 2 pages x 3 pages
  }
};

```

</details>

### Print Area

Print area is normally represented in the defined name `_xlnm.Print_Area` scoped
to the specific worksheet.  The library automatically copies and decodes the
range to the `.area` property of the worksheet.  On write, if the print area is
specified, the defined name will be re-written!

```js
/* set print area to A1:D20 */
ws["!print"].area = "A1:D20";                              // range string
ws["!print"].area = { s: { r:0, c:0 }, e: { r:19, c:3 } }; // range object
```

### Print Options

Various print options are stored in the `.props` key of the print object:

#### Orientation

`.orientation` marks the page orientation, valid values are `"landscape"` for
landscape, `"portrait"` for portrait, and `"default"` for printer default.

```js
ws["!print"].props.orientation = "landscape"; // landscape
```

#### Scaling

`.fit` enables the "Fit to Page" print option.  The fit object supports the
`width` and `height` properties indicating the number of pages to fit:

```js
ws["!print"].props.fit = { width: 3, height: 4 };
```

Other common settings are shown in the table below:

| Fit setting                   | `fit` object value        |
|:------------------------------|:--------------------------|
| "Fit Sheet on One page"       | `{ width: 1, height: 1 }` |
| "Fit All Columns on One Page" | `{ width: 1, height: 0 }` |
| "Fit All Rows on One Page"    | `{ width: 0, height: 1 }` |

`.scale` is the integral percentage print scale, defaulting to 100 (for 100%).
The scale must be between 10 (10%) and 400 (4x).  This option only applies when
fit is not specified!

```js
ws["!print"].props.scale = 50; // 50%
```

#### Paper Size / Type

`.paper` controls the Paper size.  It can be set to a numeric or string value,
as described in the table below, or to an object with `width` and `height`
properties whose values must be strings that include a distance unit.

<details>
  <summary><b>Paper Settings</b> (click to show)</summary>

| code | alias      | size              | description              |
|-----:|:-----------|:------------------|:-------------------------|
|    1 | Letter     | 8.5 in x 11 in    | Letter Paper             |
|    2 |            | 8.5 in x 11 in    | Letter Small Paper       |
|    3 | Tabloid    | 11 in x 17 in     | Tabloid Paper            |
|    4 |            | 17 in x 11 in     | Ledger Paper             |
|    5 | Legal      | 8.5 in x 14 in    | Legal Paper              |
|    6 |            | 5.5 in x 8.5 in   | Ledger Paper             |
|    7 | Executive  | 7.25 in x 10.5 in | Executive Paper          |
|    8 | A3         | 297 mm x 420 mm   | A3 Paper                 |
|    9 | A4         | 210 mm x 297 mm   | A4 Paper                 |
|   10 |            | 210 mm x 297 mm   | A4 Small Paper           |
|   11 | A5         | 148 mm x 210 mm   | A5 Paper                 |
|   12 | B4         | 257 mm x 364 mm   | B4 (JIS)                 |
|   13 | B5         | 182 mm x 257 mm   | B5 (JIS)                 |
|   14 | Folio      | 8.5 in x 13 in    | Folio paper              |
|   20 | Envelope   | 4.125 in x 9.5 in | Envelope #10             |
|   27 |            | 110 mm x 220 mm   | Envelope DL              |
|   28 |            | 162 mm x 229 mm   | Envelope C5              |
|   29 |            | 324 mm x 458 mm   | Envelope C3              |
|   30 |            | 229 mm x 324 mm   | Envelope C4              |
|   31 |            | 114 mm x 162 mm   | Envelope C6              |
|   34 |            | 175 mm x 250 mm   | Envelope B5              |
|   37 | Monarch    | 3.875 in x 7.5 in | Envelope Monarch         |
|   43 |            | 100 mm x 148 mm   | Japanese Postcard        |
|   69 |            | 200 mm x 148 mm   | Japanese Double Postcard |
|   70 | A6         | 105 mm x 148 mm   | A6 Paper                 |

</details>

For example, the standard "#10 envelope" paper can be set in a few ways:

```js
ws["!print"].props.paper = 20;                                    // code
ws["!print"].props.paper = "Envelope";                            // alias
ws["!print"].props.paper = { height: "4.125in", width: "9.5in" }; // size
```

#### Print Quality

`.dpi` controls the print quality.  The standard options differ between versions
but the standard options in Excel 2019 are 600 (default) and 1200.

```js
ws["!print"].props.dpi = 600;
```

#### First Page Number

`.first` controls the first page number.  If it is not specified or set to null,
the first page number is automatically determined:

```js
ws["!print"].props.first = 3; // First Page Number 3
```

#### Centering Content in the Pages

By default, content is aligned to the upper-left corner of the content region.

The `.centerX` and `.centerY` properties control the alignment.  If `centerX` is
true, then the content will be centered horizontally within the region.  The
`centerY` property does the same for the vertical axis.

```js
ws["!print"].props.centerX = true;  // center horizontally
ws["!print"].props.centerY = false; // align to top of content region
```

#### Print Features

There are 4 checkbox options:

| `.props` property | UI Description                         |
|:------------------|:---------------------------------------|
| `.gridlines`      | Print Gridlines                        |
| `.bw`             | Print Black and White                  |
| `.draft`          | Print Draft quality (without graphics) |
| `.headings`       | Print Row and Column Headings          |

The `.comments` property describes how cell comments are displayed on print:

| `.comments` | UI Description          |
|:------------|:------------------------|
| "displayed" | "As displayed on sheet" |
| "end"       | "At end of sheet"       |
| "none"      | "(None)"                |

The `.errors` property describes how cells with errors are displayed on print:

| `.errors`   | UI Description |
|:------------|:---------------|
| "displayed" | "displayed"    |
| "none"      | "<blank>"      |
| "dash"      | "`--`"         |
| "n/a"       | "NA"  (error)  |

For example:

```js
ws["!print"].props.comments   = "displayed";
ws["!print"].props.errors     = "displayed";
ws["!print"].props.gridlines  = false;
ws["!print"].props.bw         = true;
ws["!print"].props.draft      = false;
ws["!print"].props.headings   = true;
```

#### Page Order

The `.order` property should be set to `true` or `"over"` for "Over, then down"
page order.  The default value is "Down, then over".

```js
ws["!print"].props.order = true;  // "Over, then down"
ws["!print"].props.order = false; // "Down, then over"
```

### Page Margins

The `.margins` key of the print object is an object which stores all of the page
margins in inches.  The legacy `ws["!margins"]` object is also supported.  The
key properties are listed below:

<details>
  <summary><b>Page margin details</b> (click to show)</summary>

| key      | description            | "normal" | "wide" | "narrow" |
|----------|------------------------|:---------|:-------|:-------- |
| `left`   | left margin (inches)   | `0.7`    | `1.0`  | `0.25`   |
| `right`  | right margin (inches)  | `0.7`    | `1.0`  | `0.25`   |
| `top`    | top margin (inches)    | `0.75`   | `1.0`  | `0.75`   |
| `bottom` | bottom margin (inches) | `0.75`   | `1.0`  | `0.75`   |
| `header` | header margin (inches) | `0.3`    | `0.5`  | `0.3`    |
| `footer` | footer margin (inches) | `0.3`    | `0.5`  | `0.3`    |

```js
/* Set worksheet sheet to "normal" */
ws["!print"].margins = {left:0.7,  right:0.7,  top:0.75, bottom:0.75, header:0.3, footer:0.3 };
/* Set worksheet sheet to "wide" */
ws["!print"].margins = {left:1.0,  right:1.0,  top:1.0,  bottom:1.0,  header:0.5, footer:0.5 };
/* Set worksheet sheet to "narrow" */
ws["!print"].margins = {left:0.25, right:0.25, top:0.75, bottom:0.75, header:0.3, footer:0.3 };
```

</details>

### Row and Column Breaks

Row and Column Breaks determine the rows and columns where Excel forces a new
page.  This does not control any other natural page breaks, merely forcing
breaks at the specified points.

The `.rowBreaks` of the print object is an array of row break objects.  Each row
break has a single `R` field indicating the first zero-indexed row of the break:

```js
/* Force breaks between rows 3-4, 7-8 */
ws["!print"].rowBreaks = [ {R:3}, {R:7} ];
```

The `.colBreaks` of the print object is an array of column break objects.  Each
column break has a single `C` field for the first zero-indexed column:

```js
/* Force breaks between columns C-D, G-H */
ws["!print"].colBreaks = [ {C:3}, {C:7} ];
```

### Header and Footer

Internally, Excel stores header and footer strings for odd pages, even pages,
and the first page.  The strings include control specifiers for marking styles
and custom fields.  The parsers stick to this format as much as possible, while
the writers are a bit more flexible and support more idiomatic forms.

#### General Structure

The `.header` and `.footer` fields correspond to the header and footer.

Within each of those objects, the `.odd`, `.even`, and `.first` keys correspond
to the Odd-numbered pages, Even-numbered pages, and first page respectively.

By default, if the `.even` key is not specified, the `.odd` header/footer will
apply to even-numbered pages.  If the `.first` key is not specified, the `.odd`
header/footer will also apply to the first page.  To clear either format,
explicitly set it to an empty string:


```js
/* Apply header to every page except the first */
ws["!print"].header = {
  odd: "Hello World!",
  first: ""
};
/* Apply footer to first page and even pages but not odd pages */
ws["!print"].footer = {
  first: "First page yo!",
  even: "Even steven!",
  odd: ""
};
```

#### Header and Footer Format Representation

The internal representation uses `&`-prefixed control specifiers and is a bit
clumsy for programmatic use.  For pass-through purposes, this form is supported,
but the library also supports other forms.

**Pass-through Strings**

The Official VBA documentation describes the formatting codes in the article
["Formatting and VBA codes for headers and footers"](https://docs.microsoft.com/en-us/office/vba/excel/Concepts/Workbooks-and-Worksheets/formatting-and-vba-codes-for-headers-and-footers).

<details>
  <summary><b>Header and Footer Strings</b> (click to show)</summary>

*Position*

By default, if no positioning specifiers are included, the header/footer is
centered.  The specifiers indicate that the following text and specifiers apply
to the indicated position.

| Spec | Meaning        |
|:-----|:---------------|
| `&C` | Center section |
| `&L` | Left section   |
| `&R` | Right section  |

```js
/* Write to different sections */
ws["!print"].footer.odd= ("&L" + "Left") + ("&C" + "Center" ) + ("&R" + "Right")
//                       |-Left--------|   |-----Center-----|   |--------Right-|
```

*Special Data*

Fields like Page Number are specified as follows:

| Spec     | Meaning                        |
|:---------|:-------------------------------|
| `&P`     | Current Page number            |
| `&P+###` | Page Number + specified number |
| `&P+###` | Page Number - specified number |
| `&N`     | Page Count (`Page &P of &N)`   |
| `&D`     | Current Date                   |
| `&T`     | Current Time                   |
| `&A`     | Worksheet Name                 |
| `&F`     | File Name                      |
| `&Z`     | File Location                  |
| `&&`     | Literal `&`                    |

For example, the following code sets footer center to `Page # of #`:

```js
ws["!print"].footer = "&CPage &P of &N";
```

*Text Styles*

The text styling logic is state-based.  The individual operators toggle state.
For example, `&B` is the bold token, so the text "foo**bold**bar" is saved as

```js
ws["!print"].footer.odd = "foo&Bbold&Bbar"
// off------------------------^^on--^^off-
```

| Spec     | Meaning                                   |
|:---------|:------------------------------------------|
| `&B`     | Toggle Bold                               |
| `&I`     | Toggle Italic                             |
| `&U`     | Toggle Underline **                       |
| `&E`     | Toggle Double-Underline **                |
| `&H`     | Toggle Shadow                             |
| `&O`     | Toggle Outline                            |
| `&X`     | Toggle Superscript **                     |
| `&Y`     | Toggle Subscript **                       |
| `&S`     | Toggle Strike-through                     |
| `&K...`  | Font color (RRGGBB or theme+tint) **      |
| `&#`     | Text size in points (`&16` is 16 pt size) |
| `&"..."` | Font name                                 |
| `&"+"`   | Use the current theme Heading font        |
| `&"-"`   | Use the current theme Body font           |

The underline modes and vertical alignment are mutually exclusive, so turning on
one will disable the other.  For example, the format code `&XFoo&YBar&XBaz` will
write `Foo` in superscript, `Bar` in subscript, and `Baz` in superscript.

Font color is specified in one of two ways: standard RGB string (`FF0000` = Red)
or Theme and Tint (`##S###` where the first two digits running from `00` to `12`
specify the theme color, `S` is either "+" or "-", and `###` is a percentage
running from 000 to 100)

</details>


## International Locale Support

This build includes our SSF Pro component for reading and writing files using
locales other than `en-US`.  The locale should be an IETF language tag:

```js
XLSX.SSF.setlocale("de-DE"); // German (Germany)
XLSX.SSF.setlocale("sv-SE"); // Swedish (Sweden)
```

The support uses the built-in `Intl` framework when available to deduce field
information.  The following settings are changed when a new locale is set:

- Thousands ("grouping") and Decimal separators
- Day and Month names (`ddd` / `dddd` / `mmm` / `mmmm` / `mmmmm`)
- Default Date Style (format code 14)
- Locale-Specific formats (for `zh-CN` / `zh-TW` / `ja-JP` / `ko-KR` / `th-TH`)

Locales and other settings should be set *before* reading files.

<details>
  <summary><b>The SSF Library</b> (click to show)</summary>

The SSF Pro component is also available separately under the name `@sheet/ssf`.
Please inquire if you directly want to use the module outside of this library.

</details>

### Currency Sigil

Currency symbols are independent from standard locales and are set using ISO
4217 codes such as "USD" for US Dollars:

```js
XLSX.SSF.setcurrency("EUR");
```

### Converting from Foreign Formats to SheetJS

Excel will transparently convert format strings from other formats back to US.
For example, in German Excel `jjjj-mm-tt` is the displayed format but it is
persisted as `yyyy-mm-dd`.  Given a custom format, `SSF.normalize` will convert
back to the US form:

```js
XLSX.SSF.setlocale("de-DE");
XLSX.SSF.normalize("jjjj-mm-tt"); // "yyyy-mm-dd"
XLSX.SSF.normalize("#.##0,00"); // "#,##0.00"
```

## Data Validations

Data Validations are stored in the `!validations` array of the worksheet:

```js
ws['!validations'] = [
  /* A1:A5 show a fixed dropdown menu with specified values */
  {
    ref: 'A1:A5',
    t: 'List',
    l: ["a", "b", "c", "d", "e"],
    input: {
      title: "Letter",
      message: "Type a letter a,b,c,d,e"
    }
  },
  /* B1:B5 are restricted to integers between 0 and 10, blanks not ignored */
  {
    ref: 'B1:B5',
    t: 'Whole',
    op: 'IN',
    blank: false,
    min: 0,
    max: 10
  }
]
```

Each validation object in the array follows the schema:

| key         | description                                |
|:------------|:-------------------------------------------|
| `ref`       | Range or cell string or address object     |
| `t`         | Type of Data Validation (see table)        |
| `l`         | Array of strings for a fixed dropdown List |
| `f`         | Formula or Range for Custom or List DV     |
| `op`        | Data operator (see below)                  |
| `min/max/v` | Min / Max / Exact values for the operator  |
| `blank`     | "Ignore Blank" (set to `false` to disable) |
| `input`     | Input Message (see below)                  |
| `error`     | Error Alert (see below)                    |

The `ref` reference can be passed as a string ("A2" or "A2:C4") or address
object like `{r:1, c:0}` or range like `{s:{r:1,c:0}, e:{r:3,c:2}}`.

### Data Validation Types

The `type` refers to the "Allow" option in the Settings tab of Data Validation:

| type        | Excel interface name | Operator | Parameters           |
|:------------|:---------------------|:--------:|:---------------------|
| `"Any"`     | "Any Value"          |    No    |                      |
| `"Whole"`   | "Whole number"       |   Yes    | `min` + `max` OR `v` |
| `"Decimal"` | "Decimal"            |   Yes    | `min` + `max` OR `v` |
| `"List"`    | "List"               |    No    | `l` OR `f`           |
| `"Date"`    | "Date"               |   Yes    | `min` + `max` OR `v` |
| `"Time"`    | "Any Value"          |   Yes    | `min` + `max` OR `v` |
| `"Length"`  | "Text length"        |   Yes    | `min` + `max` OR `v` |
| `"Custom"`  | "Custom"             |    No    | `f`                  |

### Data Validation Operators

For the numeric data validations, ranges are specified in terms of operators:

| operator | Excel interface name       | min | max | v   |
|:---------|:---------------------------|:----|:----|:----|
| `"IN"`   | "between"                  | Yes | Yes |  No |
| `"OT"`   | "not between"              | Yes | Yes |  No |
| `"EQ"`   | "equal to"                 |  No |  No | Yes |
| `"NE"`   | "not equal to"             |  No |  No | Yes |
| `"GT"`   | "greater than"             |  No |  No | Yes |
| `"LT"`   | "less than"                |  No |  No | Yes |
| `"GE"`   | "greater than or equal to" |  No |  No | Yes |
| `"LE"`   | "less than or equal to"    |  No |  No | Yes |


### Input Message

By default, "Show input message when cell is selected" is enabled.  To disable,
set the `input` property to `false`.

If the `input` property is assigned a value, it is expected to be an object with
the following fields:

| key     | description     |
|:--------|:----------------|
| title   | "Title"         |
| message | "Input message" |

### Error Alert

By default, "Show error alert after invalid data is entered" is enabled.  To
disable, set the `error` property to `false`

If the `error` property is assigned a value, it is expected to be an object with
the following fields:

| key     | description     |
|:--------|:----------------|
| title   | "Title"         |
| message | "Error message" |
| style   | "Style", expected to be "stop" or "warning" or "info" |

## Conditional Formatting

Conditional Formats are stored in the `!condfmt` array of the worksheet:

```js
ws['!condfmt'] = [
  /* A1:A10 "Format only values that are above average" GT */
  {
    ref: "A1:A10",
    t: "avg",
    op: "GT",
    /* "Light Red Fill with Dark Red Text" */
    s: { color: { rgb: '9C0006' }, bgColor: { rgb: 'FFC7CE' } }
  },
  /* B1:B10 "Format all cells based on their values" (2-Color Scale) */
  {
    ref: "B1:B10",
    t: "scale",
    /* Minimum type Percent value 25 default orange */
    cmin: { v: 25, t: 'percent', color: { rgb: 'FF7128' } },
    /* Maximum type Percent value 75 default yellow */
    cmax: { v: 75, t: 'percent', color: { rgb: 'FFEF9C' } } }
  }
];
```

Each conditional format object in the array follows the schema:

| key         | description                                |
|:------------|:-------------------------------------------|
| `ref`       | Range or cell string or address object     |
| `t`         | Type of Conditional Format (listed below)  |
| `s`         | Differential Style (when relevant)         |
| `op`        | Data operator (when relevant)              |
| `f`         | Formula (when relevant)                    |
| `min/max/v` | Min / Max / Exact values (when relevant)   |
| `color`     | Data Bar Color                             |
| `cmin/cmax` | Threshold Objects (Data Bar / Color Scale) |
| `cmid`      | Middle Threshold (Color Scale)             |
| `thresh`    | Array of Threshold Objects (Icon Set)      |

The `ref` reference can be passed as a string ("A2" or "A2:C4") or address
object like `{r:1, c:0}` or range like `{s:{r:1,c:0}, e:{r:3,c:2}}`.

Conditional Formats also support disparate ranges, for example a CF over the
ranges A2:A4 and C1:D5.  Even though they appear as a comma-separated list in
Excel, they are stored as a *space*-separated list.  The previous example would
have the `ref` key set to `"A2:A4 C1:D5"`

Note: Since CF is read and written when `cellStyles: true` is set, the easiest
way to construct a specific CF is to create a test file using it, parsing using
`read` or `readFile`, and logging the conditional formatting array.

### Conditional Formatting Types

The UI mapping is explained below:

| type      | description                                            | `s` |
|:----------|:-------------------------------------------------------|:----|
| `avg`     | Format only values that are above or below average     | Yes |
| `bar`     | Format all cells based on values: Data Bars            |     |
| `blank`   | Format only cells that contain: Blanks or no Blanks    | Yes |
| `date`    | Format only cells that contain: Dates Occurring        | Yes |
| `dup`     | Format all duplicate values                            | Yes |
| `error`   | Format only cells that contain: Errors or No Errors    | Yes |
| `formula` | Format values where formula is true                    | Yes |
| `icon`    | Format all cells based on values: Icon Sets            |     |
| `rank`    | Format only top or bottom ranked values                | Yes |
| `scale`   | Format all cells based on values: 2- or 3- color scale |     |
| `text`    | Format only cells that contain: Specific Text          | Yes |
| `unique`  | Format all unique values                               | Yes |
| `val`     | Format only cells that contain: Cell Value             | Yes |

### Differential Styles

For the "classic" conditional formatting, the `s` style is interpreted as a
"differential" style, applied on top of the cell's existing style.  This extends
the standard style representation by interpreting `null` items as disabling
features:

```js
{ t: "avg", ref: "A1:A10", op: "GT",
  s: {
    b: true, /* turn on bold */
    left: null /* remove left border */
  }
}
```

The background color is set using `bgColor` (this differs from the normal cell
styles, which uses `fgColor` as the primary background color).

### Conditional Formatting Rules

The explanation is spelled out in accordance with the user interface:

#### "Format all cells based on their values": Color Scales

Type `scale` corresponds to the format styles "(2/3)-Color Scale".

<details>
  <summary><b>Color Scale details</b> (click to show)</summary>

| key    | description                                              |
|:-------|:---------------------------------------------------------|
| `cmin` | Parameters for "Minimum" threshold                       |
| `cmid` | Parameters for "Midpoint" threshold (3-Color Scale only) |
| `cmax` | Parameters for "Maximum" threshold                       |

The Color Scale thresholds include both threshold data and color info:

| thresh key | UI description             |
|:-----------|:---------------------------|
| `t`        | Threshold Type (see below) |
| `v/f`      | Threshold Value            |
| `color`    | Color specification        |

The Color Scale threshold types are described below:

| thresh type  | UI Description | thresholds  | Value |
|:-------------|:---------------|:------------|:------|
| `min`        | Lowest Value   | `cmin` only |       |
| `max`        | Highest Value  | `cmax` only |       |
| `num`        | Number         | all         | `v/f` |
| `percent`    | Percent        | all         | `v/f` |
| `formula`    | Formula        | all         | `f`   |
| `percentile` | Percentile     | all         | `v/f` |

</details>

#### "Format all cells based on their values": Data Bars

Type `bar` corresponds to the format style "Data Bar".

<details>
  <summary><b>Data Bars details</b> (click to show)</summary>

| key     | description                                              |
|:--------|:---------------------------------------------------------|
| `cmin`  | Parameters for "Minimum" threshold                       |
| `cmax`  | Parameters for "Maximum" threshold                       |
| `color` | Color specification                                      |

The Data Bar thresholds include threshold data but no color info:

| thresh key | UI description             |
|:-----------|:---------------------------|
| `t`        | Threshold Type (see below) |
| `v/f`      | Threshold Value            |

The Data Bar threshold types are described below:

| thresh type  | UI Description | thresholds  | Value |
|:-------------|:---------------|:------------|:------|
| `min`        | Lowest Value   | `cmin` only |       |
| `max`        | Highest Value  | `cmax` only |       |
| `num`        | Number         | all         | `v/f` |
| `percent`    | Percent        | all         | `v/f` |
| `formula`    | Formula        | all         | `f`   |
| `percentile` | Percentile     | all         | `v/f` |

</details>

#### "Format all cells based on their values": Icon Sets

Type `icon` corresponds to the format style "Icon Sets".

<details>
  <summary><b>Icon Sets details</b> (click to show)</summary>

| key      | description                |
|:---------|:---------------------------|
| `thresh` | Array of threshold objects |
| `v`      | Icon type (see below)      |
| `hidden` | If true, "Show Icon Only"  |

The icon types and expected number of threshold objects are listed below:

| icon type         | thresh count | UI description              |
|:------------------|:-------------|:----------------------------|
| `3Arrows`         | `3`          | 3 Arrows (Colored)          |
| `3ArrowsGray`     | `3`          | 3 Arrows (Gray)             |
| `3Flags`          | `3`          | 3 Flags                     |
| `3TrafficLights1` | `3`          | 3 Traffic Lights (unrimmed) |
| `3TrafficLights2` | `3`          | 3 Traffic Lights (rimmed)   |
| `3Signs`          | `3`          | 3 Signs                     |
| `3Symbols`        | `3`          | 3 Symbols (Circled)         |
| `3Symbols2`       | `3`          | 3 Symbols (Uncircled)       |
| `3Stars`          | `3` !!       | 3 Stars                     |
| `3Triangles`      | `3` !!       | 3 Triangles                 |
| `4Arrows`         | `4`          | 4 Arrows (Colored)          |
| `4ArrowsGray`     | `4`          | 4 Arrows (Gray)             |
| `4RedToBlack`     | `4`          | Red To Black                |
| `4Rating`         | `4`          | 4 Ratings                   |
| `4TrafficLights`  | `4`          | 4 Traffic Lights            |
| `5Arrows`         | `5`          | 5 Arrows (Colored)          |
| `5ArrowsGray`     | `5`          | 5 Arrows (Gray)             |
| `5Rating`         | `5`          | 5 Ratings                   |
| `5Quarters`       | `5`          | 5 Quarters                  |
| `5Boxes`          | `5` !!       | 5 Boxes                     |

The marked icon types are Excel extensions and may not be supported in older
versions of Excel and other spreadsheet software

The Icon Sets thresholds include threshold data but no color info:

| thresh key | UI description             |
|:-----------|:---------------------------|
| `t`        | Threshold Type (see below) |
| `v/f`      | Threshold Value            |

The Icon Set threshold types are described below:

| thresh type  | UI Description | thresholds  | Value |
|:-------------|:---------------|:------------|:------|
| `num`        | Number         | all         | `v/f` |
| `percent`    | Percent        | all         | `v/f` |
| `formula`    | Formula        | all         | `f`   |
| `percentile` | Percentile     | all         | `v/f` |

The first threshold in an icon set must be `{ t: 'percent', v: 0}` for proper
rendering of the first icon.

For example, the standard 33% / 67% Traffic Light CF is:

```js
{
  ref: ‘E2:E9',
  t: 'icon',
  thresh: [
    { v: 0,  t: 'percent' }, // This must always be 0 %
    { v: 33, t: 'percent' }, // 33 %
    { v: 67, t: 'percent' }  // 67 %
  ],
  v: ‘3TrafficLights1'
}
```

</details>

#### "Format only cells that contain": "Cell Value"

Type `val` corresponds to "Cell Value".

<details>
  <summary><b>Cell Value properties</b> (click to show)</summary>

The differential style `s` option is used.

The `op` key specifies the operator. `min/max/v` and `op` are interpreted in the
same way as the Data Validation operator:

| `op` | Excel interface name       | min | max | v   |
|:-----|:---------------------------|:----|:----|:----|
| `IN` | "between"                  | Yes | Yes |  No |
| `OT` | "not between"              | Yes | Yes |  No |
| `EQ` | "equal to"                 |  No |  No | Yes |
| `NE` | "not equal to"             |  No |  No | Yes |
| `GT` | "greater than"             |  No |  No | Yes |
| `LE` | "less than"                |  No |  No | Yes |
| `GE` | "greater than or equal to" |  No |  No | Yes |
| `LE` | "less than or equal to"    |  No |  No | Yes |

</details>

#### "Format only cells that contain": "Specific Text"

Type `text` corresponds to "Specific Text".

<details>
  <summary><b>Specific Text properties</b> (click to show)</summary>

The differential style `s` option is used.

The `v` key specifies the actual text value and the `op` key specifies operator:

| `op` | UI Definition      | Mnemonic         |
|:-----|:-------------------|:-----------------|
| `IN` | containing         | "in"             |
| `OT` | not containing     | "out" (not "in") |
| `ST` | beginning with     | start of "start" |
| `ND` | ending with        | end of "end"     |

</details>

#### "Format only cells that contain": "Dates Occurring"

Type `date` corresponds to "Dates Occurring".

<details>
  <summary><b>Dates Occurring properties</b> (click to show)</summary>

The differential style `s` option is used.

The `op` key specifies time:

| `op` | UI Definition      |
|:-----|:-------------------|
| `YS` | Yesterday          |
| `TD` | Today              |
| `TM` | Tomorrow           |
| `LS` | In the last 7 days |
| `LW` | Last week          |
| `TW` | This week          |
| `NW` | Next week          |
| `LM` | Last month         |
| `TM` | This month         |
| `NM` | Next month         |

</details>

#### "Format only cells that contain": Errors Blanks "No Errors" "No Blanks"

Type `error` corresponds to "Errors" and "No Errors".

Type `blank` corresponds to "Blanks" and "No Blanks".

<details>
  <summary><b>Errors / Blanks properties</b> (click to show)</summary>

The differential style `s` option is used.

`error` and `blank` types use the `v` key to indicate the type:

|    `v`    | type `error` | type `blank` |
|:---------:|:-------------|:-------------|
|  `true`   | "Errors"     | "Blanks"     |
| (default) | "No Errors"  | "No Blanks"  |

</details>

#### "Format only top or bottom ranked values"

Type `rank` corresponds to this rule type.

<details>
  <summary><b>Ranked Values properties</b> (click to show)</summary>

The differential style `s` option is used.

The `v` key specifies a numeric value (text box in the UI).  The `op` key
specifies rank type:

| `op` | Rank     | Value Type | Interpretation of `v = 20` |
|:-----|:---------|:-----------|:---------------------------|
| `TV` | "Top"    | Value      | Top 20                     |
| `BV` | "Bottom" | Value      | Bottom 20                  |
| `TP` | "Top"    | Percent    | Top 20%                    |
| `BP` | "Bottom" | Percent    | Bottom 20%                 |

</details>

#### "Format only values that are above or below average"

Type `avg` corresponds to this rule type.

<details>
  <summary><b>Average properties</b> (click to show)</summary>

The differential style `s` option is used.

The `op` key specifies operator type:

| `op` | UI Definition      |
|:-----|:-------------------|
| `GT` | "above"            |
| `LT` | "below"            |
| `GE` | "equal or above"   |
| `LE` | "equal or below"   |
| `G1` | "1 std dev above"  |
| `L1` | "1 std dev below"  |
| `G2` | "2 std dev above"  |
| `L2` | "2 std dev below"  |
| `G3` | "3 std dev above"  |
| `L3` | "3 std dev below"  |

</details>

#### "Format only unique or duplicate values"

Type `dup` corresponds to "Format all duplicate values in the selected range"

Type `unique` corresponds to "Format all unique values in the selected range"

The differential style `s` option is used.

#### "Use a formula to determine which cells to format"

Type `formula` corresponds to this rule type.

<details>
  <summary><b>Formula properties</b> (click to show)</summary>

The differential style `s` option is used.

The following keys are used:

| key | description                                           |
|:----|:------------------------------------------------------|
| `f` | Formula string (exactly as entered in UI formula bar) |

</details>

### Conditional Formatting Examples

A common trick for shading alternate rows in a worksheet uses a conditional
format of type `formula` using the formula `MOD(ROW(),2)=0`:

```js
{
  ref: "A1:E200", // change to the worksheet range
  t: "formula",
  f: "MOD(ROW(),2)=1", // no initial =
  s: { bgColor: { rgb: 'ECECEC' } } // light gray background
}
```

## Miscellaneous Worksheet and Workbook Properties

### Freeze Panes

Freeze panes are specified by setting the `!freeze` key of the worksheet object
to a cell reference (string or object) corresponding to the "top-left" cell of
the main pane.  This is the exact cell you would select in Excel before applying
the "Freeze pane" option in the Excel UI.

```js
ws["!freeze"] = "A2"; // Freeze first row               bottom pane starts at A2
ws["!freeze"] = "B1"; // Freeze first column             right pane starts at B1
ws["!freeze"] = "B2"; // Freeze row and column    bottom-right pane starts at B2
```


### Tab Colors

Worksheet tab colors are specified by setting the `!tabcolor` key of the
worksheet object to an RGB color object:

```js
ws["!tabcolor"] = { rgb: "FF0000" }; // Red tab
ws["!tabcolor"] = { rgb: "00FF00" }; // Green tab
ws["!tabcolor"] = { rgb: "0000FF" }; // Blue tab
```

### Gridline Visibility

Worksheet gridlines are stored in the worksheet under the `"!gridlines"` key:

```js
ws["!gridlines"] = true; // enable gridlines
ws["!gridlines"] = false; // disable gridlines
if(ws["!gridlines"] != null && !ws["!gridlines"]) { /* gridlines hidden */ }
```

### Custom XML Items

Custom XML Items are stored on the worksheet in the `CustomXML` array.  Each
Custom XML Item may contain the following properties:

| key      | description                     |
|:---------|:--------------------------------|
| `data`   | Item XML Data                   |
| `props`  | Item Properties XML             |

<details>
  <summary><b>Example code</b> (click to show)</summary>

```js
if(!wb.CustomXML) wb.CustomXML = [];
wb.CustomXML.push({
  data: `<?mso-contentType?>
<FormTemplates xmlns="http://schemas.microsoft.com/sharepoint/v3/contenttype/forms">
  <Display>DocumentLibraryForm</Display>
  <Edit>DocumentLibraryForm</Edit>
  <New>DocumentLibraryForm</New>
</FormTemplates>`,
	props: `<?xml version="1.0" encoding="UTF-8" standalone="no"?>
<ds:datastoreItem ds:itemID="{EFEDDDCB-DC54-437C-8897-1C25DE9113BA}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">
  <ds:schemaRefs>
    <ds:schemaRef ds:uri="http://schemas.microsoft.com/sharepoint/v3/contenttype/forms"/>
  </ds:schemaRefs>
</ds:datastoreItem>`
});
```

</details>


## Tables

Tables are specified as objects in the `!tables` array of the worksheet.

### Basics

The `ref` key of the Table object specifies the range.  It can be an A1-style
string like `A1:B4` or a range object.  If the range is not specified, the
entire worksheet will be considered the table.

The `name` key of the Table object specifies the name.  A default autogenerated
name will be used if omitted.

The simplest table is therefore:

```js
if(!ws["!tables"]) ws["!tables"] = [];
ws["!tables"].push({}); // spans entire worksheet
```

### Table Object Properties

| key      | description                                                  |
|:---------|:-------------------------------------------------------------|
| `ref`    | range of table                                               |
| `name`   | table name                                                   |
| `header` | "This table has headers" -- set to `0` or `false` to disable |
| `style`  | style info (see "Styling" below)                             |


### Headers

The `header` key maps to the Excel "This Table has headers" property.  If it is
set to `0` or `false`, headers and AutoFilter structures are not created.

If `header` is not set to `0` or `false`, the first row is assumed to be the
header row.  The field names are read from the cell values in the first row.
If two cells in the header row have the same value, an informative error will
tell you which cell needs to be changed (and how Excel would dedupe the name)!


### Styling

Basic style properties of the Table are stored in the `style` object:

| key          | description         |
|:-------------|:--------------------|
| `name`       | Name of Table Style |
| `rowstripe`  | Show Row Stripes    |
| `colstripe`  | Show Column Stripes |

The name of a table style takes the form `<type><id>` where `type` is `"Light"`,
`"Medium"` or `"Dark"`.  `id` is a number starting from `1`, displayed in the
Tooltip text when hovering over a style in the "Format as Table" popup.  There
are 21 Light styles, 28 Medium styles, and 11 Dark styles.

The default options use the style `Medium9` with row stripes enabled and column
stripes disabled.


### Basic Example

```js
var ws = XLSX.utils.aoa_to_sheet([
  ["Item", "Price"]
]);
XLSX.utils.sheet_add_json(ws, [
  { Item: "abc", Price: 1.23 },
  { Item: "def", Price: 4.56 },
], {origin: -1, skipHeader: true});
XLSX.utils.sheet_add_json(ws, [
  { Item: "ghi", Price: 7.89 },
], {origin: -1, skipHeader: true});

ws["!tables"] = [
  {
    "name": "MyTable",
    "ref": "A1:B4",
    "style": {
      "name": "Medium1",
      "rowstripe": true,
      "colstripe": false
    }
  }
];
```


## DOM Table Ingress

The `table_to_book` and `table_to_sheet` utility functions read CSS properties
from the specified TABLE element.

Style properties are read from the container TD cell by default:

```html
<td style="font-weight: bold">This is bold</td>
```

If the TD cell contains a single SPAN with text, text properties from the span
are interpreted as text properties for the cell:

```html
<td><span style="font-weight: 700">The cell will be bold</span></td>
```

For optimal ingress, the following styles are recommended:

```css
/* Excel draws the grid as if adjacent cells share border */
TABLE { border-collapse: collapse; }

/* table_to_book / table_to_sheet check width/height attributes */
* { box-sizing: border-box; }

/* explicitly set cell vertical-align to middle or top or bottom */
TH, TD { vertical-align: middle; }
```


### Cell Properties and Rich Text

If a cell contains multiple nodes, it will be interpreted as a rich text run:

> The cell is normal, but **this fragment is bold**

```html
<td>The cell is normal, but <b>this fragment is bold</b></td>
```

Cell properties like background color will still be read from the TD element:

```html
<td bgcolor="red">This cell will have red background <b>even in the run</b></td>
```

### Rich Text CSS Styles

Excel supports a limited subset of CSS styles in rich text:

| Excel feature  | CSS Style                            |
|:---------------|:-------------------------------------|
| Bold           | `font-weight`: `bold` or `700`       |
| Italic         | `font-style`: `italic`               |
| Underline      | `text-decoration`: `underline`       |
| Strike-through | `text-decoration`: `line-through`    |
| Text Color     | `color`: CSS Level 1 name or RGB(A)  |
| Font Name      | `font-family`: first entry           |
| Font Size      | `font-size` with unit (`px` or `pt`) |
| Subscript      | `vertical-align`: `sub`              |
| Superscript    | `vertical-align`: `super`            |

Since CSS only supports one `vertical-align` style for an element, the best way
to specify subscript/superscript with vertical alignment is to use a span:

```html
<td style="vertical-align: middle">
  <span style="vertical-align: sub; font-size: .83em">
    This text will be subscripted in a cell that is middle-aligned
</span></td>
```

Excel automatically applies font scaling to subscript and superscript, so both
should be specified with `font-size: .83em` to ensure proper sizing.

### Cell-Level CSS Styles

Other cell features are only supported at the cell or row level

| Excel feature        | CSS TD Style                                    |
|:---------------------|:------------------------------------------------|
| Horizontal alignment | `text-align`: `left` or `center` or `right`     |
| Vertical alignment   | `vertical-align`: `top` or `middle` or `bottom` |
| Background Fill      | `background-color`                              |
| Column width         | `width` for non-merge cells                     |
| Row height           | `height` of containing TR element               |

The Wrap Text property is automatically enabled when a newline is detected.
Normally `<br/>` tags are used, but if `white-space` is set to `pre` or any
related setting then newline characters are correctly interpreted.

### Borders

By default, the table borders are ignored.  Passing the option `borders: true`
to `table_to_book` or `table_to_sheet` will read the border styles:

```js
/* default ignores borders */
var ws_no_borders = XLSX.utils.table_to_sheet(table_elt);

/* pass borders: true in the options object to read border styles */
var ws_with_borders = XLSX.utils.table_to_sheet(table_elt, { borders: true });
```

### Dates

V8 (used by Chrome and NodeJS) and other JavaScript engines are extremely
aggressive in parsing date strings.  For example:

```js
/* V8 thinks this is Mon Jan 01 2001 */
new Date("This is not a date 1");
```

The text values in each cell is tested for date feasibility through the engine's
default `Date.parse` mechanism.  This is not always correct.  To specifically
suppress the Date parsing, set `rawDates: true`:

```js
var ws_no_dates = XLSX.utils.table_to_sheet(table_elt, { rawDates: true });
```

### ARIA Compliance

Nodes with the `aria-hidden="true"` attribute are automatically discarded.  All
font icons using `<I>` tags should be marked with the attribute!

### Writing to a specific location in the Worksheet

The `origin` option specifies the starting cell.  It is expected to be one of:

| `origin`         | Description                                               |
| :--------------- | :-------------------------------------------------------- |
| (cell object)    | Use specified cell (cell object)                          |
| (string)         | Use specified cell (A1-style cell)                        |
| (number >= 0)    | Start from the first column at specified row (0-indexed)  |
| (default)        | Start from cell A1                                        |

For example, to start writing the table starting at cell A3:

```
var ws_A3 = XLSX.utils.table_to_sheet(table_elt, { origin: "A3" });
```

### Adding to an existing worksheet

The `sheet_add_dom` utility function accepts three arguments: worksheet object,
DOM element and options argument.

By default, the DOM table is written starting in cell A1.  The `origin` option
specifies the starting cell, and additionally supports the value `-1` to append
the table to the existing sheet.

A small helper function can create gap rows between tables:

```js

function create_gap_rows(ws, nrows) {
  var ref = XLSX.utils.decode_range(ws["!ref"]);       // get original range
  ref.e.r += nrows;                                    // add to ending row
  ws["!ref"] = XLSX.utils.encode_range(ref);           // reassign row
}

/* first table */
var ws = XLSX.utils.table_to_sheet(document.getElementById('table1'));
create_gap_rows(ws, 1); // one row gap after first table

/* second table */
XLSX.utils.sheet_add_dom(ws, document.getElementById('table2'), {origin: -1});
create_gap_rows(ws, 3); // three rows gap after second table

/* third table */
XLSX.utils.sheet_add_dom(ws, document.getElementById('table3'), {origin: -1});
```


## Simple Examples


### File from Scratch

The most common use case involves cleaning up an export from a data store:

```js
var XLSX = require("@sheet/<replace with your ID>");

/* Build up a worksheet from your data */
var ws = XLSX.utils.aoa_to_sheet([
  ["Item", "Price"]
]);
XLSX.utils.sheet_add_json(ws, [
  { Item: "abc", Price: 1.23 },
  { Item: "def", Price: 4.56 },
]);

/* Bold the headers */
XLSX.utils.sheet_set_range_style(ws, "A1:B1", {
  bold: true
});

/* Set the format for the visible cells in the price column */
var range = XLSX.utils.decode_range(ws['!ref']);
range.s.c = 1; // start from second col
range.e.c = 1; // end on the second col
XLSX.utils.sheet_set_range_style(ws, range, {
  z: "0.00" // decimal with 2 places
});

/* Write File */
var wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, "Test");
XLSX.writeFile(wb, "e1.xlsx", {cellStyles: true});
```

### Spicing up CSV Exports

Another common use case involves cleaning up a data export from another process:

```js
var XLSX = require("@sheet/<replace with your ID>"); // node or webpack

/* Read the file */
var wb = XLSX.readFile("e2.csv");
var ws = wb.Sheets[wb.SheetNames[0]];

/* Find the header row */
var range = XLSX.utils.decode_range(ws['!ref']);
range.s.r = 0; range.e.r = 0; // restrict to the first row

/* Bold the headers */
XLSX.utils.sheet_set_range_style(ws, range, {
  bold: true
});

/* Freeze first row */
ws["!freeze"] = "A2";

/* Write to XLSX */
XLSX.writeFile(wb, "e2.xlsx", {cellStyles: true});
```

### Modifying Existing Files

Read in the file with `cellStyles:true` to initialize the styles data:

```js
var XLSX = require("@sheet/<replace with your ID>");

/* read the file from the first export */
var wb = XLSX.readFile("e1.xlsx", {cellStyles:true});
/* get the first worksheet */
var ws = wb.Sheets[wb.SheetNames[0]];
/* turn off bold */
XLSX.utils.sheet_set_range_style(ws, "A1:B1", {bold: false});

/* Write to XLSX */
XLSX.writeFile(wb, "e3.xlsx", {cellStyles: true});
```

### Exporting an HTML TABLE on the DOM

`table_to_book` and `table_to_sheet` automatically process CSS styles:

```js
var dom_elt = document.getElementById('data-table');
var wb = XLSX.utils.table_to_book(dom_elt);
XLSX.writeFile(wb, "export.xlsx", {cellStyles:true});
```

This technique works in JSDOM for server-side applications:

<details>
  <summary><b>Code and caveats</b> (click to show)</summary>

```js
const fs = require("fs");
const { JSDOM } = require('jsdom');

/* load HTML into the DOM */
const dom = new JSDOM(fs.readFileSync("table.html").toString());
const elt = dom.window.document.querySelector("#table"); // use your table id

/* to detect styles, getComputedStyle has to be visible */
getComputedStyle = dom.window.getComputedStyle;

/* generate workbook using table_to_book and export */
const wb = XLSX.utils.table_to_book(elt);
XLSX.writeFile(wb, "tablexport.xlsx", {cellStyles:true});
```

There are known issues with JSDOM and inheritance in `getComputedStyle`, so deep
nested styles should be avoided when possible.  Prefer explicit styling:

```html
<!-- works in browsers but problematic in JSDOM -->
<td>normal <b>B <i>BI</i></b></td>

<!-- works in browsers and JSDOM-->
<style>
/* vertical alignment */
TD { vertical-align: middle; }

/* styling */
.bold { font-weight: bold; }
.italic { font-style: italic; }
</style>

<td>normal <span class="bold">B</span> <span class="bold italic">BI</span></td>
```

</details>


## Worksheet and Workbook Protection

There are various forms of protection.  Some are informational, and others
require cryptography.  **This build does not include crypto features.**

### "Mark as Final"

To mark a workbook as final, set the workbook custom property `_MarkAsFinal`:

```js
if(!wb.Custprops) wb.Custprops = {};
wb.Custprops._MarkAsFinal = true; // set Mark as Final
```

If a file is not marked as final, the parsed workbook object will not have the
property set.  The `Custprops` property should also be tested for existence:

```js
if(wb.Custprops && wb.Custprops._MarkAsFinal) { /* file is marked as final */ }
```


### "Informational" Password Protection

Other forms of password protection, including "Password to Modify", are optional
insofar as the application is expected to honor the protection but nothing stops
a program from bypassing the protection.

#### Workbook Properties Protection (Protect Workbook)

The Workbook Properties Protection are found in the "Review" Ribbon tab or under
Tools ... Protect ... Protect Workbook in the menus.  They are stored in the
workbook under `.Workbook.Protection`:

| key             | description                                                |
|:----------------|:-----------------------------------------------------------|
| `lockStructure` | Sheets cannot be moved, deleted, (un)hidden, or renamed.   |
| `lockWindows`   | Windows are the same size and in the same position         |

```js
if(!wb.Workbook) wb.Workbook = {};
wb.Workbook.Protection = { lockStructure: true };
```

#### Worksheet Properties Protection (Protect Worksheet)

The Worksheet Properties Protection are found in the "Review" Ribbon tab or in
Tools ... Protect ... Protect Worksheet in the menus.  They are stored in the
worksheet under the `!protect` key:

<details>
  <summary><b>Worksheet Protection Details</b> (click to show)</summary>

| key                   | feature (true=disabled / false=enabled) | default    |
|:----------------------|:----------------------------------------|:-----------|
| `selectLockedCells`   | Select locked cells                     | enabled    |
| `selectUnlockedCells` | Select unlocked cells                   | enabled    |
| `formatCells`         | Format cells                            | disabled   |
| `formatColumns`       | Format columns                          | disabled   |
| `formatRows`          | Format rows                             | disabled   |
| `insertColumns`       | Insert columns                          | disabled   |
| `insertRows`          | Insert rows                             | disabled   |
| `insertHyperlinks`    | Insert hyperlinks                       | disabled   |
| `deleteColumns`       | Delete columns                          | disabled   |
| `deleteRows`          | Delete rows                             | disabled   |
| `sort`                | Sort                                    | disabled   |
| `autoFilter`          | Filter                                  | disabled   |
| `pivotTables`         | Use PivotTable reports                  | disabled   |
| `objects`             | Edit objects                            | enabled    |
| `scenarios`           | Edit scenarios                          | enabled    |

</details>

```js
if(!ws["!protect"]) ws["!protect"] = {};
ws["!protect"].selectLockedCells = false;
ws["!protect"].encryption = XLSX.utils.hash_password("sheetjs");
```



## Embedded Images

Images are stored in the `!images` array within the worksheet object.  The Image
support includes different strategies for working with image data and location.

### Importing and Exporting with Images

The `bookImages` option should be passed to the `read` and `write` functions:

```js
var wb = XLSX.read(buffer, {type: "buffer", cellStyles: true, bookImages: true});
XLSX.writeFile(wb, "noimage.xlsx", {cellStyles: true, bookImages: true});
```

### Basic Image Properties

The `!images` key of a worksheet object should be an array of image objects. The
image objects have the following properties:

| key         | description                     |
|:------------|:--------------------------------|
| `!pos`      | Position in worksheet           |
| `!data`     | Image Data                      |
| `!datatype` | Type of Data                    |
| `!link`     | Image Link (optional)           |
| `l`         | Hyperlink (optional)            |

### Image Data and Links

The `!datatype` field controls how the `!data` field is interpreted:

| `!datatype` | interpretation of `!data`                   |
|:------------|:--------------------------------------------|
| `"binary"`  | Binary string                               |
| `"base64"`  | Base64 string or Data URI                   |
| `"buffer"`  | NodeJS Buffer or `Uint8Array`               |
| `"remote"`  | Ignored (`!link` specifies remote address) |

#### Inserting Image Files

The "binary', "base64", and "buffer" types specify that an actual image file
should be added.  The specific type depends on how the data is acquired:

<details>
  <summary><b>NodeJS</b> (click to show)</summary>

Most standard NodeJS APIs like `fs.readFileSync` work with `Buffer` objects:

```js
var imdata = fs.readFileSync("logo.png");

ws["!images"].push({
  "!pos": { x: 100, y: 100, w: 300, h: 300 },
  "!datatype": "buffer",
  "!data": imdata
});
```

The popular [`image-size`](https://www.npmjs.com/package/image-size) module can
read the dimensions from an image directly, so maintaining aspect ratio is easy:

```js
var sizeOf = require('image-size');
var scale = 0.5; // replace this with 1 for real size, 2 for 2x size

var imdata = fs.readFileSync("logo.png");
var imdim = sizeOf(imdata);

ws["!images"].push({
  "!pos": { x: 100, y: 100, w: scale * imdim.width, h: scale * imdim.height },
  "!datatype": "buffer",
  "!data": imdata
});
```

</details>

<details>
  <summary><b>Downloading Static Image on Demand</b> (click to show)</summary>

Using `XMLHttpRequest`, the `"arraybuffer"` response type tells the browser to
return an `ArrayBuffer` which can be converted to a `Uint8Array` and stored with
type `"bufffer"`:

```js
/* send synchronous request */
var req = new XMLHttpRequest();
req.open("GET", "logo.png", false);
req.responseType = "arraybuffer";
req.send();

/* req.response is an ArrayBuffer with the data */
var data = new Uint8Array(req.response);
ws["!images"].push({
  "!pos": { x: 100, y: 100, w: 300, h: 300 },
  "!datatype": "buffer",
  "!data": imdata
});
```

</details>

<details>
  <summary><b>Angular, React, and other Frameworks</b> (click to show)</summary>

Most sites using a web framework also use a build tool like webpack to generate
the final application.  The bundlers also manage static assets, usually making
data URIs.  That is supported in the `"base64"` type:

```js
var imdata = require('../assets/pie-64x64.png'); // "data:image/png:base64,..."
ws['!images'].push({
  '!pos': {x: 700, y: 300, w: 64, h: 64},
  '!datatype': 'base64',
  '!data': imdata
});
```

</details>

To replicate the "Insert and Link" behavior, the `"!link"` key should be set to
the URL where the image is located:

<details>
  <summary><b>Example</b> (click to show)</summary>

```js
ws["!images"].push({
  "!link": "http://sheetjs.com/logo.png",
  "!pos": { x: 100, y: 100, w: 300, h: 300 },
  "!datatype": "...",
  "!data": imdata
});
```

</details>

#### Links

Excel supports pure links ("Link to File"), marked with type `"remote"`.  The
`!data` field is not checked and expected to be empty.  The actual link should
be specified in the `!link` field.  This *does not* automatically create a
hyperlink -- see the "Hyperlinks" section below for more info.

<details>
  <summary><b>Example</b> (click to show)</summary>

```js
ws["!images"].push({
  "!link": "http://sheetjs.com/logo.png",
  "!pos": { x: 100, y: 100, w: 300, h: 300 },
  "!datatype": "remote"
});
```

</details>

### Image Dimensions and Location

Image position metadata is stored in the `!pos` key.  Excel recognizes three
types of position specifications and will automatically scale images to fit.

When reading from a file, `!pos` is normalized to "Absolute Position with Size".
The `!relpos` key of the object will hold the "Relative Position of Corners"
parameters.  When writing to a file, the `!pos` object will be read and its type
will be inferred from the keys in the object.

#### Absolute Position with Size

The simplest and most straightforward position is absolute, where dimensions and
starting coordinate are specified in pixels.

| key | interpretation                             |
|:----|:-------------------------------------------|
| `x` | X-coordinate of upper-left corner (pixels) |
| `y` | Y-coordinate of upper-left corner (pixels) |
| `w` | width (pixels)                             |
| `h` | height (pixels)                            |

<details>
  <summary><b>Example</b> (click to show)</summary>

The following example specifies an image of 456 x 123 pixels starting from the
pixel position (0, 0):

```js
image["!pos"] = {
  /* location of upper-left corner (in pixels) */
  x: 0, y: 0,
  /* object size (in pixels) -- Excel will scale to fit dimensions */
  w: 456, h: 123,
};
```

</details>

#### Relative Position with Size

It is also possible to specify the upper-left corner position as a pixel offset
relative to the upper-left corner an arbitrary cell.  The "Absolute Position and
Size" is a special case starting from cell A1

| key | interpretation                                          |
|:----|:--------------------------------------------------------|
| `r` | row of upper-left corner (0-indexed)                    |
| `c` | col of upper-left corner (0-indexed)                    |
| `x` | X distance from upper-left corner of base cell (pixels) |
| `y` | Y distance from upper-left corner of base cell (pixels) |
| `w` | width (pixels)                                          |
| `h` | height (pixels)                                         |

<details>
  <summary><b>Example</b> (click to show)</summary>

The following example specifies an image of 456 x 123 pixels starting from the
point 50 pixels below and 100 pixels to the right of the upper-left corner of
cell B3:

```js
image["!pos"] = {
  /* cell address for upper-left corner B3 */
  c: 1, r: 2,
  /* offset relative to upper-left corner (in pixels) */
  x: 100, y: 50,
  /* object size (in pixels) -- Excel will scale to fit dimensions */
  w: 456, h: 123,
};
```

</details>

#### Relative Position of Corners

It is also possible to specify the upper-left corner position and lower-right
corner position as pixel offsets relative to arbitrary cells.  Excel will
determine the size based on the dynamic distance.

The keys for the upper-left corner are in lowercase, while the keys for the
lower-right corner are in uppercase:

| key | interpretation                                           |
|:----|:---------------------------------------------------------|
| `r` | row of upper-left corner (0-indexed)                     |
| `c` | col of upper-left corner (0-indexed)                     |
| `x` | X distance from upper-left corner of base cell (pixels)  |
| `y` | Y distance from upper-left corner of base cell (pixels)  |
| `R` | row of lower-right corner (0-indexed)                    |
| `C` | col of lower-right corner (0-indexed)                    |
| `X` | X distance from lower-right corner of base cell (pixels) |
| `Y` | Y distance from lower-right corner of base cell (pixels) |

<details>
  <summary><b>Example</b> (click to show)</summary>

The following example specifies an image covering over the cell B5.

- The address of the upper-left corner is `r = 4, c = 1`
- The upper-left pixel offset will be 0 in both directions.
- The address of the lower-right corner is `R = r + 1 = 5, C = c + 1 = 2`
- The upper-left pixel offset will be 0 in both directions.


```js
var addr = { r: 4, c: 1 };
var rowspan = 1, colspan = 1;
image["!pos"] = {
  /* upper-left corner cell address*/
  c: addr.c, r: addr.r,
  /* lower-right corner cell address*/
  C: addr.c + colspan, R: addr.r + rowspan,
  /* zero pixel offset for both points */
  x: 0, y: 0, X: 0, Y: 0
};
```

</details>

### Hyperlinks

Hyperlinks are independent of Image Links. The Hyperlink objects are stored in
the `l` key of the image (just like cell hyperlinks) and follow the same schema:

| key      | interpretation             |
|:---------|:---------------------------|
| `Target` | Target URL or Cell Address |

As with Cell Hyperlinks, Image links with a leading `#` are internal.

<details>
  <summary><b>Example</b> (click to show)</summary>

For example, the following snippet inserts <http://sheetjs.com/logo.png>.
Clicking the image opens a web browser window with <http://sheetjs.com>:

```js
ws["!images"].push({
  "!link": "http://sheetjs.com/logo.png",
  "!pos": { x: 100, y: 100, w: 300, h: 300 },
  "!datatype": "remote",
  l: { Target: "http://sheetjs.com" }
});
```

</details>

### Header and Footer Images

Images are stored in the `images` key of the print object:

<details open="open">
  <summary><b>Details</b> (click to show)</summary>

```js
ws["!print"].images = {
  header: {                                      // header
    odd: {                                       // +- odd pages
      left: [                                    // +--- left odd header (array)
        { /* ... image object ... */ }           // +----- image objects
      ],
      right: [/*... array of images ... */],     // +--- right odd header
      center: [/*... array of images ... */],    // +--- center odd header
    },
    even: { /* ... same format as odd ... */ },  // +- even pages
    first: { /* ... same format as odd ... */ }, // +- first page
  },

  footer: {/* ... same format as header ...*/}   // footer
}
```

The image object takes the same form as the standard image object.  However,
the absolute size must be specified.

**Image Type**

Images must be positioned as absolute!  The cell anchors are not defined in the
header and footer.

The `"remote"` type is not supported in Excel headers and footers -- the actual
file is required!

**Relative Positioning within the Header or Footer section**

The header/footer specifier `&G` marks where the image goes:

```js
/* Center Section order: "pre" <image> "post" */
ws["!print"].header.odd = "&Cpre&Gpost";
```

**Example**

The following example is appropriate for NodeJS:

```js
var fs = require("fs")
var sizeOf = require("image-size");
var imdata = fs.readFileSync("logo.png");
var imdim = sizeOf(imdata);

/* set footer even right to include the image */
ws["!print"] = {
  footer: {
    even: {
      right: "&G"
    },
  },

  images: {
    footer: {
      even: {
        right: [
          {
            "!pos": {x: 0, y: 0, w: imdim.width, h: imdim.height },
            "!datatype": "buffer",
            "!data": imdata
          }
        ]
      }
    }
  }
};
```

</details>

### Examples

- From NodeJS, images can be read with `fs.readFileSync` and added:

<details>
  <summary><b>Example Code</b> (click to show)</summary>

```js
/* read image data */
var imdata = fs.readFileSync("test.jpg");

/* create worksheet and workbook */
var wb = XLSX.utils.book_new();
var ws = XLSX.utils.aoa_to_sheet([[]]);
XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

/* initialize worksheet image array */
if(!ws["!images"]) ws["!images"] = [];

/* add image */
ws["!images"].push({
  "!pos": { x: 100, y: 100, w: 300, h: 300 },
  "!datatype": "buffer",
  "!data": imdata
});

/* write file */
XLSX.writeFile(wb, "out.xlsx", {cellStyles:true, bookImages:true});
```

</details>

- Link-only references can be added without the image data:

<details>
  <summary><b>Example Code</b> (click to show)</summary>

```js
/* create worksheet and workbook */
var wb = XLSX.utils.book_new();
var ws = XLSX.utils.aoa_to_sheet([[]]);
XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

/* initialize worksheet image array */
if(!ws["!images"]) ws["!images"] = [];

/* add image */
ws["!images"].push({
  "!pos": { x: 100, y: 100, w: 300, h: 300 },
  "!datatype": "remote",
  "!link": "http://sheetjs.com/logo.png"
});

/* write file */
XLSX.writeFile(wb, "out.xlsx", {cellStyles:true, bookImages:true});
```

</details>

## Shapes

Shapes are stored in the `!shapes` array within the worksheet object.

### Basic Shape Properties

The `!shapes` key of a worksheet object should be an array of shape objects. The
shape objects have the following properties:

| key         | description                     |
|:------------|:--------------------------------|
| `!pos`      | Position in worksheet           |
| `v`         | Plaintext value                 |
| `!shape`    | Type of shape                   |
| `s`         | Style Properties                |
| `R`         | Rich text                       |

The following shape types are supported:

| `!shape`                   | description                             |
|:---------------------------|:----------------------------------------|
| `rect`                     | Rectangle                               |
| `ellipse`                  | Ellipse                                 |
| `arc`                      | Curved Arc                              |

The following shape types only work on Excel for Windows:

| `!shape`                   | description                             |
|:---------------------------|:----------------------------------------|
| `accentBorderCallout1`     | Callout 1 with Border and Accent        |
| `accentBorderCallout2`     | Callout 2 with Border and Accent        |
| `accentBorderCallout3`     | Callout 3 with Border and Accent        |
| `accentCallout1`           | Callout 1                               |
| `accentCallout2`           | Callout 2                               |
| `accentCallout3`           | Callout 3 Shape                         |
| `actionButtonBackPrevious` | Back or Previous Button                 |
| `actionButtonBeginning`    | Beginning Button                        |
| `actionButtonBlank`        | Blank Button                            |
| `actionButtonDocument`     | Document Button                         |
| `actionButtonEnd`          | End Button                              |
| `actionButtonForwardNext`  | Forward or Next Button                  |
| `actionButtonHelp`         | Help Button                             |
| `actionButtonHome`         | Home Button                             |
| `actionButtonInformation`  | Info Button                             |
| `actionButtonMovie`        | Movie Button                            |
| `actionButtonReturn`       | Return Button                           |
| `actionButtonSound`        | Sound Button                            |
| `bentArrow`                | Bent Arrow                              |
| `bentConnector2`           | Bent Connector (2 segments)             |
| `bentConnector3`           | Bent Connector (3 segments)             |
| `bentConnector4`           | Bent Connector (4 segments)             |
| `bentConnector5`           | Bent Connector (5 segments)             |
| `bentUpArrow`              | Bent Up Arrow                           |
| `bevel`                    | Bevel                                   |
| `blockArc`                 | Block Arc                               |
| `borderCallout1`           | Callout 1 with Border                   |
| `borderCallout2`           | Callout 2 with Border                   |
| `borderCallout3`           | Callout 3 with Border                   |
| `bracePair`                | Pair of Braces                          |
| `bracketPair`              | Pair of Brackets                        |
| `callout1`                 | Callout 1                               |
| `callout2`                 | Callout 2                               |
| `callout3`                 | Callout 3                               |
| `can`                      | Can                                     |
| `chartPlus`                | Chart with centered gridlines           |
| `chartStar`                | Chart with centered isometric gridlines |
| `chartX`                   | Chart with centered diagonal gridlines  |
| `chevron`                  | Chevron                                 |
| `chord`                    | Circle Chord                            |
| `circularArrow`            | Circular Arrow                          |
| `cloud`                    | Cloud Shape                             |
| `cloudCallout`             | Callout Cloud                           |
| `corner`                   | Corner                                  |
| `cornerTabs`               | Corner Tabs                             |


