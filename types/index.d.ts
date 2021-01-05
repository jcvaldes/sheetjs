/* index.d.ts (C) 2015-present SheetJS */
// TypeScript Version: 2.2
import * as CFB from "./cfb";
import * as SSF from "./ssf";

/** Version string */
export const version: string;

/** SSF Formatter Library */
export { SSF };

/** CFB Library */
export { CFB };

/** NODE ONLY! Attempts to read filename and parse */
export function readFile(filename: string, opts?: ParsingOptions): WorkBook;
/** Attempts to parse data */
export function read(data: any, opts?: ParsingOptions): WorkBook;
/** Attempts to write or download workbook data to file */
export function writeFile(data: WorkBook, filename: string, opts?: WritingOptions): any;
/** Attempts to write the workbook data */
export function write(data: WorkBook, opts?: WritingOptions): any;

/** Change date-number handling style */
export function set_date_style(style: number): void;

/** Utility Functions */
export const utils: XLSX$Utils;
/** Stream Utility Functions */
export const stream: StreamUtils;

/** Number Format (either a string or an index to the format table) */
export type NumberFormat = string | number;

/** Worksheet specifier (string, number, worksheet) */
export type WSSpec = string | number | WorkSheet;

/** Range specifier (string or range or cell), single-cell lifted to range */
export type RangeSpec = string | Range | CellAddress;

/** Cell Address Specifier (string or cell) */
export type CellSpec = string | CellAddress;

/** Basic File Properties */
export interface Properties {
    /** Summary tab "Title" */
    Title?: string;
    /** Summary tab "Subject" */
    Subject?: string;
    /** Summary tab "Author" */
    Author?: string;
    /** Summary tab "Manager" */
    Manager?: string;
    /** Summary tab "Company" */
    Company?: string;
    /** Summary tab "Category" */
    Category?: string;
    /** Summary tab "Keywords" */
    Keywords?: string;
    /** Summary tab "Comments" */
    Comments?: string;
    /** Statistics tab "Last saved by" */
    LastAuthor?: string;
    /** Statistics tab "Created" */
    CreatedDate?: Date;
}

/** Other supported properties */
export interface FullProperties extends Properties {
    ModifiedDate?: Date;
    Application?: string;
    AppVersion?: string;
    DocSecurity?: string;
    HyperlinksChanged?: boolean;
    SharedDoc?: boolean;
    LinksUpToDate?: boolean;
    ScaleCrop?: boolean;
    Worksheets?: number;
    SheetNames?: string[];
    ContentStatus?: string;
    LastPrinted?: string;
    Revision?: string | number;
    Version?: string;
    Identifier?: string;
    Language?: string;
}

export interface CommonOptions {
    /**
     * If true, throw errors when features are not understood
     * @default false
     */
    WTF?: boolean;

    /**
     * When reading a file with VBA macros, expose CFB blob to `vbaraw` field
     * When writing BIFF8/XLSB/XLSM, reseat `vbaraw` and export to file
     * @default false
     */
    bookVBA?: boolean;

    /**
     * When reading a file, store dates as type d (default is n)
     * When writing XLSX/XLSM file, use native date (default uses date codes)
     * @default false
     */
    cellDates?: boolean;

    /**
     * Create cell objects for stub cells
     * @default false
     */
    sheetStubs?: boolean;

    /**
     * When reading a file, save style/theme info to the .s field
     * When writing a file, export style/theme info
     * @default false
     */
    cellStyles?: boolean;

    /**
     * If defined and file is encrypted, use password
     * @default ''
     */
    password?: string;

    /**
     * Text scaling (points per inch)
     * - 72  = Excel for Mac
     * - 96  = Excel for Windows 100% Text Scaling
     * - 120 = Excel for Windows 100% Text Scaling
     * - 144 = Excel for Windows 100% Text Scaling
     */
    PPI?: 72 | 96 | 120 | 144;

    /**
     * (ONLY FOR SUPPORTED BUILDS)
     * When reading a file, load images
     * When writing a file, save images
     */
    bookImages?: boolean;

    /**
     * (ONLY FOR SUPPORTED BUILDS)
     * When reading a file, load template structure
     * When writing a file, write template structure
     */
    template?: boolean;
}

export interface DateNFOption {
    /** Use specified date format */
    dateNF?: NumberFormat;
}

/** Options for read and readFile */
export interface ParsingOptions extends CommonOptions {
    /** Input data encoding */
    type?: 'base64' | 'binary' | 'buffer' | 'file' | 'array' | 'string';

    /** Default codepage */
    codepage?: number;

    /**
     * Save formulae to the .f field
     * @default true
     */
    cellFormula?: boolean;

    /**
     * Parse rich text and save HTML to the .h field
     * @default true
     */
    cellHTML?: boolean;

    /**
     * Save number format string to the .z field
     * @default false
     */
    cellNF?: boolean;

    /**
     * Generate formatted text to the .w field
     * @default true
     */
    cellText?: boolean;

    /** Override default date format (code 14) */
    dateNF?: string;

    /**
     * If >0, read the first sheetRows rows
     * @default 0
     */
    sheetRows?: number;

    /**
     * If true, parse calculation chains
     * @default false
     */
    bookDeps?: boolean;

    /**
     * If true, add raw files to book object
     * @default false
     */
    bookFiles?: boolean;

    /**
     * If true, only parse enough to get book metadata
     * @default false
     */
    bookProps?: boolean;

    /**
     * If true, only parse enough to get the sheet names
     * @default false
     */
    bookSheets?: boolean;

    /** If specified, only parse the specified sheets or sheet names */
    sheets?: number | string | Array<number | string>;

    /** If true, plaintext parsing will not parse values */
    raw?: boolean;

    /** If true, ignore "dimensions" records (some benefit for XLSX files) */
    nodim?: boolean;

    /** If true, preserve _xlfn. prefixes in formula function names */
    xlfn?: boolean;

    dense?: boolean;
}

/** Options for write and writeFile */
export interface WritingOptions extends CommonOptions {
    /** Output data encoding */
    type?: 'base64' | 'binary' | 'buffer' | 'file' | 'array' | 'string';

    /**
     * Generate Shared String Table
     * @default false
     */
    bookSST?: boolean;

    /**
     * File format of generated workbook
     * @default 'xlsx'
     */
    bookType?: BookType;

    /**
     * Name of Worksheet (for single-sheet formats)
     * @default ''
     */
    sheet?: string;

    /**
     * Use ZIP compression for ZIP-based formats
     * @default false
     */
    compression?: boolean;

    /**
     * Suppress "number stored as text" errors in generated files
     * @default true
     */
    ignoreEC?: boolean;

    /** Override theme XML (XLSX/XLSB/XLSM export) */
    themeXLSX?: string;

    /** Override workbook properties on save */
    Props?: Properties;
}

/** Workbook Object */
export interface WorkBook {
    /**
     * A dictionary of the worksheets in the workbook.
     * Use SheetNames to reference these.
     */
    Sheets: { [sheet: string]: WorkSheet };

    /** Ordered list of the sheet names in the workbook */
    SheetNames: string[];

    /** Standard workbook Properties */
    Props?: FullProperties;

    /** Custom workbook Properties */
    Custprops?: object;

    /** Custom XML */
    CustomXML?: CustomXMLItem[];

    Workbook?: WBProps;

    ExternalWB?: WorkBook[];

    vbaraw?: any;
}

export interface SheetProps {
    /** Name of Sheet */
    name?: string;

    /** Sheet Visibility (0=Visible 1=Hidden 2=VeryHidden) */
    Hidden?: 0 | 1 | 2;

    /** Name of Document Module in associated VBA Project */
    CodeName?: string;
}

/** Defined Name Object */
export interface DefinedName {
    /** Name */
    Name: string;

    /** Reference */
    Ref: string;

    /** Scope (undefined for workbook scope) */
    Sheet?: number;

    /** Name comment */
    Comment?: string;
}

/** Workbook-Level Attributes */
export interface WBProps {
    /** Sheet Properties */
    Sheets?: SheetProps[];

    /** Defined Names */
    Names?: DefinedName[];

    /** Workbook Views */
    Views?: WBView[];

    /** Workbook Modify Protection */
    Modify?: WBModify;

    /** Workbook Properties Protection */
    Protection?: WBProtect;

    /** Other Workbook Properties */
    WBProps?: WorkbookProperties;
}

/** Workbook Modify Protection */
export interface WBModify {
    /** Show a warning before permitting write access */
    warn?: boolean;

    /** Username (shown in message) */
    user?: string;

    /** Encryption info */
    encryption?: PasswordHash;
}

/** Workbook Properties Protection */
export interface WBProtect {
    /** Sheets cannot be moved, deleted, (un)hidden, or renamed */
    lockStructure?: boolean;

    /** Windows are the same size and in the same position */
    lockWindows?: boolean;

    /** Password hash */
    encryption?: PasswordHash;
}

/** Password Hash (internal) */
export interface PasswordHash {
    /* Hash algorithm */
    algo: string;

    /* Salt */
    salt: string;

    /* Spin Count */
    spin: number;

    /* Password hash */
    hash: string;
}

/** Workbook View */
export interface WBView {
    /** Right-to-left mode */
    RTL?: boolean;

    /** Zoom percentage (100 = 100%) */
    zoom?: number;
}

/** Other Workbook Properties */
export interface WorkbookProperties {
    /** Worksheet Epoch (1904 if true, 1900 if false) */
    date1904?: boolean;

    /** Warn or strip personally identifying info on save */
    filterPrivacy?: boolean;

    /** Name of Document Module in associated VBA Project */
    CodeName?: string;
}

/** Custom XML Item */
export interface CustomXMLItem {
    /** Item XML */
    data: string;

    /** Item Properties XML */
    props?: string;
}

/** Column Metrics */
export interface ColInfoCommon {
    /* --- column width --- */

    /** width in Excel's "Max Digit Width", width*256 is integral */
    width?: number;

    /** width in screen pixels */
    wpx?: number;

    /** width in "characters" */
    wch?: number;
}

/** Column Properties Object */
export interface ColInfo extends ColInfoCommon {
    /* --- visibility --- */

    /** if true, the column is hidden */
    hidden?: boolean;

    /** outline / group level */
    level?: number;

    /** Excel's "Max Digit Width" unit, always integral */
    MDW?: number;

    /** Force width recalculation and auto-fit */
    auto?: number;

    /** Number format for new cells in the column */
    z?: number | string;

    /** Style for new cells in the column */
    s?: Style;
}

/** Row Metrics */
export interface RowInfoCommon {
    /* --- visibility --- */

    /** if true, the column is hidden */
    hidden?: boolean;

    /* --- row height --- */

    /** height in screen pixels */
    hpx?: number;

    /** height in points */
    hpt?: number;
}

/** Row Properties Object */
export interface RowInfo extends RowInfoCommon {
    /** outline / group level */
    level?: number;
}

/** Default Cell Metrics */
export interface SheetFormat {
    /** Default Column Metrics */
    col?: ColInfoCommon;

    /** Default Row Metrics */
    row?: RowInfoCommon;
}

/**
 * Write sheet protection properties.
 */
export interface ProtectInfo {
    /**
     * The password for formats that support password-protected sheets
     * (XLSX/XLSB/XLS). The writer uses the XOR obfuscation method.
     */
    password?: string;
    /**
     * Select locked cells
     * @default: true
     */
    selectLockedCells?: boolean;
    /**
     * Select unlocked cells
     * @default: true
     */
    selectUnlockedCells?: boolean;
    /**
     * Format cells
     * @default: false
     */
    formatCells?: boolean;
    /**
     * Format columns
     * @default: false
     */
    formatColumns?: boolean;
    /**
     * Format rows
     * @default: false
     */
    formatRows?: boolean;
    /**
     * Insert columns
     * @default: false
     */
    insertColumns?: boolean;
    /**
     * Insert rows
     * @default: false
     */
    insertRows?: boolean;
    /**
     * Insert hyperlinks
     * @default: false
     */
    insertHyperlinks?: boolean;
    /**
     * Delete columns
     * @default: false
     */
    deleteColumns?: boolean;
    /**
     * Delete rows
     * @default: false
     */
    deleteRows?: boolean;
    /**
     * Sort
     * @default: false
     */
    sort?: boolean;
    /**
     * Filter
     * @default: false
     */
    autoFilter?: boolean;
    /**
     * Use PivotTable reports
     * @default: false
     */
    pivotTables?: boolean;
    /**
     * Edit objects
     * @default: true
     */
    objects?: boolean;
    /**
     * Edit scenarios
     * @default: true
     */
    scenarios?: boolean;

    /** Password hash */
    encryption?: PasswordHash;
}

/** Page Margins -- see Excel Page Setup .. Margins diagram for explanation */
export interface MarginInfo {
    /** Left side margin (inches) */
    left?: number;
    /** Right side margin (inches) */
    right?: number;
    /** Top side margin (inches) */
    top?: number;
    /** Bottom side margin (inches) */
    bottom?: number;
    /** Header top margin (inches) */
    header?: number;
    /** Footer bottom height (inches) */
    footer?: number;
}
export type SheetType = 'sheet' | 'chart';
export type SheetKeys = string | MarginInfo | SheetType;
/** General object representing a Sheet (worksheet or chartsheet) */
export interface Sheet {
    /**
     * Indexing with a cell address string maps to a cell object
     * Special keys start with '!'
     */
    [cell: string]: CellObject | SheetKeys | any;

    /** Sheet type */
    '!type'?: SheetType;

    /** Sheet Range */
    '!ref'?: string;

    /** Page Margins */
    '!margins'?: MarginInfo;
}

/** AutoFilter properties */
export interface AutoFilterInfo {
    /** Range of the AutoFilter table */
    ref: string;
}

/** Outline Settings */
export interface OutlineSettings {
    /**
     * Show Summary rows above detail
     * @default false
     */
    above?: boolean;
    /**
     * Show Summary columns to left of detail
     * @default false
     */
    left?: boolean;
}

export type PrintOrientation = "landscape" | "portrait" | "default";

export type PrintPaper = number | string | { height: number; width: number; };

export interface PrintProps {
    /** Page Orientation (portrait/landscape) */
    orientation?: PrintOrientation;

    fit?: any;

    scale?: number;

    /** Paper  */
    paper?: PrintPaper;

    /** Print Quality (should be 600 or 1200) */
    dpi?: number;

    /** First Page Number */
    first?: null | number;

    /** Center Horizontally */
    centerX?: boolean;

    /** Center Vertically */
    centerY?: boolean;

    /** Print Gridlines */
    gridlines?: boolean;

    /** Black and White */
    bw?: boolean;

    /** "Draft" quality (no grpahics) */
    draft?: boolean;

    /** Print Row and Column Headings */
    headings?: boolean;

    /** Print Comments */
    comments?: "displayed" | "end" | "none";

    /** Print Errors */
    errors?: "displayed" | "none" | "dash" | "n/a";

    /** Page Order ("Over, then down") */
    order?: true | false | "over" | "down";
}

export interface RowBreak {
    R: number;
}

export interface ColBreak {
    C: number;
}

export interface HFPagePart {
    /* Raw code string */
    w?: string;

    /* Plain string */
    t?: string;

    /** Text Style */
    s?: TextStyle;
}

export interface HFPage {
    /* Array of Page Parts */
    R: HFPagePart[];

    /** Text Style */
    s?: TextStyle;
}

export interface HFPageSide {
    /** Center of header/footer */
    center?: string | HFPage | HFPagePart;

    /** Left side of header/footer */
    left?: string | HFPage | HFPagePart;

    /** Right side of header/footer */
    right?: string | HFPage | HFPagePart;
}

export interface HFObject {
    /** First page header (defaults to odd header) */
    first?: HFPage | HFPageSide | string;

    /** Odd page header */
    odd?: HFPage | HFPageSide | string;

    /** Even page header */
    even?: HFPage | HFPageSide | string;
}

export type HeaderFooter = string | HFObject;

/** Print Settings */
export interface PrintSettings {
    /** Print Area Range */
    area?: string | Range;

    /** Print Properties */
    props?: PrintProps;

    /** Margin Settings */
    margins?: MarginInfo;

    /** Row Breaks */
    rowBreaks?: RowBreak[];

    /** Column Breaks */
    colBreaks?: ColBreak[];

    /** Header String */
    header?: HeaderFooter;

    /** Header String */
    footer?: HeaderFooter;
}

/** Absolute Position with Size */
export interface PositionAbsolute {
    /** X-coordinate of upper-left corner (pixels) */
    x: number;

    /** Y-coordinate of upper-left corner (pixels) */
    y: number;

    /** Width (pixels) */
    w: number;

    /** Height (pixels) */
    h: number;
}

/** Relative Position with Size  */
export interface PositionRelative extends PositionAbsolute  {
    /** row of upper-left corner (0-indexed) */
    r: number;

    /** col of upper-left corner (0-indexed) */
    c: number;
}

/** Relative Position of Corners */
export interface PositionCorners {
    /** row of upper-left corner (0-indexed) */
    r: number;

    /** col of upper-left corner (0-indexed) */
    c: number;

    /** X-coordinate of upper-left corner (pixels) */
    x: number;

    /** Y-coordinate of upper-left corner (pixels) */
    y: number;

    /** row of lower-right corner (0-indexed) */
    R: number;

    /** col of lower-right corner (0-indexed) */
    C: number;

    /** X-coordinate of lower-right corner (pixels) */
    X: number;

    /** Y-coordinate of lower-right corner (pixels) */
    Y: number;
}

/** Object (Comment, Shape, Image, Chart) Position */
export type Position = PositionAbsolute | PositionRelative | PositionCorners;


export type WSKeys = SheetKeys | ColInfo[] | RowInfo[] | Range[] | ProtectInfo | AutoFilterInfo | PrintSettings | Image[];

/** Worksheet Object */
export interface WorkSheet extends Sheet {
    /**
     * Indexing with a cell address string maps to a cell object
     * Special keys start with '!'
     */
    [cell: string]: CellObject | WSKeys | any;

    /** Column Info */
    '!cols'?: ColInfo[];

    /** Row Info */
    '!rows'?: RowInfo[];

    /** Default Cell Metrics */
    '!sheetFormat'?: SheetFormat;

    /** Merge Ranges */
    '!merges'?: Range[];

    /** Data Validations */
    '!validations'?: DataValidation[];

    /** Conditional Formats */
    '!condfmt'?: ConditionalFormat[];

    /** Worksheet Protection info */
    '!protect'?: ProtectInfo;

    /** AutoFilter info */
    '!autofilter'?: AutoFilterInfo;

    /** Print Settings */
    '!print'?: PrintSettings;

    /** Outline Settings */
    '!outline'?: OutlineSettings;

    /** Images (ONLY FOR SUPPORTED BUILDS) */
    '!images'?: Image[];

    /** Charts (ONLY FOR SUPPORTED BUILDS) */
    '!charts'?: Chart[];

    /** Pivots (ONLY FOR SUPPORTED BUILDS) */
    '!pivots'?: Pivot[];
}

/** Image Object (ONLY FOR SUPPORTED BUILDS) */
export interface ImageObject {
    /** Image position */
    '!pos'?: Position;

    /** Image link (reference to external resource) */
    '!link'?: string;

    /** Hyperlink (click on image) */
    l?: Hyperlink;
}

/** Image specified as Binary String */
export interface ImageBinary extends ImageObject {
    /** Image data */
    '!data': string;

    /** Type of data */
    '!datatype': "binary";
}

/** Image specified as Base64 String */
export interface ImageBase64 extends ImageObject {
    /** Image data */
    '!data': string;

    /** Type of data */
    '!datatype': "base64";
}

/** Image specified as Buffer */
export interface ImageBuffer extends ImageObject {
    /** Image data */
    '!data'?: any;

    /** Type of data */
    '!datatype': "buffer";
}

/** Remote image */
export interface ImageRemote extends ImageObject {
    /** Type of data */
    '!datatype': "remote";
}

/** Image Object (ONLY FOR SUPPORTED BUILDS) */
export type Image = ImageBinary | ImageBase64 | ImageBuffer | ImageRemote;

export type SeriesColumnType = "xVal" | "yVal" | "cat" | "val" | "bubbleSize";

interface SeriesCommon {
    /** Name of the Series (explicitly set to null to omit) */
    name?: string | null;

    /** Reference to Series Name (cell address string or defined name) */
    nameref?: string;

    /** Ranges for the data axes (reference with worksheet name) */
    ranges?: string[];

    /** Ranges for the data axes (reference with worksheet name) */
    cols?: SeriesColumnType[];

    /** if `true`, pull data from the `ranges`; otherwise use cache */
    raw?: boolean;
}

export type Series = SeriesCommon;

export type ChartType =
    "area" | "area3D" | "line" | "line3D" | "stock" | "radar" | "scatter" |
    "pie" | "pie3D" | "doughnut" | "bar" | "bar3D" | "ofPie" | "surface" |
    "surface3D" | "bubble" | "boxWhisker" | "clusteredColumn" | "funnel" |
    "paretoLine" | "sunburst" | "treemap" | "waterfall" | "regionMap" |
    string;

/** Plot Object */
export interface Plot {
    /** Type of chart */
    t: ChartType;

    /** Series Objects */
    ser: Series[];

    showvals?: boolean;

    raw?: boolean;

    linecolor?: Color;

    labels?: boolean;
}

/** Legend Properties */
export interface LegendProps {
    /** Legend Position */
    'pos': "b" | "l" | "r" | "t";
}

/** Axes Properties */
export interface Axes {
    /** minimum value on `y` (dependent) axis */
    y?: number;

    /** minimum value on `x` (independent) axis */
    x?: number;

    /** maximum value on `y` (dependent) axis */
    Y?: number;

    /** maximum value on `x` (independent) axis */
    X?: number;

    /** major unit step */
    ymajor?: number;

    /** minor unit step */
    yminor?: number;

    /** "Values in reverse order" */
    yrev?: boolean;

    /** number format for `y` (dependent) axis */
    ynf?: NumberFormat;

    /** `x` (independent) Axis Label Position (see below) */
    xlabelpos?: "nextTo" | "high" | "low" | "none";
}

/** Chart Object (ONLY FOR SUPPORTED BUILDS) */
export interface Chart extends Sheet {
    '!type': "chart";

    /** Chart Position */
    '!pos': Position;

    /** Chart Title */
    '!title'?: string;

    /** Legend Properties */
    '!legend'?: LegendProps;

    /** Plot Objects */
    '!plot': Plot[];

    /** Axes Properties */
    '!axes'?: null | Axes;

    /** IETF language tag, default 'en-US' */
    '!lang'?: string;
}

/** Miscellaneous PivotTable properties */
export interface PivotTableProps {
    /** Text displayed in column header */
    collabel?: string;

    /** Text displayed in row header */
    rowlabel?: string;
}

/** PivotTable style options */
export interface PivotTableStyles {
    /** Show Row Headers */
    rowhead?: boolean;

    /** Show Column Headers */
    colhead?: boolean;

    /** Show Row Stripes */
    rowstripe?: boolean;

    /** Show Column Stripes */
    colstripe?: boolean;

    /** Show Summary Column */
    lastcol?: boolean;
}

export type PivotTableOperator = "sum" | "count" | "average" | "max" | "min" |
    "product" | "countNums" | "stdDev" | "stdDevp" | "var" | "varp";

/** PivotTable Value */
export interface PivotTableValue {
    /** Field Name */
    name: string;

    /** Field index from PivotTable `fields` array */
    field: number;

    /** Data Aggregation Operator */
    op?: PivotTableOperator;

    /** Number format (string or SSF Table ID) */
    z?: number | string;
}


/** PivotTable Field */
export interface PivotTableField {
    /** Name as displayed in the UI (should be table header name) */
    name?: string;

    /** Number format for values in this field (string or SSF Table ID) */
    z?: number | string;

    /** Data type */
    t: "s" | "n" | "d";

    /** For String fields, list of values that are "checked", in order */
    l?: string[];
}

/** PivotTable Entry */
export interface PivotTableEntryObject {
    /** Field index */
    field: number;

    /**
     * Sort order
     * - "ascending" and "descending" specify sort order
     * - "manual" preserves manual `l` sort ordering
     */
   order?: "ascending" | "descending" | "manual";

   /** "Sort by" (value field) option in Excel UI */
   key?: number;

   /** Collapse subrows (show every subentry by default) */
   collapse?: boolean;
}
export type PivotTableEntry = number | PivotTableEntryObject;

/** PivotTable Source */
export interface PivotTableSource {
    /** Defined Name (if source is a defined name) */
    name?: string;

    /** A1-style range (if source is a range) */
    ref?: string | Range;

    /** Name of worksheet (if source is on a different worksheet) */
    sheet?: string;
}

/** Pivot Object (ONLY FOR SUPPORTED BUILDS) */
export interface Pivot {
    /** Data Source location */
    source?: PivotTableSource;

    /** Upper-left cell of the PivotTable */
    origin?: string | CellAddress;

    /** Data Series and Filters */
    fields?: PivotTableField[];

    /** Row Fields */
    rows?: PivotTableEntry[];

    /** Column Fields */
    cols?: PivotTableEntry[];

    /** Filters */
    filters?: number[];

    /** Summary Values */
    values?: PivotTableValue[];

    /** PivotTable style options */
    style?: PivotTableStyles;

    /** Other PivotTable properties */
    props?: PivotTableProps;
}

/**
 * Worksheet Object with CellObject type
 *
 * The normal Worksheet type uses indexer of type `any` -- this enforces CellObject
 */
export interface StrictWS { [addr: string]: CellObject; }

/**
 * The Excel data type for a cell.
 * b Boolean, n Number, e error, s String, d Date, z Stub
 */
export type ExcelDataType = 'b' | 'n' | 'e' | 's' | 'd' | 'z';

/**
 * Type of generated workbook
 * @default 'xlsx'
 */
export type BookType = 'xlsx' | 'xlsm' | 'xlsb' | 'xls' | 'xla' | 'biff8' | 'biff5' | 'biff2' | 'xlml' | 'ods' | 'fods' | 'csv' | 'txt' | 'sylk' | 'html' | 'dif' | 'rtf' | 'prn' | 'eth';

/** Comment element */
export interface Comment {
    /** Author of the comment block */
    a?: string;

    /** Plaintext of the comment */
    t?: string;

    /** Rich text */
    R?: RichTextFragment[];
}


/** Style for cell comments */
export interface CommentStyle {
    /** Primary fill color (must be RGB) */
    fgColor?: RGBColor;
}

/** Cell comments */
export interface Comments extends Array<Comment> {
    /** Hide comment by default */
    hidden?: boolean;

    /** Comment position */
    '!pos'?: Position;

    /** Comment style */
    s?: CommentStyle;
}

/** Link object */
export interface Hyperlink {
    /** Target of the link (HREF) */
    Target: string;

    /** Plaintext tooltip to display when mouse is over cell */
    Tooltip?: string;
}

/** Worksheet Cell Object */
export interface CellObject {
    /** The raw value of the cell.  Can be omitted if a formula is specified */
    v?: string | number | boolean | Date;

    /** Formatted text (if applicable) */
    w?: string;

    /**
     * The Excel Data Type of the cell.
     * b Boolean, n Number, e Error, s String, d Date, z Empty
     */
    t: ExcelDataType;

    /** Cell formula (if applicable) */
    f?: string;

    /** Range of enclosing array if formula is array formula (if applicable) */
    F?: string;

    /** Rich text run */
    R?: RichTextFragment[];

    /** HTML rendering of the rich text (if applicable) */
    h?: string;

    /** Comments associated with the cell */
    c?: Comments;

    /** Number format string associated with the cell (if requested) */
    z?: NumberFormat;

    /** Cell hyperlink object (.Target holds link, .tooltip is tooltip) */
    l?: Hyperlink;

    /** The style/theme of the cell (if applicable) */
    s?: Style;
}

export interface RichTextFragment {
    /** Should always be "s" */
    t: "s";

    /** Text of the fragment */
    v: string;

    /** Text style for the fragment */
    s?: TextStyle;
}

/** Alignment */
export interface Alignment {
    /** Indent level */
    indent?: number;

    /** Wrap text */
    wrapText?: boolean;

    /** Horizontal Alignment (left/right/center) */
    horizontal?: "left" | "right" | "center";

    /** Vertical Alignment (top/bottom/center) */
    vertical?: "top" | "bottom" | "center";

    /** Text rotation */
    textRotation?: number;

    /** Shrink to Fit */
    shrinkToFit?: boolean;
}

/** Text Style Object */
export interface TextStyle {
    /** Bold */
    bold?: boolean;

    /** Italic */
    italic?: boolean;

    /** Underline code (1=single, 2=double) */
    underline?: true | false | 0x00 | 0x01 | 0x02 | 0x21 | 0x22;

    /** Font Size */
    sz?: number;

    /** Strike-through */
    strike?: boolean;

    /** Font Name */
    name?: string;

    /** Text Color */
    color?: Color;

    /* Text alignment ("sub"-script, "super"-script) */
    valign?: "sub" | "super" | "";
}

/** Style Object */
export interface Style extends TextStyle {
    /** Top Border */
    top?: Border;

    /** Bottom Border */
    bottom?: Border;

    /** Left Border */
    left?: Border;

    /** Right Border */
    right?: Border;

    /** Pattern Type (optional) */
    patternType?: string;

    /** Cell Primary (background) Color */
    fgColor?: Color;

    /** Cell Secondary (background) Color */
    bgColor?: Color;

    /* Cell alignment properties */
    alignment?: Alignment;

    /** Cell formula hidden when worksheet is locked */
    hidden?: boolean;

    /** Cell editable when worksheet is locked (inverse of "locked" in Excel) */
    editable?: boolean;

    /** Style Name */
    style?: string;
}

/** Types of Data Validation */
export type DataValidationType = 'Any' | 'Whole' | 'Decimal' | 'List' | 'Date' | 'Time' | 'Length' | 'Custom';

/** Data Validation Operators */
export type DataValidationOp =
    'IN' | /* Between */
    'OT' | /* Not Between */
    'EQ' | /* Equal To */
    'NE' | /* Not Equal To */
    'GT' | /* Greater Than */
    'LT' | /* Less Than */
    'GE' | /* Greater Than or Equal To */
    'LE' ; /* Less Than Or Equal To */

/** Input Message metadata */
export interface DataValidationInputMessage {
    title?: string;
    message?: string;
}

/** Error Alert metadata */
export interface DataValidationErrorAlert extends DataValidationInputMessage {
    /** Error Alert style */
    style?: "stop" | "warning" | "info";
}

/** Data Validation */
export interface DataValidation {
    /** Range of cells covered by the DV */
    ref: string | CellAddress | Range;

    /** Type of Data Validation */
    t: DataValidationType;

    /** Array of strings for a fixed dropdown List (List only) */
    l?: string[];

    /** Formula or Range for Custom DV or sourced List */
    f?: string;

    /** Operator (for applicable Types) */
    op?: DataValidationOp;

    /** Operator Value (for single-value operators) */
    v?: number|Date;

    /** Minimum Value (where applicable) */
    min?: number|Date;

    /** Maximum Value (where applicable) */
    max?: number|Date;

    /** "Ignore Blank" (set to `false` to disable) */
    blank?: boolean;

    /** Input Message */
    input?: DataValidationInputMessage | false;

    /** Error Alert */
    error?: DataValidationErrorAlert | false;
}

export type ConditionalFormatTypeDiffStyle =
    'avg'     | /* Format only values that are above or below average */
    'blank'   | /* Format only cells that contain: Blanks or no Blanks */
    'date'    | /* Format only cells that contain: Dates Occurring */
    'dup'     | /* Format all duplicate values */
    'error'   | /* Format only cells that contain: Errors or No Errors */
    'formula' | /* Format values where formula is true */
    'rank'    | /* Format only top or bottom ranked values */
    'text'    | /* Format only cells that contain: Specific Text */
    'unique'  | /* Format all unique values */
    'val'     ; /* Format only cells that contain: Cell Value */

export type ConditionalFormatTypeNoDiff =
    'bar'     | /* Format all cells based on values: Data Bars */
    'icon'    | /* Format all cells based on values: Icon Sets */
    'scale'   ; /* Format all cells based on values: 2- or 3- color scale */

/** Data Validation Operators */
export type ConditionalFormatType = ConditionalFormatTypeDiffStyle | ConditionalFormatTypeNoDiff;

export type DifferentialStyle = any;

interface CFBase {
    /** Range of cells covered by the CF */
    ref: string | CellAddress | Range;

    /** Type of Data Validation */
    t: ConditionalFormatType;
}

interface CFGeneric extends CFBase {
    /** Differential Style (when relevant) */
    s?: DifferentialStyle;

    /** Data operator (when relevant) */
    op?: string;

    /** Formula (when relevant) */
    f?: string;

    min?: any;
    max?: any;
    v?: any;

    color?: any;
    cmin?: any;
    cmax?: any;
    cmid?: any;
    thresh?: any;
}

/** CF "Use a formula to determine which cells to format" */
export interface CFFormula extends CFBase {
    /** Type of Data Validation */
    t: 'formula';

    /** Formula string (exactly as entered in UI formula bar) */
    f: string;

    /** Differential Style (when relevant) */
    s?: DifferentialStyle;
}

/** Conditional Format */
export type ConditionalFormat = CFGeneric | CFFormula;

/** sheet_set_range_style Style */
export interface StyleZ extends Style {
    /** Number format string associated with the cell (if requested) */
    z?: NumberFormat;

    /** Interior Vertical Border */
    incol?: Border;

    /** Interior Horizontal Border */
    inrow?: Border;
}

/** Valid Border Style */
export type BorderStyle = 'thin' | 'medium' | 'thick' | 'dotted' | 'hair' | 'dashed' | 'mediumDashed' | 'dashDot' | 'mediumDashDot' | 'dashDotDot' | 'mediumDashDotDot' | 'slantDashDot';

/** Border object */
export interface Border {
    /** Border Style */
    style?: BorderStyle;
    /** Border Color */
    color?: Color;
}

/** Color Object */
export type Color = RGBColor | ThemeColor | IndexedColor;

/** RGB Color */
export interface RGBColor {
    /** RGB Color (hex string or number) */
    rgb: string|number;
}

/** Theme Color */
export interface ThemeColor {
    /** Theme index */
    theme: number;

    /** Tint ratio (between -1 and 1) */
    tint?: number;
}

/** Indexed Color */
export interface IndexedColor {
    /** Palette Index (integer <56) */
    index: number;
}

/** Simple Cell Address */
export interface CellAddress {
    /** Column number */
    c: number;
    /** Row number */
    r: number;
}

/** Range object (representing ranges like "A1:B2") */
export interface Range {
    /** Starting cell */
    s: CellAddress;
    /** Ending cell */
    e: CellAddress;
}

export interface Sheet2CSVOpts extends DateNFOption {
    /** Field Separator ("delimiter") */
    FS?: string;

    /** Record Separator ("row separator") */
    RS?: string;

    /** Remove trailing field separators in each record */
    strip?: boolean;

    /** Include blank lines in the CSV output */
    blankrows?: boolean;

    /** Skip hidden rows and columns in the CSV output */
    skipHidden?: boolean;
}

export interface OriginOption {
    /** Top-Left cell for operation (CellAddress or A1 string or row) */
    origin?: number | string | CellAddress;
}

export interface Sheet2HTMLOpts {
    /** TABLE element id attribute */
    id?: string;

    /** Add contenteditable to every cell */
    editable?: boolean;

    /** Header HTML */
    header?: string;

    /** Footer HTML */
    footer?: string;

    /** Override gridline color (CSS color) */
    gridcolor?: string;
}

export interface Sheet2JSONOpts extends DateNFOption {
    /** Output format */
    header?: "A"|number|string[];

    /** Override worksheet range */
    range?: any;

    /** Include or omit blank lines in the output */
    blankrows?: boolean;

    /** Default value for null/undefined values */
    defval?: any;

    /** if true, return raw data; if false, return formatted text */
    raw?: boolean;
}

export interface AOA2SheetOpts extends CommonOptions, DateNFOption, OriginOption {
    /**
     * Create cell objects for stub cells
     * @default false
     */
    sheetStubs?: boolean;
}

export interface SheetAOAOpts extends AOA2SheetOpts, OriginOption {}

export interface JSON2SheetOpts extends CommonOptions, DateNFOption {
    /** Use specified column order */
    header?: string[];

    /** Skip header row in generated sheet */
    skipHeader?: boolean;
}

export interface SheetJSONOpts extends JSON2SheetOpts, OriginOption {}

export interface Table2SheetOpts extends CommonOptions, DateNFOption, OriginOption {
    /** If true, plaintext parsing will not parse values */
    raw?: boolean;

    /** If true, values will never be guessed as Dates */
    rawDates?: boolean;

    /**
     * If >0, read the first sheetRows rows
     * @default 0
     */
    sheetRows?: number;

    /** If true, hidden rows and cells will not be parsed */
    display?: boolean;

    /** If true, HTML TABLE borders will be translated to styled borders */
    borders?: boolean;
}

export interface TemplateAOAOpts {
    /** If false, do not try to correct formulae and defined names */
    formula?: boolean;
}

/** General utilities */
export interface XLSX$Utils {
    /* --- Import Functions --- */

    /** Converts an array of arrays of JS data to a worksheet. */
    aoa_to_sheet<T>(data: T[][], opts?: AOA2SheetOpts): WorkSheet;
    aoa_to_sheet(data: any[][], opts?: AOA2SheetOpts): WorkSheet;

    /** Converts an array of JS objects to a worksheet. */
    json_to_sheet<T>(data: T[], opts?: JSON2SheetOpts): WorkSheet;
    json_to_sheet(data: any[], opts?: JSON2SheetOpts): WorkSheet;

    /** BROWSER ONLY! Converts a TABLE DOM element to a worksheet. */
    table_to_sheet(data: any,  opts?: Table2SheetOpts): WorkSheet;
    table_to_book(data: any,  opts?: Table2SheetOpts): WorkBook;
    sheet_add_dom(ws: WorkSheet, data: any, opts?: Table2SheetOpts): WorkSheet;

    /* --- Export Functions --- */

    /** Converts a worksheet object to an array of JSON objects */
    sheet_to_json<T>(worksheet: WorkSheet, opts?: Sheet2JSONOpts): T[];
    sheet_to_json(worksheet: WorkSheet, opts?: Sheet2JSONOpts): any[][];
    sheet_to_json(worksheet: WorkSheet, opts?: Sheet2JSONOpts): any[];

    /** Generates delimiter-separated-values output */
    sheet_to_csv(worksheet: WorkSheet, options?: Sheet2CSVOpts): string;

    /** Generates UTF16 Formatted Text */
    sheet_to_txt(worksheet: WorkSheet, options?: Sheet2CSVOpts): string;

    /** Generates HTML */
    sheet_to_html(worksheet: WorkSheet, options?: Sheet2HTMLOpts): string;

    /** Generates a list of the formulae (with value fallbacks) */
    sheet_to_formulae(worksheet: WorkSheet): string[];

    /** Generates DIF */
    sheet_to_dif(worksheet: WorkSheet, options?: Sheet2HTMLOpts): string;

    /** Generates SYLK (Symbolic Link) */
    sheet_to_slk(worksheet: WorkSheet, options?: Sheet2HTMLOpts): string;

    /** Generates ETH */
    sheet_to_eth(worksheet: WorkSheet, options?: Sheet2HTMLOpts): string;

    /* --- Cell Address Utilities --- */

    /** Converts 0-indexed cell address to A1 form */
    encode_cell(cell: CellAddress): string;

    /** Converts 0-indexed row to A1 form */
    encode_row(row: number): string;

    /** Converts 0-indexed column to A1 form */
    encode_col(col: number): string;

    /** Converts 0-indexed range to A1 form */
    encode_range(s: CellAddress, e: CellAddress): string;
    encode_range(r: Range): string;

    /** Converts A1 cell address to 0-indexed form */
    decode_cell(address: string): CellAddress;

    /** Converts A1 row to 0-indexed form */
    decode_row(row: string): number;

    /** Converts A1 column to 0-indexed form */
    decode_col(col: string): number;

    /** Converts A1 range to 0-indexed form */
    decode_range(range: string): Range;

    /** Format cell */
    format_cell(cell: CellObject, v?: any, opts?: any): string;

    /* --- General Utilities --- */

    /** Creates a new workbook */
    book_new(): WorkBook;

    /** Append a worksheet to a workbook */
    book_append_sheet(workbook: WorkBook, worksheet: WorkSheet, name?: string): void;

    /** Set sheet visibility (visible/hidden/very hidden) */
    book_set_sheet_visibility(workbook: WorkBook, sheet: number|string, visibility: number): void;

    /** Set number format for a cell */
    cell_set_number_format(cell: CellObject, fmt: string|number): CellObject;

    /** Set hyperlink for a cell */
    cell_set_hyperlink(cell: CellObject, target: string, tooltip?: string): CellObject;

    /** Set internal link for a cell */
    cell_set_internal_link(cell: CellObject, target: string, tooltip?: string): CellObject;

    /** Add comment to a cell */
    cell_add_comment(cell: CellObject, text: string, author?: string): void;

    /** Assign an Array Formula to a range */
    sheet_set_array_formula(ws: WorkSheet, range: Range|string, formula: string): WorkSheet;

    /** Add an array of arrays of JS data to a worksheet */
    sheet_add_aoa<T>(ws: WorkSheet, data: T[][], opts?: SheetAOAOpts): WorkSheet;
    sheet_add_aoa(ws: WorkSheet, data: any[][], opts?: SheetAOAOpts): WorkSheet;

    /** Add an array of JS objects to a worksheet */
    sheet_add_json(ws: WorkSheet, data: any[], opts?: SheetJSONOpts): WorkSheet;
    sheet_add_json<T>(ws: WorkSheet, data: T[], opts?: SheetJSONOpts): WorkSheet;

    /** Apply style to a given range */
    sheet_set_range_style(ws: WorkSheet, range: Range|string, style: StyleZ): void;

    /** Modify a style according to a differential style */
    apply_style_delta(style: StyleZ, delta: DifferentialStyle): void;

    /** Compute final style for a cell based on worksheet metadata */
    get_computed_style(ws: WorkSheet, addr: CellSpec): StyleZ;

    /**
     * (ONLY FOR SUPPORTED BUILDS -- template editor)
     * Write array of arrays to a template
     */
    template_set_aoa(wb: WorkBook, wsname: WSSpec, range: RangeSpec, aoa: any[], opts?: TemplateAOAOpts): void;

    /**
     * (ONLY FOR SUPPORTED BUILDS -- template editor)
     * Remove a worksheet from a template
     */
    template_book_delete_sheet(wb: WorkBook, wsname: WSSpec): void;

    /**
     * (ONLY FOR SUPPORTED BUILDS -- encryption)
     * Generate password hash
     */
    hash_password(password: string): PasswordHash;

    /**
     * (ONLY FOR SUPPORTED BUILDS -- encryption)
     * Verify password
     */
    test_password(enc: PasswordHash, password: string): boolean;

    consts: XLSX$Consts;
}

export interface XLSX$Consts {
    /* --- Sheet Visibility --- */

    /** Visibility: Visible */
    SHEET_VISIBLE: 0;

    /** Visibility: Hidden */
    SHEET_HIDDEN: 1;

    /** Visibility: Very Hidden */
    SHEET_VERYHIDDEN: 2;
}

/** NODE ONLY! these return Readable Streams */
export interface StreamUtils {
    /** CSV output stream, generate one line at a time */
    to_csv(sheet: WorkSheet, opts?: Sheet2CSVOpts): any;
    /** HTML output stream, generate one line at a time */
    to_html(sheet: WorkSheet, opts?: Sheet2HTMLOpts): any;
    /** JSON object stream, generate one row at a time */
    to_json(sheet: WorkSheet, opts?: Sheet2JSONOpts): any;
}
