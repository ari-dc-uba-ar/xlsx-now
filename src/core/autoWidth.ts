// Column widths worked out from what the sheet holds. Nothing here writes XML
// and nothing here knows about zip entries: it counts characters as the cells
// go by and hands back one width per column, which is all `autoWidthMax` ever
// was.
import type { NativeValue } from './valueTypes.js';

/**
 * The longest line of a text, in characters — which for a text of one line is
 * the whole of it, so there is nothing to ask before calling this.
 *
 * Walked rather than split: this is the one path that allocates per cell if
 * it is written the obvious way, and the answer is a number. A `\r\n` counts
 * as the line break it is — the carriage return is not a character the cell
 * shows, so it is not one the column has to fit. A carriage return with no
 * line feed after it is the same character and counts the same way: what a
 * reader finds there is a line break, since XML turns a lone `\r` into a `\n`
 * on its way in.
 */
function longestLine(text: string): number {
    let longest = 0;
    let from = 0;
    for (;;) {
        const at = text.indexOf('\n', from);
        const end = at < 0 ? text.length : at;
        const carriageReturn = end > from && text.charCodeAt(end - 1) === 13;
        const length = end - from - (carriageReturn ? 1 : 0);
        if (length > longest) longest = length;
        if (at < 0) return longest;
        from = at + 1;
    }
}

/**
 * How many characters a cell shows.
 *
 * A string is its own length and a number is the length of the digits it is
 * written as — not of the number format that may be shown over it, which is
 * the one thing this cannot know: a format is a style, and a style is a name
 * the workbook resolves at the end. A boolean is `TRUE` or `FALSE`, as Excel
 * spells it.
 *
 * `shown` is the answer a value's own type gave, for the values whose written
 * form says nothing about their shown one — a date is a serial here and ten
 * characters on the screen, and only the type that made it a serial knows
 * that.
 *
 * `wraps` is whether the cell's style wraps its text, and it is what decides
 * what a line break in the value means. A cell that wraps shows one line per
 * break, so what the column has to fit is the longest of them; a cell that
 * does not shows the text on one line, break and all, and there the whole
 * length is the answer — which is why this cannot be settled by looking at
 * the value alone. Excel is the one drawing the distinction: a `CHAR(10)` in
 * a cell without wrap text is not shown as a line break at all, which is why
 * Alt+Enter turns wrapping on as it inserts one.
 *
 * An empty cell measures 0: it takes part in no width, so a column of blanks
 * is a column nobody asked to resize.
 */
export function cellTextLength(value: NativeValue, shown?: number, wraps?: boolean): number {
    if (value === null || value === undefined) return 0;
    if (shown !== undefined) return shown;
    if (typeof value === 'boolean') return value ? 4 : 5;
    const text = String(value);
    // A cell that does not wrap — which is nearly every cell — is out before
    // its text is looked at at all. The one that does is walked, break in it
    // or not: a text of one line is its own longest line.
    if (!wraps) return text.length;
    return longestLine(text);
}

/**
 * The widest digit of the normal font, in pixels — Calibri 11 at 96 dpi,
 * which is the font a workbook has until one of its styles says otherwise —
 * and the padding a column carries around its text: two pixels of margin on
 * each side, and one more for the gridline.
 */
const MAX_DIGIT_WIDTH = 7;
const COLUMN_PADDING_PIXELS = 5;

/**
 * A count of characters as the `width` a `<col>` carries.
 *
 * The two are not the same number, which is the whole of this function.
 * ECMA-376 §18.3.1.13 measures a column in multiples of the widest digit of
 * the normal font *plus the padding*, and stores it in 1/256ths:
 *
 * ```
 * width = Truncate([{characters} * {digit width} + {5px padding}] / {digit width} * 256) / 256
 * ```
 *
 * which is why Excel writes `8.7109375` for a column it autofitted to eight
 * characters, and not `8`. Writing the count itself leaves every column short
 * by that padding — the text ends up clipped, and a number or a date under it
 * comes out as `##########`, which is the visible half of the same bug.
 */
export function columnWidth(characters: number): number {
    const pixels = characters * MAX_DIGIT_WIDTH + COLUMN_PADDING_PIXELS;
    return Math.trunc((pixels / MAX_DIGIT_WIDTH) * 256) / 256;
}

/** Why a max nobody can size a column by was refused. */
function badMaxError(max: number): Error {
    return new Error(
        `"${max}" is not an autoWidthMax: it is the widest a column may get, in characters, ` +
            'so it has to be a number above 0.',
    );
}

/** Why a min nobody can size a column by was refused. */
function badMinError(min: number, max: number | undefined): Error {
    return new Error(
        `"${min}" is not an autoWidthMin: it is the narrowest a measured column may get, ` +
            'in characters, so it has to be a number above 0' +
            (max === undefined ? '.' : ` and no more than the autoWidthMax of ${max}.`),
    );
}

/** Why a count of rows nobody can measure was refused. */
function badRowsError(rows: number): Error {
    return new Error(
        `"${rows}" is not an autoWidthRows: it is how many rows of the sheet are measured, ` +
            'so it has to be a whole number above 0.',
    );
}

/** What a sheet asks of its meter besides the maximum. */
export interface WidthMeterOptions {
    /** The narrowest a column that measured something may get, in characters. */
    min?: number | undefined;
    /** How many rows are measured before the widths are settled. */
    rows?: number | undefined;
}

/**
 * The widths of a sheet, as its cells go by: every cell is measured into the
 * column it lands in, and the column keeps the longest of them, between the
 * minimum and the maximum it was opened with.
 *
 * One meter per worksheet, and every worksheet has one — a sheet with no
 * `autoWidthMax` gets a meter that measures nothing, so nothing downstream
 * has to ask whether there is one. What it does have to ask is `measures`:
 * that is what says whether the sheet can go out as it is written or has to
 * wait for more rows. A meter given `rows` stops measuring once that many
 * rows are in, and from there on the sheet goes out as it is written.
 */
export class WidthMeter {
    /** The longest cell seen per 0-based column; a hole is a column with nothing in it. */
    private readonly widths: number[] = [];
    private readonly max: number;
    private readonly min: number;
    /** The rows left to measure; `Infinity` when every row is. */
    private rowsLeft: number;
    private measuring: boolean;

    constructor(max: number | undefined, options: WidthMeterOptions = {}) {
        const { min, rows } = options;
        if (max !== undefined && !(Number.isFinite(max) && max > 0)) throw badMaxError(max);
        if (min !== undefined && !(Number.isFinite(min) && min > 0 && min <= (max ?? min))) {
            throw badMinError(min, max);
        }
        if (rows !== undefined && !(Number.isInteger(rows) && rows > 0)) throw badRowsError(rows);
        this.measuring = max !== undefined;
        this.max = max ?? 0;
        this.min = min ?? 0;
        this.rowsLeft = rows ?? Number.POSITIVE_INFINITY;
    }

    /**
     * Whether the cells are still being measured — and so whether the sheet
     * has to wait before its `<cols>` can be written.
     */
    get measures(): boolean {
        return this.measuring;
    }

    /** One cell, in the column it was written in. */
    see(column: number, value: NativeValue, shown?: number, wraps?: boolean): void {
        if (!this.measuring) return;
        const length = cellTextLength(value, shown, wraps);
        if (!length) return;
        const width = length > this.max ? this.max : length;
        if (width > (this.widths[column] ?? 0)) this.widths[column] = width;
    }

    /**
     * One row is in. The last one the meter was asked to measure settles the
     * widths: what comes after it is written, not measured.
     */
    endRow(): void {
        if (!this.measuring) return;
        this.rowsLeft--;
        if (this.rowsLeft <= 0) this.measuring = false;
    }

    /**
     * What every column measured, as the `width` a `<col>` is written with —
     * the characters it counted, raised to the minimum, plus the padding
     * Excel measures a column by. A column nobody wrote anything in is a
     * hole: it keeps whatever `columnFormats` says about it, and Excel's
     * default width when that says nothing either — the minimum is for the
     * columns that measured something, not for every column of the sheet.
     */
    columnWidths(): readonly number[] {
        // `map` keeps the holes as holes, which is what the sparse array is for.
        return this.widths.map((width) => columnWidth(width < this.min ? this.min : width));
    }
}
