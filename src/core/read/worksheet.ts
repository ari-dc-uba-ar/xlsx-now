// `xl/worksheets/sheetN.xml`: the data itself.
//
// This is the part that can be big — a sheet at Excel's limit is a quarter of
// a gigabyte of XML — and the only one of them the reader never holds whole.
// The chunks come out of the zip, go through the parser, and rows come out
// the other end as they close.
//
// Everything above it is small by comparison and read in one piece: the
// workbook, the relationships, the styles, and the shared strings. That last
// one is the honest limit of how little a reader can hold, and it is the
// format's doing — a cell says `<v>7</v>` and means the seventh entry of a
// table it does not carry.
import { columnIndex } from '../cell.js';
import type { CellType, StyledCell } from '../types.js';
import { readDate, readDuration, readIsoDate, ZONE, type ReadDates } from './dates.js';
import {
    codeOf,
    formatDateSerial,
    formatDateText,
    formatElapsed,
    formatNumber,
    generalText,
    isoOfSerial,
    localeFor,
} from './displayText.js';
import type { NumberFormats } from './numberFormats.js';
import type { ReadCell, ReadFields, ReadRow, ReadValue } from './types.js';
import { XmlParser } from './xml.js';

/** Days between the 1900 epoch and the 1904 one a Macintosh workbook uses. */
const DAYS_1904_TO_1900 = 1462;

/** A cell as the file spells it, before anything has been made of it. */
export interface RawCell {
    /** The `t`: what the cell says it holds. Absent means a number. */
    type: string | undefined;
    /** The `s`: which entry of `cellXfs` says how it is shown. */
    style: number | undefined;
    /** The text of `<v>`. */
    value: string | undefined;
    /** The expression of `<f>`, without the `=` a sheet does not store. */
    formula: string | undefined;
    /**
     * The text of the `<is>` a cell carries its own string in, its runs
     * joined. Left out of a `RawCell` built by hand, where `value` holds it.
     */
    inline?: string | undefined;
}

/** What a cell is read against: everything the worksheet itself does not say. */
export interface CellContext {
    sharedStrings: readonly string[];
    formats: NumberFormats;
    date1904: boolean;
    /** What a date is built as; see `ReadDates`. */
    dates: ReadDates;
    /**
     * The language the text of a cell is shown in, for `w`: `es` or an `es-…`
     * for Spanish, anything else for English. Left out, a date is shown in ISO.
     */
    lang?: string | undefined;
}

/** `B12` as the coordinates it names, counting columns from 0 and rows from 1. */
export function parseCellReference(
    reference: string,
): { column: number; row: number } | undefined {
    const match = /^([A-Za-z]+)([0-9]+)$/.exec(reference);
    if (!match) return undefined;
    const column = columnIndex(match[1] as string);
    if (column === undefined) return undefined;
    return { column, row: Number(match[2]) };
}

/** The number in a `<v>`, or a failure that names what was there instead. */
function numberOf(raw: RawCell): number {
    const value = Number(raw.value);
    if (!Number.isFinite(value)) {
        throw new Error(`A numeric cell holds "${raw.value}", which is not a number.`);
    }
    return value;
}

/** The date a serial means, under whichever epoch the workbook counts from. */
function dateOf(serial: number, context: CellContext): ReadValue {
    return readDate(context.date1904 ? serial + DAYS_1904_TO_1900 : serial, context.dates);
}

/**
 * A cell as its value.
 *
 * The types are the file's own: `s` points into the shared strings, `b` is a
 * boolean written as `1` or `0`, `str` is the cached result of a formula,
 * `inlineStr` is text the cell carries itself, `e` is an error, `d` is a date
 * written out in full — rare, but in the spec — and no type at all is a
 * number. Which is where the format comes in: a number under a date format is
 * a date, and that is the only thing the styles are read for.
 *
 * An error comes back as its own text — `#DIV/0!` — rather than as `null`.
 * It is what the cell shows, and a `null` would be a value that went missing
 * without anyone saying so.
 */
export function cellValue(raw: RawCell, context: CellContext): ReadValue {
    if (raw.type === 'inlineStr' && raw.inline !== undefined) return raw.inline;
    if (raw.value === undefined) return null;
    switch (raw.type) {
        case undefined:
        case 'n': {
            const value = numberOf(raw);
            const kind = context.formats.kind(raw.style);
            // A length of time is a number of days, the same in either epoch:
            // thirty hours are thirty hours whatever day the workbook starts on.
            if (kind === 'elapsed') return readDuration(value, context.dates);
            return kind === 'date' ? dateOf(value, context) : value;
        }
        case 's': {
            const index = numberOf(raw);
            const text = context.sharedStrings[index];
            if (text === undefined) {
                throw new Error(
                    `A cell points at the shared string ${index}, and the table has ${context.sharedStrings.length}.`,
                );
            }
            return text;
        }
        case 'b':
            return raw.value === '1';
        case 'str':
        case 'inlineStr':
        case 'e':
            return raw.value;
        case 'd':
            // The cell that spells its date out instead of numbering it, which
            // is what a day the serial cannot number has to be written as. It
            // never becomes one: see `readIsoDate` — and the 1904 epoch has
            // nothing to shift in a text that names its own day.
            return readIsoDate(raw.value, context.dates);
        default:
            throw new Error(`A cell says it holds "${raw.type}", which is not a type a sheet has.`);
    }
}

/**
 * A cell as the writer would take it back.
 *
 * `t` is only written where the writer would not work it out for itself: a
 * string, a number, a boolean and a date all say what they are by being what
 * they are, and what is left is the cached result of a formula — text that
 * has to stay text — and an error, which is text that is not a value.
 *
 * The format goes in `s`, as a style written out, because that is where the
 * writer reads one: `{ v: 45306, s: { numFmt: 14 } }` is a cell that can go
 * straight back into a workbook and come out the same.
 */
export function styledCell(raw: RawCell, context: CellContext): StyledCell {
    const cell: StyledCell = { v: cellValue(raw, context) };
    const type: CellType | undefined =
        raw.type === 'e' ? 'e' : raw.type === 'str' ? 'str' : undefined;
    if (type !== undefined) cell.t = type;
    if (raw.formula !== undefined) cell.f = raw.formula;
    const numFmt = context.formats.forStyle(raw.style);
    if (numFmt !== undefined) cell.s = { numFmt };
    return cell;
}

/**
 * The text a cell shows — its `w`. See `displayText.ts` for how close that
 * comes to what the application showed, and why it is never left out.
 *
 * A date is the one value that depends on whether a language was given: with
 * one, it is shown under its own code, in that language; without one, in ISO.
 * A length of time is shown under its code either way, since `30:00` says the
 * same thing in every language.
 */
export function cellText(raw: RawCell, context: CellContext): string {
    const locale = localeFor(context.lang);
    switch (raw.type) {
        case undefined:
        case 'n': {
            if (raw.value === undefined) return '';
            const value = numberOf(raw);
            const numFmt = context.formats.forStyle(raw.style);
            const kind = context.formats.kind(raw.style);
            if (kind === 'elapsed') return formatElapsed(value, codeOf(numFmt, locale), locale);
            if (kind === 'date') {
                const serial = context.date1904 ? value + DAYS_1904_TO_1900 : value;
                return context.lang === undefined
                    ? isoOfSerial(serial)
                    : formatDateSerial(serial, codeOf(numFmt, locale), locale);
            }
            const formatted =
                numFmt === undefined ? undefined : formatNumber(value, codeOf(numFmt, locale), locale);
            return formatted ?? generalText(value, locale);
        }
        case 'b':
            if (raw.value === undefined) return '';
            return raw.value === '1' ? locale.true : locale.false;
        case 'd': {
            if (raw.value === undefined) return '';
            const wall = raw.value.replace(ZONE, '');
            if (context.lang === undefined) return wall;
            // A date written out under a format that is not a date one is
            // still a date, and is shown as the language's short date.
            const code =
                context.formats.kind(raw.style) === 'date'
                    ? codeOf(context.formats.forStyle(raw.style), locale)
                    : locale.shortDate + (wall.includes('T') ? ' hh:mm:ss' : '');
            return formatDateText(wall, code, locale);
        }
        default: {
            const value = cellValue(raw, context);
            return value === null ? '' : String(value);
        }
    }
}

/**
 * A cell as the fields a mode asked for — see `ReadCell`.
 *
 * A field the cell has nothing for is left out rather than set to
 * `undefined`: a cell with no formula has no `f`, as a cell of the writer's
 * would not. The one exception is `w`, which is always there when it was
 * asked for, and `v`, which is `null` for a cell that holds nothing.
 */
export function readCell(raw: RawCell, context: CellContext, fields: ReadFields): Partial<ReadCell> {
    const cell: Partial<ReadCell> = {};
    if (fields.v !== false) cell.v = cellValue(raw, context);
    if (fields.s) {
        const numFmt = context.formats.forStyle(raw.style);
        if (numFmt !== undefined) cell.s = { numFmt };
    }
    if (fields.f && raw.formula !== undefined) cell.f = raw.formula;
    if (fields.t) {
        const type: CellType | undefined =
            raw.type === 'e' ? 'e' : raw.type === 'str' ? 'str' : undefined;
        if (type !== undefined) cell.t = type;
    }
    if (fields.w) cell.w = cellText(raw, context);
    if (fields.kind) {
        if (raw.type === 'd') cell.kind = 'date';
        else if ((raw.type === undefined || raw.type === 'n') && raw.value !== undefined) {
            cell.kind = context.formats.kind(raw.style);
        }
    }
    if (fields._t && raw.type !== undefined) cell._t = raw.type;
    if (fields._s && raw.style !== undefined) cell._s = raw.style;
    if (fields._v && raw.value !== undefined) cell._v = raw.value;
    if (fields._f && raw.formula !== undefined) cell._f = raw.formula;
    if (fields._is && raw.inline !== undefined) cell._is = raw.inline;
    return cell;
}

/**
 * The rows of a worksheet, as the chunks of it go by.
 *
 * `saxes` calls back while a chunk is being written and a generator cannot
 * yield from inside a callback, so a chunk's rows are collected and handed
 * out right after it — the same order, with the rows in hand instead of on
 * the stack. Nothing accumulates past one chunk.
 */
export async function* readRows<C>(
    chunks: AsyncIterable<string>,
    convert: (raw: RawCell) => C,
    partName: string,
): AsyncIterable<ReadRow<C>> {
    const ready: ReadRow<C>[] = [];

    let inSheetData = false;
    let rowIndex = 0;
    let cells: (C | undefined)[] = [];
    let column = 0;
    let raw: RawCell | undefined;
    /** Where the text arriving now belongs, if anywhere. */
    let target: 'value' | 'inline' | 'formula' | undefined;

    const parser = new XmlParser(
        {
            open(name, attributes) {
                if (name === 'sheetData') {
                    inSheetData = true;
                    return;
                }
                if (!inSheetData) return;
                switch (name) {
                    case 'row': {
                        // A row is allowed to leave its number out, and then
                        // it is simply the next one.
                        const declared = Number(attributes['r']);
                        rowIndex = Number.isInteger(declared) ? declared : rowIndex + 1;
                        cells = [];
                        column = 0;
                        break;
                    }
                    case 'c': {
                        const reference = attributes['r'];
                        const at = reference === undefined ? undefined : parseCellReference(reference);
                        if (reference !== undefined && at === undefined) {
                            throw new Error(`"${reference}" is not a cell reference.`);
                        }
                        column = at?.column ?? column;
                        const style = Number(attributes['s']);
                        raw = {
                            type: attributes['t'],
                            style: Number.isInteger(style) ? style : undefined,
                            value: undefined,
                            formula: undefined,
                        };
                        break;
                    }
                    case 'v':
                        if (raw) {
                            raw.value ??= '';
                            target = 'value';
                        }
                        break;
                    case 't':
                        // A `<t>` in a sheet is the text of an `<is>`, which
                        // is kept apart from `<v>`: they are two fields of
                        // the file, `_is` and `_v`.
                        if (raw) {
                            // `??=` and not `=`: the runs of an `<is>` are
                            // several `<t>` of one value, and they join.
                            raw.inline ??= '';
                            target = 'inline';
                        }
                        break;
                    case 'f':
                        if (raw) {
                            raw.formula ??= '';
                            target = 'formula';
                        }
                        break;
                }
            },
            text(text) {
                if (!raw || target === undefined) return;
                if (target === 'value') raw.value += text;
                else if (target === 'inline') raw.inline += text;
                else raw.formula += text;
            },
            close(name) {
                switch (name) {
                    case 'sheetData':
                        inSheetData = false;
                        break;
                    case 'v':
                    case 't':
                    case 'f':
                        target = undefined;
                        break;
                    case 'c':
                        if (raw) {
                            cells[column] = convert(raw);
                            raw = undefined;
                            column++;
                        }
                        break;
                    case 'row':
                        if (inSheetData) ready.push({ index: rowIndex, cells });
                        break;
                }
            },
        },
        partName,
    );

    for await (const chunk of chunks) {
        parser.write(chunk);
        yield* ready.splice(0);
    }
    parser.close();
    yield* ready.splice(0);
}
