// `xl/styles.xml`, read for the one thing a reader of data cannot do without.
//
// A date in a sheet is a number. Nothing in the cell says otherwise: `45306`
// is `45306` whether it means forty-five thousand of something or the 15th of
// January of 2024, and the only thing that tells them apart is the number
// format its style points at. So the reader parses the styles it was told not
// to care about — but only this far: which format each style shows, and
// whether that format is a date. Fonts, fills and borders are not read.
import { parseXml } from './xml.js';

const PART_NAME = 'xl/styles.xml';

/** The general format: a number shown as whatever it is. Never a date. */
const GENERAL = 0;

/**
 * The formats every reader is born knowing, of the ones that are dates.
 *
 * 14 to 22 are the short dates and the times, 45 and 47 the minutes and
 * seconds of one, and the two runs in between — 27 to 36 and 50 to 58 — are
 * the same dates as written in Japanese, Chinese and Korean locales. A file
 * made in one of those locales holds ordinary dates under those ids, so
 * leaving them out would come back as bare serial numbers rather than as an
 * error.
 */
const BUILTIN_DATE_IDS = new Set([
    14, 15, 16, 17, 18, 19, 20, 21, 22,
    27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
    45, 47,
    50, 51, 52, 53, 54, 55, 56, 57, 58,
]);

/** The one built-in format that is a length of time: `[h]:mm:ss`. */
const BUILTIN_ELAPSED_IDS = new Set([46]);

/** The letters a format code shows a date or a time with. */
const DATE_LETTERS = new Set(['y', 'd', 'h', 's', 'm']);

/**
 * What a number under a format is:
 *
 * - `number` — a number, shown however the format says.
 * - `date` — a day, a day and its hour, or an hour of the day: a point on the
 *   calendar or on the clock.
 * - `elapsed` — a length of time, which is what `[h]:mm` shows: `1.25` under
 *   it is thirty hours, not the sixth hour of the first of January of 1900.
 */
export type FormatKind = 'number' | 'date' | 'elapsed';

/**
 * What a format code writes.
 *
 * A code is a template with literals in it, and the literals are the whole
 * difference between a format and its text: `#,##0 "días"` shows a number and
 * says `d` twice, and `[Red]0.00` says nothing at all through its `Red`. So
 * the scan walks the code and skips what a reader of it would skip — quoted
 * text, escaped characters, the padding of `_` and the fill of `*`, and
 * whatever is in brackets.
 *
 * Brackets are the one exception to being skipped: `[h]`, `[mm]` and `[ss]`
 * are elapsed time, which is the only thing in brackets that is a value
 * rather than an instruction. The rest — a colour, a condition, the `[$-409]`
 * of a locale — is not. And one of them is enough to make the whole code a
 * length of time, whatever else it says: the `mm:ss` after `[h]` are the
 * minutes and seconds of those hours, not of a day.
 *
 * `m` counts as a date letter although it is minutes as often as it is
 * months. Both are time, and which one it is only matters to whoever formats
 * the value, not to the reader deciding it is one.
 */
export function formatKind(code: string): FormatKind {
    let kind: FormatKind = 'number';
    for (let index = 0; index < code.length; index++) {
        const char = code[index] as string;
        if (char === '"') {
            const end = code.indexOf('"', index + 1);
            index = end === -1 ? code.length : end;
        } else if (char === '\\' || char === '_' || char === '*') {
            index++;
        } else if (char === '[') {
            const end = code.indexOf(']', index + 1);
            const inside = code.slice(index + 1, end === -1 ? code.length : end);
            if (/^[hms]+$/i.test(inside)) return 'elapsed';
            index = end === -1 ? code.length : end;
        } else if (DATE_LETTERS.has(char.toLowerCase())) kind = 'date';
    }
    return kind;
}

/** Whether a format code writes a date or a time of day — see `formatKind`. */
export function isDateFormat(code: string): boolean {
    return formatKind(code) === 'date';
}

/**
 * What the styles of a workbook say about the values under them.
 *
 * Indexed by the `s` of a cell, which is a position in `cellXfs` — the one
 * list of the part that matters here. `cellStyleXfs` is a list of the same
 * shape right next to it, holding the named styles that cell formats inherit
 * from, and reading the two as one would shift every index by however many
 * are in the first.
 */
export interface NumberFormats {
    /**
     * The format a cell with this style shows, as the writer takes one: the
     * code itself when the workbook declared it, the id when it is built in,
     * and nothing at all for `General`. What comes back can be handed
     * straight back to the writer.
     */
    forStyle(style: number | undefined): string | number | undefined;
    /** What a number under this style is: a number, a date, or a length of time. */
    kind(style: number | undefined): FormatKind;
    /** Whether a number under this style is a date rather than a number. */
    isDate(style: number | undefined): boolean;
}

/** A workbook with no styles part, where nothing is a date. */
export const NO_FORMATS: NumberFormats = {
    forStyle: () => undefined,
    kind: () => 'number',
    isDate: () => false,
};

export function readNumberFormats(xml: string): NumberFormats {
    /** The codes the workbook declared, by the id it gave them. */
    const codes = new Map<number, string>();
    /** The format id of each `cellXfs` entry, in order. */
    const styles: number[] = [];
    let inCellXfs = false;

    parseXml(
        xml,
        {
            open(name, attributes) {
                if (name === 'numFmt') {
                    const id = Number(attributes['numFmtId']);
                    const code = attributes['formatCode'];
                    if (Number.isInteger(id) && code !== undefined) codes.set(id, code);
                } else if (name === 'cellXfs') inCellXfs = true;
                else if (name === 'xf' && inCellXfs) {
                    const id = Number(attributes['numFmtId'] ?? GENERAL);
                    styles.push(Number.isInteger(id) ? id : GENERAL);
                }
            },
            close(name) {
                if (name === 'cellXfs') inCellXfs = false;
            },
        },
        PART_NAME,
    );

    /** The format id a style points at; `General` for a style there is none of. */
    function formatId(style: number | undefined): number {
        if (style === undefined) return GENERAL;
        return styles[style] ?? GENERAL;
    }

    // Worked out once per format rather than per cell: a sheet of a million
    // dates asks this a million times and there are only ever a handful of
    // answers.
    const kindById = new Map<number, FormatKind>();
    function kindOfId(id: number): FormatKind {
        const known = kindById.get(id);
        if (known !== undefined) return known;
        const code = codes.get(id);
        const answer =
            code !== undefined
                ? formatKind(code)
                : BUILTIN_ELAPSED_IDS.has(id)
                  ? 'elapsed'
                  : BUILTIN_DATE_IDS.has(id)
                    ? 'date'
                    : 'number';
        kindById.set(id, answer);
        return answer;
    }

    return {
        forStyle(style) {
            const id = formatId(style);
            if (id === GENERAL) return undefined;
            return codes.get(id) ?? id;
        },
        kind: (style) => kindOfId(formatId(style)),
        isDate: (style) => kindOfId(formatId(style)) === 'date',
    };
}
