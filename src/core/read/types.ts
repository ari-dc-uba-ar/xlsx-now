import type { Duration, PlainDate, PlainDateTime, PlainTime } from '../temporal.js';
import type { CellType, StyledCell } from '../types.js';
import type { FormatKind } from './numberFormats.js';

/**
 * A value as the reader gives it back: the four things a sheet holds, plus
 * whatever a number under a date format was asked to become.
 *
 * Deliberately narrower than the writer's `CellValue`, which is open to any
 * class the caller registered a type for. Nothing like that comes back out of
 * a file: what a sheet stores is a number, a string, a boolean or nothing,
 * and a date is the one of them the format gives a meaning to.
 *
 * Which of the date types it is, is the reader's `dates` option and nothing
 * else: they are all four in the union because the option is read at run time,
 * and a caller who picked one knows which one they picked. A `string` is in it
 * twice over, being both what a text cell holds and what `dates: 'isoString'`
 * gives back.
 *
 * `null` is a cell that is there and empty. A cell that is not there at all
 * is `undefined`, and the two are different on purpose — the same difference
 * the writer makes on the way in.
 */
export type ReadValue =
    | string
    | number
    | boolean
    | Date
    | PlainDate
    | PlainDateTime
    | PlainTime
    | Duration
    | null;

/** The style a cell is read with: for now, only its number format. */
export interface ReadStyle {
    /** The code the workbook declared, or the id of a built-in format. */
    numFmt: string | number;
}

/**
 * Everything the reader can say about a cell, of which a mode picks a part.
 *
 * Two kinds of field, and they never share a name. The ones that start with
 * `_` are the file's own, as the `<c>` spells them and with nothing made of
 * them: `_v` of a shared string is the index into the table, and `_v` of a
 * date is its serial in whatever epoch the workbook counts from. The rest are
 * what the reader makes of those, and they are named as the writer names
 * them, so that `v`, `s`, `f` and `t` go straight back into a workbook.
 */
export interface ReadCell {
    /** The value, built as `dates` says. */
    v: ReadValue;
    /** The number format, when it is not `General`. */
    s?: ReadStyle;
    /** The formula, without the `=` a sheet does not store. */
    f?: string;
    /** The type, only where the writer would not work it out: `str` or `e`. */
    t?: CellType;
    /**
     * The text the cell shows: the value under its format, as close as this
     * reader can come to what the person who made the file saw. Always there
     * when it is asked for. See `lang`.
     */
    w: string;
    /**
     * What the number of the cell is, by its format: a number, a date, or a
     * length of time. Only for a cell that holds a number or a date written
     * out; it is how a caller who reads dates as serials tells `45306` the day
     * from `45306` the amount.
     */
    kind?: FormatKind;
    /** The `t` attribute: what the cell says it holds. Absent means a number. */
    _t?: string;
    /** The `s` attribute: which entry of `cellXfs` the cell is shown with. */
    _s?: number;
    /** The text of `<v>`. */
    _v?: string;
    /** The text of `<f>`. */
    _f?: string;
    /** The text of `<is>`, the string a cell carries itself, runs joined. */
    _is?: string;
}

/** The name of one field of a `ReadCell`. */
export type ReadField = keyof ReadCell;

/**
 * Which fields of a `ReadCell` to read, as a mode: `{ w: true }` is a cell
 * with its value and its text. `v` is there unless it is turned off with
 * `v: false`, since a cell is read for its value nearly every time; every
 * other field is there only when it is asked for.
 */
export type ReadFields = { readonly [K in ReadField]?: boolean };

/**
 * What comes back for each cell, chosen by `mode`:
 *
 * - `values` — the value alone, which is what most callers want.
 * - `cells` — the `StyledCell` the writer takes, carrying the value in `v`,
 *   the formula in `f`, and the number format in `s`. Which means what a
 *   reader gives back can be handed straight to a writer.
 * - an object of `ReadFields` — the cell with the fields it names. See
 *   `ReadCell`.
 */
export type ReadMode = 'values' | 'cells' | ReadFields;

/** What each of the two named modes gives back per cell. */
export interface ReadModes {
    values: ReadValue;
    cells: StyledCell;
}

/** The part of a `ReadCell` a `ReadFields` asks for. */
export type PickedCell<M extends ReadFields> = {
    [K in keyof ReadCell as K extends 'v'
        ? M extends { readonly v: false }
            ? never
            : K
        : M extends { readonly [P in K]: true }
          ? K
          : never]: ReadCell[K];
};

/** What a mode gives back per cell. */
export type CellOf<M extends ReadMode> = M extends 'values'
    ? ReadValue
    : M extends 'cells'
      ? StyledCell
      : M extends ReadFields
        ? PickedCell<M>
        : never;

/** The mode `readXlsx` reads with when it was told nothing: the value alone. */
export const READ_VALUES = 'values';
/** The mode whose cells go straight back into a writer. */
export const READ_CELLS = 'cells';
/** The cell as the file spells it, and nothing made of it. */
export const READ_RAW = { v: false, _t: true, _s: true, _v: true, _f: true, _is: true } as const;

/** One row of a sheet, as it is read. */
export interface ReadRow<C> {
    /** The row number the sheet gives it, counting from 1. */
    index: number;
    /**
     * The cells, by column index counting from 0. A position no cell was
     * written in is `undefined`, so a row with a hole in the middle has one
     * here too, and the array ends at the last cell the row actually has.
     */
    cells: (C | undefined)[];
}

/** One worksheet, read whole. */
export interface SheetData<C> {
    /** The name the workbook gives the sheet. */
    name: string;
    /**
     * The grid, by row and then by column, both counting from 0 — so `A1` is
     * `cells[0]?.[0]`.
     *
     * Dense in rows: there is an entry for every row up to the last one that
     * holds anything, and a row that holds nothing is an empty array. Ragged
     * in columns: each row ends at its own last cell, and `maxCol` is what
     * says how wide the sheet is as a whole.
     */
    cells: (C | undefined)[][];
    /** How many columns the widest row of the sheet reaches. */
    maxCol: number;
    /** How many rows the sheet reaches: the same as `cells.length`. */
    maxRow: number;
    /**
     * Whether the workbook counts its days from 1904 — which only matters to
     * whoever reads a serial out of `_v`, since every other field has already
     * taken it into account.
     */
    date1904: boolean;
}
