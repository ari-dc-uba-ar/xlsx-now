// How a date comes out of a file.
//
// A sheet stores a date as a number and nothing else, so what a reader gives
// back for one is a choice, not a reading: the same `45306.5` is a
// `Temporal.PlainDateTime`, a `Date`, the text `2024-01-15T12:00:00`, or
// the number itself, depending on what the caller asked for. This is the
// whole of that choice — five answers to one number, told apart by five `if`s.
import { fromExcelSerial, fromExcelSerialUtc, isoSerial, serialKind } from '../cell.js';
import { requireTemporal } from '../temporal.js';
import type { ReadValue } from './types.js';

/**
 * What a date cell is built as, as the `dates` option of the reader says it:
 *
 * - `temporal` — a `Temporal.PlainDate`, `PlainDateTime` or `PlainTime`,
 *   whichever the number turns out to be. The default, and the only one of the
 *   four where nothing about the value is left to the reader's own clock: a
 *   day is a day, an hour is an hour, and neither carries an instant it never
 *   had. It needs a `Temporal` in the environment — native or a polyfill — and
 *   says so at `openXlsx` when there is none.
 * - `utcDate` — a `Date` whose *UTC* reading is what the cell shows:
 *   `getUTCHours()` gives the hour in the sheet.
 * - `localDate` — a `Date` whose *local* reading is what the cell shows:
 *   `getHours()` gives the hour in the sheet. It is the same date the writer
 *   takes back by default, and it is the only one of the four that depends on
 *   the zone the reader happens to run in.
 * - `isoString` — the wall clock as text: `2024-01-15`, `2024-01-15T12:00:00`
 *   or `12:00:00`. What `temporal` builds its values out of, for whoever wants
 *   the text and no class at all.
 * - `serial` — the number itself, which is what the sheet stores: nothing is
 *   built, so there is nothing to convert twice. Always counted from 1900,
 *   whatever epoch the workbook declared, and a date written out in text
 *   (`t="d"`) becomes the number it would have had — below zero for a day
 *   before 31/12/1899.
 *
 * A length of time — a number under `[h]:mm` — is not a date and is built as
 * one of its own kind: a `Temporal.Duration` under `temporal`, the ISO
 * duration (`PT30H`) under `isoString`, and the number under the other three,
 * since a `Date` has nothing to say a length with.
 */
export type ReadDates = 'temporal' | 'utcDate' | 'localDate' | 'isoString' | 'serial';

const MODES: readonly ReadDates[] = ['temporal', 'utcDate', 'localDate', 'isoString', 'serial'];

/**
 * The `dates` a package was opened with, checked once and up front — including
 * the `Temporal` that `temporal` needs, so a workbook that cannot be read the
 * way it was asked for fails at `openXlsx` and not at the first date of the
 * first sheet.
 */
export function readDates(option: ReadDates | undefined): ReadDates {
    const dates = option ?? 'temporal';
    if (!MODES.includes(dates)) {
        throw new Error(
            `dates: "${String(dates)}" is not how a date can be read: say ${MODES.join(', ')}.`,
        );
    }
    if (dates === 'temporal') requireTemporal();
    return dates;
}

/**
 * The serial of a date cell, as the value the caller asked for.
 *
 * The three modes that are not `localDate` are built from the one `Date` that
 * does not depend on where it is read — the wall clock taken as UTC — and from
 * its ISO text, which is already exactly what the sheet shows. Which is why
 * there is no arithmetic below and no time zone in it: past `fromExcelSerialUtc`
 * there is nothing left to move.
 */
export function readDate(serial: number, dates: ReadDates): ReadValue {
    if (dates === 'serial') return serial;
    if (dates === 'localDate') return fromExcelSerial(serial);
    const utc = fromExcelSerialUtc(serial);
    if (dates === 'utcDate') return utc;
    // `2024-01-15T12:00:00.000Z`, of which every mode below takes a slice.
    const iso = utc.toISOString();
    const kind = serialKind(serial);
    const text =
        kind === 'date'
            ? iso.slice(0, 10)
            : kind === 'time'
              ? timeOf(iso)
              : `${iso.slice(0, 10)}T${timeOf(iso)}`;
    if (dates === 'isoString') return text;
    const temporal = requireTemporal();
    if (kind === 'date') return temporal.PlainDate.from(text);
    if (kind === 'time') return temporal.PlainTime.from(text);
    return temporal.PlainDateTime.from(text);
}

/** The time of an ISO instant, without the milliseconds when it has none. */
function timeOf(iso: string): string {
    const time = iso.slice(11, 23);
    return time.endsWith('.000') ? time.slice(0, -4) : time;
}

/** A zone written next to a date: a `Z`, or an offset from it. */
export const ZONE = /(?:[Zz]|[+-]\d\d:?\d\d)$/;

/**
 * The other kind of date cell: the one that spells its day out in ISO text
 * instead of numbering it.
 *
 * It goes nowhere near a serial, and that is the point of it. A serial is what
 * a sheet stores a date as and what this reader is built around, but it cannot
 * number a day before 31/12/1899 — there is no negative serial — so a file
 * that has one has nothing else to write it as. Turning the text into a serial
 * here would throw away the one thing the text can do that the number cannot.
 * Except under `serial`, where the caller asked for the number in so many
 * words, and gets the one the numbering would have reached carried on below
 * zero — see `isoSerial`.
 *
 * So the text is what each mode is built from, the same way the serial is
 * above: `Temporal` takes it as it is written, and `Date` is what checks it.
 * A zone on the text is dropped rather than applied — a sheet has no time
 * zone, so a `Z` there is not something the cell can mean — and what is left
 * is the wall clock every other date in the file is too.
 */
export function readIsoDate(text: string, dates: ReadDates): ReadValue {
    const wall = text.replace(ZONE, '');
    if (dates === 'temporal') {
        const temporal = requireTemporal();
        return wall.includes('T')
            ? temporal.PlainDateTime.from(wall)
            : temporal.PlainDate.from(wall);
    }
    // A day on its own gets a midnight, so that `localDate` reads it as one:
    // `Date` takes a bare date as UTC and a date with an hour on it as local,
    // and only the second of those is what that mode means.
    const full = wall.includes('T') ? wall : `${wall}T00:00:00`;
    const date = new Date(dates === 'utcDate' ? `${full}Z` : full);
    if (Number.isNaN(date.getTime())) {
        throw new Error(`A date cell holds "${text}", which is not a date.`);
    }
    if (dates === 'serial') return isoSerial(wall);
    return dates === 'isoString' ? wall : date;
}

const MS_PER_DAY = 86400000;
const MS_PER_HOUR = 3600000;
const MS_PER_MINUTE = 60000;
const MS_PER_SECOND = 1000;

/**
 * A length of time, given in days, as an ISO duration: `1.25` is `PT30H`.
 *
 * In hours and not in days — `PT30H` rather than `P1DT6H` — because hours are
 * what an elapsed format counts in, and a day in a duration is a calendar day
 * that some days is not twenty-four hours long. Rounded to the millisecond,
 * for the same reason a date is: a fraction of a day taken at face value comes
 * back a millisecond short of what went in.
 */
export function isoDuration(days: number): string {
    let rest = Math.round(Math.abs(days) * MS_PER_DAY);
    const hours = Math.floor(rest / MS_PER_HOUR);
    rest -= hours * MS_PER_HOUR;
    const minutes = Math.floor(rest / MS_PER_MINUTE);
    rest -= minutes * MS_PER_MINUTE;
    const seconds = Math.floor(rest / MS_PER_SECOND);
    const milliseconds = rest - seconds * MS_PER_SECOND;
    let text = '';
    if (hours) text += `${hours}H`;
    if (minutes) text += `${minutes}M`;
    if (seconds || milliseconds) {
        const fraction = milliseconds ? `.${String(milliseconds).padStart(3, '0').replace(/0+$/, '')}` : '';
        text += `${seconds}${fraction}S`;
    }
    return `${days < 0 ? '-' : ''}PT${text || '0S'}`;
}

/**
 * The number under an elapsed format, as the value the caller asked for.
 *
 * Not a date in any of the modes: a `Duration` where there are Temporal
 * values, its ISO text where there is text, and where the mode is a `Date` —
 * which has no way to say a length — or the number itself, the number.
 */
export function readDuration(days: number, dates: ReadDates): ReadValue {
    if (dates === 'isoString') return isoDuration(days);
    if (dates === 'temporal') return requireTemporal().Duration.from(isoDuration(days));
    return days;
}
