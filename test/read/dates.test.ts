// The four answers to one number.
//
// The environment these tests run in has no `Temporal` of its own, so the
// polyfill is installed for the whole suite by `.mocharc.json` — before any of
// the package loads, which is the one order this package counts on. What is
// imported here is the same classes the global carries, for their types: the
// polyfill installs the global but declares nothing for it, and these tests
// are the one place that needs to name `Temporal.PlainDate` in TypeScript.
import assert from 'node:assert/strict';
import { Temporal } from 'temporal-polyfill';
import { isoDuration, readDate, readDates, readDuration, readIsoDate } from '../../src/core/read/dates.js';

/** 15/01/2024, and the same day at half past twelve. */
const DAY = 45306;
const NOON_AND_A_HALF = 45306.5;
/** Half past ten in the morning, of no day at all. */
const TIME = 0.4375;

describe('readDates', () => {
    it('reads dates as Temporal values when nothing was asked for', () => {
        assert.equal(readDates(undefined), 'temporal');
    });

    it('says what the modes are when it is handed something else', () => {
        assert.throws(
            () => readDates('plainDate' as 'temporal'),
            /"plainDate" is not how a date can be read.*temporal, utcDate, localDate, isoString/s,
        );
    });

    it('says so up front when temporal was asked for and there is no Temporal', async () => {
        const global = globalThis as { Temporal?: unknown };
        const temporal = global.Temporal;
        delete global.Temporal;
        try {
            // A second copy of the module, loaded the way an environment with
            // no `Temporal` would load it — which is the only way to be
            // without one, since the global is read when the module loads and
            // never again.
            const path = '../../src/core/temporal.js';
            const fresh = (await import(`${path}?no-temporal`)) as typeof import(
                '../../src/core/temporal.js'
            );
            assert.equal(fresh.temporalApi, undefined);
            assert.throws(() => fresh.requireTemporal(), /Temporal\.PlainDate/);
            assert.throws(
                () => fresh.requireTemporal(),
                /dates: "utcDate", "localDate" or "isoString"/,
            );
        } finally {
            global.Temporal = temporal;
        }

        // And the three modes that need no Temporal never ask for one.
        assert.equal(readDate(DAY, 'isoString'), '2024-01-15');
        assert.ok(readDate(DAY, 'utcDate') instanceof Date);
    });
});

describe('readDate: temporal', () => {
    it('builds a PlainDate out of a whole day', () => {
        const date = readDate(DAY, 'temporal') as Temporal.PlainDate;
        assert.ok(date instanceof Temporal.PlainDate);
        assert.equal(date.toString(), '2024-01-15');
    });

    it('builds a PlainDateTime out of a day with an hour in it', () => {
        const when = readDate(NOON_AND_A_HALF, 'temporal') as Temporal.PlainDateTime;
        assert.ok(when instanceof Temporal.PlainDateTime);
        assert.equal(when.toString(), '2024-01-15T12:00:00');
    });

    it('builds a PlainTime out of a serial with no day left in it', () => {
        const time = readDate(TIME, 'temporal') as Temporal.PlainTime;
        assert.ok(time instanceof Temporal.PlainTime);
        assert.equal(time.toString(), '10:30:00');
    });

    it('keeps the seconds a serial carries', () => {
        // 12:34:56 of the same day, as the fraction of a day it is.
        const serial = DAY + (12 * 3600 + 34 * 60 + 56) / 86400;
        assert.equal(String(readDate(serial, 'temporal')), '2024-01-15T12:34:56');
    });
});

describe('readDate: isoString', () => {
    it('is the text the Temporal values are built from', () => {
        assert.equal(readDate(DAY, 'isoString'), '2024-01-15');
        assert.equal(readDate(NOON_AND_A_HALF, 'isoString'), '2024-01-15T12:00:00');
        assert.equal(readDate(TIME, 'isoString'), '10:30:00');
    });

    it('keeps a millisecond where there is one, and writes none where there is not', () => {
        assert.equal(readDate(DAY + 0.5 + 1 / 86400000, 'isoString'), '2024-01-15T12:00:00.001');
    });
});

describe('readDate: the two Dates', () => {
    it('reads the wall clock in UTC, whatever zone the reader is in', () => {
        const utc = readDate(NOON_AND_A_HALF, 'utcDate') as Date;
        assert.equal(utc.toISOString(), '2024-01-15T12:00:00.000Z');
    });

    it('reads the same wall clock locally', () => {
        const local = readDate(NOON_AND_A_HALF, 'localDate') as Date;
        assert.equal(local.getFullYear(), 2024);
        assert.equal(local.getMonth(), 0);
        assert.equal(local.getDate(), 15);
        assert.equal(local.getHours(), 12);
    });

    it('is the same clock said twice: the two differ by the zone and nothing else', () => {
        const local = readDate(NOON_AND_A_HALF, 'localDate') as Date;
        const utc = readDate(NOON_AND_A_HALF, 'utcDate') as Date;
        assert.equal(local.getTime() - local.getTimezoneOffset() * 60000, utc.getTime());
        assert.equal(local.getHours(), utc.getUTCHours());
    });

    it('refuses the day the calendar does not have, whichever mode asked', () => {
        for (const mode of ['temporal', 'utcDate', 'localDate', 'isoString'] as const) {
            assert.throws(() => readDate(60, mode), RangeError, mode);
        }
    });

    it('reads a serial below zero as the day it counts back to, whichever mode asked', () => {
        assert.equal(String(readDate(-18091, 'temporal')), '1850-06-20');
        assert.equal(readDate(-18091, 'isoString'), '1850-06-20');
        assert.equal(readDate(-0.5, 'isoString'), '1899-12-30T12:00:00');
        assert.equal((readDate(-18091, 'utcDate') as Date).toISOString(), '1850-06-20T00:00:00.000Z');
        const local = readDate(-18091, 'localDate') as Date;
        assert.deepEqual([local.getFullYear(), local.getMonth(), local.getDate()], [1850, 5, 20]);
    });
});

describe('the serial mode', () => {
    it('gives the number back as it is, which is what the sheet stores', () => {
        assert.equal(readDate(DAY, 'serial'), DAY);
        assert.equal(readDate(NOON_AND_A_HALF, 'serial'), NOON_AND_A_HALF);
        assert.equal(readDate(TIME, 'serial'), TIME);
        // The day Excel has and the calendar does not has a number all the same.
        assert.equal(readDate(60, 'serial'), 60);
    });

    it('numbers a date written out in text, below zero for a day no serial reaches', () => {
        assert.equal(readIsoDate('2024-01-15', 'serial'), DAY);
        assert.equal(readIsoDate('2024-01-15T12:00:00Z', 'serial'), NOON_AND_A_HALF);
        assert.equal(readIsoDate('1900-03-01', 'serial'), 61);
        assert.equal(readIsoDate('1900-01-01', 'serial'), 1);
        assert.equal(readIsoDate('1899-12-31', 'serial'), 0);
        assert.equal(readIsoDate('1899-12-30', 'serial'), -1);
        // 18091 days before 31/12/1899.
        assert.equal(readIsoDate('1850-06-20', 'serial'), -18091);
    });

    it('refuses a date written out in text that is not one', () => {
        assert.throws(() => readIsoDate('ayer', 'serial'), /not a date/);
    });
});

describe('isoDuration', () => {
    it('is a length in days as an ISO duration, counted in hours', () => {
        assert.equal(isoDuration(1.25), 'PT30H');
        assert.equal(isoDuration(0.5 / 24), 'PT30M');
        assert.equal(isoDuration(1 + 1 / 24 + 1 / 1440 + 1.5 / 86400), 'PT25H1M1.5S');
        assert.equal(isoDuration(0), 'PT0S');
        assert.equal(isoDuration(-1.25), '-PT30H');
    });

    it('rounds to the millisecond, so a fraction of a day does not come back a hair short', () => {
        assert.equal(isoDuration(0.5 / 24 + 0.001 / 86400), 'PT30M0.001S');
        assert.equal(isoDuration(1 / 3), 'PT8H');
    });
});

describe('readDuration', () => {
    it('is a Duration, its ISO text, or the number, and never a date', () => {
        const duration = readDuration(1.25, 'temporal');
        assert.ok(duration instanceof Temporal.Duration);
        assert.equal(duration.hours, 30);
        assert.equal(readDuration(1.25, 'isoString'), 'PT30H');
        for (const mode of ['utcDate', 'localDate', 'serial'] as const) {
            assert.equal(readDuration(1.25, mode), 1.25, mode);
        }
    });
});
