// What a cell shows: its value under its format, as text.
//
// A file does not store this. It stores the value and the code of the format,
// and the text is what the application that opens it works out from the two —
// with the language of the machine it runs on filling in whatever the code
// leaves to it: the names of the months, the separators of a number, and the
// whole of the short date, which is built-in format 14 and has no code at all.
//
// So this is an approximation, and it is one on purpose rather than by
// omission: a cell asked for its text always has one. Where the code says
// something this does not take apart — a fraction, a condition — the value is
// written plainly instead, and where a date is shown and nobody said in which
// language, it is written in ISO, which is the one way of writing a date that
// nobody can read as a different one.
import { serialKind } from '../cell.js';

const MS_PER_DAY = 86400000;
const MS_PER_HOUR = 3600000;
const MS_PER_MINUTE = 60000;
const MS_PER_SECOND = 1000;

/** Days between 1899-12-31, which is serial 0, and 1970-01-01. */
const EXCEL_EPOCH_OFFSET_DAYS = 25569;
/** Serial 60, 29/02/1900: the day Excel has and the calendar does not. */
const PHANTOM_LEAP_DAY_SERIAL = 60;

/** What a language says that a format leaves to it. */
export interface Locale {
    /** The separator of the decimals. */
    decimal: string;
    /** The separator of the thousands. */
    group: string;
    months: readonly string[];
    monthsShort: readonly string[];
    /** From Sunday, as `getUTCDay` counts. */
    days: readonly string[];
    daysShort: readonly string[];
    am: string;
    pm: string;
    /** The code of built-in format 14, the short date of the machine. */
    shortDate: string;
    true: string;
    false: string;
}

const ENGLISH: Locale = {
    decimal: '.',
    group: ',',
    months: [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December',
    ],
    monthsShort: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
    days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
    daysShort: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
    am: 'AM',
    pm: 'PM',
    shortDate: 'mm/dd/yyyy',
    true: 'TRUE',
    false: 'FALSE',
};

const SPANISH: Locale = {
    decimal: ',',
    group: '.',
    months: [
        'enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
        'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre',
    ],
    monthsShort: ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'],
    days: ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'],
    daysShort: ['dom', 'lun', 'mar', 'mié', 'jue', 'vie', 'sáb'],
    am: 'a. m.',
    pm: 'p. m.',
    shortDate: 'dd/mm/yyyy',
    true: 'VERDADERO',
    false: 'FALSO',
};

/**
 * The locale a `lang` names: Spanish for `es` and any `es-…`, English for
 * every other. With no `lang` at all it is English too, but the dates are
 * not shown through it — see `displayDate`.
 */
export function localeFor(lang: string | undefined): Locale {
    return lang !== undefined && /^es(?:$|[-_])/i.test(lang) ? SPANISH : ENGLISH;
}

/**
 * The code of a built-in format, which a file names by its id alone.
 *
 * The ones that depend on the machine — the short date, and the dates of the
 * Japanese, Chinese and Korean locales — are taken as the short date of the
 * language asked for. An id this does not know is shown as `General`.
 */
export function builtinCode(id: number, locale: Locale): string {
    switch (id) {
        case 1: return '0';
        case 2: return '0.00';
        case 3: return '#,##0';
        case 4: return '#,##0.00';
        case 5: return '"$"#,##0_);("$"#,##0)';
        case 6: return '"$"#,##0_);[Red]("$"#,##0)';
        case 7: return '"$"#,##0.00_);("$"#,##0.00)';
        case 8: return '"$"#,##0.00_);[Red]("$"#,##0.00)';
        case 9: return '0%';
        case 10: return '0.00%';
        case 11: return '0.00E+00';
        case 12: return '# ?/?';
        case 13: return '# ??/??';
        case 15: return 'd-mmm-yy';
        case 16: return 'd-mmm';
        case 17: return 'mmm-yy';
        case 18: return 'h:mm AM/PM';
        case 19: return 'h:mm:ss AM/PM';
        case 20: return 'h:mm';
        case 21: return 'h:mm:ss';
        case 22: return `${locale.shortDate} hh:mm`;
        case 37: return '#,##0 ;(#,##0)';
        case 38: return '#,##0 ;[Red](#,##0)';
        case 39: return '#,##0.00;(#,##0.00)';
        case 40: return '#,##0.00;[Red](#,##0.00)';
        case 45: return 'mm:ss';
        case 46: return '[h]:mm:ss';
        case 47: return 'mmss.0';
        case 48: return '##0.0E+0';
        case 49: return '@';
    }
    if (id === 14 || (id >= 27 && id <= 36) || (id >= 50 && id <= 58)) return locale.shortDate;
    return 'General';
}

/** The code a style's `numFmt` stands for. */
export function codeOf(numFmt: string | number | undefined, locale: Locale): string {
    if (numFmt === undefined) return 'General';
    return typeof numFmt === 'number' ? builtinCode(numFmt, locale) : numFmt;
}

/**
 * A code cut into its sections — positive, negative, zero, text — at the
 * semicolons that are not text of its own.
 */
function sections(code: string): string[] {
    const found: string[] = [];
    let start = 0;
    for (let index = 0; index < code.length; index++) {
        const char = code[index];
        if (char === '"') {
            const end = code.indexOf('"', index + 1);
            index = end === -1 ? code.length : end;
        } else if (char === '\\' || char === '_' || char === '*') {
            index++;
        } else if (char === '[') {
            const end = code.indexOf(']', index + 1);
            index = end === -1 ? code.length : end;
        } else if (char === ';') {
            found.push(code.slice(start, index));
            start = index + 1;
        }
    }
    found.push(code.slice(start));
    return found;
}

/** A number as `General` shows it: as it is, with the language's decimals. */
export function generalText(value: number, locale: Locale): string {
    return String(value).replace('e', 'E').replace('.', locale.decimal);
}

// ---------------------------------------------------------------------------
// Dates, times and lengths of time.
// ---------------------------------------------------------------------------

type DateToken =
    | { kind: 'literal'; text: string }
    | { kind: 'year' | 'month' | 'day' | 'hour' | 'minute' | 'second' | 'fraction'; length: number }
    | { kind: 'monthOrMinute'; length: number }
    | { kind: 'elapsed'; unit: 'h' | 'm' | 's'; length: number }
    | { kind: 'ampm'; am: string; pm: string };

/** How many times the character at `index` repeats from there. */
function runOf(code: string, index: number): number {
    const char = (code[index] as string).toLowerCase();
    let end = index;
    while (end < code.length && (code[end] as string).toLowerCase() === char) end++;
    return end - index;
}

/** A date code, taken apart. Only its first section is read: a date is never negative. */
function dateTokens(code: string): DateToken[] {
    const text = sections(code)[0] ?? '';
    const tokens: DateToken[] = [];
    for (let index = 0; index < text.length; index++) {
        const char = text[index] as string;
        const lower = char.toLowerCase();
        if (char === '"') {
            const end = text.indexOf('"', index + 1);
            const stop = end === -1 ? text.length : end;
            tokens.push({ kind: 'literal', text: text.slice(index + 1, stop) });
            index = stop;
        } else if (char === '\\') {
            tokens.push({ kind: 'literal', text: text[index + 1] ?? '' });
            index++;
        } else if (char === '_') {
            tokens.push({ kind: 'literal', text: ' ' });
            index++;
        } else if (char === '*') {
            index++;
        } else if (char === '[') {
            const end = text.indexOf(']', index + 1);
            const stop = end === -1 ? text.length : end;
            const inside = text.slice(index + 1, stop);
            const elapsed = /^(h+|m+|s+)$/i.exec(inside);
            if (elapsed) {
                const unit = (inside[0] as string).toLowerCase() as 'h' | 'm' | 's';
                tokens.push({ kind: 'elapsed', unit, length: inside.length });
            } else if (inside.startsWith('$')) {
                // `[$€-C0A]`: a currency symbol, and the locale it belongs to.
                tokens.push({ kind: 'literal', text: inside.slice(1).split('-')[0] ?? '' });
            }
            index = stop;
        } else if (text.slice(index, index + 5).toUpperCase() === 'AM/PM') {
            tokens.push({ kind: 'ampm', am: 'AM', pm: 'PM' });
            index += 4;
        } else if (text.slice(index, index + 3).toUpperCase() === 'A/P') {
            tokens.push({ kind: 'ampm', am: text[index] as string, pm: text[index + 2] as string });
            index += 2;
        } else if (lower === 'y' || lower === 'e') {
            const length = runOf(text, index);
            tokens.push({ kind: 'year', length: lower === 'e' ? 4 : length });
            index += length - 1;
        } else if (lower === 'm') {
            const length = runOf(text, index);
            tokens.push({ kind: length >= 3 ? 'month' : 'monthOrMinute', length });
            index += length - 1;
        } else if (lower === 'd' || lower === 'h' || lower === 's') {
            const length = runOf(text, index);
            tokens.push({ kind: lower === 'd' ? 'day' : lower === 'h' ? 'hour' : 'second', length });
            index += length - 1;
        } else if (
            char === '.' &&
            text[index + 1] === '0' &&
            (tokens.at(-1)?.kind === 'second' || isElapsed(tokens.at(-1), 's'))
        ) {
            const length = runOf(text, index + 1);
            tokens.push({ kind: 'fraction', length });
            index += length;
        } else {
            tokens.push({ kind: 'literal', text: char });
        }
    }
    // `m` is the month unless it sits next to an hour or a second, which is
    // how Excel itself tells `mm/dd` from `hh:mm`.
    return tokens.map((token, index) => {
        if (token.kind !== 'monthOrMinute') return token;
        const before = tokens.slice(0, index).reverse().find((other) => other.kind !== 'literal');
        const after = tokens.slice(index + 1).find((other) => other.kind !== 'literal');
        const minute =
            before?.kind === 'hour' ||
            isElapsed(before, 'h') ||
            after?.kind === 'second' ||
            isElapsed(after, 's');
        return { kind: minute ? 'minute' : 'month', length: token.length };
    });
}

function isElapsed(token: DateToken | undefined, unit: 'h' | 'm' | 's'): boolean {
    return token?.kind === 'elapsed' && token.unit === unit;
}

/** A day of the calendar, as a date code shows one. */
interface CalendarDay {
    year: number;
    month: number;
    day: number;
    /** From Sunday, 0 to 6. */
    weekday: number;
}

/**
 * The day a whole serial is, from 1900. Serial 60 is the 29th of February of
 * 1900, which the calendar does not have and a sheet does — so it is shown,
 * as Excel shows it, rather than refused. A serial below zero is the count
 * carried on backwards, the same one `isoSerial` gives: Excel shows `######`
 * there, but the day is known, and a text that can say it does.
 */
function dayOfSerial(serial: number): CalendarDay {
    if (serial === PHANTOM_LEAP_DAY_SERIAL) return { year: 1900, month: 2, day: 29, weekday: 3 };
    const days = serial < PHANTOM_LEAP_DAY_SERIAL ? serial + 1 : serial;
    const date = new Date((days - EXCEL_EPOCH_OFFSET_DAYS) * MS_PER_DAY);
    return {
        year: date.getUTCFullYear(),
        month: date.getUTCMonth() + 1,
        day: date.getUTCDate(),
        weekday: date.getUTCDay(),
    };
}

const pad = (value: number, length: number): string => String(value).padStart(length, '0');

/**
 * The milliseconds a code rounds to: the decimals of its seconds when it
 * shows them, and the second otherwise — which is what Excel does, so
 * `10:20:59.6` under `hh:mm` is `10:21` and not `10:20`.
 */
function roundingOf(tokens: readonly DateToken[]): number {
    for (const token of tokens) {
        if (token.kind === 'fraction') return 10 ** Math.max(0, 3 - token.length);
    }
    return MS_PER_SECOND;
}

/**
 * Tokens filled in with a day and a time.
 *
 * `time` is the milliseconds the time part holds: the hour of the day for a
 * date, and the whole length for an elapsed code. The same arithmetic serves
 * both, because the one thing that changes is which field counts in full —
 * the `[h]` of an elapsed code — and that is a token of its own.
 */
function renderDate(
    tokens: readonly DateToken[],
    day: CalendarDay | undefined,
    time: number,
    locale: Locale,
): string {
    const twelve = tokens.some((token) => token.kind === 'ampm');
    const hours = Math.floor(time / MS_PER_HOUR);
    let text = '';
    for (const token of tokens) {
        switch (token.kind) {
            case 'literal':
                text += token.text;
                break;
            case 'year':
                if (day) text += token.length <= 2 ? pad(day.year % 100, 2) : pad(day.year, 4);
                break;
            case 'month':
                if (!day) break;
                if (token.length <= 2) text += pad(day.month, token.length);
                else if (token.length === 3) text += locale.monthsShort[day.month - 1];
                else if (token.length === 4) text += locale.months[day.month - 1];
                else text += (locale.months[day.month - 1] ?? '').charAt(0);
                break;
            case 'day':
                if (!day) break;
                if (token.length <= 2) text += pad(day.day, token.length);
                else if (token.length === 3) text += locale.daysShort[day.weekday];
                else text += locale.days[day.weekday];
                break;
            case 'hour': {
                const hour = hours % 24;
                const shown = twelve ? hour % 12 || 12 : hour;
                text += pad(shown, Math.min(token.length, 2));
                break;
            }
            case 'minute':
            case 'monthOrMinute':
                text += pad(Math.floor(time / MS_PER_MINUTE) % 60, Math.min(token.length, 2));
                break;
            case 'second':
                text += pad(Math.floor(time / MS_PER_SECOND) % 60, Math.min(token.length, 2));
                break;
            case 'fraction': {
                const digits = pad(time % MS_PER_SECOND, 3).slice(0, token.length);
                text += locale.decimal + digits;
                break;
            }
            case 'elapsed': {
                const unit =
                    token.unit === 'h' ? MS_PER_HOUR : token.unit === 'm' ? MS_PER_MINUTE : MS_PER_SECOND;
                text += pad(Math.floor(time / unit), token.length);
                break;
            }
            case 'ampm': {
                const morning = hours % 24 < 12;
                const spelled = token.am === 'AM' ? (morning ? locale.am : locale.pm) : morning ? token.am : token.pm;
                text += spelled;
                break;
            }
        }
    }
    return text;
}

/** A serial under a date code, as the code shows it. */
export function formatDateSerial(serial: number, code: string, locale: Locale): string {
    const tokens = dateTokens(code);
    const unit = roundingOf(tokens);
    const total = Math.round((serial * MS_PER_DAY) / unit) * unit;
    const whole = Math.floor(total / MS_PER_DAY);
    return renderDate(tokens, dayOfSerial(whole), total - whole * MS_PER_DAY, locale);
}

/** A length of time, in days, under an elapsed code: `1.25` under `[h]:mm` is `30:00`. */
export function formatElapsed(days: number, code: string, locale: Locale): string {
    const tokens = dateTokens(code);
    const unit = roundingOf(tokens);
    const total = Math.round((Math.abs(days) * MS_PER_DAY) / unit) * unit;
    return (days < 0 ? '-' : '') + renderDate(tokens, undefined, total, locale);
}

/** An ISO wall clock, `2024-01-15` or `2024-01-15T10:30:00.5`, taken apart. */
const ISO_WALL = /^(-?\d{4,})-(\d\d)-(\d\d)(?:T(\d\d):(\d\d)(?::(\d\d)(?:\.(\d+))?)?)?$/;

/**
 * A date written out in ISO text (`t="d"`) under a date code. The text says
 * the day itself, so this is the one date that has no serial to go through —
 * which is the point of a cell like that: the day may be one no serial
 * reaches.
 */
export function formatDateText(wall: string, code: string, locale: Locale): string {
    const match = ISO_WALL.exec(wall);
    if (!match) return wall;
    const [, year, month, day, hour, minute, second, fraction] = match;
    const date = new Date(0);
    date.setUTCFullYear(Number(year), Number(month) - 1, Number(day));
    const time =
        Number(hour ?? 0) * MS_PER_HOUR +
        Number(minute ?? 0) * MS_PER_MINUTE +
        Number(second ?? 0) * MS_PER_SECOND +
        Math.round(Number(`0.${fraction ?? '0'}`) * MS_PER_SECOND);
    return renderDate(
        dateTokens(code),
        { year: Number(year), month: Number(month), day: Number(day), weekday: date.getUTCDay() },
        time,
        locale,
    );
}

/**
 * A serial as ISO text, the way a date is shown when no language was given:
 * the day, the day and the hour, or the hour, as the number says — the same
 * text `dates: 'isoString'` reads it as.
 */
export function isoOfSerial(serial: number): string {
    const total = Math.round(serial * MS_PER_DAY);
    const whole = Math.floor(total / MS_PER_DAY);
    const day = dayOfSerial(whole);
    const ms = total - whole * MS_PER_DAY;
    const date = `${pad(day.year, 4)}-${pad(day.month, 2)}-${pad(day.day, 2)}`;
    const millis = ms % MS_PER_SECOND;
    const time =
        `${pad(Math.floor(ms / MS_PER_HOUR), 2)}:${pad(Math.floor(ms / MS_PER_MINUTE) % 60, 2)}:` +
        `${pad(Math.floor(ms / MS_PER_SECOND) % 60, 2)}${millis ? `.${pad(millis, 3)}` : ''}`;
    const kind = serialKind(serial);
    return kind === 'date' ? date : kind === 'time' ? time : `${date}T${time}`;
}

// ---------------------------------------------------------------------------
// Numbers.
// ---------------------------------------------------------------------------

type NumberToken =
    | { kind: 'literal'; text: string }
    | { kind: 'digit'; char: '0' | '#' | '?' }
    | { kind: 'point' }
    | { kind: 'exponent'; letter: string; sign: '+' | '-' }
    | { kind: 'general' };

/** A number section, taken apart; `undefined` for what this does not read. */
interface NumberSection {
    tokens: NumberToken[];
    /** Whether the integer part is grouped in thousands. */
    grouped: boolean;
    /** How many thousands the value is divided by: one per trailing comma. */
    scale: number;
    /** How many hundreds it is multiplied by: one per `%`. */
    percent: number;
}

const DIGIT = /[0#?]/;

function numberSection(text: string): NumberSection | undefined {
    const tokens: NumberToken[] = [];
    let grouped = false;
    let scale = 0;
    let percent = 0;
    let seenPoint = false;
    for (let index = 0; index < text.length; index++) {
        const char = text[index] as string;
        if (char === '"') {
            const end = text.indexOf('"', index + 1);
            const stop = end === -1 ? text.length : end;
            tokens.push({ kind: 'literal', text: text.slice(index + 1, stop) });
            index = stop;
        } else if (char === '\\') {
            tokens.push({ kind: 'literal', text: text[index + 1] ?? '' });
            index++;
        } else if (char === '_') {
            tokens.push({ kind: 'literal', text: ' ' });
            index++;
        } else if (char === '*') {
            index++;
        } else if (char === '[') {
            const end = text.indexOf(']', index + 1);
            const stop = end === -1 ? text.length : end;
            const inside = text.slice(index + 1, stop);
            // A condition decides which section a value goes to by something
            // other than its sign, and that is not read here.
            if (/^[<>=]/.test(inside)) return undefined;
            if (inside.startsWith('$')) {
                tokens.push({ kind: 'literal', text: inside.slice(1).split('-')[0] ?? '' });
            }
            index = stop;
        } else if (text.slice(index, index + 7).toLowerCase() === 'general') {
            tokens.push({ kind: 'general' });
            index += 6;
        } else if (DIGIT.test(char)) {
            tokens.push({ kind: 'digit', char: char as '0' | '#' | '?' });
        } else if (char === '.' && !seenPoint) {
            seenPoint = true;
            tokens.push({ kind: 'point' });
        } else if (char === ',') {
            const previous = tokens.at(-1);
            if (previous?.kind === 'digit' && DIGIT.test(text[index + 1] ?? '')) {
                grouped = true;
            } else if (previous?.kind === 'digit' || (previous?.kind === 'point' && scale > 0)) {
                scale++;
            } else {
                tokens.push({ kind: 'literal', text: ',' });
            }
        } else if (char === '%') {
            percent++;
            tokens.push({ kind: 'literal', text: '%' });
        } else if ((char === 'E' || char === 'e') && (text[index + 1] === '+' || text[index + 1] === '-')) {
            tokens.push({ kind: 'exponent', letter: char, sign: text[index + 1] as '+' | '-' });
            index++;
        } else if (char === '/' && tokens.some((token) => token.kind === 'digit')) {
            // A fraction: `# ?/?`. Not taken apart here.
            return undefined;
        } else if (char === '@') {
            // The text of the cell: a number under it is shown as it is.
            tokens.push({ kind: 'general' });
        } else {
            tokens.push({ kind: 'literal', text: char });
        }
    }
    return { tokens, grouped, scale, percent };
}

/** Thousands separators put into a run of digits. */
function group(digits: string, separator: string): string {
    return digits.replace(/\B(?=(\d{3})+(?!\d))/g, separator);
}

/**
 * Placeholders filled with digits, from the right for the integer part: the
 * leftmost placeholder takes every digit that is left, and one that gets none
 * shows a `0`, a space or nothing, as it is a `0`, a `?` or a `#`.
 */
function fillInteger(placeholders: readonly string[], digits: string): string[] {
    const filled: string[] = [];
    let rest = digits;
    for (let index = placeholders.length - 1; index >= 0; index--) {
        const placeholder = placeholders[index] as string;
        if (index === 0) {
            filled[index] = rest || (placeholder === '0' ? '0' : placeholder === '?' ? ' ' : '');
        } else if (rest) {
            filled[index] = rest.slice(-1);
            rest = rest.slice(0, -1);
        } else {
            filled[index] = placeholder === '0' ? '0' : placeholder === '?' ? ' ' : '';
        }
    }
    return filled;
}

/** The decimals, from the left: a trailing zero is kept by a `0`, a space for a `?`, dropped for a `#`. */
function fillDecimals(placeholders: readonly string[], digits: string): string[] {
    const filled = placeholders.map((_, index) => digits[index] ?? '0');
    for (let index = placeholders.length - 1; index >= 0; index--) {
        if (filled[index] !== '0') break;
        const placeholder = placeholders[index];
        if (placeholder === '0') break;
        filled[index] = placeholder === '?' ? ' ' : '';
    }
    return filled;
}

function renderNumber(section: NumberSection, value: number, locale: Locale): string | undefined {
    let number = value * 100 ** section.percent / 1000 ** section.scale;
    const { tokens } = section;
    const pointAt = tokens.findIndex((token) => token.kind === 'point');
    const exponentAt = tokens.findIndex((token) => token.kind === 'exponent');
    const mantissaEnd = exponentAt === -1 ? tokens.length : exponentAt;
    const integerEnd = pointAt === -1 ? mantissaEnd : pointAt;
    const digitsIn = (from: number, to: number): string[] =>
        tokens.slice(from, to).flatMap((token) => (token.kind === 'digit' ? [token.char] : []));
    const integer = digitsIn(0, integerEnd);
    const decimals = pointAt === -1 ? [] : digitsIn(pointAt + 1, mantissaEnd);
    const exponentDigits = exponentAt === -1 ? [] : digitsIn(exponentAt + 1, tokens.length);

    if (integer.length === 0 && decimals.length === 0) {
        return tokens
            .map((token) => (token.kind === 'general' ? generalText(number, locale) : token.kind === 'literal' ? token.text : ''))
            .join('');
    }

    let exponent = 0;
    if (exponentAt !== -1 && number !== 0) {
        exponent = Math.floor(Math.log10(Math.abs(number)));
        // `##0.0E+0` keeps the exponent a multiple of three: engineering notation.
        if (integer.length > 1) exponent = Math.floor(exponent / integer.length) * integer.length;
        number /= 10 ** exponent;
    }
    if (Math.abs(number) >= 1e21) return undefined;
    const fixed = Math.abs(number).toFixed(decimals.length);
    const [whole = '', fraction = ''] = fixed.split('.');
    const integerDigits = whole === '0' ? '' : whole;

    // Where the integer placeholders sit next to each other, the digits go in
    // as one run, which is the only shape grouping makes sense in; with text
    // between them (`000-000`) each placeholder takes its own digit.
    const integerTokens = tokens.slice(0, integerEnd);
    const firstDigit = integerTokens.findIndex((token) => token.kind === 'digit');
    const lastDigit = integerTokens.length - 1 - [...integerTokens].reverse().findIndex((token) => token.kind === 'digit');
    const contiguous =
        firstDigit === -1 ||
        integerTokens.slice(firstDigit, lastDigit + 1).every((token) => token.kind === 'digit');
    const integerFilled = fillInteger(integer, integerDigits);
    const integerRun = contiguous
        ? (section.grouped ? group(integerFilled.join(''), locale.group) : integerFilled.join(''))
        : '';
    const decimalFilled = fillDecimals(decimals, fraction);
    const exponentToken = tokens[exponentAt];
    const exponentText =
        exponentToken?.kind !== 'exponent'
            ? ''
            : (exponent < 0 ? '-' : exponentToken.sign === '+' ? '+' : '') +
              String(Math.abs(exponent)).padStart(exponentDigits.filter((digit) => digit === '0').length, '0');

    let text = '';
    let integerIndex = 0;
    let decimalIndex = 0;
    let exponentWritten = false;
    tokens.forEach((token, index) => {
        if (token.kind === 'literal') text += token.text;
        else if (token.kind === 'general') text += generalText(number, locale);
        // A code with no integer placeholder (`.00`) still shows the integer
        // part of a number that has one, right before the point.
        else if (token.kind === 'point') text += (integer.length ? '' : integerDigits) + locale.decimal;
        else if (token.kind === 'exponent') text += token.letter.toUpperCase();
        else if (index < integerEnd) {
            if (contiguous) {
                if (integerIndex === 0) text += integerRun;
            } else {
                text += integerFilled[integerIndex] ?? '';
            }
            integerIndex++;
        } else if (index < mantissaEnd) {
            text += decimalFilled[decimalIndex++] ?? '';
        } else if (!exponentWritten) {
            text += exponentText;
            exponentWritten = true;
        }
    });
    return text;
}

/**
 * A number under a number code, or `undefined` for a code this does not take
 * apart — a fraction, a condition — which the caller writes plainly instead.
 */
export function formatNumber(value: number, code: string, locale: Locale): string | undefined {
    const found = sections(code);
    let section: string;
    let shown = value;
    let sign = '';
    if (found.length >= 3 && value === 0) section = found[2] as string;
    else if (found.length >= 2 && value < 0) {
        section = found[1] as string;
        shown = -value;
    } else {
        section = found[0] as string;
        if (value < 0) {
            sign = '-';
            shown = -value;
        }
    }
    const parsed = numberSection(section);
    if (parsed === undefined) return undefined;
    const text = renderNumber(parsed, shown, locale);
    return text === undefined ? undefined : sign + text;
}
