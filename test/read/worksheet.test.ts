import assert from 'node:assert/strict';
import { NO_FORMATS, readNumberFormats } from '../../src/core/read/numberFormats.js';
import type { ReadRow, ReadValue } from '../../src/core/read/types.js';
import {
    cellText,
    cellValue,
    parseCellReference,
    readCell,
    readRows,
    styledCell,
    type CellContext,
    type RawCell,
} from '../../src/core/read/worksheet.js';
import { stylesOf } from '../helpers/package.js';
import { asAsyncIterable } from '../helpers/streams.js';

const PLAIN: CellContext = {
    sharedStrings: [],
    formats: NO_FORMATS,
    date1904: false,
    // The tests below are about what a cell holds, not about what a date is
    // built as, so they read the one that is a `Date` on the caller's clock.
    dates: 'localDate',
};

function context(overrides: Partial<CellContext>): CellContext {
    return { ...PLAIN, ...overrides };
}

/** The same, reading its dates as Temporal values. */
const TEMPORAL: CellContext = { ...PLAIN, dates: 'temporal' };

/** A cell as the file spells it, with everything it does not say left out. */
function raw(cell: Partial<RawCell>): RawCell {
    return { type: undefined, style: undefined, value: undefined, formula: undefined, ...cell };
}

/** The rows a `<sheetData>` holds, read whole. */
async function rowsOf<C>(
    body: string,
    convert: (cell: RawCell) => C,
    chunks = 1,
): Promise<ReadRow<C>[]> {
    const xml = `<worksheet><sheetData>${body}</sheetData></worksheet>`;
    const pieces: string[] = [];
    const size = Math.ceil(xml.length / chunks);
    for (let at = 0; at < xml.length; at += size) pieces.push(xml.slice(at, at + size));
    const rows: ReadRow<C>[] = [];
    for await (const row of readRows(asAsyncIterable(pieces), convert, 'sheet1.xml')) rows.push(row);
    return rows;
}

const values = (body: string, ctx = PLAIN, chunks?: number): Promise<ReadRow<ReadValue>[]> =>
    rowsOf(body, (cell) => cellValue(cell, ctx), chunks);

describe('parseCellReference', () => {
    it('is the coordinates a cell is named by', () => {
        assert.deepEqual(parseCellReference('A1'), { column: 0, row: 1 });
        assert.deepEqual(parseCellReference('B12'), { column: 1, row: 12 });
        assert.deepEqual(parseCellReference('AA100'), { column: 26, row: 100 });
    });

    it('is nothing for what is not a reference', () => {
        for (const text of ['', '1A', 'A', '12', 'A1:B2', '$A$1']) {
            assert.equal(parseCellReference(text), undefined, `${text} was read as a reference`);
        }
    });
});

describe('cellValue', () => {
    it('reads a number, which is what a cell that says nothing holds', () => {
        assert.equal(cellValue(raw({ value: '42.5' }), PLAIN), 42.5);
        assert.equal(cellValue(raw({ type: 'n', value: '-3' }), PLAIN), -3);
    });

    it('reads a boolean out of the 1 or 0 it is written as', () => {
        assert.equal(cellValue(raw({ type: 'b', value: '1' }), PLAIN), true);
        assert.equal(cellValue(raw({ type: 'b', value: '0' }), PLAIN), false);
    });

    it('reads a shared string through the table it points into', () => {
        const ctx = context({ sharedStrings: ['uno', 'dos'] });
        assert.equal(cellValue(raw({ type: 's', value: '1' }), ctx), 'dos');
    });

    it('reads the string a cell carries itself, and the one a formula left', () => {
        assert.equal(cellValue(raw({ type: 'inlineStr', value: 'hola' }), PLAIN), 'hola');
        assert.equal(cellValue(raw({ type: 'str', value: 'hola' }), PLAIN), 'hola');
    });

    it('reads an error as what the cell shows, not as a value that went missing', () => {
        assert.equal(cellValue(raw({ type: 'e', value: '#DIV/0!' }), PLAIN), '#DIV/0!');
    });

    it('reads a date written out in full, which the spec allows and Excel does not write', () => {
        // The text is a wall clock like every other date in a sheet, so what
        // comes back reads half past twelve by whichever clock was asked for —
        // and a `Z` on the text does not make it an instant the sheet had.
        const local = cellValue(raw({ type: 'd', value: '2024-01-15T12:30:00Z' }), PLAIN);
        assert.ok(local instanceof Date);
        assert.deepEqual([local.getDate(), local.getHours(), local.getMinutes()], [15, 12, 30]);

        const utc = cellValue(
            raw({ type: 'd', value: '2024-01-15T12:30:00' }),
            context({ dates: 'utcDate' }),
        );
        assert.ok(utc instanceof Date);
        assert.equal(utc.toISOString(), '2024-01-15T12:30:00.000Z');
    });

    it('reads a date written out as a day alone', () => {
        assert.equal(
            cellValue(raw({ type: 'd', value: '2024-01-15' }), context({ dates: 'isoString' })),
            '2024-01-15',
        );
        assert.equal(
            String(cellValue(raw({ type: 'd', value: '2024-01-15' }), TEMPORAL)),
            '2024-01-15',
        );
        assert.equal(
            String(cellValue(raw({ type: 'd', value: '2024-01-15T12:30:00' }), TEMPORAL)),
            '2024-01-15T12:30:00',
        );
    });

    it('reads a day the serial cannot number, which is what that cell is for', () => {
        // There is no negative serial, so a date before 31/12/1899 has nothing
        // to be numbered as and the text is the only way a file can hold it.
        assert.equal(
            String(cellValue(raw({ type: 'd', value: '1850-06-20' }), TEMPORAL)),
            '1850-06-20',
        );
        assert.equal(
            cellValue(raw({ type: 'd', value: '1850-06-20' }), context({ dates: 'isoString' })),
            '1850-06-20',
        );
        const local = cellValue(raw({ type: 'd', value: '1850-06-20' }), PLAIN);
        assert.ok(local instanceof Date);
        assert.deepEqual([local.getFullYear(), local.getMonth(), local.getDate()], [1850, 5, 20]);
    });

    it('says what a date cell holds when it does not hold a date', () => {
        assert.throws(
            () => cellValue(raw({ type: 'd', value: 'ayer' }), PLAIN),
            /A date cell holds "ayer"/,
        );
    });

    it('is a Date when the format under the number says it is one', () => {
        const ctx = context({ formats: readNumberFormats(stylesOf([14, 0])) });
        const date = cellValue(raw({ value: '45306', style: 0 }), ctx);
        assert.ok(date instanceof Date);
        assert.equal(date.getFullYear(), 2024);
        assert.equal(date.getMonth(), 0);
        assert.equal(date.getDate(), 15);
        // The same number under a format that is not a date stays a number.
        assert.equal(cellValue(raw({ value: '45306', style: 1 }), ctx), 45306);
    });

    it('counts from 1904 when the workbook says it does', () => {
        const ctx = context({ formats: readNumberFormats(stylesOf([14])), date1904: true });
        const date = cellValue(raw({ value: '43844', style: 0 }), ctx);
        assert.ok(date instanceof Date);
        assert.equal(date.getFullYear(), 2024);
        assert.equal(date.getMonth(), 0);
        assert.equal(date.getDate(), 15);
    });

    it('is null for a cell that is there and holds nothing', () => {
        assert.equal(cellValue(raw({ style: 3 }), PLAIN), null);
    });

    it('refuses a cell whose value is not what it says it is', () => {
        assert.throws(() => cellValue(raw({ value: 'no es un número' }), PLAIN), /not a number/);
        assert.throws(() => cellValue(raw({ type: 'd', value: 'ayer' }), PLAIN), /not a date/);
        assert.throws(() => cellValue(raw({ type: 'x', value: '1' }), PLAIN), /not a type/);
    });

    it('refuses a shared string the table does not have', () => {
        const ctx = context({ sharedStrings: ['uno'] });
        assert.throws(() => cellValue(raw({ type: 's', value: '5' }), ctx), /shared string 5/);
    });
});

describe('styledCell', () => {
    it('carries the value, and says no more than it has to', () => {
        assert.deepEqual(styledCell(raw({ value: '42' }), PLAIN), { v: 42 });
        assert.deepEqual(styledCell(raw({ type: 'b', value: '1' }), PLAIN), { v: true });
    });

    it('carries the number format as a style the writer takes back', () => {
        const ctx = context({ formats: readNumberFormats(stylesOf(['#,##0.00', 14])) });
        assert.deepEqual(styledCell(raw({ value: '1234.5', style: 0 }), ctx), {
            v: 1234.5,
            s: { numFmt: '#,##0.00' },
        });
        const date = styledCell(raw({ value: '45306', style: 1 }), ctx);
        assert.deepEqual(date.s, { numFmt: 14 });
    });

    it('carries the formula, and the result the file cached for it', () => {
        assert.deepEqual(styledCell(raw({ value: '3', formula: '1+2' }), PLAIN), {
            v: 3,
            f: '1+2',
        });
    });

    it('says the type only where the writer would not work it out', () => {
        // Text that is the cached result of a formula has to stay text, and
        // an error is text that is not a value; a number, a boolean, a date
        // and a plain string all say what they are by being what they are.
        assert.equal(styledCell(raw({ type: 'str', value: 'x', formula: 'A1' }), PLAIN).t, 'str');
        assert.equal(styledCell(raw({ type: 'e', value: '#N/A' }), PLAIN).t, 'e');
        assert.equal(styledCell(raw({ type: 'inlineStr', value: 'x' }), PLAIN).t, undefined);
        assert.equal(styledCell(raw({ value: '1' }), PLAIN).t, undefined);
    });
});

describe('readRows', () => {
    it('reads the cells of a row into the columns they name', async () => {
        const rows = await values('<row r="1"><c r="A1"><v>1</v></c><c r="C1"><v>3</v></c></row>');
        assert.equal(rows.length, 1);
        assert.equal(rows[0]?.index, 1);
        assert.deepEqual([...(rows[0]?.cells ?? [])], [1, undefined, 3]);
    });

    it('carries on from the last column for a cell that does not name one', async () => {
        const rows = await values('<row><c><v>1</v></c><c><v>2</v></c><c r="E1"><v>5</v></c><c><v>6</v></c></row>');
        assert.equal(rows[0]?.cells.length, 6);
        assert.deepEqual(rows[0]?.cells[4], 5);
        assert.deepEqual(rows[0]?.cells[5], 6);
    });

    it('numbers a row that does not number itself as the next one', async () => {
        const rows = await values('<row><c><v>1</v></c></row><row><c><v>2</v></c></row>');
        assert.deepEqual(rows.map((row) => row.index), [1, 2]);
    });

    it('keeps the numbers of rows the sheet skipped over', async () => {
        const rows = await values('<row r="1"><c><v>1</v></c></row><row r="9"><c><v>9</v></c></row>');
        assert.deepEqual(rows.map((row) => row.index), [1, 9]);
    });

    it('joins the runs of a string the cell carries itself', async () => {
        const rows = await values(
            '<row><c t="inlineStr"><is><r><t>en </t></r><r><t>partes</t></r></is></c></row>',
        );
        assert.equal(rows[0]?.cells[0], 'en partes');
    });

    it('reads the same rows however the file is cut into chunks', async () => {
        const body =
            '<row r="1"><c r="A1" t="inlineStr"><is><t>una frase larga</t></is></c>' +
            '<c r="B1"><v>2</v></c></row><row r="2"><c r="A2"><v>3</v></c></row>';
        const whole = await values(body);
        for (const chunks of [2, 5, 20, 100]) {
            assert.deepEqual(await values(body, PLAIN, chunks), whole, `cut into ${chunks}`);
        }
    });

    it('reads a file whose elements carry a namespace prefix', async () => {
        const xml =
            '<x:worksheet xmlns:x="urn:x"><x:sheetData><x:row r="1">' +
            '<x:c r="A1"><x:v>7</x:v></x:c></x:row></x:sheetData></x:worksheet>';
        const rows: ReadRow<ReadValue>[] = [];
        for await (const row of readRows(
            asAsyncIterable([xml]),
            (cell) => cellValue(cell, PLAIN),
            'sheet1.xml',
        )) {
            rows.push(row);
        }
        assert.equal(rows[0]?.cells[0], 7);
    });

    it('leaves alone what is not the data', async () => {
        // `<dimension>` and `<mergeCells>` sit outside `<sheetData>` and have
        // an `r` of their own, which a reader that went by tag names alone
        // would take for a row.
        const xml =
            '<worksheet><dimension ref="A1:B2"/><sheetData><row r="1"><c><v>1</v></c></row>' +
            '</sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells></worksheet>';
        const rows: ReadRow<ReadValue>[] = [];
        for await (const row of readRows(
            asAsyncIterable([xml]),
            (cell) => cellValue(cell, PLAIN),
            'sheet1.xml',
        )) {
            rows.push(row);
        }
        assert.equal(rows.length, 1);
    });

    it('reads a row that holds nothing as a row with no cells', async () => {
        const rows = await values('<row r="3"/>');
        assert.deepEqual(rows, [{ index: 3, cells: [] }]);
    });

    it('refuses a cell whose reference is not one', async () => {
        await assert.rejects(values('<row r="1"><c r="hola"><v>1</v></c></row>'), /not a cell reference/);
    });

    it('refuses a part that does not parse', async () => {
        await assert.rejects(values('<row r="1"><c><v>1</v></row>'));
    });
});

describe('cellValue: a length of time', () => {
    const ctx = (dates: CellContext['dates']): CellContext =>
        context({ formats: readNumberFormats(stylesOf(['[h]:mm', 46])), dates });

    it('is not a date, under a declared code or the built-in one', () => {
        assert.equal(cellValue(raw({ value: '1.25', style: 0 }), ctx('isoString')), 'PT30H');
        assert.equal(cellValue(raw({ value: '1.25', style: 1 }), ctx('isoString')), 'PT30H');
        assert.equal(cellValue(raw({ value: '1.25', style: 0 }), ctx('localDate')), 1.25);
        assert.equal(cellValue(raw({ value: '1.25', style: 0 }), ctx('serial')), 1.25);
        assert.equal(String(cellValue(raw({ value: '1.25', style: 0 }), ctx('temporal'))), 'PT30H');
    });

    it('is the same length in a workbook that counts from 1904', () => {
        const in1904 = { ...ctx('isoString'), date1904: true };
        assert.equal(cellValue(raw({ value: '1.25', style: 0 }), in1904), 'PT30H');
    });
});

describe('cellValue: the serial mode', () => {
    it('is the number under a date format, counted from 1900 whatever the workbook says', () => {
        const ctx = context({ formats: readNumberFormats(stylesOf([14])), dates: 'serial' });
        assert.equal(cellValue(raw({ value: '45306', style: 0 }), ctx), 45306);
        assert.equal(cellValue(raw({ value: '43844', style: 0 }), { ...ctx, date1904: true }), 45306);
    });
});

describe('cellText', () => {
    const es = (overrides: Partial<CellContext> = {}): CellContext => context({ lang: 'es', ...overrides });

    it('is the text of a string, an error or the string a formula left', () => {
        assert.equal(cellText(raw({ type: 's', value: '0' }), es({ sharedStrings: ['hola'] })), 'hola');
        assert.equal(cellText(raw({ type: 'inlineStr', inline: 'hola' }), es()), 'hola');
        assert.equal(cellText(raw({ type: 'e', value: '#DIV/0!' }), es()), '#DIV/0!');
        assert.equal(cellText(raw({ type: 'str', value: 'x', formula: 'A1' }), es()), 'x');
    });

    it('is the word of the language for a boolean', () => {
        assert.equal(cellText(raw({ type: 'b', value: '1' }), es()), 'VERDADERO');
        assert.equal(cellText(raw({ type: 'b', value: '0' }), PLAIN), 'FALSE');
    });

    it('is a number under its format, or as it is under General', () => {
        const formats = readNumberFormats(stylesOf(['#,##0.00', 4]));
        assert.equal(cellText(raw({ value: '1234.5', style: 0 }), es({ formats })), '1.234,50');
        assert.equal(cellText(raw({ value: '1234.5', style: 1 }), es({ formats })), '1.234,50');
        assert.equal(cellText(raw({ value: '1234.5' }), es()), '1234,5');
        assert.equal(cellText(raw({ value: '1234.5', style: 0 }), context({ formats })), '1,234.50');
    });

    it('is a number as it is under a format it does not take apart', () => {
        const formats = readNumberFormats(stylesOf(['# ?/?']));
        assert.equal(cellText(raw({ value: '1.5', style: 0 }), es({ formats })), '1,5');
    });

    it('is a date under its code in the language given, and in ISO with no language', () => {
        const formats = readNumberFormats(stylesOf(['yyyy-mm-dd hh:mm', 14, 'd-mmm-yy']));
        const when = raw({ value: '44560.43055555556', style: 0 });
        assert.equal(cellText(when, es({ formats })), '2021-12-30 10:20');
        assert.equal(cellText(when, context({ formats })), '2021-12-30T10:20:00');
        assert.equal(cellText(raw({ value: '45306', style: 1 }), es({ formats })), '15/01/2024');
        assert.equal(cellText(raw({ value: '45306', style: 1 }), context({ formats, lang: 'en' })), '01/15/2024');
        assert.equal(cellText(raw({ value: '45306', style: 2 }), es({ formats })), '15-ene-24');
        assert.equal(cellText(raw({ value: '45306', style: 2 }), context({ formats })), '2024-01-15');
    });

    it('counts a date from 1904 when the workbook does', () => {
        const formats = readNumberFormats(stylesOf([14]));
        assert.equal(cellText(raw({ value: '43844', style: 0 }), es({ formats, date1904: true })), '15/01/2024');
    });

    it('is a length of time under its code, with a language or without one', () => {
        const formats = readNumberFormats(stylesOf(['[h]:mm']));
        assert.equal(cellText(raw({ value: '1.25', style: 0 }), es({ formats })), '30:00');
        assert.equal(cellText(raw({ value: '1.25', style: 0 }), context({ formats })), '30:00');
    });

    it('is a date written out in text under its code, or the text itself with no language', () => {
        const formats = readNumberFormats(stylesOf(['dd/mm/yyyy']));
        assert.equal(cellText(raw({ type: 'd', value: '1850-06-20', style: 0 }), es({ formats })), '20/06/1850');
        assert.equal(cellText(raw({ type: 'd', value: '1850-06-20' }), es()), '20/06/1850');
        assert.equal(cellText(raw({ type: 'd', value: '1850-06-20T10:30:00Z' }), PLAIN), '1850-06-20T10:30:00');
    });

    it('is empty for a cell that holds nothing', () => {
        assert.equal(cellText(raw({ style: 1 }), es()), '');
        assert.equal(cellText(raw({ type: 'b' }), es()), '');
    });
});

describe('readCell', () => {
    const ctx = context({
        sharedStrings: ['uno', 'dos'],
        formats: readNumberFormats(stylesOf(['#,##0.00', 14])),
        lang: 'es',
    });

    it('is the value alone when nothing else was asked for', () => {
        assert.deepEqual(readCell(raw({ value: '42' }), ctx, {}), { v: 42 });
    });

    it('carries the fields asked for, and leaves out the ones the cell has nothing for', () => {
        const all = { s: true, f: true, t: true, w: true } as const;
        assert.deepEqual(readCell(raw({ value: '1234.5', style: 0 }), ctx, all), {
            v: 1234.5,
            s: { numFmt: '#,##0.00' },
            w: '1.234,50',
        });
        assert.deepEqual(readCell(raw({ type: 'str', value: 'x', formula: 'A1' }), ctx, all), {
            v: 'x',
            f: 'A1',
            t: 'str',
            w: 'x',
        });
    });

    it('gives the file its own names, with nothing made of them', () => {
        const fields = { v: false, _t: true, _s: true, _v: true, _f: true, _is: true } as const;
        assert.deepEqual(readCell(raw({ type: 's', value: '1', style: 1 }), ctx, fields), {
            _t: 's',
            _s: 1,
            _v: '1',
        });
        assert.deepEqual(readCell(raw({ type: 'inlineStr', inline: 'hola' }), ctx, fields), {
            _t: 'inlineStr',
            _is: 'hola',
        });
        assert.deepEqual(readCell(raw({ value: '3', formula: '1+2' }), ctx, fields), { _v: '3', _f: '1+2' });
    });

    it('can say a cell both ways at once', () => {
        assert.deepEqual(readCell(raw({ type: 's', value: '1' }), ctx, { _v: true }), { v: 'dos', _v: '1' });
    });

    it('says what kind of number a cell holds, by its format', () => {
        const kinds = context({ formats: readNumberFormats(stylesOf(['0.00', 14, '[h]:mm'])), dates: 'serial' });
        assert.deepEqual(readCell(raw({ value: '92', style: 0 }), kinds, { kind: true }), { v: 92, kind: 'number' });
        assert.deepEqual(readCell(raw({ value: '92', style: 1 }), kinds, { kind: true }), { v: 92, kind: 'date' });
        assert.deepEqual(readCell(raw({ value: '92', style: 2 }), kinds, { kind: true }), { v: 92, kind: 'elapsed' });
        assert.deepEqual(readCell(raw({ value: '92' }), kinds, { kind: true }), { v: 92, kind: 'number' });
        assert.deepEqual(readCell(raw({ type: 'd', value: '1850-06-20' }), kinds, { kind: true }), { v: -18091, kind: 'date' });
        assert.deepEqual(readCell(raw({ type: 'inlineStr', inline: 'x' }), kinds, { kind: true }), { v: 'x' });
        assert.deepEqual(readCell(raw({ style: 1 }), kinds, { kind: true }), { v: null });
    });

    it('always has a w when it was asked for, even for a cell that holds nothing', () => {
        assert.deepEqual(readCell(raw({ style: 0 }), ctx, { w: true }), { v: null, w: '' });
    });
});

describe('readRows: the string a cell carries itself', () => {
    it('keeps the text of an <is> apart from the <v> the cell does not have', async () => {
        const rows = await rowsOf(
            '<row><c t="inlineStr"><is><r><t>en </t></r><r><t>partes</t></r></is></c></row>',
            (cell) => cell,
        );
        const cell = rows[0]?.cells[0];
        assert.equal(cell?.inline, 'en partes');
        assert.equal(cell?.value, undefined);
    });
});
