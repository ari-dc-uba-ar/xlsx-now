import assert from 'node:assert/strict';
import {
    builtinCode,
    codeOf,
    formatDateSerial,
    formatDateText,
    formatElapsed,
    formatNumber,
    generalText,
    isoOfSerial,
    localeFor,
} from '../../src/core/read/displayText.js';

const ES = localeFor('es');
const EN = localeFor('en');

describe('localeFor', () => {
    it('is Spanish for es and any es-…, and English for everything else', () => {
        for (const lang of ['es', 'ES', 'es-AR', 'es_ES']) assert.equal(localeFor(lang), ES, lang);
        for (const lang of ['en', 'en-US', 'fr', 'esperanto', undefined]) {
            assert.equal(localeFor(lang), EN, String(lang));
        }
    });
});

describe('builtinCode', () => {
    it('is the short date of the language for 14, which a file leaves to the machine', () => {
        assert.equal(builtinCode(14, ES), 'dd/mm/yyyy');
        assert.equal(builtinCode(14, EN), 'mm/dd/yyyy');
        assert.equal(builtinCode(22, ES), 'dd/mm/yyyy hh:mm');
        // The dates of the Japanese, Chinese and Korean locales, the same way.
        assert.equal(builtinCode(30, ES), 'dd/mm/yyyy');
    });

    it('is the code every reader knows for the rest, and General for an id it does not', () => {
        assert.equal(builtinCode(4, ES), '#,##0.00');
        assert.equal(builtinCode(46, ES), '[h]:mm:ss');
        assert.equal(builtinCode(163, ES), 'General');
        assert.equal(codeOf(undefined, ES), 'General');
        assert.equal(codeOf('0.0', ES), '0.0');
    });
});

describe('generalText', () => {
    it('is the number as it is, with the decimals of the language', () => {
        assert.equal(generalText(1234.5, ES), '1234,5');
        assert.equal(generalText(1234.5, EN), '1234.5');
        assert.equal(generalText(1e21, EN), '1E+21');
    });
});

describe('formatNumber', () => {
    it('shows the decimals and the thousands a code asks for, in the language given', () => {
        assert.equal(formatNumber(1234.5, '#,##0.00', ES), '1.234,50');
        assert.equal(formatNumber(1234.5, '#,##0.00', EN), '1,234.50');
        assert.equal(formatNumber(1234567, '#,##0', EN), '1,234,567');
        assert.equal(formatNumber(0.5, '0', EN), '1');
        assert.equal(formatNumber(3, '0.00', EN), '3.00');
    });

    it('pads with zeros where the code has them, and drops what a # does not need', () => {
        assert.equal(formatNumber(123, '00000', EN), '00123');
        assert.equal(formatNumber(0.5, '#.##', EN), '.5');
        assert.equal(formatNumber(1.5, '0.0#', EN), '1.5');
        assert.equal(formatNumber(1.25, '0.0#', EN), '1.25');
        assert.equal(formatNumber(5.5, '.00', EN), '5.50');
    });

    it('multiplies a percentage, and divides by a thousand for each comma at the end', () => {
        assert.equal(formatNumber(0.125, '0.0%', ES), '12,5%');
        assert.equal(formatNumber(1234567, '#,##0,', EN), '1,235');
        assert.equal(formatNumber(1234567, '0.0,,"M"', EN), '1.2M');
    });

    it('puts a minus sign in front with one section, and none where the negative one says', () => {
        assert.equal(formatNumber(-1234.5, '#,##0.00', EN), '-1,234.50');
        assert.equal(formatNumber(-1234.5, '#,##0.00;(#,##0.00)', EN), '(1,234.50)');
        assert.equal(formatNumber(0, '0.00;-0.00;"cero"', EN), 'cero');
        assert.equal(formatNumber(-5, '[Red]0;[Blue]-0', EN), '-5');
    });

    it('writes the literals of a code where the code puts them', () => {
        assert.equal(formatNumber(30, '0 "días"', ES), '30 días');
        assert.equal(formatNumber(1234.5, '"$"#,##0.00_)', EN), '$1,234.50 ');
        assert.equal(formatNumber(5, '[$€-C0A] 0.00', ES), '€ 5,00');
        assert.equal(formatNumber(1234567, '000-0000', EN), '123-4567');
    });

    it('writes scientific notation', () => {
        assert.equal(formatNumber(12345, '0.00E+00', EN), '1.23E+04');
        assert.equal(formatNumber(0.00012, '0.00E+00', EN), '1.20E-04');
        assert.equal(formatNumber(12345, '##0.0E+0', EN), '12.3E+3');
    });

    it('shows a number under the text format as it is', () => {
        assert.equal(formatNumber(1234.5, '@', ES), '1234,5');
        assert.equal(formatNumber(7, 'General', EN), '7');
    });

    it('says nothing for what it does not take apart, so the caller writes it plainly', () => {
        assert.equal(formatNumber(1.5, '# ?/?', EN), undefined);
        assert.equal(formatNumber(150, '[>100]0;0.00', EN), undefined);
    });
});

describe('formatDateSerial', () => {
    it('shows a day under its code, in the language given', () => {
        assert.equal(formatDateSerial(45306, 'dd/mm/yyyy', ES), '15/01/2024');
        assert.equal(formatDateSerial(45306, 'd-mmm-yy', ES), '15-ene-24');
        assert.equal(formatDateSerial(45306, 'd-mmm-yy', EN), '15-Jan-24');
        assert.equal(formatDateSerial(45306, 'dddd d "de" mmmm', ES), 'lunes 15 de enero');
        assert.equal(formatDateSerial(45306, 'ddd mmmmm', EN), 'Mon J');
    });

    it('tells the minutes from the month by what is next to them', () => {
        assert.equal(formatDateSerial(45306.43125, 'yyyy-mm-dd hh:mm', ES), '2024-01-15 10:21');
        assert.equal(formatDateSerial(45306.43125, 'mm:ss', ES), '21:00');
        assert.equal(formatDateSerial(45306.43125, 'm/d/yy h:mm', EN), '1/15/24 10:21');
    });

    it('shows a twelve-hour clock where the code has AM/PM', () => {
        assert.equal(formatDateSerial(0.75, 'h:mm AM/PM', EN), '6:00 PM');
        assert.equal(formatDateSerial(0.75, 'h:mm AM/PM', ES), '6:00 p. m.');
        assert.equal(formatDateSerial(0.25, 'hh:mm a/p', EN), '06:00 a');
        assert.equal(formatDateSerial(0, 'h AM/PM', EN), '12 AM');
    });

    it('rounds to the second it shows, the way Excel does', () => {
        // 10:20:59.6
        const serial = (10 * 3600 + 20 * 60 + 59.6) / 86400;
        assert.equal(formatDateSerial(serial, 'hh:mm', EN), '10:21');
        assert.equal(formatDateSerial(serial, 'hh:mm:ss.0', ES), '10:20:59,6');
    });

    it('shows the day Excel has and the calendar does not, and the day a negative serial is', () => {
        assert.equal(formatDateSerial(60, 'dd/mm/yyyy', ES), '29/02/1900');
        assert.equal(formatDateSerial(59, 'dd/mm/yyyy', ES), '28/02/1900');
        assert.equal(formatDateSerial(61, 'dd/mm/yyyy', ES), '01/03/1900');
        assert.equal(formatDateSerial(-1, 'dd/mm/yyyy', ES), '30/12/1899');
        assert.equal(formatDateSerial(-18091, 'dd/mm/yyyy', ES), '20/06/1850');
        assert.equal(formatDateSerial(-0.5, 'dd/mm/yyyy hh:mm', ES), '30/12/1899 12:00');
    });
});

describe('formatElapsed', () => {
    it('counts the bracketed unit in full and the rest within it', () => {
        assert.equal(formatElapsed(1.25, '[h]:mm', ES), '30:00');
        assert.equal(formatElapsed(1.25, '[h]:mm:ss', ES), '30:00:00');
        assert.equal(formatElapsed(1.25, '[mm]:ss', ES), '1800:00');
        assert.equal(formatElapsed(0.5 / 24, '[h]:mm', ES), '0:30');
        assert.equal(formatElapsed(-1.25, '[h]:mm', ES), '-30:00');
    });
});

describe('formatDateText', () => {
    it('shows a date written out in text under a code, a day no serial reaches included', () => {
        assert.equal(formatDateText('1850-06-20', 'dd/mm/yyyy', ES), '20/06/1850');
        assert.equal(formatDateText('1850-06-20', 'dddd', ES), 'jueves');
        assert.equal(formatDateText('2024-01-15T10:30:00', 'dd/mm/yyyy hh:mm', ES), '15/01/2024 10:30');
    });
});

describe('isoOfSerial', () => {
    it('is the ISO text of a day, a day and its hour, or an hour', () => {
        assert.equal(isoOfSerial(45306), '2024-01-15');
        assert.equal(isoOfSerial(45306.5), '2024-01-15T12:00:00');
        assert.equal(isoOfSerial(0.4375), '10:30:00');
        assert.equal(isoOfSerial(60), '1900-02-29');
        assert.equal(isoOfSerial(-18091), '1850-06-20');
        assert.equal(isoOfSerial(-0.5), '1899-12-30T12:00:00');
    });
});
