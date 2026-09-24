import { describe, expect, it } from 'vitest';
import { formatValue, parseInput, partsToSerial, serialToParts } from '../src/sheets/format/numfmt';

const f = (v: unknown, fmt: string) => formatValue(v, fmt).text;

describe('number formats', () => {
  it('formats numbers', () => {
    expect(f(1234.567, '#,##0.00')).toBe('1,234.57');
    expect(f(3.1, '0.00')).toBe('3.10');
    expect(f(0.5, '#.##')).toBe('.5');
    expect(f(5, '#.##')).toBe('5.');
    expect(f(1234567, '#,##0,')).toBe('1,235');
    expect(f(1234567, '0.0,,"M"')).toBe('1.2M');
    expect(f(5551234567, '(###) ###-####')).toBe('(555) 123-4567');
    expect(f(7, '000')).toBe('007');
  });
  it('formats negatives and sections', () => {
    expect(f(-1234.5, '#,##0.00;(#,##0.00)')).toBe('(1,234.50)');
    expect(formatValue(-5, '0;[Red]-0')).toEqual({ text: '-5', color: '#ff0000' });
    expect(f(-1234.5, '$#,##0.00')).toBe('-$1,234.50');
    expect(f(0, '0;-0;"zero"')).toBe('zero');
    expect(f('abc', '0;0;0;"x"@')).toBe('xabc');
    expect(f(150, '[>100]"big";"small"')).toBe('big');
    expect(f(50, '[>100]"big";"small"')).toBe('small');
  });
  it('formats percent, currency and scientific', () => {
    expect(f(0.1234, '0.0%')).toBe('12.3%');
    expect(f(0.75, '0%')).toBe('75%');
    expect(f(1234.5, '$#,##0.00')).toBe('$1,234.50');
    expect(f(1234.5, '[$€-407] #,##0.00')).toBe('€ 1,234.50');
    expect(f(12345.678, '0.00E+00')).toBe('1.23E+04');
    expect(f(0.000123, '0.00E+00')).toBe('1.23E-04');
  });
  it('formats fractions', () => {
    expect(f(1.75, '# ?/?').trim()).toBe('1 3/4');
    expect(f(0.333, '?/?').trim()).toBe('1/3');
    expect(f(2.5, '# ?/8').trim()).toBe('2 4/8');
  });
  it('formats General', () => {
    expect(f(0.1 + 0.2, 'General')).toBe('0.3');
    expect(f(1 / 3, 'General')).toBe('0.3333333333');
    expect(f(123456789012, 'General')).toBe('1.23457E+11');
    expect(f(42, 'General')).toBe('42');
    expect(f(-0.5, 'General')).toBe('-0.5');
    expect(f(true, 'General')).toBe('TRUE');
  });
  it('formats dates and times', () => {
    expect(partsToSerial(2023, 3, 15)).toBe(45000);
    expect(f(45000, 'yyyy-mm-dd')).toBe('2023-03-15');
    expect(f(45000, 'dddd, mmmm d, yyyy')).toBe('Wednesday, March 15, 2023');
    expect(f(45000, 'd-mmm-yy')).toBe('15-Mar-23');
    expect(f(1, 'm/d/yyyy')).toBe('1/1/1900');
    expect(f(60, 'yyyy-mm-dd')).toBe('1900-02-29');
    expect(f(61, 'yyyy-mm-dd')).toBe('1900-03-01');
    expect(f(45000.5, 'h:mm AM/PM')).toBe('12:00 PM');
    expect(f(45000.75, 'hh:mm')).toBe('18:00');
    expect(f(1.5, '[h]:mm')).toBe('36:00');
    expect(f(1 / 86400, 'h:mm:ss')).toBe('0:00:01');
    expect(f(0.5 + 1.5 / 86400, 'h:mm:ss.0')).toBe('12:00:01.5');
    expect(serialToParts(45000.25).H).toBe(6);
  });
});

describe('input parsing', () => {
  it('parses numbers', () => {
    expect(parseInput('1,234.5')).toEqual({ value: 1234.5, format: '#,##0.00' });
    expect(parseInput('12%')).toEqual({ value: 0.12, format: '0%' });
    expect(parseInput('$5')).toEqual({ value: 5, format: '$#,##0' });
    expect(parseInput('(50)').value).toBe(-50);
    expect(parseInput('-3.5').value).toBe(-3.5);
    expect(parseInput('1e3').value).toBe(1000);
  });
  it('parses dates, times and text', () => {
    expect(parseInput('2023-03-15')).toEqual({ value: 45000, format: 'yyyy-mm-dd' });
    expect(parseInput('14:30').value).toBeCloseTo(14.5 / 24);
    expect(parseInput('2:30 PM').value).toBeCloseTo(14.5 / 24);
    expect(parseInput('15-Mar-2023').value).toBe(45000);
    expect(parseInput('March 15, 2023').value).toBe(45000);
    expect(parseInput('TRUE').value).toBe(true);
    expect(parseInput('hello').value).toBe('hello');
    expect(parseInput("'123").value).toBe('123');
    expect(parseInput('').value).toBe(null);
  });
});
