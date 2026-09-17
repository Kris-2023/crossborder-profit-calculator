'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const engine = require('../engine.js');
const exporter = require('../export.js');

function makeSnapshot(overrides = {}) {
  const model = engine.createDefaultModel();
  Object.assign(model, overrides);
  const result = engine.calculate(model);
  assert.equal(result.complete, true);
  return exporter.buildSnapshot(model, result, {
    platformLabel: model.platform,
    createdAt: '2026-09-17T00:00:00.000Z'
  });
}

function readStoredZip(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const decoder = new TextDecoder();
  const entries = new Map();
  let offset = 0;
  while (offset + 30 <= bytes.length && view.getUint32(offset, true) === 0x04034b50) {
    const method = view.getUint16(offset + 8, true);
    const crc = view.getUint32(offset + 14, true);
    const size = view.getUint32(offset + 18, true);
    const nameLength = view.getUint16(offset + 26, true);
    const extraLength = view.getUint16(offset + 28, true);
    assert.equal(method, 0, '测试解析器只接受无压缩 ZIP 条目');
    const nameStart = offset + 30;
    const dataStart = nameStart + nameLength + extraLength;
    const data = bytes.slice(dataStart, dataStart + size);
    const name = decoder.decode(bytes.slice(nameStart, nameStart + nameLength));
    assert.equal(exporter.crc32(data), crc, `${name} CRC32 应匹配`);
    entries.set(name, data);
    offset = dataStart + size;
  }
  return entries;
}

test('导出快照与页面结果保持一致', () => {
  const snapshot = makeSnapshot();
  assert.equal(snapshot.netProfit, 58.26);
  assert.equal(snapshot.grossProfit, 88.16);
  assert.equal(snapshot.rows.length, 18);
  assert.equal(snapshot.platform, 'Amazon');
});

test('生成真正的 XLSX ZIP，并包含两个工作表的必需部件', () => {
  const bytes = exporter.buildXlsxBytes(makeSnapshot());
  assert.deepEqual(Array.from(bytes.slice(0, 4)), [0x50, 0x4b, 0x03, 0x04]);
  const entries = readStoredZip(bytes);
  for (const name of [
    '[Content_Types].xml',
    '_rels/.rels',
    'xl/workbook.xml',
    'xl/_rels/workbook.xml.rels',
    'xl/styles.xml',
    'xl/worksheets/sheet1.xml',
    'xl/worksheets/sheet2.xml'
  ]) assert.ok(entries.has(name), `缺少 ${name}`);
  const workbook = new TextDecoder().decode(entries.get('xl/workbook.xml'));
  assert.match(workbook, /name="利润模型"/);
  assert.match(workbook, /name="计算说明"/);
});

test('用户文本按 inlineStr 写入，避免被 Excel 当作公式执行', () => {
  const snapshot = makeSnapshot({ productModel: '=1+1 & <SKU>' });
  const entries = readStoredZip(exporter.buildXlsxBytes(snapshot));
  const sheet = new TextDecoder().decode(entries.get('xl/worksheets/sheet1.xml'));
  assert.doesNotMatch(sheet, /<f(?:\s|>)/);
  assert.match(sheet, /t="inlineStr"/);
  assert.match(sheet, /=1\+1 &amp; &lt;SKU&gt;/);
});

test('Excel Blob 使用官方 xlsx MIME', () => {
  const blob = exporter.buildXlsxBlob(makeSnapshot());
  assert.equal(blob.type, exporter.XLSX_MIME);
  assert.ok(blob.size > 1000);
});

test('导出文件名会过滤 Windows 非法字符', () => {
  const filename = exporter.sanitizeFilename('Amazon:SKU/01*利润?', 'xlsx');
  assert.equal(filename, 'Amazon-SKU-01-利润-.xlsx');
});
