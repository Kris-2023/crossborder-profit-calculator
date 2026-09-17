'use strict';

const assert = require('node:assert/strict');
const engine = require('../engine.js');

let passed = 0;
let failed = 0;

function test(name, fn) {
  try {
    fn();
    passed += 1;
    console.log(`\u2713 ${name}`);
  } catch (error) {
    failed += 1;
    console.error(`\u2717 ${name}`);
    console.error(error && error.stack ? error.stack : error);
  }
}

function approx(actual, expected, epsilon = 1e-9) {
  assert.equal(typeof actual, 'number');
  assert.ok(
    Math.abs(actual - expected) <= epsilon,
    `expected ${actual} to be within ${epsilon} of ${expected}`
  );
}

function roundMoney(value) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function baseModel(overrides = {}) {
  return {
    platform: 'Amazon',
    productModel: 'TEST-SKU',
    price: { amount: 100, currency: 'USD' },
    settlementCurrency: 'USD',
    usdCnyRate: 6.78,
    settlementCnyRate: 6.78,
    directCosts: [],
    productCost: { amount: 0, currency: 'USD' },
    allocations: [],
    ...overrides
  };
}

test('299 美元示例得到毛利 88.16、净利润 58.26', () => {
  const result = engine.calculate(engine.createDefaultModel());

  assert.equal(result.complete, true);
  assert.equal(roundMoney(result.grossProfit), 88.16);
  assert.equal(roundMoney(result.netProfit), 58.26);
  assert.equal(roundMoney(result.directTotal), 210.84);
  assert.equal(roundMoney(result.allocationTotal), 29.90);
});

test('人民币产品成本按 6.78 换算成美元且不重复累计', () => {
  const model = engine.createDefaultModel();
  model.productCost = { amount: 508.50, currency: 'CNY' };
  const result = engine.calculate(model);

  assert.equal(result.complete, true);
  approx(result.productCost.amount, 75);
  assert.equal(roundMoney(result.grossProfit), 88.16);
  assert.equal(roundMoney(result.netProfit), 58.26);
});

test('美元产品成本可换算到人民币结算币种', () => {
  const result = engine.calculate(baseModel({
    settlementCurrency: 'CNY',
    price: { amount: 100, currency: 'USD' },
    productCost: { amount: 75, currency: 'USD' }
  }));

  assert.equal(result.complete, true);
  approx(result.revenue, 678);
  approx(result.productCost.amount, 508.5);
  approx(result.grossProfit, 169.5);
});

test('金额模式保持金额，百分比模式随售价联动', () => {
  const costs = [
    { key: 'onsiteAds', mode: 'amount', value: 30 },
    { key: 'offsiteAds', mode: 'percent', value: 15 }
  ];
  const atTwoHundred = engine.calculate(baseModel({
    price: { amount: 200, currency: 'USD' },
    directCosts: costs
  }));
  const atThreeHundred = engine.calculate(baseModel({
    price: { amount: 300, currency: 'USD' },
    directCosts: costs
  }));

  const amount200 = atTwoHundred.directCosts.find(line => line.key === 'onsiteAds');
  const percent200 = atTwoHundred.directCosts.find(line => line.key === 'offsiteAds');
  const amount300 = atThreeHundred.directCosts.find(line => line.key === 'onsiteAds');
  const percent300 = atThreeHundred.directCosts.find(line => line.key === 'offsiteAds');

  approx(amount200.amount, 30);
  approx(amount200.percent, 15);
  approx(percent200.amount, 30);
  approx(percent200.percent, 15);
  approx(amount300.amount, 30);
  approx(amount300.percent, 10);
  approx(percent300.amount, 45);
  approx(percent300.percent, 15);
  approx(atThreeHundred.directTotal, 75);
});

test('非 USD/CNY 结算币使用手动平台币兑人民币汇率', () => {
  const result = engine.calculate(baseModel({
    price: { amount: 100, currency: 'USD' },
    settlementCurrency: 'EUR',
    usdCnyRate: 6.8,
    settlementCnyRate: 8.5,
    directCosts: [
      { key: 'platformCommission', mode: 'percent', value: 10 }
    ],
    productCost: { amount: 68, currency: 'CNY' },
    allocations: [
      { key: 'peopleAndRd', mode: 'percent', value: 10 }
    ]
  }));

  assert.equal(result.complete, true);
  approx(result.revenue, 80);
  approx(result.productCost.amount, 8);
  approx(result.directTotal, 16);
  approx(result.grossProfit, 64);
  approx(result.allocationTotal, 8);
  approx(result.netProfit, 56);
});

test('空白可选成本按 0 处理', () => {
  const result = engine.calculate(baseModel({
    directCosts: [
      { key: 'onsiteAds', mode: 'amount', value: '' },
      { key: 'storage', mode: 'percent', value: null }
    ],
    productCost: { amount: '', currency: 'CNY' },
    allocations: [
      { key: 'peopleAndRd', mode: 'percent', value: undefined }
    ]
  }));

  assert.equal(result.complete, true);
  approx(result.directTotal, 0);
  approx(result.allocationTotal, 0);
  approx(result.grossProfit, 100);
  approx(result.netProfit, 100);
});

test('总成本超过售价时保留负毛利与负净利润', () => {
  const result = engine.calculate(baseModel({
    directCosts: [
      { key: 'firstMile', mode: 'amount', value: 30 },
      { key: 'platformCommission', mode: 'percent', value: 15 }
    ],
    productCost: { amount: 80, currency: 'USD' },
    allocations: [
      { key: 'peopleAndRd', mode: 'percent', value: 10 }
    ]
  }));

  assert.equal(result.complete, true);
  approx(result.directTotal, 125);
  approx(result.grossProfit, -25);
  approx(result.grossMargin, -25);
  approx(result.netProfit, -35);
  approx(result.netMargin, -35);
  assert.ok(result.warnings.some(message => message.includes('毛利为负')));
  assert.ok(result.warnings.some(message => message.includes('净利润为负')));
});

test('售价为 0 或负数时报错', () => {
  for (const amount of [0, -1, '']) {
    const result = engine.calculate(baseModel({
      price: { amount, currency: 'USD' }
    }));
    assert.equal(result.complete, false);
    assert.match(result.errors.join('\n'), /定价/);
  }
});

test('USD/CNY 汇率非法时报错', () => {
  for (const usdCnyRate of [0, -6.78, '', 'not-a-rate']) {
    const result = engine.calculate(baseModel({ usdCnyRate }));
    assert.equal(result.complete, false);
    assert.match(result.errors.join('\n'), /美元兑人民币汇率/);
  }
});

test('非 USD/CNY 结算币缺少有效手动汇率时报错', () => {
  for (const settlementCnyRate of [0, -8, '', null]) {
    const result = engine.calculate(baseModel({
      settlementCurrency: 'GBP',
      settlementCnyRate
    }));
    assert.equal(result.complete, false);
    assert.match(result.errors.join('\n'), /平台币兑人民币汇率/);
  }
});

test('百分比费用计算过程不提前按分币四舍五入', () => {
  const result = engine.calculate(baseModel({
    price: { amount: 10, currency: 'USD' },
    directCosts: [
      { key: 'onsiteAds', mode: 'percent', value: 3.333 },
      { key: 'offsiteAds', mode: 'percent', value: 3.333 },
      { key: 'storage', mode: 'percent', value: 3.333 }
    ]
  }));

  assert.equal(result.complete, true);
  approx(result.directTotal, 0.9999, 1e-12);
  approx(result.grossProfit, 9.0001, 1e-12);
  assert.notEqual(result.grossProfit, 9.01);
  assert.equal(roundMoney(result.grossProfit), 9.00);
});

console.log(`\n${passed} passed, ${failed} failed`);
if (failed > 0) process.exitCode = 1;
