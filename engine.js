(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.CrossborderProfitEngine = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  const PLATFORM_OPTIONS = [
    'Amazon', 'Walmart', 'Shopify', 'Wayfair', 'Temu', 'TikTok Shop', 'eBay', '其他'
  ];

  const CURRENCY_META = {
    USD: { symbol: '$', name: '美元', decimals: 2 },
    CNY: { symbol: '¥', name: '人民币', decimals: 2 },
    EUR: { symbol: '€', name: '欧元', decimals: 2 },
    GBP: { symbol: '£', name: '英镑', decimals: 2 },
    CAD: { symbol: 'CA$', name: '加元', decimals: 2 },
    JPY: { symbol: '¥', name: '日元', decimals: 0 },
    AUD: { symbol: 'A$', name: '澳元', decimals: 2 },
    MXN: { symbol: 'MX$', name: '墨西哥比索', decimals: 2 }
  };

  const DIRECT_COST_DEFINITIONS = [
    { key: 'onsiteAds', group: '推广费用', label: '站内投放' },
    { key: 'offsiteAds', group: '推广费用', label: '站外投放' },
    { key: 'firstMile', group: '物流仓储', label: '头程运费' },
    { key: 'customs', group: '物流仓储', label: '清关关税' },
    { key: 'lastMile', group: '物流仓储', label: '尾程物流' },
    { key: 'storage', group: '物流仓储', label: '仓储成本' },
    { key: 'afterSales', group: '物流仓储', label: '售后成本' },
    { key: 'platformCommission', group: '平台及产品成本', label: '平台佣金' },
    { key: 'otherDirect', group: '平台及产品成本', label: '其他直接成本' }
  ];

  const ALLOCATION_DEFINITIONS = [
    { key: 'peopleAndRd', group: '分摊成本', label: '人力 & 研发' },
    { key: 'brandInvestment', group: '分摊成本', label: '品牌投入' },
    { key: 'otherAllocation', group: '分摊成本', label: '其他分摊' }
  ];

  function finiteNumber(value) {
    if (value === '' || value === null || value === undefined) return null;
    const number = Number(value);
    return Number.isFinite(number) ? number : null;
  }

  function nonNegative(value, field, errors) {
    const number = finiteNumber(value);
    if (number === null || number < 0) {
      errors.push(`${field}需要是大于或等于 0 的数字`);
      return null;
    }
    return number;
  }

  function optionalNonNegative(value, field, errors) {
    if (value === '' || value === null || value === undefined) return 0;
    return nonNegative(value, field, errors);
  }

  function positive(value, field, errors) {
    const number = finiteNumber(value);
    if (number === null || number <= 0) {
      errors.push(`${field}需要是大于 0 的数字`);
      return null;
    }
    return number;
  }

  function getSettlementCnyRate(currency, usdCnyRate, manualSettlementCnyRate) {
    if (currency === 'CNY') return 1;
    if (currency === 'USD') return usdCnyRate;
    return manualSettlementCnyRate;
  }

  function convertToSettlement(amount, sourceCurrency, settlementCurrency, rates) {
    const value = finiteNumber(amount);
    const usdCny = finiteNumber(rates && rates.usdCnyRate);
    const settlementCny = getSettlementCnyRate(
      settlementCurrency,
      usdCny,
      finiteNumber(rates && rates.settlementCnyRate)
    );
    if (value === null || value < 0 || !CURRENCY_META[sourceCurrency] || !CURRENCY_META[settlementCurrency]) return null;
    if (!(usdCny > 0) || !(settlementCny > 0)) return null;
    let cnyValue;
    if (sourceCurrency === 'CNY') cnyValue = value;
    else if (sourceCurrency === 'USD') cnyValue = value * usdCny;
    else if (sourceCurrency === settlementCurrency) cnyValue = value * settlementCny;
    else return null;
    return cnyValue / settlementCny;
  }

  function convertBetweenSourceCurrencies(amount, fromCurrency, toCurrency, usdCnyRate) {
    const value = finiteNumber(amount);
    const rate = finiteNumber(usdCnyRate);
    if (value === null || !(rate > 0)) return null;
    if (fromCurrency === toCurrency) return value;
    if (fromCurrency === 'USD' && toCurrency === 'CNY') return value * rate;
    if (fromCurrency === 'CNY' && toCurrency === 'USD') return value / rate;
    return null;
  }

  function resolveLine(line, price, field, errors) {
    const mode = line && line.mode === 'percent' ? 'percent' : 'amount';
    const raw = optionalNonNegative(line && line.value, field, errors);
    if (raw === null) return null;
    const amount = mode === 'percent' ? price * raw / 100 : raw;
    const percent = price > 0 ? amount / price * 100 : null;
    return { ...line, mode, value: raw, amount, percent };
  }

  function calculate(model) {
    const errors = [];
    const warnings = [];
    const settlementCurrency = model && CURRENCY_META[model.settlementCurrency]
      ? model.settlementCurrency
      : 'USD';
    const priceAmount = positive(model && model.price && model.price.amount, '定价', errors);
    const priceCurrency = model && model.price && ['USD', 'CNY'].includes(model.price.currency)
      ? model.price.currency
      : 'USD';
    const usdCnyRate = positive(model && model.usdCnyRate, '美元兑人民币汇率', errors);
    const settlementCnyRate = getSettlementCnyRate(
      settlementCurrency,
      usdCnyRate,
      positive(
        settlementCurrency === 'USD' || settlementCurrency === 'CNY'
          ? (settlementCurrency === 'USD' ? model && model.usdCnyRate : 1)
          : model && model.settlementCnyRate,
        '平台币兑人民币汇率',
        settlementCurrency === 'USD' || settlementCurrency === 'CNY' ? [] : errors
      )
    );
    if (errors.length) return { complete: false, errors, warnings, settlementCurrency };

    const revenue = convertToSettlement(priceAmount, priceCurrency, settlementCurrency, {
      usdCnyRate,
      settlementCnyRate
    });
    if (!(revenue > 0)) {
      errors.push('定价无法按当前汇率折算为平台币');
      return { complete: false, errors, warnings, settlementCurrency };
    }

    const directCosts = [];
    for (const definition of DIRECT_COST_DEFINITIONS) {
      const line = (model.directCosts || []).find(item => item.key === definition.key) || {
        ...definition, mode: 'amount', value: 0
      };
      const resolved = resolveLine({ ...definition, ...line }, revenue, definition.label, errors);
      if (resolved) directCosts.push(resolved);
    }

    const productCostAmount = optionalNonNegative(model && model.productCost && model.productCost.amount, '产品成本', errors);
    const productCostCurrency = model && model.productCost && ['USD', 'CNY'].includes(model.productCost.currency)
      ? model.productCost.currency
      : 'USD';
    const productCostValue = convertToSettlement(productCostAmount, productCostCurrency, settlementCurrency, {
      usdCnyRate,
      settlementCnyRate
    });
    if (productCostValue === null) errors.push('产品成本无法按当前汇率折算');

    const allocations = [];
    for (const definition of ALLOCATION_DEFINITIONS) {
      const line = (model.allocations || []).find(item => item.key === definition.key) || {
        ...definition, mode: 'percent', value: 0
      };
      const resolved = resolveLine({ ...definition, ...line }, revenue, definition.label, errors);
      if (resolved) allocations.push(resolved);
    }
    if (errors.length) return { complete: false, errors, warnings, settlementCurrency };

    for (const line of [...directCosts, ...allocations]) {
      if (line.percent > 100) warnings.push(`${line.label}超过售价的 100%`);
    }

    const directLineTotal = directCosts.reduce((sum, line) => sum + line.amount, 0);
    const directTotal = directLineTotal + productCostValue;
    const grossProfit = revenue - directTotal;
    const grossMargin = grossProfit / revenue * 100;
    const allocationTotal = allocations.reduce((sum, line) => sum + line.amount, 0);
    const netProfit = grossProfit - allocationTotal;
    const netMargin = netProfit / revenue * 100;
    const totalCost = directTotal + allocationTotal;
    if (grossProfit < 0) warnings.push('直接成本已超过售价，毛利为负');
    if (netProfit < 0) warnings.push('总成本已超过售价，净利润为负');

    return {
      complete: true,
      errors,
      warnings,
      platform: model.platform || 'Amazon',
      productModel: model.productModel || '',
      settlementCurrency,
      usdCnyRate,
      settlementCnyRate,
      revenue,
      directCosts,
      productCost: {
        label: '产品成本',
        sourceAmount: productCostAmount,
        sourceCurrency: productCostCurrency,
        amount: productCostValue,
        percent: productCostValue / revenue * 100
      },
      directTotal,
      grossProfit,
      grossMargin,
      allocations,
      allocationTotal,
      netProfit,
      netMargin,
      totalCost,
      totalCostRate: totalCost / revenue * 100
    };
  }

  function createDefaultModel() {
    return {
      platform: 'Amazon',
      customPlatform: '',
      productModel: 'DEMO-SKU-299',
      price: { amount: 299, currency: 'USD' },
      settlementCurrency: 'USD',
      usdCnyRate: 6.78,
      settlementCnyRate: 6.78,
      directCosts: [
        { key: 'onsiteAds', mode: 'amount', value: 0 },
        { key: 'offsiteAds', mode: 'amount', value: 0 },
        { key: 'firstMile', mode: 'amount', value: 56 },
        { key: 'customs', mode: 'amount', value: 0 },
        { key: 'lastMile', mode: 'amount', value: 32 },
        { key: 'storage', mode: 'amount', value: 2.99 },
        { key: 'afterSales', mode: 'percent', value: 3 },
        { key: 'platformCommission', mode: 'percent', value: 12 },
        { key: 'otherDirect', mode: 'amount', value: 0 }
      ],
      productCost: { amount: 75, currency: 'USD' },
      allocations: [
        { key: 'peopleAndRd', mode: 'percent', value: 10 },
        { key: 'brandInvestment', mode: 'percent', value: 0 },
        { key: 'otherAllocation', mode: 'percent', value: 0 }
      ]
    };
  }

  function formatMoney(value, currency) {
    const meta = CURRENCY_META[currency] || CURRENCY_META.USD;
    return new Intl.NumberFormat('zh-CN', {
      style: 'currency',
      currency: CURRENCY_META[currency] ? currency : 'USD',
      minimumFractionDigits: meta.decimals,
      maximumFractionDigits: meta.decimals
    }).format(Number(value) || 0);
  }

  function formatPercent(value) {
    return `${(Number(value) || 0).toFixed(2)}%`;
  }

  return {
    PLATFORM_OPTIONS,
    CURRENCY_META,
    DIRECT_COST_DEFINITIONS,
    ALLOCATION_DEFINITIONS,
    finiteNumber,
    getSettlementCnyRate,
    convertToSettlement,
    convertBetweenSourceCurrencies,
    calculate,
    createDefaultModel,
    formatMoney,
    formatPercent
  };
});
