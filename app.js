(function () {
  'use strict';

  const Engine = window.CrossborderProfitEngine;
  const Exporter = window.CrossborderProfitExport;
  const byId = id => document.getElementById(id);
  const clone = value => typeof structuredClone === 'function'
    ? structuredClone(value)
    : JSON.parse(JSON.stringify(value));

  let state = Engine.createDefaultModel();
  let lastResult = null;
  let announceTimer = null;
  let toastTimer = null;
  let pendingCurrencyConversion = null;
  const currencyCnyRates = { USD: 6.78, CNY: 1, EUR: '', GBP: '', CAD: '', JPY: '', AUD: '', MXN: '' };

  const elements = {
    platform: byId('platformSelect'),
    customPlatformField: byId('customPlatformField'),
    customPlatform: byId('customPlatformInput'),
    productModel: byId('productModelInput'),
    priceAmount: byId('priceAmountInput'),
    priceCurrency: byId('priceCurrencySelect'),
    settlementCurrency: byId('settlementCurrencySelect'),
    usdCnyRate: byId('usdCnyRateInput'),
    settlementRateField: byId('settlementRateField'),
    settlementRateCode: byId('settlementRateCode'),
    settlementCnyRate: byId('settlementCnyRateInput'),
    priceConverted: byId('priceConvertedText'),
    directLedger: byId('directCostLedger'),
    allocationLedger: byId('allocationLedger'),
    productCost: byId('productCostInput'),
    productCostCurrency: byId('productCostCurrencySelect'),
    productCostConverted: byId('productCostConverted'),
    productCostPercent: byId('productCostPercent'),
    validationBanner: byId('validationBanner'),
    resultState: byId('resultState'),
    revenue: byId('revenueValue'),
    netProfitHero: byId('netProfitHero'),
    netProfit: byId('netProfitValue'),
    netMargin: byId('netMarginValue'),
    profitStateText: byId('profitStateText'),
    grossProfit: byId('grossProfitValue'),
    grossMargin: byId('grossMarginValue'),
    directTotal: byId('directTotalValue'),
    directRate: byId('directRateValue'),
    allocationTotal: byId('allocationTotalValue'),
    allocationRate: byId('allocationRateValue'),
    directBar: byId('directBar'),
    allocationBar: byId('allocationBar'),
    profitBar: byId('profitBar'),
    costTrack: byId('costTrack'),
    breakdownBody: byId('breakdownBody'),
    warningList: byId('warningList'),
    reset: byId('resetSampleButton'),
    copyImage: byId('copyImageButton'),
    exportExcel: byId('exportExcelButton'),
    toast: byId('toast'),
    liveRegion: byId('liveRegion')
  };

  function getPlatformLabel() {
    return state.platform === '其他'
      ? (state.customPlatform.trim() || '其他平台')
      : state.platform;
  }

  function getLine(collection, key) {
    return collection.find(line => line.key === key);
  }

  function getResolvedLine(key, allocation) {
    if (!lastResult || !lastResult.complete) return null;
    const collection = allocation ? lastResult.allocations : lastResult.directCosts;
    return collection.find(line => line.key === key) || null;
  }

  function renderLedger(container, definitions, collectionName, allocation) {
    container.textContent = '';
    let previousGroup = '';
    for (const definition of definitions) {
      const row = document.createElement('div');
      row.className = 'cost-row';
      row.dataset.key = definition.key;
      row.dataset.collection = collectionName;
      row.dataset.groupStart = previousGroup && previousGroup !== definition.group ? 'true' : 'false';
      previousGroup = definition.group;

      const name = document.createElement('div');
      name.className = 'cost-name';
      const group = document.createElement('span');
      group.className = 'cost-group';
      group.textContent = definition.group;
      const strong = document.createElement('strong');
      strong.textContent = definition.label;
      name.append(group, strong);

      const modes = document.createElement('div');
      modes.className = 'mode-toggle';
      modes.setAttribute('role', 'group');
      modes.setAttribute('aria-label', `${definition.label}输入方式`);
      if (!allocation) modes.append(createModeButton(definition, collectionName, 'amount', '金额'));
      modes.append(createModeButton(definition, collectionName, 'percent', '比例'));

      const inputWrap = document.createElement('label');
      inputWrap.className = 'cost-input-wrap';
      const inputLabel = document.createElement('span');
      inputLabel.className = 'sr-only';
      inputLabel.textContent = definition.label;
      const input = document.createElement('input');
      input.type = 'number';
      input.min = '0';
      input.step = '0.01';
      input.inputMode = 'decimal';
      input.dataset.role = 'cost-input';
      input.addEventListener('input', () => {
        const line = getLine(state[collectionName], definition.key);
        line.value = input.value;
        update();
      });
      const suffix = document.createElement('span');
      suffix.className = 'cost-input-suffix';
      suffix.dataset.role = 'cost-suffix';
      inputWrap.append(inputLabel, input, suffix);

      const derived = document.createElement('div');
      derived.className = 'cost-derived';
      const amount = document.createElement('strong');
      amount.dataset.role = 'derived-amount';
      const percent = document.createElement('span');
      percent.dataset.role = 'derived-percent';
      derived.append(amount, percent);

      row.append(name, modes, inputWrap, derived);
      container.append(row);
    }
  }

  function createModeButton(definition, collectionName, mode, label) {
    const button = document.createElement('button');
    button.type = 'button';
    button.dataset.mode = mode;
    button.textContent = label;
    button.addEventListener('click', () => {
      const line = getLine(state[collectionName], definition.key);
      if (line.mode === mode) return;
      const resolved = getResolvedLine(definition.key, collectionName === 'allocations');
      line.mode = mode;
      if (resolved) line.value = mode === 'amount' ? resolved.amount : resolved.percent;
      update();
    });
    return button;
  }

  function syncLedgerRows(container, collectionName, allocation) {
    for (const row of container.querySelectorAll('.cost-row')) {
      const key = row.dataset.key;
      const line = getLine(state[collectionName], key);
      const resolved = getResolvedLine(key, allocation);
      for (const button of row.querySelectorAll('[data-mode]')) {
        const active = button.dataset.mode === line.mode;
        button.setAttribute('aria-pressed', String(active));
        button.disabled = Boolean(pendingCurrencyConversion);
      }
      const input = row.querySelector('[data-role="cost-input"]');
      if (document.activeElement !== input) input.value = line.value === '' ? '' : trimNumber(line.value);
      input.disabled = Boolean(pendingCurrencyConversion);
      row.querySelector('[data-role="cost-suffix"]').textContent = line.mode === 'percent'
        ? '%'
        : state.settlementCurrency;
      row.querySelector('[data-role="derived-amount"]').textContent = resolved
        ? Engine.formatMoney(resolved.amount, state.settlementCurrency)
        : '—';
      row.querySelector('[data-role="derived-percent"]').textContent = resolved
        ? `${Engine.formatPercent(resolved.percent)} · 占售价`
        : '—';
    }
  }

  function trimNumber(value) {
    const number = Number(value);
    if (!Number.isFinite(number)) return value;
    return String(Math.round(number * 1000000) / 1000000);
  }

  function readFormIntoState() {
    state.platform = elements.platform.value;
    state.customPlatform = elements.customPlatform.value;
    state.productModel = elements.productModel.value;
    state.price.amount = elements.priceAmount.value;
    state.price.currency = elements.priceCurrency.value;
    state.settlementCurrency = elements.settlementCurrency.value;
    state.usdCnyRate = elements.usdCnyRate.value;
    state.settlementCnyRate = state.settlementCurrency === 'USD'
      ? elements.usdCnyRate.value
      : state.settlementCurrency === 'CNY'
        ? 1
        : elements.settlementCnyRate.value;
    state.productCost.amount = elements.productCost.value;
    state.productCost.currency = elements.productCostCurrency.value;
  }

  function syncFormFromState() {
    elements.platform.value = state.platform;
    elements.customPlatform.value = state.customPlatform || '';
    elements.customPlatformField.hidden = state.platform !== '其他';
    elements.productModel.value = state.productModel || '';
    elements.priceAmount.value = state.price.amount;
    elements.priceCurrency.value = state.price.currency;
    elements.settlementCurrency.value = state.settlementCurrency;
    elements.usdCnyRate.value = state.usdCnyRate;
    elements.productCost.value = state.productCost.amount;
    elements.productCostCurrency.value = state.productCost.currency;
    const manualRate = !['USD', 'CNY'].includes(state.settlementCurrency);
    elements.settlementRateField.hidden = !manualRate;
    elements.settlementRateCode.textContent = state.settlementCurrency;
    elements.settlementCnyRate.value = manualRate ? (state.settlementCnyRate || '') : '';
  }

  function update() {
    readFormIntoState();
    lastResult = Engine.calculate(state);
    syncLedgerRows(elements.directLedger, 'directCosts', false);
    syncLedgerRows(elements.allocationLedger, 'allocations', true);
    renderResult(lastResult);
  }

  function renderResult(result) {
    const complete = result && result.complete;
    elements.validationBanner.hidden = complete;
    elements.resultState.classList.toggle('is-error', !complete);
    elements.resultState.textContent = complete ? '已完成' : '输入不完整';
    elements.copyImage.disabled = !complete;
    elements.exportExcel.disabled = !complete;
    if (!complete) {
      elements.validationBanner.textContent = (result.errors || []).join('；');
      for (const target of [
        elements.revenue, elements.netProfit, elements.netMargin, elements.grossProfit,
        elements.grossMargin, elements.directTotal, elements.directRate,
        elements.allocationTotal, elements.allocationRate,
        elements.productCostConverted, elements.productCostPercent
      ]) target.textContent = '—';
      elements.profitStateText.textContent = '等待有效定价与汇率';
      elements.breakdownBody.innerHTML = '<tr><td colspan="3">补齐定价和汇率后生成利润明细。</td></tr>';
      setBarWidths(0, 0, 0, false);
      elements.warningList.hidden = true;
      return;
    }

    const currency = result.settlementCurrency;
    elements.revenue.textContent = Engine.formatMoney(result.revenue, currency);
    elements.netProfit.textContent = Engine.formatMoney(result.netProfit, currency);
    elements.netMargin.textContent = Engine.formatPercent(result.netMargin);
    elements.grossProfit.textContent = Engine.formatMoney(result.grossProfit, currency);
    elements.grossMargin.textContent = Engine.formatPercent(result.grossMargin);
    elements.directTotal.textContent = Engine.formatMoney(result.directTotal, currency);
    elements.directRate.textContent = Engine.formatPercent(result.directTotal / result.revenue * 100);
    elements.allocationTotal.textContent = Engine.formatMoney(result.allocationTotal, currency);
    elements.allocationRate.textContent = Engine.formatPercent(result.allocationTotal / result.revenue * 100);
    elements.productCostConverted.textContent = Engine.formatMoney(result.productCost.amount, currency);
    elements.productCostPercent.textContent = `${Engine.formatPercent(result.productCost.percent)} · 占售价`;
    elements.priceConverted.textContent = `折算为 ${Engine.formatMoney(result.revenue, currency)}`;
    const profitable = result.netProfit >= 0;
    elements.netProfitHero.classList.toggle('is-loss', !profitable);
    elements.profitStateText.textContent = profitable ? '盈利 · 扣除全部分摊后' : '亏损 · 总成本高于售价';

    setBarWidths(
      result.directTotal / result.revenue * 100,
      result.allocationTotal / result.revenue * 100,
      Math.max(0, result.netMargin),
      profitable
    );
    renderBreakdown(result);
    renderWarnings(result.warnings);
    announceResult(result);
  }

  function setBarWidths(direct, allocation, profit, profitable) {
    const total = Math.max(100, direct + allocation + profit);
    elements.directBar.style.width = `${Math.max(0, direct) / total * 100}%`;
    elements.allocationBar.style.width = `${Math.max(0, allocation) / total * 100}%`;
    elements.profitBar.style.width = `${Math.max(0, profit) / total * 100}%`;
    elements.costTrack.style.background = profitable ? 'var(--surface)' : '#f2c5c2';
    elements.costTrack.setAttribute(
      'aria-label',
      `直接成本 ${direct.toFixed(2)}%，分摊成本 ${allocation.toFixed(2)}%，净利润 ${profit.toFixed(2)}%`
    );
  }

  function renderBreakdown(result) {
    const rows = [];
    for (const line of result.directCosts) {
      if (line.amount === 0) continue;
      rows.push({ label: line.label, amount: line.amount, percent: line.percent });
    }
    rows.push({ label: '产品成本', amount: result.productCost.amount, percent: result.productCost.percent });
    rows.push({ label: '毛利（分摊前）', amount: result.grossProfit, percent: result.grossMargin, className: 'is-subtotal' });
    for (const line of result.allocations) {
      if (line.amount === 0) continue;
      rows.push({ label: line.label, amount: line.amount, percent: line.percent });
    }
    rows.push({ label: '净利润', amount: result.netProfit, percent: result.netMargin, className: 'is-net' });
    elements.breakdownBody.textContent = '';
    for (const row of rows) {
      const tr = document.createElement('tr');
      if (row.className) tr.className = row.className;
      const label = document.createElement('td');
      const amount = document.createElement('td');
      const percent = document.createElement('td');
      label.textContent = row.label;
      amount.textContent = Engine.formatMoney(row.amount, result.settlementCurrency);
      percent.textContent = Engine.formatPercent(row.percent);
      tr.append(label, amount, percent);
      elements.breakdownBody.append(tr);
    }
  }

  function renderWarnings(warnings) {
    if (!warnings || !warnings.length) {
      elements.warningList.hidden = true;
      elements.warningList.textContent = '';
      return;
    }
    const list = document.createElement('ul');
    for (const warning of warnings) {
      const item = document.createElement('li');
      item.textContent = warning;
      list.append(item);
    }
    elements.warningList.textContent = '';
    elements.warningList.append(list);
    elements.warningList.hidden = false;
  }

  function announceResult(result) {
    clearTimeout(announceTimer);
    announceTimer = setTimeout(() => {
      elements.liveRegion.textContent = `毛利 ${Engine.formatMoney(result.grossProfit, result.settlementCurrency)}，净利润 ${Engine.formatMoney(result.netProfit, result.settlementCurrency)}，净利率 ${Engine.formatPercent(result.netMargin)}`;
    }, 450);
  }

  function convertSourceInput(amountElement, fromCurrency, toCurrency) {
    const converted = Engine.convertBetweenSourceCurrencies(amountElement.value, fromCurrency, toCurrency, elements.usdCnyRate.value);
    if (converted !== null) amountElement.value = trimNumber(converted);
  }

  function snapshotAmountLines() {
    return Object.fromEntries(state.directCosts
      .filter(line => line.mode === 'amount')
      .map(line => [line.key, Number(line.value) || 0]));
  }

  function convertAmountLines(values, oldRate, newRate) {
    if (!(oldRate > 0) || !(newRate > 0)) return;
    for (const line of state.directCosts) {
      if (line.mode !== 'amount' || !Object.prototype.hasOwnProperty.call(values, line.key)) continue;
      line.value = values[line.key] * oldRate / newRate;
    }
  }

  function handleSettlementCurrencyChange() {
    const oldCurrency = state.settlementCurrency;
    const currentUsdCnyRate = Number(elements.usdCnyRate.value);
    const oldRate = Engine.getSettlementCnyRate(
      oldCurrency,
      currentUsdCnyRate,
      oldCurrency === 'USD' ? currentUsdCnyRate : Number(state.settlementCnyRate)
    );
    const baseValues = pendingCurrencyConversion ? pendingCurrencyConversion.values : snapshotAmountLines();
    const baseRate = pendingCurrencyConversion ? pendingCurrencyConversion.oldRate : oldRate;
    const nextCurrency = elements.settlementCurrency.value;
    const nextRate = nextCurrency === 'USD'
      ? Number(elements.usdCnyRate.value)
      : nextCurrency === 'CNY'
        ? 1
        : Number(currencyCnyRates[nextCurrency]);

    state.settlementCurrency = nextCurrency;
    state.usdCnyRate = elements.usdCnyRate.value;
    if (nextRate > 0) {
      convertAmountLines(baseValues, baseRate, nextRate);
      state.settlementCnyRate = nextRate;
      pendingCurrencyConversion = null;
    } else {
      pendingCurrencyConversion = { oldRate: baseRate, values: baseValues };
      state.settlementCnyRate = '';
    }
    syncFormFromState();
    update();
  }

  function handleManualSettlementRate() {
    const rate = Number(elements.settlementCnyRate.value);
    currencyCnyRates[state.settlementCurrency] = elements.settlementCnyRate.value;
    if (pendingCurrencyConversion && rate > 0) {
      convertAmountLines(pendingCurrencyConversion.values, pendingCurrencyConversion.oldRate, rate);
      pendingCurrencyConversion = null;
    }
    state.settlementCnyRate = elements.settlementCnyRate.value;
    update();
  }

  function showToast(message) {
    clearTimeout(toastTimer);
    elements.toast.textContent = message;
    elements.toast.hidden = false;
    toastTimer = setTimeout(() => { elements.toast.hidden = true; }, 3200);
  }

  function buildCurrentSnapshot() {
    if (!lastResult || !lastResult.complete) throw new Error('请先补齐有效定价与汇率');
    return Exporter.buildSnapshot(state, lastResult, {
      platformLabel: getPlatformLabel(),
      title: '跨境电商利润模型计算器'
    });
  }

  function exportFilename(extension) {
    const now = new Date();
    const date = [
      now.getFullYear(),
      String(now.getMonth() + 1).padStart(2, '0'),
      String(now.getDate()).padStart(2, '0')
    ].join('');
    const model = state.productModel.trim() || '未填写型号';
    return Exporter.sanitizeFilename(
      `跨境利润模型_${getPlatformLabel()}_${model}_${date}`,
      extension
    );
  }

  async function handleCopyImage() {
    elements.copyImage.disabled = true;
    elements.copyImage.setAttribute('aria-busy', 'true');
    try {
      const snapshot = buildCurrentSnapshot();
      const outcome = await Exporter.copyOrDownloadImage(snapshot, {
        filename: exportFilename('png'),
        renderOptions: {
          palette: {
            page: '#f3f5ef',
            card: '#ffffff',
            ink: '#11140f',
            muted: '#596159',
            line: '#cfd5cc',
            accent: '#6bcd37',
            accentSoft: '#e4f6d8',
            success: '#125c35',
            successSoft: '#e4f6e9',
            danger: '#b43c3c',
            dangerSoft: '#fff0ef'
          }
        }
      });
      showToast(outcome.action === 'copied'
        ? '利润结果已复制为图片'
        : `当前浏览器不支持图片剪贴板，已下载 ${outcome.filename}`);
    } catch (error) {
      showToast(`图片导出失败：${error && error.message ? error.message : '请重试'}`);
    } finally {
      elements.copyImage.removeAttribute('aria-busy');
      elements.copyImage.disabled = !(lastResult && lastResult.complete);
    }
  }

  function handleExportExcel() {
    elements.exportExcel.disabled = true;
    try {
      const snapshot = buildCurrentSnapshot();
      const outcome = Exporter.downloadXlsx(snapshot, { filename: exportFilename('xlsx') });
      showToast(`已导出 ${outcome.filename}`);
    } catch (error) {
      showToast(`Excel 导出失败：${error && error.message ? error.message : '请重试'}`);
    } finally {
      elements.exportExcel.disabled = !(lastResult && lastResult.complete);
    }
  }

  function conciseResult(result) {
    return {
      platform: getPlatformLabel(),
      productModel: state.productModel || '',
      currency: result.settlementCurrency,
      revenue: result.revenue,
      directCost: result.directTotal,
      grossProfit: result.grossProfit,
      grossMargin: result.grossMargin,
      allocationCost: result.allocationTotal,
      netProfit: result.netProfit,
      netMargin: result.netMargin,
      warnings: result.warnings || []
    };
  }

  function registerWebMcpTools() {
    const context = document.modelContext;
    if (!context || typeof context.registerTool !== 'function') return;
    const lifecycle = new AbortController();
    const costSchema = Object.fromEntries(Engine.DIRECT_COST_DEFINITIONS.map(definition => [
      definition.key,
      {
        type: 'object',
        properties: {
          mode: { type: 'string', enum: ['amount', 'percent'] },
          value: { type: 'number', minimum: 0 }
        },
        required: ['mode', 'value'],
        additionalProperties: false
      }
    ]));
    const allocationSchema = Object.fromEntries(Engine.ALLOCATION_DEFINITIONS.map(definition => [
      definition.key,
      { type: 'number', minimum: 0 }
    ]));
    const register = tool => {
      try {
        void Promise.resolve(context.registerTool(tool, { signal: lifecycle.signal }))
          .catch(error => console.warn('WebMCP tool registration failed', error));
      } catch (error) {
        console.warn('WebMCP tool registration failed', error);
      }
    };

    register({
      name: 'read_profit_summary',
      title: '读取利润测算结果',
      description: '读取当前跨境电商单品利润模型中已经显示的收入、毛利、净利和成本摘要。',
      inputSchema: { type: 'object', properties: {}, additionalProperties: false },
      annotations: { readOnlyHint: true, untrustedContentHint: false },
      execute() {
        if (!lastResult || !lastResult.complete) throw new Error('当前输入尚未形成有效利润结果');
        return conciseResult(lastResult);
      }
    });

    register({
      name: 'update_profit_model',
      title: '更新利润模型',
      description: '更新页面中的单品定价、币种和成本输入并重新计算。金额型直接成本按平台/报告币种解释，比例均以售价为基数。',
      inputSchema: {
        type: 'object',
        properties: {
          platform: { type: 'string', enum: Engine.PLATFORM_OPTIONS },
          customPlatform: { type: 'string' },
          productModel: { type: 'string' },
          price: {
            type: 'object',
            properties: {
              amount: { type: 'number', exclusiveMinimum: 0 },
              currency: { type: 'string', enum: ['USD', 'CNY'] }
            },
            additionalProperties: false
          },
          settlementCurrency: { type: 'string', enum: Object.keys(Engine.CURRENCY_META) },
          usdCnyRate: { type: 'number', exclusiveMinimum: 0 },
          settlementCnyRate: { type: 'number', exclusiveMinimum: 0 },
          productCost: {
            type: 'object',
            properties: {
              amount: { type: 'number', minimum: 0 },
              currency: { type: 'string', enum: ['USD', 'CNY'] }
            },
            additionalProperties: false
          },
          directCosts: { type: 'object', properties: costSchema, additionalProperties: false },
          allocations: { type: 'object', properties: allocationSchema, additionalProperties: false }
        },
        additionalProperties: false
      },
      annotations: { readOnlyHint: false, untrustedContentHint: false },
      execute(input) {
        if (!input || typeof input !== 'object' || Array.isArray(input)) throw new Error('输入必须是对象');
        const next = clone(state);
        if (input.platform !== undefined) next.platform = input.platform;
        if (input.customPlatform !== undefined) next.customPlatform = String(input.customPlatform);
        if (input.productModel !== undefined) next.productModel = String(input.productModel);
        if (input.price) next.price = { ...next.price, ...input.price };
        if (input.usdCnyRate !== undefined) next.usdCnyRate = input.usdCnyRate;
        if (input.settlementCurrency !== undefined) next.settlementCurrency = input.settlementCurrency;
        if (input.settlementCnyRate !== undefined) next.settlementCnyRate = input.settlementCnyRate;
        if (input.productCost) next.productCost = { ...next.productCost, ...input.productCost };
        if (input.directCosts) {
          next.directCosts = next.directCosts.map(line => input.directCosts[line.key]
            ? { ...line, ...input.directCosts[line.key] }
            : line);
        }
        if (input.allocations) {
          next.allocations = next.allocations.map(line => Object.prototype.hasOwnProperty.call(input.allocations, line.key)
            ? { ...line, mode: 'percent', value: input.allocations[line.key] }
            : line);
        }
        const calculated = Engine.calculate(next);
        if (!calculated.complete) throw new Error(calculated.errors.join('；'));
        state = next;
        currencyCnyRates.USD = state.usdCnyRate;
        if (!['USD', 'CNY'].includes(state.settlementCurrency)) {
          currencyCnyRates[state.settlementCurrency] = state.settlementCnyRate;
        }
        pendingCurrencyConversion = null;
        syncFormFromState();
        update();
        return conciseResult(lastResult);
      }
    });
  }

  function bindEvents() {
    elements.platform.addEventListener('change', () => {
      state.platform = elements.platform.value;
      elements.customPlatformField.hidden = state.platform !== '其他';
      update();
    });
    elements.customPlatform.addEventListener('input', update);
    elements.productModel.addEventListener('input', update);
    elements.priceAmount.addEventListener('input', update);
    elements.usdCnyRate.addEventListener('input', () => {
      currencyCnyRates.USD = elements.usdCnyRate.value;
      update();
    });
    elements.priceCurrency.addEventListener('change', event => {
      const previous = state.price.currency;
      const next = event.target.value;
      convertSourceInput(elements.priceAmount, previous, next);
      state.price.currency = next;
      update();
    });
    elements.productCost.addEventListener('input', update);
    elements.productCostCurrency.addEventListener('change', event => {
      const previous = state.productCost.currency;
      const next = event.target.value;
      convertSourceInput(elements.productCost, previous, next);
      state.productCost.currency = next;
      update();
    });
    elements.settlementCurrency.addEventListener('change', handleSettlementCurrencyChange);
    elements.settlementCnyRate.addEventListener('input', handleManualSettlementRate);
    elements.reset.addEventListener('click', () => {
      state = Engine.createDefaultModel();
      currencyCnyRates.USD = state.usdCnyRate;
      pendingCurrencyConversion = null;
      syncFormFromState();
      update();
      showToast('已恢复公开合成示例');
    });
    elements.copyImage.addEventListener('click', handleCopyImage);
    elements.exportExcel.addEventListener('click', handleExportExcel);
  }

  function init() {
    renderLedger(elements.directLedger, Engine.DIRECT_COST_DEFINITIONS, 'directCosts', false);
    renderLedger(elements.allocationLedger, Engine.ALLOCATION_DEFINITIONS, 'allocations', true);
    syncFormFromState();
    bindEvents();
    update();
    window.CrossborderProfitApp = {
      getState: () => clone(state),
      getResult: () => clone(lastResult),
      getPlatformLabel
    };
    registerWebMcpTools();
  }

  init();
})();
