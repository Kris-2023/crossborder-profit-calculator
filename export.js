(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  if (root) root.CrossborderProfitExport = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const PNG_MIME = 'image/png';
  const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const CURRENCY_META = {
    USD: { symbol: '$', decimals: 2 },
    CNY: { symbol: '¥', decimals: 2 },
    EUR: { symbol: '€', decimals: 2 },
    GBP: { symbol: '£', decimals: 2 },
    CAD: { symbol: 'CA$', decimals: 2 },
    JPY: { symbol: '¥', decimals: 0 },
    AUD: { symbol: 'A$', decimals: 2 },
    MXN: { symbol: 'MX$', decimals: 2 }
  };

  const DEFAULT_PALETTE = {
    page: '#f7f3ea',
    card: '#fffdf8',
    ink: '#171715',
    muted: '#666159',
    line: '#ddd6ca',
    accent: '#a84b32',
    accentSoft: '#efe0d8',
    success: '#2d6a4f',
    successSoft: '#dcebe2',
    danger: '#b42318',
    dangerSoft: '#f4d8d5',
    white: '#ffffff'
  };

  function normalizeText(value, maxLength) {
    const source = value === null || value === undefined ? '' : String(value);
    let output = '';
    for (let index = 0; index < source.length; index += 1) {
      const code = source.charCodeAt(index);
      if (code === 0x09 || code === 0x0a || code === 0x0d) {
        output += source[index];
        continue;
      }
      if (code < 0x20 || code === 0xfffe || code === 0xffff) continue;
      if (code >= 0xd800 && code <= 0xdbff) {
        const next = source.charCodeAt(index + 1);
        if (next >= 0xdc00 && next <= 0xdfff) {
          output += source[index] + source[index + 1];
          index += 1;
        } else {
          output += '\ufffd';
        }
        continue;
      }
      if (code >= 0xdc00 && code <= 0xdfff) {
        output += '\ufffd';
        continue;
      }
      output += source[index];
    }
    const limit = Number.isFinite(maxLength) ? Math.max(0, maxLength) : 32767;
    return output.length > limit ? output.slice(0, limit) : output;
  }

  function xmlEscape(value) {
    return normalizeText(value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function finiteNumber(value, field, fallback) {
    const number = Number(value);
    if (Number.isFinite(number)) return Object.is(number, -0) ? 0 : number;
    if (fallback !== undefined) return fallback;
    throw new TypeError(`${field || '数值'}必须是有限数字`);
  }

  function safeCurrency(value) {
    const code = normalizeText(value || 'USD', 8).toUpperCase();
    return CURRENCY_META[code] ? code : 'USD';
  }

  function formatMoney(value, currency) {
    const code = safeCurrency(currency);
    const meta = CURRENCY_META[code];
    const number = finiteNumber(value, '金额', 0);
    try {
      return new Intl.NumberFormat('zh-CN', {
        style: 'currency',
        currency: code,
        minimumFractionDigits: meta.decimals,
        maximumFractionDigits: meta.decimals
      }).format(number);
    } catch (_) {
      return `${meta.symbol}${number.toFixed(meta.decimals)}`;
    }
  }

  function formatPercent(value) {
    return `${finiteNumber(value, '百分比', 0).toFixed(2)}%`;
  }

  function findInputLine(model, collection, key) {
    const lines = model && Array.isArray(model[collection]) ? model[collection] : [];
    return lines.find(line => line && line.key === key) || null;
  }

  function resolveBuildArguments(modelOrConfig, resultArg, optionsArg) {
    if (modelOrConfig && typeof modelOrConfig === 'object' &&
        (Object.prototype.hasOwnProperty.call(modelOrConfig, 'model') ||
         Object.prototype.hasOwnProperty.call(modelOrConfig, 'result'))) {
      return {
        model: modelOrConfig.model || {},
        result: modelOrConfig.result || resultArg,
        options: { ...modelOrConfig, ...(optionsArg || {}) }
      };
    }
    if (modelOrConfig && modelOrConfig.complete === true && !resultArg) {
      return { model: {}, result: modelOrConfig, options: optionsArg || {} };
    }
    return { model: modelOrConfig || {}, result: resultArg, options: optionsArg || {} };
  }

  function snapshotLine(line, inputLine, fallbackGroup) {
    const source = line || {};
    const input = inputLine || {};
    return {
      key: normalizeText(source.key || input.key, 128),
      group: normalizeText(source.group || fallbackGroup || '成本', 128),
      label: normalizeText(source.label || input.label || '未命名项目', 256),
      inputMode: input.mode === 'percent' || source.mode === 'percent' ? 'percent' : 'amount',
      inputValue: finiteNumber(input.value !== undefined ? input.value : source.value, '输入值', 0),
      amount: finiteNumber(source.amount, `${source.label || '成本'}金额`),
      percent: finiteNumber(source.percent, `${source.label || '成本'}占比`, 0)
    };
  }

  function buildSnapshot(modelOrConfig, resultArg, optionsArg) {
    const resolved = resolveBuildArguments(modelOrConfig, resultArg, optionsArg);
    const model = resolved.model;
    const result = resolved.result;
    const options = resolved.options;
    if (!result || result.complete !== true) {
      throw new TypeError('只能导出已经完成的利润计算结果');
    }

    const settlementCurrency = safeCurrency(result.settlementCurrency || model.settlementCurrency);
    const rawPlatform = options.platformLabel ||
      (model.platform === '其他' ? model.customPlatform : '') ||
      result.platform || model.platform || '未指定平台';
    const platform = normalizeText(rawPlatform, 256).trim() || '未指定平台';
    const productModel = normalizeText(result.productModel || model.productModel || '未填写型号', 512).trim() || '未填写型号';
    const createdDate = options.createdAt ? new Date(options.createdAt) : new Date();
    const createdAt = Number.isNaN(createdDate.getTime()) ? new Date().toISOString() : createdDate.toISOString();
    const revenue = finiteNumber(result.revenue, '销售收入');

    const directCosts = (Array.isArray(result.directCosts) ? result.directCosts : []).map(line =>
      snapshotLine(line, findInputLine(model, 'directCosts', line && line.key), '直接成本')
    );
    const allocations = (Array.isArray(result.allocations) ? result.allocations : []).map(line =>
      snapshotLine(line, findInputLine(model, 'allocations', line && line.key), '分摊成本')
    );
    const sourceProductCost = result.productCost || {};
    const productCost = {
      label: normalizeText(sourceProductCost.label || '产品成本', 256),
      sourceAmount: finiteNumber(
        sourceProductCost.sourceAmount !== undefined
          ? sourceProductCost.sourceAmount
          : model && model.productCost && model.productCost.amount,
        '产品成本原币金额',
        0
      ),
      sourceCurrency: safeCurrency(
        sourceProductCost.sourceCurrency || (model && model.productCost && model.productCost.currency)
      ),
      amount: finiteNumber(sourceProductCost.amount, '产品成本'),
      percent: finiteNumber(sourceProductCost.percent, '产品成本占比', revenue ? sourceProductCost.amount / revenue * 100 : 0)
    };

    const snapshot = {
      schemaVersion: 1,
      title: normalizeText(options.title || '跨境电商利润模型', 256),
      createdAt,
      platform,
      productModel,
      settlementCurrency,
      price: {
        sourceAmount: finiteNumber(model && model.price && model.price.amount, '原始定价', revenue),
        sourceCurrency: safeCurrency((model && model.price && model.price.currency) || settlementCurrency),
        amount: revenue
      },
      usdCnyRate: finiteNumber(result.usdCnyRate !== undefined ? result.usdCnyRate : model.usdCnyRate, '美元兑人民币汇率'),
      settlementCnyRate: finiteNumber(
        result.settlementCnyRate !== undefined ? result.settlementCnyRate : model.settlementCnyRate,
        '平台币兑人民币汇率'
      ),
      directCosts,
      productCost,
      directTotal: finiteNumber(result.directTotal, '直接成本合计'),
      grossProfit: finiteNumber(result.grossProfit, '毛利'),
      grossMargin: finiteNumber(result.grossMargin, '毛利率'),
      allocations,
      allocationTotal: finiteNumber(result.allocationTotal, '分摊成本合计'),
      netProfit: finiteNumber(result.netProfit, '净利润'),
      netMargin: finiteNumber(result.netMargin, '净利率'),
      totalCost: finiteNumber(result.totalCost, '总成本'),
      totalCostRate: finiteNumber(result.totalCostRate, '总成本率', revenue ? result.totalCost / revenue * 100 : 0),
      warnings: (Array.isArray(result.warnings) ? result.warnings : [])
        .map(item => normalizeText(item, 1024))
        .filter(Boolean)
    };

    const directRows = directCosts.map(line => ({
      section: line.group,
      label: line.label,
      amount: line.amount,
      percent: line.percent,
      kind: 'direct'
    }));
    const allocationRows = allocations.map(line => ({
      section: line.group,
      label: line.label,
      amount: line.amount,
      percent: line.percent,
      kind: 'allocation'
    }));
    snapshot.rows = [
      { section: '价格策略', label: '销售定价', amount: revenue, percent: 100, kind: 'revenue' },
      ...directRows,
      {
        section: '平台及产品成本',
        label: `产品成本（${productCost.sourceCurrency}）`,
        amount: productCost.amount,
        percent: productCost.percent,
        kind: 'direct'
      },
      { section: '直接成本', label: '直接成本合计', amount: snapshot.directTotal, percent: revenue ? snapshot.directTotal / revenue * 100 : 0, kind: 'subtotal' },
      { section: '利润', label: '毛利（分摊前）', amount: snapshot.grossProfit, percent: snapshot.grossMargin, kind: 'gross' },
      ...allocationRows,
      { section: '分摊成本', label: '分摊成本合计', amount: snapshot.allocationTotal, percent: revenue ? snapshot.allocationTotal / revenue * 100 : 0, kind: 'subtotal' },
      { section: '利润', label: '净利润', amount: snapshot.netProfit, percent: snapshot.netMargin, kind: 'net' }
    ];
    return snapshot;
  }

  function buildPlainSummary(snapshot) {
    return [
      snapshot.title,
      `平台：${snapshot.platform}`,
      `产品型号：${snapshot.productModel}`,
      `销售收入：${formatMoney(snapshot.price.amount, snapshot.settlementCurrency)}`,
      `毛利：${formatMoney(snapshot.grossProfit, snapshot.settlementCurrency)}（${formatPercent(snapshot.grossMargin)}）`,
      `净利润：${formatMoney(snapshot.netProfit, snapshot.settlementCurrency)}（${formatPercent(snapshot.netMargin)}）`
    ].join('\n');
  }

  function roundedRect(ctx, x, y, width, height, radius) {
    const r = Math.max(0, Math.min(radius, width / 2, height / 2));
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.lineTo(x + width - r, y);
    ctx.quadraticCurveTo(x + width, y, x + width, y + r);
    ctx.lineTo(x + width, y + height - r);
    ctx.quadraticCurveTo(x + width, y + height, x + width - r, y + height);
    ctx.lineTo(x + r, y + height);
    ctx.quadraticCurveTo(x, y + height, x, y + height - r);
    ctx.lineTo(x, y + r);
    ctx.quadraticCurveTo(x, y, x + r, y);
    ctx.closePath();
  }

  function truncateCanvasText(ctx, value, maxWidth) {
    const text = normalizeText(value).replace(/[\r\n\t]+/g, ' ').trim();
    if (ctx.measureText(text).width <= maxWidth) return text;
    let low = 0;
    let high = text.length;
    while (low < high) {
      const middle = Math.ceil((low + high) / 2);
      if (ctx.measureText(`${text.slice(0, middle)}…`).width <= maxWidth) low = middle;
      else high = middle - 1;
    }
    return `${text.slice(0, low)}…`;
  }

  function wrapCanvasText(ctx, value, maxWidth, maxLines) {
    const text = normalizeText(value).replace(/\r/g, '').trim();
    if (!text) return [''];
    const explicitLines = text.split('\n');
    const lines = [];
    for (const explicitLine of explicitLines) {
      let current = '';
      for (const character of explicitLine) {
        const candidate = current + character;
        if (current && ctx.measureText(candidate).width > maxWidth) {
          lines.push(current);
          current = character;
        } else {
          current = candidate;
        }
        if (lines.length >= maxLines) break;
      }
      if (lines.length >= maxLines) break;
      if (current || !explicitLine) lines.push(current);
      if (lines.length >= maxLines) break;
    }
    if (lines.length === maxLines) {
      const consumed = lines.join('').length;
      if (consumed < text.replace(/\n/g, '').length) lines[maxLines - 1] = truncateCanvasText(ctx, `${lines[maxLines - 1]}…`, maxWidth);
    }
    return lines.slice(0, maxLines);
  }

  function renderSummaryCanvas(snapshot, options) {
    const settings = options || {};
    const doc = settings.document || (typeof document !== 'undefined' ? document : null);
    const canvas = settings.canvas || (doc && doc.createElement ? doc.createElement('canvas') : null);
    if (!canvas || typeof canvas.getContext !== 'function') {
      throw new Error('当前环境不支持 Canvas；请在浏览器中调用图片导出 API');
    }
    const ctx = canvas.getContext('2d');
    if (!ctx) throw new Error('无法创建 Canvas 2D 绘图上下文');

    const palette = { ...DEFAULT_PALETTE, ...(settings.palette || {}) };
    const width = Math.max(840, finiteNumber(settings.width, '图片宽度', 1200));
    const rowHeight = 58;
    const margin = 48;
    const headerHeight = 230;
    const metricsHeight = 150;
    const tableHeaderHeight = 60;
    const warningHeight = snapshot.warnings.length ? 84 : 50;
    const height = margin * 2 + headerHeight + 28 + metricsHeight + 28 + tableHeaderHeight + snapshot.rows.length * rowHeight + warningHeight;
    const fallbackScale = typeof devicePixelRatio === 'number' ? devicePixelRatio : 2;
    const scale = Math.max(1, Math.min(3, finiteNumber(settings.scale, '图片缩放倍率', Math.max(2, fallbackScale))));
    canvas.width = Math.round(width * scale);
    canvas.height = Math.round(height * scale);
    if (canvas.style) {
      canvas.style.width = `${width}px`;
      canvas.style.height = `${height}px`;
    }
    if (typeof ctx.setTransform === 'function') ctx.setTransform(scale, 0, 0, scale, 0, 0);
    else ctx.scale(scale, scale);
    ctx.textBaseline = 'middle';

    ctx.fillStyle = palette.page;
    ctx.fillRect(0, 0, width, height);

    const cardX = margin;
    const cardWidth = width - margin * 2;
    roundedRect(ctx, cardX, margin, cardWidth, height - margin * 2, 24);
    ctx.fillStyle = palette.card;
    ctx.fill();
    ctx.strokeStyle = palette.line;
    ctx.lineWidth = 1;
    ctx.stroke();

    const innerX = cardX + 44;
    const innerWidth = cardWidth - 88;
    let y = margin + 44;
    ctx.fillStyle = palette.accent;
    ctx.font = '700 16px "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif';
    ctx.fillText('CROSS-BORDER COMMERCE', innerX, y + 10);
    ctx.fillStyle = palette.ink;
    ctx.font = '700 44px Georgia, "Noto Serif SC", "Songti SC", serif';
    ctx.fillText(truncateCanvasText(ctx, snapshot.title, innerWidth * 0.63), innerX, y + 65);

    ctx.font = '600 19px "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif';
    ctx.fillStyle = palette.ink;
    ctx.fillText(truncateCanvasText(ctx, `${snapshot.platform} · ${snapshot.productModel}`, innerWidth * 0.62), innerX, y + 116);
    ctx.font = '400 15px "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif';
    ctx.fillStyle = palette.muted;
    ctx.fillText(`平台币种 ${snapshot.settlementCurrency} · USD/CNY ${snapshot.usdCnyRate.toFixed(4)}`, innerX, y + 150);

    const profitBoxWidth = 350;
    const profitBoxX = innerX + innerWidth - profitBoxWidth;
    roundedRect(ctx, profitBoxX, y, profitBoxWidth, 178, 18);
    ctx.fillStyle = snapshot.netProfit >= 0 ? palette.successSoft : palette.dangerSoft;
    ctx.fill();
    ctx.fillStyle = snapshot.netProfit >= 0 ? palette.success : palette.danger;
    ctx.font = '700 16px "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif';
    ctx.fillText('净利润', profitBoxX + 28, y + 32);
    ctx.font = '700 37px "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif';
    ctx.fillText(truncateCanvasText(ctx, formatMoney(snapshot.netProfit, snapshot.settlementCurrency), profitBoxWidth - 56), profitBoxX + 28, y + 82);
    ctx.font = '700 21px "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif';
    ctx.fillText(formatPercent(snapshot.netMargin), profitBoxX + 28, y + 132);

    y = margin + headerHeight;
    ctx.strokeStyle = palette.line;
    ctx.beginPath();
    ctx.moveTo(innerX, y);
    ctx.lineTo(innerX + innerWidth, y);
    ctx.stroke();
    y += 28;

    const gap = 18;
    const metricWidth = (innerWidth - gap * 2) / 3;
    const metrics = [
      ['销售收入', snapshot.price.amount, 100],
      ['毛利（分摊前）', snapshot.grossProfit, snapshot.grossMargin],
      ['总成本', snapshot.totalCost, snapshot.totalCostRate]
    ];
    for (let index = 0; index < metrics.length; index += 1) {
      const metricX = innerX + index * (metricWidth + gap);
      roundedRect(ctx, metricX, y, metricWidth, metricsHeight, 16);
      ctx.fillStyle = index === 1 ? palette.accentSoft : palette.page;
      ctx.fill();
      ctx.fillStyle = palette.muted;
      ctx.font = '600 15px "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif';
      ctx.fillText(metrics[index][0], metricX + 24, y + 30);
      ctx.fillStyle = palette.ink;
      ctx.font = '700 27px "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif';
      ctx.fillText(truncateCanvasText(ctx, formatMoney(metrics[index][1], snapshot.settlementCurrency), metricWidth - 48), metricX + 24, y + 78);
      ctx.fillStyle = index === 1 ? palette.accent : palette.muted;
      ctx.font = '600 16px "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif';
      ctx.fillText(formatPercent(metrics[index][2]), metricX + 24, y + 119);
    }

    y += metricsHeight + 28;
    const columnWidths = [innerWidth * 0.21, innerWidth * 0.35, innerWidth * 0.27, innerWidth * 0.17];
    const columnX = [
      innerX,
      innerX + columnWidths[0],
      innerX + columnWidths[0] + columnWidths[1],
      innerX + columnWidths[0] + columnWidths[1] + columnWidths[2]
    ];
    roundedRect(ctx, innerX, y, innerWidth, tableHeaderHeight, 12);
    ctx.fillStyle = palette.ink;
    ctx.fill();
    ctx.fillStyle = palette.white;
    ctx.font = '700 15px "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif';
    ['类别', '项目', '金额', '占售价'].forEach((label, index) => {
      ctx.textAlign = index >= 2 ? 'right' : 'left';
      const x = index >= 2 ? columnX[index] + columnWidths[index] - 18 : columnX[index] + 18;
      ctx.fillText(label, x, y + tableHeaderHeight / 2);
    });
    ctx.textAlign = 'left';
    y += tableHeaderHeight;

    snapshot.rows.forEach((row, rowIndex) => {
      const isNet = row.kind === 'net';
      const isProfit = row.kind === 'gross' || row.kind === 'subtotal';
      if (isNet || isProfit) {
        ctx.fillStyle = isNet
          ? (row.amount >= 0 ? palette.successSoft : palette.dangerSoft)
          : palette.accentSoft;
        ctx.fillRect(innerX, y, innerWidth, rowHeight);
      } else if (rowIndex % 2 === 1) {
        ctx.fillStyle = palette.page;
        ctx.fillRect(innerX, y, innerWidth, rowHeight);
      }
      ctx.strokeStyle = palette.line;
      ctx.beginPath();
      ctx.moveTo(innerX, y + rowHeight);
      ctx.lineTo(innerX + innerWidth, y + rowHeight);
      ctx.stroke();

      const emphasisColor = isNet
        ? (row.amount >= 0 ? palette.success : palette.danger)
        : isProfit ? palette.accent : palette.ink;
      ctx.font = `${isNet || isProfit ? '700' : '500'} 15px "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif`;
      ctx.fillStyle = isNet || isProfit ? emphasisColor : palette.muted;
      ctx.fillText(truncateCanvasText(ctx, row.section, columnWidths[0] - 36), columnX[0] + 18, y + rowHeight / 2);
      ctx.fillStyle = isNet || isProfit ? emphasisColor : palette.ink;
      ctx.fillText(truncateCanvasText(ctx, row.label, columnWidths[1] - 36), columnX[1] + 18, y + rowHeight / 2);
      ctx.textAlign = 'right';
      ctx.fillText(formatMoney(row.amount, snapshot.settlementCurrency), columnX[2] + columnWidths[2] - 18, y + rowHeight / 2);
      ctx.fillText(formatPercent(row.percent), columnX[3] + columnWidths[3] - 18, y + rowHeight / 2);
      ctx.textAlign = 'left';
      y += rowHeight;
    });

    y += 24;
    ctx.fillStyle = palette.muted;
    ctx.font = '400 14px "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif';
    if (snapshot.warnings.length) {
      const warningText = `提示：${snapshot.warnings.join('；')}`;
      const warningLines = wrapCanvasText(ctx, warningText, innerWidth, 2);
      warningLines.forEach((line, index) => ctx.fillText(line, innerX, y + 10 + index * 22));
    } else {
      ctx.fillText('结果由当前输入自动测算；导出内容仅在本地浏览器生成。', innerX, y + 10);
    }
    return canvas;
  }

  function canvasToPngBlob(canvas) {
    if (canvas && typeof canvas.convertToBlob === 'function') {
      return canvas.convertToBlob({ type: PNG_MIME });
    }
    if (!canvas || typeof canvas.toBlob !== 'function') {
      return Promise.reject(new Error('当前浏览器不支持 Canvas PNG 导出'));
    }
    return new Promise((resolve, reject) => {
      canvas.toBlob(blob => {
        if (blob) resolve(blob);
        else reject(new Error('无法生成 PNG 图片'));
      }, PNG_MIME);
    });
  }

  async function renderSummaryPng(snapshot, options) {
    const settings = options || {};
    const doc = settings.document || (typeof document !== 'undefined' ? document : null);
    if (doc && doc.fonts && doc.fonts.ready) {
      try { await doc.fonts.ready; } catch (_) { /* Use system font fallback. */ }
    }
    return canvasToPngBlob(renderSummaryCanvas(snapshot, settings));
  }

  const renderSummaryBlob = renderSummaryPng;

  function sanitizeFilename(value, extension) {
    const fallback = '跨境电商利润模型';
    let name = normalizeText(value || fallback, 120)
      .replace(/[<>:"/\\|?*\u0000-\u001f]/g, '-')
      .replace(/\s+/g, ' ')
      .replace(/[. ]+$/g, '')
      .trim();
    if (!name) name = fallback;
    if (/^(con|prn|aux|nul|com[1-9]|lpt[1-9])$/i.test(name)) name = `_${name}`;
    const ext = normalizeText(extension || '', 10).replace(/[^a-z0-9]/gi, '').toLowerCase();
    return ext && !name.toLowerCase().endsWith(`.${ext}`) ? `${name}.${ext}` : name;
  }

  function defaultFilename(snapshot, extension) {
    const model = snapshot && snapshot.productModel && snapshot.productModel !== '未填写型号'
      ? snapshot.productModel
      : snapshot && snapshot.platform;
    return sanitizeFilename(`${model || '跨境电商'}-利润模型`, extension);
  }

  function downloadBlob(blob, filename, options) {
    const settings = options || {};
    if (typeof settings.download === 'function') {
      settings.download(blob, filename);
      return filename;
    }
    const doc = settings.document || (typeof document !== 'undefined' ? document : null);
    const urlApi = settings.URL || (typeof URL !== 'undefined' ? URL : null);
    if (!doc || !doc.createElement || !urlApi || !urlApi.createObjectURL) {
      throw new Error('当前环境无法触发文件下载');
    }
    const objectUrl = urlApi.createObjectURL(blob);
    const anchor = doc.createElement('a');
    anchor.href = objectUrl;
    anchor.download = filename;
    anchor.rel = 'noopener';
    if (anchor.style) anchor.style.display = 'none';
    const host = doc.body || doc.documentElement;
    if (host && host.appendChild) host.appendChild(anchor);
    anchor.click();
    if (anchor.remove) anchor.remove();
    else if (host && host.removeChild) host.removeChild(anchor);
    const revoke = () => urlApi.revokeObjectURL(objectUrl);
    if (typeof setTimeout === 'function') setTimeout(revoke, finiteNumber(settings.revokeDelay, '回收延迟', 1500));
    else revoke();
    return filename;
  }

  async function copyOrDownloadImage(snapshot, options) {
    const settings = options || {};
    const filename = sanitizeFilename(settings.filename || defaultFilename(snapshot, 'png'), 'png');
    const pngPromise = renderSummaryPng(snapshot, settings.renderOptions || settings);
    const nav = settings.navigator || (typeof navigator !== 'undefined' ? navigator : null);
    const ClipboardCtor = settings.ClipboardItem || (typeof ClipboardItem !== 'undefined' ? ClipboardItem : null);
    const secure = settings.isSecureContext !== undefined
      ? Boolean(settings.isSecureContext)
      : (typeof isSecureContext !== 'undefined' && isSecureContext);
    let pngSupported = true;
    if (ClipboardCtor && typeof ClipboardCtor.supports === 'function') {
      try { pngSupported = ClipboardCtor.supports(PNG_MIME); } catch (_) { pngSupported = true; }
    }
    const canCopy = settings.forceDownload !== true && secure && pngSupported &&
      nav && nav.clipboard && typeof nav.clipboard.write === 'function' && ClipboardCtor;

    if (canCopy) {
      const textPromise = Promise.resolve(new Blob([buildPlainSummary(snapshot)], { type: 'text/plain' }));
      try {
        await nav.clipboard.write([new ClipboardCtor({
          [PNG_MIME]: pngPromise,
          'text/plain': textPromise
        })]);
        return { action: 'copied', blob: await pngPromise, filename };
      } catch (error) {
        const blob = await pngPromise;
        downloadBlob(blob, filename, settings);
        return { action: 'downloaded', blob, filename, clipboardError: error };
      }
    }

    const blob = await pngPromise;
    downloadBlob(blob, filename, settings);
    return { action: 'downloaded', blob, filename };
  }

  function columnName(number) {
    let value = number;
    let name = '';
    while (value > 0) {
      value -= 1;
      name = String.fromCharCode(65 + value % 26) + name;
      value = Math.floor(value / 26);
    }
    return name;
  }

  function cellRef(column, row) {
    return `${columnName(column)}${row}`;
  }

  function textCell(column, row, value, style) {
    const ref = cellRef(column, row);
    const text = normalizeText(value, 32767);
    return `<c r="${ref}" s="${style || 0}" t="inlineStr"><is><t xml:space="preserve">${xmlEscape(text)}</t></is></c>`;
  }

  function numberCell(column, row, value, style) {
    const ref = cellRef(column, row);
    const number = finiteNumber(value, `单元格 ${ref}`);
    return `<c r="${ref}" s="${style || 0}" t="n"><v>${String(Object.is(number, -0) ? 0 : number)}</v></c>`;
  }

  function rowXml(rowNumber, cells, height) {
    const heightAttribute = height ? ` ht="${height}" customHeight="1"` : '';
    return `<row r="${rowNumber}"${heightAttribute}>${cells.join('')}</row>`;
  }

  function buildProfitSheet(snapshot) {
    const rows = [];
    let row = 1;
    rows.push(rowXml(row, [textCell(1, row, snapshot.title, 1)], 30));
    row += 1;
    rows.push(rowXml(row, [textCell(1, row, '平台', 2), textCell(2, row, snapshot.platform, 3), textCell(3, row, '产品型号', 2), textCell(4, row, snapshot.productModel, 3)], 22));
    row += 1;
    rows.push(rowXml(row, [textCell(1, row, '平台币种', 2), textCell(2, row, snapshot.settlementCurrency, 3), textCell(3, row, '导出时间', 2), textCell(4, row, snapshot.createdAt, 3)], 22));
    row += 1;
    rows.push(rowXml(row, [textCell(1, row, '原始定价币种', 2), textCell(2, row, snapshot.price.sourceCurrency, 3), textCell(3, row, '原始定价', 2), numberCell(4, row, snapshot.price.sourceAmount, 4)], 22));
    row += 1;
    rows.push(rowXml(row, [textCell(1, row, 'USD/CNY', 2), numberCell(2, row, snapshot.usdCnyRate, 18), textCell(3, row, `${snapshot.settlementCurrency}/CNY`, 2), numberCell(4, row, snapshot.settlementCnyRate, 18)], 22));
    row += 2;
    const tableHeaderRow = row;
    rows.push(rowXml(row, [
      textCell(1, row, '类别', 5),
      textCell(2, row, '项目', 5),
      textCell(3, row, `金额（${snapshot.settlementCurrency}）`, 5),
      textCell(4, row, '占售价', 5)
    ], 24));
    row += 1;

    for (const item of snapshot.rows) {
      let textStyle = 6;
      let amountStyle = 7;
      let percentStyle = 8;
      if (item.kind === 'gross' || item.kind === 'subtotal') {
        textStyle = 9;
        amountStyle = 10;
        percentStyle = 11;
      } else if (item.kind === 'net') {
        const profitable = item.amount >= 0;
        textStyle = profitable ? 12 : 15;
        amountStyle = profitable ? 13 : 16;
        percentStyle = profitable ? 14 : 17;
      }
      rows.push(rowXml(row, [
        textCell(1, row, item.section, textStyle),
        textCell(2, row, item.label, textStyle),
        numberCell(3, row, item.amount, amountStyle),
        numberCell(4, row, item.percent / 100, percentStyle)
      ], 21));
      row += 1;
    }

    const warningStart = row + 1;
    if (snapshot.warnings.length) {
      rows.push(rowXml(warningStart, [textCell(1, warningStart, '风险提示', 2), textCell(2, warningStart, snapshot.warnings.join('；'), 19)], 32));
    } else {
      rows.push(rowXml(warningStart, [textCell(1, warningStart, '说明', 2), textCell(2, warningStart, '当前计算未产生额外风险提示。', 19)], 28));
    }
    const lastRow = warningStart;

    return `${XML_HEADER}
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:D${lastRow}"/>
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="${tableHeaderRow}" topLeftCell="A${tableHeaderRow + 1}" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="20"/>
  <cols>
    <col min="1" max="1" width="22" customWidth="1"/>
    <col min="2" max="2" width="34" customWidth="1"/>
    <col min="3" max="3" width="21" customWidth="1"/>
    <col min="4" max="4" width="24" customWidth="1"/>
  </cols>
  <sheetData>${rows.join('')}</sheetData>
  <autoFilter ref="A${tableHeaderRow}:D${tableHeaderRow + snapshot.rows.length}"/>
  <mergeCells count="1"><mergeCell ref="A1:D1"/></mergeCells>
  <pageMargins left="0.45" right="0.45" top="0.55" bottom="0.55" header="0.2" footer="0.2"/>
</worksheet>`;
  }

  function buildNotesSheet(snapshot) {
    const notes = [
      ['本文件', '由跨境电商利润模型计算器在本地浏览器生成，不会自动上传数据。'],
      ['平台币种', `所有成本、毛利和净利润最终统一折算为 ${snapshot.settlementCurrency}。`],
      ['定价换算', `原始定价为 ${formatMoney(snapshot.price.sourceAmount, snapshot.price.sourceCurrency)}；USD/CNY 汇率为 ${snapshot.usdCnyRate.toFixed(4)}。`],
      ['金额模式', '费用输入为具体金额时，直接计入对应成本。'],
      ['比例模式', '费用输入为百分比时，费用金额 = 平台币售价 × 输入比例。'],
      ['产品成本', `原币成本为 ${formatMoney(snapshot.productCost.sourceAmount, snapshot.productCost.sourceCurrency)}，按汇率折算到平台币。`],
      ['毛利', '毛利 = 销售收入 − 直接成本合计；毛利率 = 毛利 ÷ 销售收入。'],
      ['净利润', '净利润 = 毛利 − 分摊成本合计；净利率 = 净利润 ÷ 销售收入。'],
      ['舍入规则', '计算使用完整精度，表格仅通过单元格格式显示约定的小数位。'],
      ['安全说明', '产品型号与备注均按纯文本写入，不会作为 Excel 公式执行。']
    ];
    const rows = [];
    rows.push(rowXml(1, [textCell(1, 1, '计算说明', 1)], 30));
    rows.push(rowXml(3, [textCell(1, 3, '项目', 5), textCell(2, 3, '说明', 5)], 24));
    notes.forEach((note, index) => {
      const row = index + 4;
      rows.push(rowXml(row, [textCell(1, row, note[0], 6), textCell(2, row, note[1], 19)], 34));
    });
    const lastRow = notes.length + 3;
    return `${XML_HEADER}
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B${lastRow}"/>
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="20"/>
  <cols><col min="1" max="1" width="20" customWidth="1"/><col min="2" max="2" width="86" customWidth="1"/></cols>
  <sheetData>${rows.join('')}</sheetData>
  <mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>
  <pageMargins left="0.45" right="0.45" top="0.55" bottom="0.55" header="0.2" footer="0.2"/>
</worksheet>`;
  }

  function buildStylesXml(snapshot) {
    const currency = snapshot.settlementCurrency;
    const meta = CURRENCY_META[currency];
    const decimals = meta.decimals > 0 ? `.${'0'.repeat(meta.decimals)}` : '';
    const currencyFormat = `&quot;${xmlEscape(meta.symbol)}&quot;#,##0${decimals};-&quot;${xmlEscape(meta.symbol)}&quot;#,##0${decimals}`;
    return `${XML_HEADER}
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="3">
    <numFmt numFmtId="164" formatCode="${currencyFormat}"/>
    <numFmt numFmtId="165" formatCode="0.00%"/>
    <numFmt numFmtId="166" formatCode="0.0000"/>
  </numFmts>
  <fonts count="5">
    <font><sz val="11"/><color rgb="FF171715"/><name val="Aptos"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FF171715"/><name val="Aptos"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Aptos"/><family val="2"/></font>
    <font><b/><sz val="18"/><color rgb="FF171715"/><name val="Aptos Display"/><family val="2"/></font>
    <font><sz val="11"/><color rgb="FF5F5B53"/><name val="Aptos"/><family val="2"/></font>
  </fonts>
  <fills count="6">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFA84B32"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF7F3EA"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF2D6A4F"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFB42318"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FFDDD6CA"/></left><right style="thin"><color rgb="FFDDD6CA"/></right><top style="thin"><color rgb="FFDDD6CA"/></top><bottom style="thin"><color rgb="FFDDD6CA"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="20">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"><alignment vertical="center"/></xf>
    <xf numFmtId="4" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"><alignment vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment vertical="center"/></xf>
    <xf numFmtId="164" fontId="1" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="165" fontId="1" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment vertical="center"/></xf>
    <xf numFmtId="164" fontId="2" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="165" fontId="2" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment vertical="center"/></xf>
    <xf numFmtId="164" fontId="2" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="165" fontId="2" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0" fontId="4" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"><alignment vertical="center" wrapText="1"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>`;
  }

  function workbookXml() {
    return `${XML_HEADER}
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews><workbookView xWindow="0" yWindow="0" windowWidth="19200" windowHeight="10800"/></bookViews>
  <sheets>
    <sheet name="利润模型" sheetId="1" r:id="rId1"/>
    <sheet name="计算说明" sheetId="2" r:id="rId2"/>
  </sheets>
  <calcPr calcId="191029" fullCalcOnLoad="1"/>
</workbook>`;
  }

  function workbookRelationshipsXml() {
    return `${XML_HEADER}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
  }

  function rootRelationshipsXml() {
    return `${XML_HEADER}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
  }

  function contentTypesXml() {
    return `${XML_HEADER}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;
  }

  let crcTable = null;
  function getCrcTable() {
    if (crcTable) return crcTable;
    crcTable = new Uint32Array(256);
    for (let index = 0; index < 256; index += 1) {
      let value = index;
      for (let bit = 0; bit < 8; bit += 1) value = (value & 1) ? (0xedb88320 ^ (value >>> 1)) : (value >>> 1);
      crcTable[index] = value >>> 0;
    }
    return crcTable;
  }

  function crc32(bytes) {
    const table = getCrcTable();
    let crc = 0xffffffff;
    for (let index = 0; index < bytes.length; index += 1) crc = table[(crc ^ bytes[index]) & 0xff] ^ (crc >>> 8);
    return (crc ^ 0xffffffff) >>> 0;
  }

  function encodeUtf8(value) {
    if (typeof TextEncoder === 'undefined') throw new Error('当前环境不支持 UTF-8 编码');
    return new TextEncoder().encode(value);
  }

  function writeUint16(bytes, offset, value) {
    bytes[offset] = value & 0xff;
    bytes[offset + 1] = (value >>> 8) & 0xff;
  }

  function writeUint32(bytes, offset, value) {
    bytes[offset] = value & 0xff;
    bytes[offset + 1] = (value >>> 8) & 0xff;
    bytes[offset + 2] = (value >>> 16) & 0xff;
    bytes[offset + 3] = (value >>> 24) & 0xff;
  }

  function dosDateTime(input) {
    const date = input instanceof Date && !Number.isNaN(input.getTime()) ? input : new Date();
    const year = Math.max(1980, Math.min(2107, date.getUTCFullYear()));
    const month = date.getUTCMonth() + 1;
    const day = date.getUTCDate();
    const hours = date.getUTCHours();
    const minutes = date.getUTCMinutes();
    const seconds = Math.floor(date.getUTCSeconds() / 2);
    return {
      time: ((hours << 11) | (minutes << 5) | seconds) & 0xffff,
      date: (((year - 1980) << 9) | (month << 5) | day) & 0xffff
    };
  }

  function concatBytes(chunks, totalLength) {
    const output = new Uint8Array(totalLength === undefined
      ? chunks.reduce((sum, chunk) => sum + chunk.length, 0)
      : totalLength);
    let offset = 0;
    for (const chunk of chunks) {
      output.set(chunk, offset);
      offset += chunk.length;
    }
    return output;
  }

  function buildStoredZip(entries, timestamp) {
    if (!Array.isArray(entries) || entries.length === 0 || entries.length > 0xffff) {
      throw new Error('ZIP 条目数量无效');
    }
    const dos = dosDateTime(timestamp);
    const localChunks = [];
    const centralChunks = [];
    let localOffset = 0;

    for (const entry of entries) {
      const nameBytes = encodeUtf8(entry.name);
      const dataBytes = entry.data instanceof Uint8Array ? entry.data : encodeUtf8(String(entry.data));
      const checksum = crc32(dataBytes);
      const localHeader = new Uint8Array(30 + nameBytes.length);
      writeUint32(localHeader, 0, 0x04034b50);
      writeUint16(localHeader, 4, 20);
      writeUint16(localHeader, 6, 0x0800);
      writeUint16(localHeader, 8, 0);
      writeUint16(localHeader, 10, dos.time);
      writeUint16(localHeader, 12, dos.date);
      writeUint32(localHeader, 14, checksum);
      writeUint32(localHeader, 18, dataBytes.length);
      writeUint32(localHeader, 22, dataBytes.length);
      writeUint16(localHeader, 26, nameBytes.length);
      writeUint16(localHeader, 28, 0);
      localHeader.set(nameBytes, 30);
      localChunks.push(localHeader, dataBytes);

      const centralHeader = new Uint8Array(46 + nameBytes.length);
      writeUint32(centralHeader, 0, 0x02014b50);
      writeUint16(centralHeader, 4, 20);
      writeUint16(centralHeader, 6, 20);
      writeUint16(centralHeader, 8, 0x0800);
      writeUint16(centralHeader, 10, 0);
      writeUint16(centralHeader, 12, dos.time);
      writeUint16(centralHeader, 14, dos.date);
      writeUint32(centralHeader, 16, checksum);
      writeUint32(centralHeader, 20, dataBytes.length);
      writeUint32(centralHeader, 24, dataBytes.length);
      writeUint16(centralHeader, 28, nameBytes.length);
      writeUint16(centralHeader, 30, 0);
      writeUint16(centralHeader, 32, 0);
      writeUint16(centralHeader, 34, 0);
      writeUint16(centralHeader, 36, 0);
      writeUint32(centralHeader, 38, 0);
      writeUint32(centralHeader, 42, localOffset);
      centralHeader.set(nameBytes, 46);
      centralChunks.push(centralHeader);
      localOffset += localHeader.length + dataBytes.length;
    }

    const centralLength = centralChunks.reduce((sum, chunk) => sum + chunk.length, 0);
    const end = new Uint8Array(22);
    writeUint32(end, 0, 0x06054b50);
    writeUint16(end, 4, 0);
    writeUint16(end, 6, 0);
    writeUint16(end, 8, entries.length);
    writeUint16(end, 10, entries.length);
    writeUint32(end, 12, centralLength);
    writeUint32(end, 16, localOffset);
    writeUint16(end, 20, 0);
    return concatBytes([...localChunks, ...centralChunks, end], localOffset + centralLength + end.length);
  }

  function buildXlsxBytes(snapshot) {
    const entries = [
      { name: '[Content_Types].xml', data: contentTypesXml() },
      { name: '_rels/.rels', data: rootRelationshipsXml() },
      { name: 'xl/workbook.xml', data: workbookXml() },
      { name: 'xl/_rels/workbook.xml.rels', data: workbookRelationshipsXml() },
      { name: 'xl/styles.xml', data: buildStylesXml(snapshot) },
      { name: 'xl/worksheets/sheet1.xml', data: buildProfitSheet(snapshot) },
      { name: 'xl/worksheets/sheet2.xml', data: buildNotesSheet(snapshot) }
    ];
    return buildStoredZip(entries, new Date(snapshot.createdAt));
  }

  function buildXlsxBlob(snapshot, options) {
    const settings = options || {};
    const BlobCtor = settings.Blob || (typeof Blob !== 'undefined' ? Blob : null);
    if (!BlobCtor) throw new Error('当前环境不支持 Blob 文件生成');
    return new BlobCtor([buildXlsxBytes(snapshot)], { type: XLSX_MIME });
  }

  function downloadXlsx(snapshot, options) {
    const settings = options || {};
    const blob = buildXlsxBlob(snapshot, settings);
    const filename = sanitizeFilename(settings.filename || defaultFilename(snapshot, 'xlsx'), 'xlsx');
    downloadBlob(blob, filename, settings);
    return { blob, filename };
  }

  return {
    XLSX_MIME,
    PNG_MIME,
    buildSnapshot,
    buildPlainSummary,
    renderSummaryCanvas,
    renderSummaryPng,
    renderSummaryBlob,
    copyOrDownloadImage,
    buildXlsxBytes,
    buildXlsxBlob,
    downloadXlsx,
    downloadBlob,
    sanitizeFilename,
    xmlEscape,
    crc32
  };
});
