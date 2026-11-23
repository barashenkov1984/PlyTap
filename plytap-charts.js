// =========================================================================
// PLYTAP-CHARTS.JS — Генерация PNG-графиков (порт veneer_report.py) и PDF
// =========================================================================

const PLY_CHART_COLORS = ['#2E75B6', '#ED7D31', '#70AD47', '#FFC000', '#7030A0', '#C00000'];
const PLY_GROUP1_NAME = 'Дефекты обработки';
const PLY_GROUP2_NAME = 'Дефекты древесины';
const PLY_TOP_N = 10;

// ------------------- 1. ПОДГОТОВКА И РАСЧЁТЫ ДАННЫХ -------------------

function plyChartDefectTitle(key, customTitlesMap) {
    if (customTitlesMap && customTitlesMap[key] && String(customTitlesMap[key]).trim()) {
        return String(customTitlesMap[key]).trim();
    }
    if (typeof getDefectTitle === 'function') {
        return getDefectTitle(key, customTitlesMap || {});
    }
    const col = (typeof EXCEL_COLUMNS !== 'undefined' ? EXCEL_COLUMNS : []).find(c => c.key === key);
    return col ? col.title : key;
}

function plyFormatDateRuShort(iso) {
    if (!iso || !String(iso).includes('-')) return String(iso || '');
    return String(iso).split('-').reverse().join('.');
}

function plyNormalizeCycles(cycles) {
    if (!Array.isArray(cycles)) return [];
    return cycles
        .filter(c => c && Number(c.checked) > 0)
        .map(c => {
            const veneer = c.veneerThickness === '0,9' ? '0,9' : '1,41';
            const tu = c.tuStandard === 'ТУ-Р' ? 'ТУ-Р' : 'ТУ-Э';
            const counts = {};
            const customTitles = c.customTitles || {};
            let b1 = 0;
            let b2 = 0;
            let maxVal = 0;
            let maxName = 'нет брака';

            ALL_DEFECT_KEYS.forEach(k => {
                const v = Number(c.counts ? c.counts[k] : 0) || 0;
                counts[k] = v;
                if (DEFECT_KEYS_CN.includes(k)) b1 += v;
                else b2 += v;
                if (v > maxVal) {
                    maxVal = v;
                    maxName = plyChartDefectTitle(k, customTitles);
                }
            });

            const checked = Number(c.checked) || 0;
            const bt = b1 + b2;
            const pct = checked > 0 ? (bt / checked) * 100 : 0;
            const dateIso = c.date || (typeof appState !== 'undefined' ? appState.date : '');
            const shiftStr = String(c.shift || (typeof appState !== 'undefined' ? appState.shift : '1'));

            return {
                veneer,
                tu,
                group: `Рубашка ${veneer} ${tu}`,
                position: String(c.sort || '—').trim(),
                checked,
                counts,
                customTitles,
                b1,
                b2,
                bt,
                pct,
                goodPct: 100 - pct,
                maxDefVal: maxVal,
                maxDefName: maxName,
                date: dateIso,
                shift: shiftStr,
                period: `${plyFormatDateRuShort(dateIso)} см.${shiftStr}`
            };
        });
}

// Объединение одноимённых позиций внутри одного (Рубашка + ТУ) — как в итоговом Excel
function plyAggregatePositions(records) {
    const map = new Map();
    records.forEach(r => {
        const key = `${r.veneer}|${r.tu}|${r.position.toUpperCase().replace(/\s+/g, ' ')}`;
        if (!map.has(key)) {
            map.set(key, {
                veneer: r.veneer,
                tu: r.tu,
                group: r.group,
                position: r.position,
                checked: r.checked,
                counts: { ...r.counts },
                customTitles: { ...r.customTitles }
            });
        } else {
            const ex = map.get(key);
            ex.checked += r.checked;
            ALL_DEFECT_KEYS.forEach(k => {
                ex.counts[k] = (ex.counts[k] || 0) + (r.counts[k] || 0);
            });
            Object.assign(ex.customTitles, r.customTitles);
        }
    });

    const out = [];
    map.forEach(item => {
        let b1 = 0, b2 = 0, maxVal = 0, maxName = 'нет брака';
        ALL_DEFECT_KEYS.forEach(k => {
            const v = item.counts[k] || 0;
            if (DEFECT_KEYS_CN.includes(k)) b1 += v;
            else b2 += v;
            if (v > maxVal) {
                maxVal = v;
                maxName = plyChartDefectTitle(k, item.customTitles);
            }
        });
        const bt = b1 + b2;
        const pct = item.checked > 0 ? (bt / item.checked) * 100 : 0;
        out.push({
            ...item,
            b1, b2, bt, pct,
            goodPct: 100 - pct,
            maxDefVal: maxVal,
            maxDefName: maxName
        });
    });
    return out;
}

function plyComputeDefectTable(records) {
    const mergedTitles = {};
    const totals = {};
    ALL_DEFECT_KEYS.forEach(k => totals[k] = 0);

    let totalBt = 0;
    records.forEach(r => {
        Object.assign(mergedTitles, r.customTitles || {});
        ALL_DEFECT_KEYS.forEach(k => {
            const v = r.counts[k] || 0;
            totals[k] += v;
            totalBt += v;
        });
    });

    const rows = ALL_DEFECT_KEYS
        .map(k => ({
            key: k,
            defect: plyChartDefectTitle(k, mergedTitles),
            group: DEFECT_KEYS_CN.includes(k) ? PLY_GROUP1_NAME : PLY_GROUP2_NAME,
            count: totals[k]
        }))
        .filter(r => r.count > 0)
        .sort((a, b) => b.count - a.count);

    let cum = 0;
    rows.forEach(r => {
        r.share = totalBt > 0 ? (r.count / totalBt) * 100 : 0;
        cum += r.share;
        r.cumShare = cum;
    });
    return rows;
}

function plyComputeGroupSummary(records) {
    const order = [
        { veneer: '0,9', tu: 'ТУ-Э' },
        { veneer: '0,9', tu: 'ТУ-Р' },
        { veneer: '1,41', tu: 'ТУ-Э' },
        { veneer: '1,41', tu: 'ТУ-Р' }
    ];
    const result = [];
    order.forEach(cfg => {
        const sub = records.filter(r => r.veneer === cfg.veneer && r.tu === cfg.tu);
        if (sub.length === 0) return; // Пропускаем пустые блоки!
        const checked = sub.reduce((a, r) => a + r.checked, 0);
        const b1 = sub.reduce((a, r) => a + r.b1, 0);
        const b2 = sub.reduce((a, r) => a + r.b2, 0);
        const bt = b1 + b2;
        result.push({
            veneer: cfg.veneer,
            tu: cfg.tu,
            labelTop: cfg.veneer,
            labelBottom: `Рубашка ${cfg.veneer} ${cfg.tu}`,
            checked, b1, b2, bt,
            pct: checked > 0 ? (bt / checked) * 100 : 0
        });
    });
    return result;
}

function plyComputeDynamics(records) {
    const map = new Map();
    records.forEach(r => {
        const sortKey = `${r.date || '0000-00-00'}_${String(r.shift || '0').padStart(2, '0')}`;
        if (!map.has(sortKey)) {
            map.set(sortKey, {
                sortKey,
                period: r.period,
                checked: 0, bt: 0,
                chk09: 0, bt09: 0,
                chk141: 0, bt141: 0
            });
        }
        const row = map.get(sortKey);
        row.checked += r.checked;
        row.bt += r.bt;
        if (r.veneer === '0,9') {
            row.chk09 += r.checked;
            row.bt09 += r.bt;
        } else {
            row.chk141 += r.checked;
            row.bt141 += r.bt;
        }
    });

    return Array.from(map.values())
        .sort((a, b) => a.sortKey.localeCompare(b.sortKey))
        .map(r => ({
            period: r.period,
            pctTotal: r.checked > 0 ? (r.bt / r.checked) * 100 : 0,
            pct09: r.chk09 > 0 ? (r.bt09 / r.chk09) * 100 : null,
            pct141: r.chk141 > 0 ? (r.bt141 / r.chk141) * 100 : null
        }));
}

// ------------------- 2. БАЗОВЫЕ УТИЛИТЫ CANVAS -------------------

function plyCreateCanvas(width, height) {
    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(0, 0, width, height);
    return { canvas, ctx };
}

function plyDrawRoundRect(ctx, x, y, w, h, r) {
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.lineTo(x + w - r, y);
    ctx.quadraticCurveTo(x + w, y, x + w, y + r);
    ctx.lineTo(x + w, y + h - r);
    ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
    ctx.lineTo(x + r, y + h);
    ctx.quadraticCurveTo(x, y + h, x, y + h - r);
    ctx.lineTo(x, y + r);
    ctx.quadraticCurveTo(x, y, x + r, y);
    ctx.closePath();
}

function plyNiceMax(val, minVal = 5) {
    if (!val || val <= 0) return minVal;
    const mag = Math.pow(10, Math.floor(Math.log10(val)));
    const norm = val / mag;
    let nice;
    if (norm <= 1.2) nice = 1.2;
    else if (norm <= 1.5) nice = 1.5;
    else if (norm <= 2) nice = 2;
    else if (norm <= 2.5) nice = 2.5;
    else if (norm <= 3) nice = 3;
    else if (norm <= 4) nice = 4;
    else if (norm <= 5) nice = 5;
    else if (norm <= 6) nice = 6;
    else if (norm <= 8) nice = 8;
    else nice = 10;
    return nice * mag;
}

function plyExportChartObj(canvas, id, filename, title) {
    return {
        id,
        filename,
        title,
        width: canvas.width,
        height: canvas.height,
        dataUrl: canvas.toDataURL('image/png')
    };
}

// ------------------- 3. ОТРИСОВКА ГРАФИКОВ VENEER_REPORT -------------------

// 0. Сводная табличка (0_svodka.png)
function drawSummaryTableChart(records, periodStr, filename = '0_svodka.png') {
    if (!records || records.length === 0) return null;
    const W = 1140, H = 640;
    const { canvas, ctx } = plyCreateCanvas(W, H);

    const totChk = records.reduce((a, r) => a + r.checked, 0);
    const totBt = records.reduce((a, r) => a + r.bt, 0);
    const totGood = totChk - totBt;
    const overallGoodPct = totChk > 0 ? (totGood / totChk) * 100 : 0;

    const sheetYieldText = (thick) => {
        const sub = records.filter(r => r.veneer === thick);
        if (sub.length === 0) return 'нет данных';
        const chk = sub.reduce((a, r) => a + r.checked, 0);
        const bt = sub.reduce((a, r) => a + r.bt, 0);
        const good = chk - bt;
        const pct = chk > 0 ? (good / chk) * 100 : 0;
        return `${pct.toFixed(2)}%  (${Math.round(good)} из ${Math.round(chk)} л)`;
    };

    // Заголовок
    ctx.fillStyle = '#1F4E78';
    ctx.font = 'bold 26px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText('Сводка по контролю качества шпона', W / 2, 42);

    // Таблица
    const rows = [
        ['Дата / Период', String(periodStr)],
        ['Количество отсортированных листов', `${Math.round(totChk)} л  (годных: ${Math.round(totGood)} л, брак: ${Math.round(totBt)} л)`],
        ['Выход годного на 0,9', sheetYieldText('0,9')],
        ['Выход годного на 1,41', sheetYieldText('1,41')]
    ];

    const tblX = 48, tblY = 82, tblW = W - 96;
    const col1W = Math.round(tblW * 0.43);
    const col2W = tblW - col1W;
    const rowH = 52;

    // Шапка таблицы
    ctx.fillStyle = '#1F4E78';
    ctx.fillRect(tblX, tblY, tblW, rowH);
    ctx.strokeStyle = '#BFBFBF';
    ctx.lineWidth = 1.5;
    ctx.strokeRect(tblX, tblY, col1W, rowH);
    ctx.strokeRect(tblX + col1W, tblY, col2W, rowH);

    ctx.fillStyle = '#FFFFFF';
    ctx.font = 'bold 19px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText('Показатель', tblX + col1W / 2, tblY + rowH / 2);
    ctx.fillText('Значение', tblX + col1W + col2W / 2, tblY + rowH / 2);

    // Строки таблицы
    rows.forEach((r, idx) => {
        const y = tblY + rowH * (idx + 1);
        ctx.fillStyle = idx % 2 === 0 ? '#F2F6FA' : '#FFFFFF';
        ctx.fillRect(tblX, y, tblW, rowH);
        ctx.strokeStyle = '#BFBFBF';
        ctx.strokeRect(tblX, y, col1W, rowH);
        ctx.strokeRect(tblX + col1W, y, col2W, rowH);

        ctx.textAlign = 'left';
        ctx.fillStyle = '#1F4E78';
        ctx.font = 'bold 18px "Segoe UI", Arial, sans-serif';
        ctx.fillText(r[0], tblX + 16, y + rowH / 2);

        ctx.fillStyle = '#111827';
        ctx.font = '18px "Segoe UI", Arial, sans-serif';
        ctx.fillText(r[1], tblX + col1W + 16, y + rowH / 2);
    });

    // Нижний зелёный KPI-блок
    const boxX = 110, boxY = 382, boxW = W - 220, boxH = 216;
    plyDrawRoundRect(ctx, boxX, boxY, boxW, boxH, 18);
    ctx.fillStyle = '#EBF5EE';
    ctx.fill();
    ctx.strokeStyle = '#70AD47';
    ctx.lineWidth = 3;
    ctx.stroke();

    ctx.textAlign = 'center';
    ctx.fillStyle = '#1F4E78';
    ctx.font = 'bold 21px "Segoe UI", Arial, sans-serif';
    ctx.fillText('ОБЩИЙ ПРОЦЕНТ ГОДНОГО ВСЕГО', W / 2, boxY + 42);

    ctx.fillStyle = '#2E7D32';
    ctx.font = 'bold 68px "Segoe UI", Arial, sans-serif';
    ctx.fillText(`${overallGoodPct.toFixed(2)}%`, W / 2, boxY + 122);

    ctx.fillStyle = '#404040';
    ctx.font = '18px "Segoe UI", Arial, sans-serif';
    ctx.fillText(`Всего годного: ${Math.round(totGood)} л из ${Math.round(totChk)} л`, W / 2, boxY + 184);

    return plyExportChartObj(canvas, 'svodka', filename, `Сводка (${periodStr})`);
}

// 1. Парето дефектов РАЗДЕЛЬНО по (Толщина рубашки + ТУ) — пустые блоки пропускаются!
function drawParetoChart(records, veneer, tu, filename) {
    const sub = records.filter(r => r.veneer === veneer && r.tu === tu);
    if (sub.length === 0) return null; // Нет данных в блоке — не строим

    const dt = plyComputeDefectTable(sub).slice(0, PLY_TOP_N);
    if (dt.length === 0) return null; // 0 дефектов в блоке — не строим

    const W = 1320, H = 680;
    const { canvas, ctx } = plyCreateCanvas(W, H);

    const padL = 90, padR = 95, padT = 76, padB = 165;
    const plotW = W - padL - padR;
    const plotH = H - padT - padB;

    const maxCnt = Math.max(...dt.map(d => d.count));
    const yMax = plyNiceMax(maxCnt * 1.15, 5);

    // Заголовок
    ctx.fillStyle = '#111827';
    ctx.font = 'bold 22px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(`Парето дефектов — Рубашка ${veneer} ${tu} (топ-${dt.length})`, W / 2, 34);

    // Сетка и левая ось Y
    const yTicks = 5;
    ctx.font = '15px "Segoe UI", Arial, sans-serif';
    for (let i = 0; i <= yTicks; i++) {
        const val = (yMax / yTicks) * i;
        const y = padT + plotH - (val / yMax) * plotH;
        ctx.strokeStyle = i === 0 ? '#374151' : 'rgba(156,163,175,0.45)';
        ctx.lineWidth = i === 0 ? 1.5 : 1;
        ctx.beginPath();
        ctx.moveTo(padL, y);
        ctx.lineTo(padL + plotW, y);
        ctx.stroke();

        ctx.fillStyle = '#374151';
        ctx.textAlign = 'right';
        ctx.fillText(Math.round(val).toString(), padL - 12, y);
    }

    // Правая ось Y (0..105%)
    const pctTicks = [0, 20, 40, 60, 80, 100];
    pctTicks.forEach(p => {
        const y = padT + plotH - (p / 105) * plotH;
        ctx.fillStyle = '#C00000';
        ctx.textAlign = 'left';
        ctx.fillText(`${p}%`, padL + plotW + 12, y);
    });

    // Подписи осей Y
    ctx.save();
    ctx.translate(28, padT + plotH / 2);
    ctx.rotate(-Math.PI / 2);
    ctx.fillStyle = '#111827';
    ctx.font = 'bold 16px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText('Количество, л', 0, 0);
    ctx.restore();

    ctx.save();
    ctx.translate(W - 24, padT + plotH / 2);
    ctx.rotate(Math.PI / 2);
    ctx.fillStyle = '#C00000';
    ctx.font = 'bold 16px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText('Накопленная доля, %', 0, 0);
    ctx.restore();

    // Столбцы
    const slotW = plotW / dt.length;
    const barW = Math.min(82, slotW * 0.66);

    dt.forEach((d, idx) => {
        const cx = padL + slotW * idx + slotW / 2;
        const barH = (d.count / yMax) * plotH;
        const x = cx - barW / 2;
        const y = padT + plotH - barH;

        ctx.fillStyle = d.group === PLY_GROUP1_NAME ? PLY_CHART_COLORS[0] : PLY_CHART_COLORS[1];
        ctx.fillRect(x, y, barW, barH);

        // Число над столбцом
        ctx.fillStyle = '#111827';
        ctx.font = 'bold 16px "Segoe UI", Arial, sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'bottom';
        ctx.fillText(String(Math.round(d.count)), cx, y - 5);

        // Подпись X под углом 35°
        ctx.save();
        ctx.translate(cx, padT + plotH + 12);
        ctx.rotate(-35 * Math.PI / 180);
        ctx.fillStyle = '#1F2937';
        ctx.font = '15px "Segoe UI", Arial, sans-serif';
        ctx.textAlign = 'right';
        ctx.textBaseline = 'middle';
        ctx.fillText(d.defect, 0, 0);
        ctx.restore();
    });

    // Пунктирная линия 80%
    const y80 = padT + plotH - (80 / 105) * plotH;
    ctx.save();
    ctx.strokeStyle = 'rgba(192, 0, 0, 0.55)';
    ctx.lineWidth = 1.5;
    ctx.setLineDash([8, 6]);
    ctx.beginPath();
    ctx.moveTo(padL, y80);
    ctx.lineTo(padL + plotW, y80);
    ctx.stroke();
    ctx.restore();

    // Линия накопленной доли
    ctx.strokeStyle = '#C00000';
    ctx.lineWidth = 3;
    ctx.beginPath();
    dt.forEach((d, idx) => {
        const cx = padL + slotW * idx + slotW / 2;
        const cy = padT + plotH - (d.cumShare / 105) * plotH;
        if (idx === 0) ctx.moveTo(cx, cy);
        else ctx.lineTo(cx, cy);
    });
    ctx.stroke();

    dt.forEach((d, idx) => {
        const cx = padL + slotW * idx + slotW / 2;
        const cy = padT + plotH - (d.cumShare / 105) * plotH;
        ctx.fillStyle = '#C00000';
        ctx.beginPath();
        ctx.arc(cx, cy, 5.5, 0, Math.PI * 2);
        ctx.fill();
        ctx.strokeStyle = '#FFFFFF';
        ctx.lineWidth = 1.5;
        ctx.stroke();
    });

    // Рамка области графика
    ctx.strokeStyle = '#374151';
    ctx.lineWidth = 1.5;
    ctx.strokeRect(padL, padT, plotW, plotH);

    // Легенда (справа по центру)
    const legW = 225, legH = 68;
    const legX = padL + plotW - legW - 14;
    const legY = padT + plotH * 0.42;
    ctx.fillStyle = 'rgba(255,255,255,0.92)';
    ctx.fillRect(legX, legY, legW, legH);
    ctx.strokeStyle = '#D1D5DB';
    ctx.lineWidth = 1;
    ctx.strokeRect(legX, legY, legW, legH);

    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';
    ctx.font = '15px "Segoe UI", Arial, sans-serif';

    ctx.fillStyle = PLY_CHART_COLORS[0];
    ctx.fillRect(legX + 12, legY + 14, 22, 14);
    ctx.fillStyle = '#111827';
    ctx.fillText(PLY_GROUP1_NAME, legX + 42, legY + 21);

    ctx.fillStyle = PLY_CHART_COLORS[1];
    ctx.fillRect(legX + 12, legY + 40, 22, 14);
    ctx.fillStyle = '#111827';
    ctx.fillText(PLY_GROUP2_NAME, legX + 42, legY + 47);

    const safeV = veneer.replace(',', '_');
    const id = `pareto_${safeV}_${tu}`;
    return plyExportChartObj(canvas, id, filename || `1_pareto_${safeV}_${tu}.png`, `Парето: Рубашка ${veneer} ${tu}`);
}

// 2. Структура брака по группам (2_gruppy.png)
function drawGroupsChart(records, filename = '2_gruppy.png') {
    const groups = plyComputeGroupSummary(records);
    if (groups.length === 0) return null;

    const W = 1280, H = 640;
    const { canvas, ctx } = plyCreateCanvas(W, H);

    const padL = 90, padR = 45, padT = 74, padB = 105;
    const plotW = W - padL - padR;
    const plotH = H - padT - padB;

    const maxBt = Math.max(...groups.map(g => g.bt), 1);
    const yMax = plyNiceMax(maxBt * 1.28, 5);

    ctx.fillStyle = '#111827';
    ctx.font = 'bold 22px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText('Структура брака по группам (листам / ТУ)', W / 2, 34);

    // Сетка
    const ticks = 5;
    ctx.font = '15px "Segoe UI", Arial, sans-serif';
    for (let i = 0; i <= ticks; i++) {
        const val = (yMax / ticks) * i;
        const y = padT + plotH - (val / yMax) * plotH;
        ctx.strokeStyle = i === 0 ? '#374151' : 'rgba(156,163,175,0.45)';
        ctx.lineWidth = i === 0 ? 1.5 : 1;
        ctx.beginPath();
        ctx.moveTo(padL, y);
        ctx.lineTo(padL + plotW, y);
        ctx.stroke();

        ctx.fillStyle = '#374151';
        ctx.textAlign = 'right';
        ctx.fillText(Math.round(val).toString(), padL - 12, y);
    }

    ctx.save();
    ctx.translate(28, padT + plotH / 2);
    ctx.rotate(-Math.PI / 2);
    ctx.fillStyle = '#111827';
    ctx.font = 'bold 16px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText('Брак, л', 0, 0);
    ctx.restore();

    const slotW = plotW / groups.length;
    const barW = Math.min(130, slotW * 0.52);

    groups.forEach((g, idx) => {
        const cx = padL + slotW * idx + slotW / 2;
        const x = cx - barW / 2;
        const h1 = (g.b1 / yMax) * plotH;
        const h2 = (g.b2 / yMax) * plotH;

        ctx.fillStyle = PLY_CHART_COLORS[0];
        ctx.fillRect(x, padT + plotH - h1, barW, h1);

        ctx.fillStyle = PLY_CHART_COLORS[1];
        ctx.fillRect(x, padT + plotH - h1 - h2, barW, h2);

        // Текст над столбцом
        const topY = padT + plotH - h1 - h2;
        ctx.fillStyle = '#111827';
        ctx.font = 'bold 16px "Segoe UI", Arial, sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'bottom';
        ctx.fillText(`${Math.round(g.bt)} л`, cx, topY - 22);
        ctx.font = '15px "Segoe UI", Arial, sans-serif';
        ctx.fillText(`(${g.pct.toFixed(2)}%)`, cx, topY - 4);

        // Подпись X
        ctx.fillStyle = '#1F2937';
        ctx.font = 'bold 16px "Segoe UI", Arial, sans-serif';
        ctx.textBaseline = 'top';
        ctx.fillText(g.labelTop, cx, padT + plotH + 14);
        ctx.font = '15px "Segoe UI", Arial, sans-serif';
        ctx.fillText(g.labelBottom, cx, padT + plotH + 38);
    });

    ctx.strokeStyle = '#374151';
    ctx.lineWidth = 1.5;
    ctx.strokeRect(padL, padT, plotW, plotH);

    // Легенда
    const legW = 225, legH = 66;
    const legX = padL + plotW - legW - 14, legY = padT + 14;
    ctx.fillStyle = 'rgba(255,255,255,0.92)';
    ctx.fillRect(legX, legY, legW, legH);
    ctx.strokeStyle = '#D1D5DB';
    ctx.strokeRect(legX, legY, legW, legH);

    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';
    ctx.font = '15px "Segoe UI", Arial, sans-serif';
    ctx.fillStyle = PLY_CHART_COLORS[0];
    ctx.fillRect(legX + 12, legY + 13, 22, 14);
    ctx.fillStyle = '#111827';
    ctx.fillText(PLY_GROUP1_NAME, legX + 42, legY + 20);

    ctx.fillStyle = PLY_CHART_COLORS[1];
    ctx.fillRect(legX + 12, legY + 39, 22, 14);
    ctx.fillStyle = '#111827';
    ctx.fillText(PLY_GROUP2_NAME, legX + 42, legY + 46);

    return plyExportChartObj(canvas, 'gruppy', filename, 'Структура брака по группам');
}

// 3. % брака по позициям (3_pozicii.png)
function drawByPositionChart(aggRecords, filename = '3_pozicii.png') {
    if (!aggRecords || aggRecords.length === 0) return null;

    const W = 1320;
    const rowStep = 44;
    const padT = 74, padB = 74, padR = 115;
    const H = Math.max(460, padT + padB + aggRecords.length * rowStep + 20);
    const { canvas, ctx } = plyCreateCanvas(W, H);

    ctx.font = 'bold 15px "Segoe UI", Arial, sans-serif';
    const labels = aggRecords.map(r => `${r.position}  [${r.veneer}]`);
    const maxLblW = Math.max(...labels.map(l => ctx.measureText(l).width), 120);
    const padL = Math.min(300, Math.ceil(maxLblW) + 34);
    const plotW = W - padL - padR;
    const plotH = H - padT - padB;

    const sheets = Array.from(new Set(aggRecords.map(r => r.veneer)));
    const totChk = aggRecords.reduce((a, r) => a + r.checked, 0);
    const totBt = aggRecords.reduce((a, r) => a + r.bt, 0);
    const avgPct = totChk > 0 ? (totBt / totChk) * 100 : 0;

    const maxPct = Math.max(...aggRecords.map(r => r.pct), avgPct, 1);
    const xMax = plyNiceMax(maxPct * 1.18 + 0.5, 5);

    ctx.fillStyle = '#111827';
    ctx.font = 'bold 22px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText('% брака по позициям (сорт, толщина) [лист]', W / 2, 34);

    // Вертикальная сетка
    const ticks = 5;
    ctx.font = '15px "Segoe UI", Arial, sans-serif';
    for (let i = 0; i <= ticks; i++) {
        const val = (xMax / ticks) * i;
        const x = padL + (val / xMax) * plotW;
        ctx.strokeStyle = i === 0 ? '#374151' : 'rgba(156,163,175,0.45)';
        ctx.lineWidth = i === 0 ? 1.5 : 1;
        ctx.beginPath();
        ctx.moveTo(x, padT);
        ctx.lineTo(x, padT + plotH);
        ctx.stroke();

        ctx.fillStyle = '#374151';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'top';
        ctx.fillText(val.toFixed(1).replace(/\.0$/, ''), x, padT + plotH + 10);
    }

    ctx.fillStyle = '#111827';
    ctx.font = 'bold 16px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText('% брака', padL + plotW / 2, H - 24);

    // Горизонтальные столбцы
    const slotH = plotH / aggRecords.length;
    const barH = Math.min(30, slotH * 0.68);

    aggRecords.forEach((r, idx) => {
        const cy = padT + slotH * idx + slotH / 2;
        const bw = (r.pct / xMax) * plotW;
        const cIdx = sheets.indexOf(r.veneer) % PLY_CHART_COLORS.length;

        ctx.fillStyle = PLY_CHART_COLORS[cIdx];
        ctx.fillRect(padL, cy - barH / 2, bw, barH);

        // Подпись слева
        ctx.fillStyle = '#1F2937';
        ctx.font = 'bold 15px "Segoe UI", Arial, sans-serif';
        ctx.textAlign = 'right';
        ctx.textBaseline = 'middle';
        ctx.fillText(labels[idx], padL - 12, cy);

        // Значение справа
        ctx.fillStyle = '#111827';
        ctx.font = 'bold 15px "Segoe UI", Arial, sans-serif';
        ctx.textAlign = 'left';
        ctx.fillText(`${r.pct.toFixed(1)}%`, padL + bw + 8, cy);
    });

    // Красная пунктирная линия среднего
    const avgX = padL + (avgPct / xMax) * plotW;
    ctx.save();
    ctx.strokeStyle = '#C00000';
    ctx.lineWidth = 2.5;
    ctx.setLineDash([8, 6]);
    ctx.beginPath();
    ctx.moveTo(avgX, padT);
    ctx.lineTo(avgX, padT + plotH);
    ctx.stroke();
    ctx.restore();

    ctx.strokeStyle = '#374151';
    ctx.lineWidth = 1.5;
    ctx.strokeRect(padL, padT, plotW, plotH);

    // Легенда (справа снизу)
    const legW = 265, legH = 42;
    const legX = padL + plotW - legW - 12;
    const legY = padT + plotH - legH - 12;
    ctx.fillStyle = 'rgba(255,255,255,0.94)';
    ctx.fillRect(legX, legY, legW, legH);
    ctx.strokeStyle = '#D1D5DB';
    ctx.strokeRect(legX, legY, legW, legH);

    ctx.save();
    ctx.strokeStyle = '#C00000';
    ctx.lineWidth = 2.5;
    ctx.setLineDash([6, 4]);
    ctx.beginPath();
    ctx.moveTo(legX + 12, legY + 21);
    ctx.lineTo(legX + 42, legY + 21);
    ctx.stroke();
    ctx.restore();

    ctx.fillStyle = '#111827';
    ctx.font = '15px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';
    ctx.fillText(`Средний (общий) ${avgPct.toFixed(2)}%`, legX + 50, legY + 21);

    return plyExportChartObj(canvas, 'pozicii', filename, '% брака по позициям');
}

// 3b. Наибольший дефект по позициям (3b_max_defekt_pozicii.png)
function drawMaxDefectByPositionChart(aggRecords, filename = '3b_max_defekt_pozicii.png') {
    if (!aggRecords || aggRecords.length === 0) return null;

    const W = 1320;
    const rowStep = 44;
    const padT = 74, padB = 74, padR = 60;
    const H = Math.max(460, padT + padB + aggRecords.length * rowStep + 20);
    const { canvas, ctx } = plyCreateCanvas(W, H);

    ctx.font = 'bold 15px "Segoe UI", Arial, sans-serif';
    const labels = aggRecords.map(r => `${r.position}  [${r.veneer}]`);
    const maxLblW = Math.max(...labels.map(l => ctx.measureText(l).width), 120);
    const padL = Math.min(300, Math.ceil(maxLblW) + 34);
    const plotW = W - padL - padR;
    const plotH = H - padT - padB;

    const sheets = Array.from(new Set(aggRecords.map(r => r.veneer)));
    const maxVal = Math.max(...aggRecords.map(r => r.maxDefVal), 0);
    const xMax = maxVal > 0 ? plyNiceMax(maxVal * 1.55, 5) : 10;

    ctx.fillStyle = '#111827';
    ctx.font = 'bold 22px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText('Наибольший дефект по позициям (в листах) [лист]', W / 2, 34);

    // Сетка
    const ticks = 5;
    ctx.font = '15px "Segoe UI", Arial, sans-serif';
    for (let i = 0; i <= ticks; i++) {
        const val = (xMax / ticks) * i;
        const x = padL + (val / xMax) * plotW;
        ctx.strokeStyle = i === 0 ? '#374151' : 'rgba(156,163,175,0.45)';
        ctx.lineWidth = i === 0 ? 1.5 : 1;
        ctx.beginPath();
        ctx.moveTo(x, padT);
        ctx.lineTo(x, padT + plotH);
        ctx.stroke();

        ctx.fillStyle = '#374151';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'top';
        ctx.fillText(Math.round(val).toString(), x, padT + plotH + 10);
    }

    ctx.fillStyle = '#111827';
    ctx.font = 'bold 16px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText('Количество брака, л', padL + plotW / 2, H - 24);

    const slotH = plotH / aggRecords.length;
    const barH = Math.min(30, slotH * 0.68);

    aggRecords.forEach((r, idx) => {
        const cy = padT + slotH * idx + slotH / 2;
        const bw = (r.maxDefVal / xMax) * plotW;
        const cIdx = sheets.indexOf(r.veneer) % PLY_CHART_COLORS.length;

        ctx.fillStyle = PLY_CHART_COLORS[cIdx];
        ctx.fillRect(padL, cy - barH / 2, bw, barH);

        ctx.fillStyle = '#1F2937';
        ctx.font = 'bold 15px "Segoe UI", Arial, sans-serif';
        ctx.textAlign = 'right';
        ctx.textBaseline = 'middle';
        ctx.fillText(labels[idx], padL - 12, cy);

        const barText = r.maxDefVal > 0
            ? `${Math.round(r.maxDefVal)} л (${r.maxDefName}) из ${Math.round(r.checked)} л`
            : `0 л из ${Math.round(r.checked)} л`;
        ctx.fillStyle = '#111827';
        ctx.font = '14px "Segoe UI", Arial, sans-serif';
        ctx.textAlign = 'left';
        ctx.fillText(barText, padL + bw + 8, cy);
    });

    ctx.strokeStyle = '#374151';
    ctx.lineWidth = 1.5;
    ctx.strokeRect(padL, padT, plotW, plotH);

    // Легенда по листам
    const legW = 145, legH = sheets.length * 26 + 16;
    const legX = padL + plotW - legW - 12;
    const legY = padT + plotH - legH - 12;
    ctx.fillStyle = 'rgba(255,255,255,0.94)';
    ctx.fillRect(legX, legY, legW, legH);
    ctx.strokeStyle = '#D1D5DB';
    ctx.strokeRect(legX, legY, legW, legH);

    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';
    ctx.font = '14px "Segoe UI", Arial, sans-serif';
    sheets.forEach((sh, i) => {
        ctx.fillStyle = PLY_CHART_COLORS[i % PLY_CHART_COLORS.length];
        ctx.fillRect(legX + 12, legY + 12 + i * 26, 20, 14);
        ctx.fillStyle = '#111827';
        ctx.fillText(`Лист ${sh}`, legX + 40, legY + 19 + i * 26);
    });

    return plyExportChartObj(canvas, 'max_defekt_pozicii', filename, 'Наибольший дефект по позициям');
}

// 4. Внутрисменная динамика (4_vnutrismennaya.png)
function drawIntrashiftChart(rawRecords, filename = '4_vnutrismennaya.png') {
    if (!rawRecords || rawRecords.length === 0) return null;

    const W = 1320, H = 620;
    const { canvas, ctx } = plyCreateCanvas(W, H);

    const padL = 88, padR = 45, padT = 74, padB = 135;
    const plotW = W - padL - padR;
    const plotH = H - padT - padB;

    let cumBt = 0, cumChk = 0;
    const pts = rawRecords.map(r => {
        cumBt += r.bt;
        cumChk += r.checked;
        return {
            label: r.position,
            pct: r.pct,
            cumPct: cumChk > 0 ? (cumBt / cumChk) * 100 : 0
        };
    });

    const meanPct = pts.reduce((a, p) => a + p.pct, 0) / pts.length;
    const maxVal = Math.max(...pts.map(p => Math.max(p.pct, p.cumPct)), meanPct, 1);
    const yMax = plyNiceMax(maxVal * 1.2, 5);

    ctx.fillStyle = '#111827';
    ctx.font = 'bold 22px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText('Внутрисменная динамика брака (позиции в порядке контроля)', W / 2, 34);

    // Сетка Y
    const ticks = 5;
    ctx.font = '15px "Segoe UI", Arial, sans-serif';
    for (let i = 0; i <= ticks; i++) {
        const val = (yMax / ticks) * i;
        const y = padT + plotH - (val / yMax) * plotH;
        ctx.strokeStyle = i === 0 ? '#374151' : 'rgba(156,163,175,0.45)';
        ctx.lineWidth = i === 0 ? 1.5 : 1;
        ctx.beginPath();
        ctx.moveTo(padL, y);
        ctx.lineTo(padL + plotW, y);
        ctx.stroke();

        ctx.fillStyle = '#374151';
        ctx.textAlign = 'right';
        ctx.fillText(val.toFixed(1).replace(/\.0$/, ''), padL - 12, y);
    }

    ctx.save();
    ctx.translate(28, padT + plotH / 2);
    ctx.rotate(-Math.PI / 2);
    ctx.fillStyle = '#111827';
    ctx.font = 'bold 16px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText('%', 0, 0);
    ctx.restore();

    const getX = (idx) => pts.length === 1
        ? padL + plotW / 2
        : padL + (plotW / (pts.length - 1)) * idx * 0.92 + plotW * 0.04;
    const getY = (val) => padT + plotH - (val / yMax) * plotH;

    // Серая пунктирная линия среднего
    const meanY = getY(meanPct);
    ctx.save();
    ctx.strokeStyle = '#6B7280';
    ctx.lineWidth = 1.8;
    ctx.setLineDash([4, 5]);
    ctx.beginPath();
    ctx.moveTo(padL, meanY);
    ctx.lineTo(padL + plotW, meanY);
    ctx.stroke();
    ctx.restore();

    // Синяя линия (% брака позиции)
    ctx.strokeStyle = PLY_CHART_COLORS[0];
    ctx.lineWidth = 2.5;
    ctx.beginPath();
    pts.forEach((p, i) => {
        const x = getX(i), y = getY(p.pct);
        if (i === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
    });
    ctx.stroke();

    // Оранжевая линия (Накопленный % брака)
    ctx.strokeStyle = PLY_CHART_COLORS[1];
    ctx.lineWidth = 3.2;
    ctx.beginPath();
    pts.forEach((p, i) => {
        const x = getX(i), y = getY(p.cumPct);
        if (i === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
    });
    ctx.stroke();

    // Маркеры и подписи X
    pts.forEach((p, i) => {
        const x = getX(i);
        const y1 = getY(p.pct);
        const y2 = getY(p.cumPct);

        ctx.fillStyle = PLY_CHART_COLORS[0];
        ctx.beginPath();
        ctx.arc(x, y1, 5.5, 0, Math.PI * 2);
        ctx.fill();

        ctx.fillStyle = PLY_CHART_COLORS[1];
        ctx.fillRect(x - 5.5, y2 - 5.5, 11, 11);

        ctx.save();
        ctx.translate(x, padT + plotH + 12);
        ctx.rotate(-40 * Math.PI / 180);
        ctx.fillStyle = '#1F2937';
        ctx.font = '14px "Segoe UI", Arial, sans-serif';
        ctx.textAlign = 'right';
        ctx.textBaseline = 'middle';
        ctx.fillText(p.label, 0, 0);
        ctx.restore();
    });

    ctx.strokeStyle = '#374151';
    ctx.lineWidth = 1.5;
    ctx.strokeRect(padL, padT, plotW, plotH);

    // Легенда
    const legW = 285, legH = 92;
    const legX = padL + plotW - legW - 14, legY = padT + 14;
    ctx.fillStyle = 'rgba(255,255,255,0.93)';
    ctx.fillRect(legX, legY, legW, legH);
    ctx.strokeStyle = '#D1D5DB';
    ctx.strokeRect(legX, legY, legW, legH);

    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';
    ctx.font = '14px "Segoe UI", Arial, sans-serif';

    ctx.strokeStyle = PLY_CHART_COLORS[0]; ctx.lineWidth = 2.5;
    ctx.beginPath(); ctx.moveTo(legX + 12, legY + 20); ctx.lineTo(legX + 40, legY + 20); ctx.stroke();
    ctx.fillStyle = '#111827'; ctx.fillText('% брака позиции', legX + 48, legY + 20);

    ctx.strokeStyle = PLY_CHART_COLORS[1]; ctx.lineWidth = 3;
    ctx.beginPath(); ctx.moveTo(legX + 12, legY + 46); ctx.lineTo(legX + 40, legY + 46); ctx.stroke();
    ctx.fillStyle = '#111827'; ctx.fillText('Накопленный % брака', legX + 48, legY + 46);

    ctx.save();
    ctx.strokeStyle = '#6B7280'; ctx.lineWidth = 2; ctx.setLineDash([4, 4]);
    ctx.beginPath(); ctx.moveTo(legX + 12, legY + 72); ctx.lineTo(legX + 40, legY + 72); ctx.stroke();
    ctx.restore();
    ctx.fillStyle = '#111827'; ctx.fillText(`Ср. по позициям ${meanPct.toFixed(2)}%`, legX + 48, legY + 72);

    return plyExportChartObj(canvas, 'vnutrismennaya', filename, 'Внутрисменная динамика брака');
}

// 5. Динамика по сменам / датам (5_dinamika.png)
function drawDynamicsChart(dynRows, filename = '5_dinamika.png') {
    if (!dynRows || dynRows.length <= 1) return null;

    const W = 1320, H = 640;
    const { canvas, ctx } = plyCreateCanvas(W, H);

    const padL = 88, padR = 45, padT = 74, padB = 145;
    const plotW = W - padL - padR;
    const plotH = H - padT - padB;

    const allVals = [];
    dynRows.forEach(r => {
        allVals.push(r.pctTotal);
        if (r.pct09 !== null) allVals.push(r.pct09);
        if (r.pct141 !== null) allVals.push(r.pct141);
    });
    const yMax = plyNiceMax(Math.max(...allVals, 1) * 1.25, 5);

    ctx.fillStyle = '#111827';
    ctx.font = 'bold 22px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText('Динамика % брака по сменам / датам', W / 2, 34);

    const ticks = 5;
    ctx.font = '15px "Segoe UI", Arial, sans-serif';
    for (let i = 0; i <= ticks; i++) {
        const val = (yMax / ticks) * i;
        const y = padT + plotH - (val / yMax) * plotH;
        ctx.strokeStyle = i === 0 ? '#374151' : 'rgba(156,163,175,0.45)';
        ctx.lineWidth = i === 0 ? 1.5 : 1;
        ctx.beginPath();
        ctx.moveTo(padL, y);
        ctx.lineTo(padL + plotW, y);
        ctx.stroke();

        ctx.fillStyle = '#374151';
        ctx.textAlign = 'right';
        ctx.fillText(val.toFixed(1).replace(/\.0$/, ''), padL - 12, y);
    }

    ctx.save();
    ctx.translate(28, padT + plotH / 2);
    ctx.rotate(-Math.PI / 2);
    ctx.fillStyle = '#111827';
    ctx.font = 'bold 16px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText('% брака', 0, 0);
    ctx.restore();

    const getX = (idx) => padL + (plotW / (dynRows.length - 1)) * idx * 0.92 + plotW * 0.04;
    const getY = (val) => padT + plotH - (val / yMax) * plotH;

    const drawSeries = (getter, color, width, dashed, annotate) => {
        ctx.save();
        ctx.strokeStyle = color;
        ctx.lineWidth = width;
        if (dashed) ctx.setLineDash([8, 6]);
        ctx.beginPath();
        let started = false;
        dynRows.forEach((r, i) => {
            const v = getter(r);
            if (v === null || v === undefined) {
                started = false;
                return;
            }
            const x = getX(i), y = getY(v);
            if (!started) { ctx.moveTo(x, y); started = true; }
            else ctx.lineTo(x, y);
        });
        ctx.stroke();
        ctx.restore();

        dynRows.forEach((r, i) => {
            const v = getter(r);
            if (v === null || v === undefined) return;
            const x = getX(i), y = getY(v);
            ctx.fillStyle = color;
            ctx.beginPath();
            ctx.arc(x, y, 5, 0, Math.PI * 2);
            ctx.fill();

            if (annotate) {
                ctx.fillStyle = '#111827';
                ctx.font = 'bold 14px "Segoe UI", Arial, sans-serif';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'bottom';
                ctx.fillText(v.toFixed(2), x, y - 9);
            }
        });
    };

    const has09 = dynRows.some(r => r.pct09 !== null);
    const has141 = dynRows.some(r => r.pct141 !== null);

    if (has09) drawSeries(r => r.pct09, PLY_CHART_COLORS[0], 2, true, false);
    if (has141) drawSeries(r => r.pct141, PLY_CHART_COLORS[1], 2, true, false);
    drawSeries(r => r.pctTotal, '#000000', 3.2, false, true);

    // Подписи дат
    dynRows.forEach((r, i) => {
        const x = getX(i);
        ctx.save();
        ctx.translate(x, padT + plotH + 12);
        ctx.rotate(-35 * Math.PI / 180);
        ctx.fillStyle = '#1F2937';
        ctx.font = '14px "Segoe UI", Arial, sans-serif';
        ctx.textAlign = 'right';
        ctx.textBaseline = 'middle';
        ctx.fillText(r.period, 0, 0);
        ctx.restore();
    });

    ctx.strokeStyle = '#374151';
    ctx.lineWidth = 1.5;
    ctx.strokeRect(padL, padT, plotW, plotH);

    // Легенда
    const legItems = [{ label: 'Всего', color: '#000000', dashed: false }];
    if (has09) legItems.push({ label: 'лист «0,9»', color: PLY_CHART_COLORS[0], dashed: true });
    if (has141) legItems.push({ label: 'лист «1,41»', color: PLY_CHART_COLORS[1], dashed: true });

    const legW = 180, legH = legItems.length * 26 + 16;
    const legX = padL + plotW - legW - 14, legY = padT + 14;
    ctx.fillStyle = 'rgba(255,255,255,0.93)';
    ctx.fillRect(legX, legY, legW, legH);
    ctx.strokeStyle = '#D1D5DB';
    ctx.strokeRect(legX, legY, legW, legH);

    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';
    ctx.font = '14px "Segoe UI", Arial, sans-serif';
    legItems.forEach((it, idx) => {
        const y = legY + 20 + idx * 26;
        ctx.save();
        ctx.strokeStyle = it.color;
        ctx.lineWidth = it.dashed ? 2 : 3;
        if (it.dashed) ctx.setLineDash([6, 4]);
        ctx.beginPath(); ctx.moveTo(legX + 12, y); ctx.lineTo(legX + 40, y); ctx.stroke();
        ctx.restore();
        ctx.fillStyle = '#111827';
        ctx.fillText(it.label, legX + 48, y);
    });

    return plyExportChartObj(canvas, 'dinamika', filename, 'Динамика % брака по сменам');
}

// ------------------- 4. СБОРЩИКИ ПАКЕТОВ ГРАФИКОВ И PDF -------------------

// Полный пакет графиков для одной смены (с раздельными Парето ТУ-Э / ТУ-Р и пропуском пустых)
function generateShiftChartsBundle(shiftCycles, periodLabel, monthAllCycles = []) {
    const rawRecords = plyNormalizeCycles(shiftCycles);
    if (rawRecords.length === 0) return [];

    const aggRecords = plyAggregatePositions(rawRecords);
    const charts = [];

    const add = (ch) => { if (ch) charts.push(ch); };

    add(drawSummaryTableChart(aggRecords, periodLabel, '0_svodka.png'));

    // Раздельные Парето по 4 блокам (если данных нет — drawParetoChart вернёт null)
    const paretoBlocks = [
        ['0,9', 'ТУ-Э', '1_pareto_0_9_ТУ-Э.png'],
        ['0,9', 'ТУ-Р', '1_pareto_0_9_ТУ-Р.png'],
        ['1,41', 'ТУ-Э', '1_pareto_1_41_ТУ-Э.png'],
        ['1,41', 'ТУ-Р', '1_pareto_1_41_ТУ-Р.png']
    ];
    paretoBlocks.forEach(([v, tu, fname]) => {
        add(drawParetoChart(aggRecords, v, tu, fname));
    });

    add(drawGroupsChart(aggRecords, '2_gruppy.png'));
    add(drawByPositionChart(aggRecords, '3_pozicii.png'));
    add(drawMaxDefectByPositionChart(aggRecords, '3b_max_defekt_pozicii.png'));
    add(drawIntrashiftChart(rawRecords, '4_vnutrismennaya.png'));

    if (Array.isArray(monthAllCycles) && monthAllCycles.length > 0) {
        const monthNorm = plyNormalizeCycles(monthAllCycles);
        const dyn = plyComputeDynamics(monthNorm);
        add(drawDynamicsChart(dyn, '5_dinamika.png'));
    }

    return charts;
}

// Пакет накопительных графиков за месяц (аналог 00_Итого_за_месяц_графики)
function generateMonthSummaryChartsBundle(monthCycles, monthLabel) {
    const rawRecords = plyNormalizeCycles(monthCycles);
    if (rawRecords.length === 0) return [];

    const aggRecords = plyAggregatePositions(rawRecords);
    const dyn = plyComputeDynamics(rawRecords);
    const charts = [];
    const add = (ch) => { if (ch) charts.push(ch); };

    add(drawSummaryTableChart(aggRecords, `ИТОГО: ${monthLabel}`, '0_svodka_mesyac.png'));

    const paretoBlocks = [
        ['0,9', 'ТУ-Э', '1_pareto_0_9_ТУ-Э_mesyac.png'],
        ['0,9', 'ТУ-Р', '1_pareto_0_9_ТУ-Р_mesyac.png'],
        ['1,41', 'ТУ-Э', '1_pareto_1_41_ТУ-Э_mesyac.png'],
        ['1,41', 'ТУ-Р', '1_pareto_1_41_ТУ-Р_mesyac.png']
    ];
    paretoBlocks.forEach(([v, tu, fname]) => {
        add(drawParetoChart(aggRecords, v, tu, fname));
    });

    add(drawGroupsChart(aggRecords, '2_gruppy_mesyac.png'));
    add(drawDynamicsChart(dyn, '5_dinamika_mesyac.png'));

    return charts;
}

// Объединение массива PNG-графиков в один PDF-файл (альбомный A4, 1 график = 1 страница)
function combineChartsToPdf(charts) {
    if (!charts || charts.length === 0) return null;
    if (!window.jspdf || !window.jspdf.jsPDF) {
        throw new Error('Библиотека jsPDF не загружена');
    }

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    const pageW = doc.internal.pageSize.getWidth();
    const pageH = doc.internal.pageSize.getHeight();
    const margin = 10;
    const maxW = pageW - margin * 2;
    const maxH = pageH - margin * 2 - 6;

    charts.forEach((ch, idx) => {
        if (idx > 0) doc.addPage('a4', 'landscape');
        const ratio = Math.min(maxW / ch.width, maxH / ch.height);
        const drawW = ch.width * ratio;
        const drawH = ch.height * ratio;
        const x = (pageW - drawW) / 2;
        const y = (pageH - drawH) / 2 - 2;

        doc.addImage(ch.dataUrl, 'PNG', x, y, drawW, drawH, undefined, 'FAST');
        doc.setFontSize(8);
        doc.setTextColor(120);
        doc.text(`${idx + 1} / ${charts.length}`, pageW - margin, pageH - 5, { align: 'right' });
    });

    const blob = doc.output('blob');
    const dataUri = doc.output('datauristring');
    const base64 = dataUri.split(',')[1] || '';
    return { blob, base64 };
}
// ------------------- 5. ЖИВАЯ ПАНЕЛЬ ГРАФИКОВ ТЕКУЩЕЙ СМЕНЫ -------------------

let currentLiveShiftCharts = [];
let lastLiveChartsSignature = '';

function getCurrentMonthOtherShiftCycles() {
    const activeDate = (appState.shiftStartDate || appState.date || '').slice(0, 7);
    if (!activeDate) return [];

    const collected = [];
    const currentFileId = appState.currentShiftFileId || null;

    if (typeof savedShiftReports !== 'undefined' && Array.isArray(savedShiftReports)) {
        savedShiftReports.forEach(rep => {
            if (rep.monthKey === activeDate && rep.fileId !== currentFileId && Array.isArray(rep.cycles)) {
                collected.push(...rep.cycles);
            }
        });
    }

    if (collected.length === 0 && typeof savedFiles !== 'undefined' && Array.isArray(savedFiles)) {
        savedFiles.forEach(f => {
            if (f.id !== currentFileId && f.date && f.date.slice(0, 7) === activeDate && Array.isArray(f.cycles)) {
                collected.push(...f.cycles);
            }
        });
    }

    return collected;
}

function renderLiveShiftCharts(force = false) {
    const grid = document.getElementById('liveShiftChartsGrid');
    const badge = document.getElementById('liveChartsCountBadge');
    const btnMax = document.getElementById('btnLiveChartsMaxPdf');
    const btnDl = document.getElementById('btnLiveChartsDownloadPdf');
    if (!grid) return;

    const cycles = Array.isArray(appState.accumulatedCycles) ? appState.accumulatedCycles : [];

    if (cycles.length === 0) {
        currentLiveShiftCharts = [];
        lastLiveChartsSignature = '';
        if (badge) badge.textContent = '0 графиков';
        if (btnMax) btnMax.disabled = true;
        if (btnDl) btnDl.disabled = true;
        grid.innerHTML = `
            <div class="col-span-full p-5 rounded-xl bg-slate-900/70 border border-slate-800 text-center text-xs text-slate-400">
                Завершите хотя бы один цикл точковки — PNG-графики текущей смены сформируются здесь автоматически.
            </div>
        `;
        return;
    }

    const shiftDateISO = appState.shiftStartDate || (cycles[0] && cycles[0].date) || appState.date;
    const shiftNum = appState.shift || (cycles[0] && cycles[0].shift) || '1';
    const periodLabel = `${plyFormatDateRuShort(shiftDateISO)} см.${shiftNum}`;

    const otherMonthCycles = getCurrentMonthOtherShiftCycles();
    const signature = JSON.stringify({ periodLabel, cycles, otherCount: otherMonthCycles.length });

    if (!force && signature === lastLiveChartsSignature && currentLiveShiftCharts.length > 0) {
        return;
    }
    lastLiveChartsSignature = signature;

    const monthAllCycles = [...otherMonthCycles, ...cycles];
    currentLiveShiftCharts = generateShiftChartsBundle(cycles, periodLabel, monthAllCycles);

    if (badge) badge.textContent = `${currentLiveShiftCharts.length} графиков`;
    if (btnMax) btnMax.disabled = currentLiveShiftCharts.length === 0;
    if (btnDl) btnDl.disabled = currentLiveShiftCharts.length === 0;

    grid.innerHTML = currentLiveShiftCharts.map((ch, idx) => `
        <div class="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden flex flex-col shadow">
            <div class="px-3 py-2 bg-slate-900 border-b border-slate-800 flex items-center justify-between gap-2">
                <div class="truncate">
                    <span class="text-xs font-extrabold text-white block truncate">${ch.title}</span>
                    <span class="text-[10px] text-slate-400 mono">${ch.filename}</span>
                </div>
                <div class="flex items-center gap-1.5 shrink-0">
                    <button
                        type="button"
                        onclick="openLiveChartZoom(${idx})"
                        class="px-2 py-1 rounded-lg bg-slate-800 hover:bg-slate-700 text-sky-300 text-xs font-bold flex items-center gap-1 transition"
                        title="Открыть на весь экран"
                    >
                        <i class="ph-bold ph-magnifying-glass-plus"></i>
                    </button>
                    <button
                        type="button"
                        onclick="downloadSingleLiveChartPng(${idx})"
                        class="px-2.5 py-1 rounded-lg bg-emerald-500/15 hover:bg-emerald-500 hover:text-slate-950 text-emerald-300 border border-emerald-500/30 text-xs font-bold flex items-center gap-1 transition"
                        title="Скачать картинку PNG"
                    >
                        <i class="ph-bold ph-download-simple"></i>
                        <span>PNG</span>
                    </button>
                </div>
            </div>
            <div onclick="openLiveChartZoom(${idx})" class="bg-white p-1.5 cursor-pointer hover:opacity-95 transition">
                <img src="${ch.dataUrl}" alt="${ch.title}" class="w-full h-auto block rounded">
            </div>
        </div>
    `).join('');
}

function downloadSingleLiveChartPng(idx) {
    const ch = currentLiveShiftCharts[idx];
    if (!ch) return;
    const a = document.createElement('a');
    a.href = ch.dataUrl;
    a.download = ch.filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    showToast(`Скачан график: ${ch.filename}`, 'emerald');
}

function buildLiveChartsPdfInfo() {
    if (!currentLiveShiftCharts || currentLiveShiftCharts.length === 0) {
        showToast('Нет графиков для формирования PDF!', 'amber');
        return null;
    }
    const cycles = appState.accumulatedCycles || [];
    const shiftDateISO = appState.shiftStartDate || (cycles[0] && cycles[0].date) || appState.date;
    const shiftNum = appState.shift || (cycles[0] && cycles[0].shift) || '1';
    const dateRu = plyFormatDateRuShort(shiftDateISO);
    const filename = `Графики_${dateRu}_см${shiftNum}.pdf`;

    const totSheets = cycles.reduce((a, c) => a + (Number(c.checked) || 0), 0);
    const totDef = cycles.reduce((a, c) => a + (Number(c.sumAA) || 0), 0);
    const goodPct = totSheets > 0 ? (((totSheets - totDef) / totSheets) * 100).toFixed(2) : '0.00';

    const summaryText = `Графики точковки PlyTap (${dateRu}, Смена ${shiftNum})\nОператор: ${appState.operator || '-'}\nПроверено: ${totSheets} л., Брак: ${totDef} л., Годных: ${goodPct}%`;
    const pdf = combineChartsToPdf(currentLiveShiftCharts);
    return { pdf, filename, summaryText };
}

function downloadLiveChartsPdf() {
    try {
        const info = buildLiveChartsPdfInfo();
        if (!info || !info.pdf) return;
        const url = URL.createObjectURL(info.pdf.blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = info.filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 1500);
        showToast(`PDF сохранён: ${info.filename}`, 'emerald');
    } catch (e) {
        console.error(e);
        showToast('Ошибка при создании PDF из графиков', 'amber');
    }
}

async function sendLiveChartsToMaxPdf() {
    try {
        const info = buildLiveChartsPdfInfo();
        if (!info || !info.pdf) return;
        await sharePdfToMax(info.pdf.blob, info.filename, info.summaryText);
    } catch (e) {
        console.error(e);
        showToast('Ошибка при отправке PDF в MAX', 'amber');
    }
}

function openChartLightbox(dataUrl, title, filename) {
    let modal = document.getElementById('plyChartLightboxModal');
    if (!modal) {
        modal = document.createElement('div');
        modal.id = 'plyChartLightboxModal';
        modal.className = 'fixed inset-0 z-50 hidden items-center justify-center bg-slate-950/90 backdrop-blur-sm p-3 sm:p-6';
        modal.innerHTML = `
            <div class="bg-slate-900 border-2 border-slate-700 rounded-2xl max-w-6xl w-full max-h-[94vh] flex flex-col overflow-hidden shadow-2xl">
                <div class="px-4 py-3 border-b border-slate-800 flex items-center justify-between gap-2">
                    <div class="truncate">
                        <h4 id="plyLightboxTitle" class="text-sm sm:text-base font-extrabold text-white truncate"></h4>
                        <span id="plyLightboxFile" class="text-xs text-slate-400 mono"></span>
                    </div>
                    <div class="flex items-center gap-2">
                        <a id="plyLightboxDlLink" href="#" download="" class="px-3 py-1.5 rounded-xl bg-emerald-500 hover:bg-emerald-400 text-slate-950 text-xs font-extrabold flex items-center gap-1.5">
                            <i class="ph-bold ph-download-simple"></i>
                            <span>Скачать PNG</span>
                        </a>
                        <button type="button" onclick="closeChartLightbox()" class="p-1.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white">
                            <i class="ph-bold ph-x text-lg"></i>
                        </button>
                    </div>
                </div>
                <div class="p-3 overflow-auto bg-white flex items-center justify-center">
                    <img id="plyLightboxImg" src="" alt="" class="max-w-full h-auto block">
                </div>
            </div>
        `;
        modal.addEventListener('click', (e) => {
            if (e.target === modal) closeChartLightbox();
        });
        document.body.appendChild(modal);
    }

    document.getElementById('plyLightboxTitle').textContent = title || 'График';
    document.getElementById('plyLightboxFile').textContent = filename || '';
    document.getElementById('plyLightboxImg').src = dataUrl;
    const dl = document.getElementById('plyLightboxDlLink');
    dl.href = dataUrl;
    dl.download = filename || 'chart.png';

    modal.classList.remove('hidden');
    modal.classList.add('flex');
}

function closeChartLightbox() {
    const modal = document.getElementById('plyChartLightboxModal');
    if (modal) {
        modal.classList.add('hidden');
        modal.classList.remove('flex');
    }
}

function openLiveChartZoom(idx) {
    const ch = currentLiveShiftCharts[idx];
    if (!ch) return;
    openChartLightbox(ch.dataUrl, ch.title, ch.filename);
}