// =========================================================================
// PLYTAP-PDF-REPORT.JS — Аналитический PDF-отчёт (порт monthly_report.py)
// =========================================================================

const MR_BLUE   = '#3274A1';
const MR_LBLUE  = '#72A8D4';
const MR_GREEN  = '#548B54';
const MR_RED    = '#C0392B';
const MR_ORANGE = '#E6A100';

// ------------------- 1. ПОДГОТОВКА ДАННЫХ -------------------

function mrFmtInt(n) {
    return Math.round(Number(n) || 0).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
}

function mrFmtFlt(n, digits = 2) {
    return (Number(n) || 0).toFixed(digits).replace('.', ',');
}

function mrExtractThickness(posName) {
    const matches = String(posName || '').match(/\d+(?:[.,]\d+)?/g);
    if (!matches || matches.length === 0) return 0;
    const val = parseFloat(matches[matches.length - 1].replace(',', '.'));
    return Number.isFinite(val) ? val : 0;
}

function mrFormatThickness(th) {
    return Number.isInteger(th) ? String(th) : String(th).replace('.', ',');
}

function mrPrepareRecords(cycles) {
    const raw = plyNormalizeCycles(cycles);
    const agg = plyAggregatePositions(raw);

    return agg.map(item => {
        const defectsMap = {};
        ALL_DEFECT_KEYS.forEach(k => {
            const v = Number(item.counts[k]) || 0;
            if (v > 0) {
                const name = plyChartDefectTitle(k, item.customTitles).toLowerCase();
                defectsMap[name] = (defectsMap[name] || 0) + v;
            }
        });

        return {
            veneer: item.veneer,
            tu: item.tu,
            group: `${item.veneer} мм ${item.tu}`,
            position: item.position,
            posFull: `${item.position} (${item.veneer} мм)`,
            posShort: `${item.position} (${item.veneer})`,
            thickness: mrExtractThickness(item.position),
            checked: item.checked,
            techDef: item.b1,
            natDef: item.b2,
            totalDef: item.bt,
            defRate: item.pct,
            qualRate: item.goodPct,
            defects: defectsMap
        };
    });
}

function mrGetNaturalDefectNames(cycles) {
    const natSet = new Set();
    const mergedTitles = {};
    (cycles || []).forEach(c => {
        if (c && c.customTitles) Object.assign(mergedTitles, c.customTitles);
    });
    DEFECT_KEYS_PY.forEach(k => {
        natSet.add(plyChartDefectTitle(k, mergedTitles).toLowerCase());
    });
    return natSet;
}

function mrBuildParetoData(subRecords, natSet) {
    const subChecked = subRecords.reduce((a, r) => a + r.checked, 0);
    const subDef = subRecords.reduce((a, r) => a + r.totalDef, 0);
    const counts = {};

    subRecords.forEach(r => {
        Object.entries(r.defects).forEach(([k, v]) => {
            counts[k] = (counts[k] || 0) + v;
        });
    });

    const rows = Object.entries(counts)
        .filter(([, v]) => v > 0)
        .sort((a, b) => b[1] - a[1])
        .map(([defect, count]) => ({
            defect,
            nature: natSet.has(defect) ? 'природный' : 'технолог.',
            count,
            share: subDef > 0 ? (count / subDef) * 100 : 0,
            per1000: subChecked > 0 ? (count / subChecked) * 1000 : 0
        }));

    let cum = 0;
    rows.forEach(r => {
        cum += r.share;
        r.cumShare = cum;
    });
    return rows;
}

// ------------------- 2. ОТРИСОВКА 8 ГРАФИКОВ MONTHLY_REPORT -------------------

// График 1: Объём контроля и уровень дефектности по группам
function mrDrawChart1Groups(grpRows) {
    const W = 1200, H = 560;
    const { canvas, ctx } = plyCreateCanvas(W, H);
    const padL = 100, padR = 95, padT = 68, padB = 82;
    const plotW = W - padL - padR, plotH = H - padT - padB;

    const maxChk = Math.max(...grpRows.map(g => g.checked), 1);
    const y1Max = plyNiceMax(maxChk * 1.22, 10);
    const maxRate = Math.max(...grpRows.map(g => g.defRate), 1);
    const y2Max = Math.max(10, Math.ceil(maxRate * 1.35));

    ctx.fillStyle = '#111827';
    ctx.font = 'bold 21px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText('Объём контроля и уровень дефектности по группам продукции', W / 2, 32);

    // Сетка и ось Y1
    ctx.font = '14px "Segoe UI", Arial, sans-serif';
    for (let i = 0; i <= 5; i++) {
        const val = (y1Max / 5) * i;
        const y = padT + plotH - (val / y1Max) * plotH;
        ctx.save();
        ctx.strokeStyle = i === 0 ? '#374151' : 'rgba(156,163,175,0.45)';
        ctx.lineWidth = i === 0 ? 1.5 : 1;
        if (i > 0) ctx.setLineDash([3, 4]);
        ctx.beginPath(); ctx.moveTo(padL, y); ctx.lineTo(padL + plotW, y); ctx.stroke();
        ctx.restore();

        ctx.fillStyle = '#374151';
        ctx.textAlign = 'right';
        ctx.fillText(mrFmtInt(val), padL - 10, y);

        const rVal = (y2Max / 5) * i;
        ctx.fillStyle = MR_RED;
        ctx.textAlign = 'left';
        ctx.fillText(rVal.toFixed(1).replace(/\.0$/, ''), padL + plotW + 10, y);
    }

    ctx.save();
    ctx.translate(26, padT + plotH / 2); ctx.rotate(-Math.PI / 2);
    ctx.fillStyle = '#111827'; ctx.font = 'bold 15px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center'; ctx.fillText('Проверено, листов', 0, 0);
    ctx.restore();

    ctx.save();
    ctx.translate(W - 22, padT + plotH / 2); ctx.rotate(Math.PI / 2);
    ctx.fillStyle = MR_RED; ctx.font = 'bold 15px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center'; ctx.fillText('Доля дефектных листов, %', 0, 0);
    ctx.restore();

    const slotW = plotW / grpRows.length;
    const barW = Math.min(130, slotW * 0.54);

    grpRows.forEach((g, idx) => {
        const cx = padL + slotW * idx + slotW / 2;
        const bh = (g.checked / y1Max) * plotH;
        ctx.fillStyle = g.group.includes('ТУ-Р') ? MR_BLUE : MR_LBLUE;
        ctx.fillRect(cx - barW / 2, padT + plotH - bh, barW, bh);

        ctx.fillStyle = '#111827';
        ctx.font = 'bold 15px "Segoe UI", Arial, sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'bottom';
        ctx.fillText(mrFmtInt(g.checked), cx, padT + plotH - bh - 6);

        const parts = g.group.split(' ТУ');
        ctx.fillStyle = '#1F2937';
        ctx.font = 'bold 15px "Segoe UI", Arial, sans-serif';
        ctx.textBaseline = 'top';
        ctx.fillText(parts[0], cx, padT + plotH + 12);
        if (parts[1]) ctx.fillText('ТУ' + parts[1], cx, padT + plotH + 34);
    });

    // Красная пунктирная линия дефектности
    ctx.save();
    ctx.strokeStyle = MR_RED;
    ctx.lineWidth = 2.8;
    ctx.setLineDash([8, 6]);
    ctx.beginPath();
    grpRows.forEach((g, idx) => {
        const cx = padL + slotW * idx + slotW / 2;
        const cy = padT + plotH - (g.defRate / y2Max) * plotH;
        if (idx === 0) ctx.moveTo(cx, cy);
        else ctx.lineTo(cx, cy);
    });
    ctx.stroke();
    ctx.restore();

    grpRows.forEach((g, idx) => {
        const cx = padL + slotW * idx + slotW / 2;
        const cy = padT + plotH - (g.defRate / y2Max) * plotH;
        ctx.fillStyle = MR_RED;
        ctx.beginPath(); ctx.arc(cx, cy, 7, 0, Math.PI * 2); ctx.fill();

        ctx.font = 'bold 15px "Segoe UI", Arial, sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'bottom';
        ctx.fillText(`${g.defRate.toFixed(2)}%`, cx, cy - 11);
    });

    ctx.strokeStyle = '#374151'; ctx.lineWidth = 1.5;
    ctx.strokeRect(padL, padT, plotW, plotH);
    return canvas.toDataURL('image/png');
}

// График 2: 3 бублика природы дефектов (Всего, 0,9 мм, 1,41 мм)
function mrDrawChart2Donuts(records) {
    const configs = [
        {
            title: 'Всего',
            tech: records.reduce((a, r) => a + r.techDef, 0),
            nat: records.reduce((a, r) => a + r.natDef, 0)
        },
        {
            title: '0,9 мм',
            tech: records.filter(r => r.veneer === '0,9').reduce((a, r) => a + r.techDef, 0),
            nat: records.filter(r => r.veneer === '0,9').reduce((a, r) => a + r.natDef, 0)
        },
        {
            title: '1,41 мм',
            tech: records.filter(r => r.veneer === '1,41').reduce((a, r) => a + r.techDef, 0),
            nat: records.filter(r => r.veneer === '1,41').reduce((a, r) => a + r.natDef, 0)
        }
    ];

    const W = 1200, H = 390;
    const { canvas, ctx } = plyCreateCanvas(W, H);
    const colW = W / 3;

    configs.forEach((cfg, idx) => {
        const cx = colW * idx + colW / 2;
        const cy = 225;
        const rOut = 108, rIn = 62;
        const tot = cfg.tech + cfg.nat;

        ctx.fillStyle = '#111827';
        ctx.font = 'bold 19px "Segoe UI", Arial, sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(cfg.title, cx, 36);
        ctx.fillText(`${mrFmtInt(tot)} деф.`, cx, 62);

        if (tot <= 0) {
            ctx.beginPath();
            ctx.arc(cx, cy, rOut, 0, Math.PI * 2);
            ctx.arc(cx, cy, rIn, 0, Math.PI * 2, true);
            ctx.fillStyle = '#E5E7EB';
            ctx.fill();
            ctx.fillStyle = '#6B7280';
            ctx.font = '15px "Segoe UI", Arial, sans-serif';
            ctx.fillText('нет данных', cx, cy);
            return;
        }

        const techAngle = (cfg.tech / tot) * Math.PI * 2;
        const startA = -Math.PI / 2;

        // Сектор Технологические (влево, против часовой — как в matplotlib)
        const slices = [
            { val: cfg.tech, pct: (cfg.tech / tot) * 100, color: MR_BLUE, label: ['Техно-', 'логические'], a0: startA - techAngle, a1: startA },
            { val: cfg.nat,  pct: (cfg.nat / tot) * 100,  color: MR_GREEN, label: ['Природные'], a0: startA, a1: startA + (Math.PI * 2 - techAngle) }
        ];

        slices.forEach(sl => {
            if (sl.val <= 0) return;
            ctx.beginPath();
            ctx.arc(cx, cy, rOut, sl.a0, sl.a1);
            ctx.arc(cx, cy, rIn, sl.a1, sl.a0, true);
            ctx.closePath();
            ctx.fillStyle = sl.color;
            ctx.fill();
            ctx.strokeStyle = '#FFFFFF';
            ctx.lineWidth = 2.5;
            ctx.stroke();

            const midA = (sl.a0 + sl.a1) / 2;
            const rPct = (rOut + rIn) / 2;
            const px = cx + Math.cos(midA) * rPct;
            const py = cy + Math.sin(midA) * rPct;
            ctx.fillStyle = '#FFFFFF';
            ctx.font = 'bold 14px "Segoe UI", Arial, sans-serif';
            ctx.textAlign = 'center';
            ctx.textBaseline = 'middle';
            ctx.fillText(`${sl.pct.toFixed(1)}%`, px, py);

            const lx = cx + Math.cos(midA) * (rOut + 16);
            const ly = cy + Math.sin(midA) * (rOut + 16);
            ctx.fillStyle = '#111827';
            ctx.font = '14px "Segoe UI", Arial, sans-serif';
            ctx.textAlign = Math.cos(midA) >= 0 ? 'left' : 'right';
            sl.label.forEach((line, li) => {
                ctx.fillText(line, lx, ly + (li - (sl.label.length - 1) / 2) * 17);
            });
        });
    });

    return canvas.toDataURL('image/png');
}

// Графики 3а / 3б: Парето-анализ (синий = технолог., зелёный = природн., красная кривая)
function mrDrawChart3Pareto(pRows, veneerTitle) {
    if (!pRows || pRows.length === 0) return null;
    const W = 1320, H = 580;
    const { canvas, ctx } = plyCreateCanvas(W, H);
    const padL = 86, padR = 90, padT = 68, padB = 152;
    const plotW = W - padL - padR, plotH = H - padT - padB;

    const n80 = pRows.filter(r => r.cumShare < 80).length + (pRows.some(r => r.cumShare >= 80) ? 1 : 0);
    const maxCnt = Math.max(...pRows.map(r => r.count), 1);
    const yMax = plyNiceMax(maxCnt * 1.18, 5);

    ctx.fillStyle = '#111827';
    ctx.font = 'bold 20px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(`Парето-анализ дефектов (${veneerTitle}): 80 % отклонений дают ${Math.min(n80, pRows.length)} видов из ${pRows.length}`, W / 2, 32);

    ctx.font = '14px "Segoe UI", Arial, sans-serif';
    for (let i = 0; i <= 5; i++) {
        const val = (yMax / 5) * i;
        const y = padT + plotH - (val / yMax) * plotH;
        ctx.save();
        ctx.strokeStyle = i === 0 ? '#374151' : 'rgba(156,163,175,0.45)';
        ctx.lineWidth = i === 0 ? 1.5 : 1;
        if (i > 0) ctx.setLineDash([3, 4]);
        ctx.beginPath(); ctx.moveTo(padL, y); ctx.lineTo(padL + plotW, y); ctx.stroke();
        ctx.restore();

        ctx.fillStyle = '#374151';
        ctx.textAlign = 'right';
        ctx.fillText(Math.round(val).toString(), padL - 10, y);
    }

    [0, 20, 40, 60, 80, 100].forEach(p => {
        const y = padT + plotH - (p / 105) * plotH;
        ctx.fillStyle = MR_RED;
        ctx.textAlign = 'left';
        ctx.fillText(String(p), padL + plotW + 10, y);
    });

    ctx.save();
    ctx.translate(24, padT + plotH / 2); ctx.rotate(-Math.PI / 2);
    ctx.fillStyle = '#111827'; ctx.font = 'bold 15px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center'; ctx.fillText('Количество дефектов, шт.', 0, 0);
    ctx.restore();

    ctx.save();
    ctx.translate(W - 22, padT + plotH / 2); ctx.rotate(Math.PI / 2);
    ctx.fillStyle = MR_RED; ctx.font = 'bold 15px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center'; ctx.fillText('Накопленная доля, %', 0, 0);
    ctx.restore();

    const slotW = plotW / pRows.length;
    const barW = Math.min(72, slotW * 0.68);

    pRows.forEach((r, idx) => {
        const cx = padL + slotW * idx + slotW / 2;
        const bh = (r.count / yMax) * plotH;
        ctx.fillStyle = r.nature === 'природный' ? MR_GREEN : MR_BLUE;
        ctx.fillRect(cx - barW / 2, padT + plotH - bh, barW, bh);

        ctx.fillStyle = '#111827';
        ctx.font = '13px "Segoe UI", Arial, sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'bottom';
        ctx.fillText(String(Math.round(r.count)), cx, padT + plotH - bh - 4);

        ctx.save();
        ctx.translate(cx, padT + plotH + 10);
        ctx.rotate(-35 * Math.PI / 180);
        ctx.fillStyle = '#1F2937';
        ctx.font = '14px "Segoe UI", Arial, sans-serif';
        ctx.textAlign = 'right';
        ctx.textBaseline = 'middle';
        ctx.fillText(r.defect, 0, 0);
        ctx.restore();
    });

    // Линия 80%
    const y80 = padT + plotH - (80 / 105) * plotH;
    ctx.save();
    ctx.strokeStyle = '#6B7280'; ctx.lineWidth = 1.4; ctx.setLineDash([4, 4]);
    ctx.beginPath(); ctx.moveTo(padL, y80); ctx.lineTo(padL + plotW, y80); ctx.stroke();
    ctx.restore();

    // Кривая Парето
    ctx.strokeStyle = MR_RED; ctx.lineWidth = 2.6;
    ctx.beginPath();
    pRows.forEach((r, idx) => {
        const cx = padL + slotW * idx + slotW / 2;
        const cy = padT + plotH - (r.cumShare / 105) * plotH;
        if (idx === 0) ctx.moveTo(cx, cy); else ctx.lineTo(cx, cy);
    });
    ctx.stroke();

    pRows.forEach((r, idx) => {
        const cx = padL + slotW * idx + slotW / 2;
        const cy = padT + plotH - (r.cumShare / 105) * plotH;
        ctx.fillStyle = MR_RED;
        ctx.beginPath(); ctx.arc(cx, cy, 4.5, 0, Math.PI * 2); ctx.fill();
    });

    ctx.strokeStyle = '#374151'; ctx.lineWidth = 1.5;
    ctx.strokeRect(padL, padT, plotW, plotH);

    // Легенда
    const legW = 225, legH = 62;
    const legX = padL + plotW - legW - 12, legY = padT + 12;
    ctx.fillStyle = 'rgba(255,255,255,0.92)'; ctx.fillRect(legX, legY, legW, legH);
    ctx.strokeStyle = '#D1D5DB'; ctx.strokeRect(legX, legY, legW, legH);
    ctx.textAlign = 'left'; ctx.textBaseline = 'middle'; ctx.font = '14px "Segoe UI", Arial, sans-serif';
    ctx.fillStyle = MR_BLUE; ctx.fillRect(legX + 12, legY + 13, 22, 13);
    ctx.fillStyle = '#111827'; ctx.fillText('Технологические', legX + 42, legY + 19);
    ctx.fillStyle = MR_GREEN; ctx.fillRect(legX + 12, legY + 37, 22, 13);
    ctx.fillStyle = '#111827'; ctx.fillText('Природные (шпон)', legX + 42, legY + 43);

    return canvas.toDataURL('image/png');
}

// График 4: Профиль ТОП-10 дефектов (0,9 мм против 1,41 мм на 1000 л.)
function mrDrawChart4Profile(records, paretoTotal) {
    const top10 = paretoTotal.slice(0, 10).map(r => r.defect);
    if (top10.length === 0) return null;

    const c09 = records.filter(r => r.veneer === '0,9').reduce((a, r) => a + r.checked, 0);
    const c141 = records.filter(r => r.veneer === '1,41').reduce((a, r) => a + r.checked, 0);

    const rows = top10.map(dname => {
        const cnt09 = records.filter(r => r.veneer === '0,9').reduce((a, r) => a + (r.defects[dname] || 0), 0);
        const cnt141 = records.filter(r => r.veneer === '1,41').reduce((a, r) => a + (r.defects[dname] || 0), 0);
        return {
            defect: dname,
            r09: c09 > 0 ? (cnt09 / c09) * 1000 : 0,
            r141: c141 > 0 ? (cnt141 / c141) * 1000 : 0
        };
    });

    const W = 1220, H = Math.max(460, top10.length * 46 + 130);
    const { canvas, ctx } = plyCreateCanvas(W, H);
    const padL = 220, padR = 70, padT = 66, padB = 66;
    const plotW = W - padL - padR, plotH = H - padT - padB;

    const maxRate = Math.max(...rows.map(r => Math.max(r.r09, r.r141)), 1);
    const xMax = plyNiceMax(maxRate * 1.15, 5);

    ctx.fillStyle = '#111827';
    ctx.font = 'bold 20px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText('Профиль ТОП-10 дефектов: 0,9 мм против 1,41 мм (норм. на 1000 л.)', W / 2, 30);

    ctx.font = '14px "Segoe UI", Arial, sans-serif';
    for (let i = 0; i <= 5; i++) {
        const val = (xMax / 5) * i;
        const x = padL + (val / xMax) * plotW;
        ctx.save();
        ctx.strokeStyle = i === 0 ? '#374151' : 'rgba(156,163,175,0.45)';
        ctx.lineWidth = i === 0 ? 1.5 : 1;
        if (i > 0) ctx.setLineDash([3, 4]);
        ctx.beginPath(); ctx.moveTo(x, padT); ctx.lineTo(x, padT + plotH); ctx.stroke();
        ctx.restore();

        ctx.fillStyle = '#374151'; ctx.textAlign = 'center'; ctx.textBaseline = 'top';
        ctx.fillText(val.toFixed(1).replace(/\.0$/, ''), x, padT + plotH + 8);
    }

    ctx.fillStyle = '#111827'; ctx.font = 'bold 15px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center'; ctx.fillText('Дефектов на 1000 проверенных листов', padL + plotW / 2, H - 22);

    const slotH = plotH / rows.length;
    const barH = Math.min(18, slotH * 0.38);

    rows.forEach((r, idx) => {
        const cy = padT + slotH * idx + slotH / 2;
        const w09 = (r.r09 / xMax) * plotW;
        const w141 = (r.r141 / xMax) * plotW;

        ctx.fillStyle = MR_BLUE;
        ctx.fillRect(padL, cy - barH - 1, w09, barH);
        ctx.fillStyle = MR_ORANGE;
        ctx.fillRect(padL, cy + 1, w141, barH);

        ctx.fillStyle = '#1F2937'; ctx.font = 'bold 14px "Segoe UI", Arial, sans-serif';
        ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
        ctx.fillText(r.defect, padL - 12, cy);

        ctx.fillStyle = '#111827'; ctx.font = '13px "Segoe UI", Arial, sans-serif';
        ctx.textAlign = 'left';
        ctx.fillText(r.r09.toFixed(1), padL + w09 + 6, cy - barH / 2 - 1);
        ctx.fillText(r.r141.toFixed(1), padL + w141 + 6, cy + barH / 2 + 1);
    });

    ctx.strokeStyle = '#374151'; ctx.lineWidth = 1.5;
    ctx.strokeRect(padL, padT, plotW, plotH);

    // Легенда
    const legW = 140, legH = 60;
    const legX = padL + plotW - legW - 12, legY = padT + plotH - legH - 12;
    ctx.fillStyle = 'rgba(255,255,255,0.94)'; ctx.fillRect(legX, legY, legW, legH);
    ctx.strokeStyle = '#D1D5DB'; ctx.strokeRect(legX, legY, legW, legH);
    ctx.textAlign = 'left'; ctx.textBaseline = 'middle'; ctx.font = '14px "Segoe UI", Arial, sans-serif';
    ctx.fillStyle = MR_BLUE; ctx.fillRect(legX + 12, legY + 12, 22, 13);
    ctx.fillStyle = '#111827'; ctx.fillText('0,9 мм', legX + 42, legY + 19);
    ctx.fillStyle = MR_ORANGE; ctx.fillRect(legX + 12, legY + 36, 22, 13);
    ctx.fillStyle = '#111827'; ctx.fillText('1,41 мм', legX + 42, legY + 42);

    return canvas.toDataURL('image/png');
}

// График 5: Влияние толщины плиты
function mrDrawChart5Thickness(records, minChecked = 1) {
    const map = new Map();
    records.forEach(r => {
        if (!r.thickness || r.thickness <= 0) return;
        if (!map.has(r.thickness)) {
            map.set(r.thickness, { thickness: r.thickness, checked: 0, techDef: 0, natDef: 0, totalDef: 0 });
        }
        const g = map.get(r.thickness);
        g.checked += r.checked;
        g.techDef += r.techDef;
        g.natDef += r.natDef;
        g.totalDef += r.totalDef;
    });

    let list = Array.from(map.values()).filter(g => g.checked >= minChecked);
    if (list.length === 0) list = Array.from(map.values());
    list.sort((a, b) => a.thickness - b.thickness);
    if (list.length === 0) return null;

    list.forEach(g => {
        g.techRate = g.checked > 0 ? (g.techDef / g.checked) * 100 : 0;
        g.natRate = g.checked > 0 ? (g.natDef / g.checked) * 100 : 0;
        g.totalRate = g.checked > 0 ? (g.totalDef / g.checked) * 100 : 0;
    });

    const W = 1220, H = 520;
    const { canvas, ctx } = plyCreateCanvas(W, H);
    const padL = 82, padR = 40, padT = 68, padB = 74;
    const plotW = W - padL - padR, plotH = H - padT - padB;

    const maxRate = Math.max(...list.map(g => g.totalRate), 1);
    const yMax = plyNiceMax(maxRate * 1.32, 5);

    const titleSuffix = minChecked > 1 ? ` (партии от ${minChecked} листов)` : '';
    ctx.fillStyle = '#111827';
    ctx.font = 'bold 20px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(`Уровень дефектности по толщине плиты${titleSuffix}`, W / 2, 30);

    ctx.font = '14px "Segoe UI", Arial, sans-serif';
    for (let i = 0; i <= 5; i++) {
        const val = (yMax / 5) * i;
        const y = padT + plotH - (val / yMax) * plotH;
        ctx.save();
        ctx.strokeStyle = i === 0 ? '#374151' : 'rgba(156,163,175,0.45)';
        ctx.lineWidth = i === 0 ? 1.5 : 1;
        if (i > 0) ctx.setLineDash([3, 4]);
        ctx.beginPath(); ctx.moveTo(padL, y); ctx.lineTo(padL + plotW, y); ctx.stroke();
        ctx.restore();

        ctx.fillStyle = '#374151'; ctx.textAlign = 'right';
        ctx.fillText(val.toFixed(1).replace(/\.0$/, ''), padL - 10, y);
    }

    ctx.save();
    ctx.translate(24, padT + plotH / 2); ctx.rotate(-Math.PI / 2);
    ctx.fillStyle = '#111827'; ctx.font = 'bold 15px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center'; ctx.fillText('Доля дефектных листов, %', 0, 0);
    ctx.restore();

    ctx.fillStyle = '#111827'; ctx.font = 'bold 15px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center'; ctx.fillText('Толщина плиты, мм', padL + plotW / 2, H - 20);

    const slotW = plotW / list.length;
    const barW = Math.min(72, slotW * 0.58);

    list.forEach((g, idx) => {
        const cx = padL + slotW * idx + slotW / 2;
        const hT = (g.techRate / yMax) * plotH;
        const hN = (g.natRate / yMax) * plotH;

        ctx.fillStyle = MR_BLUE;
        ctx.fillRect(cx - barW / 2, padT + plotH - hT, barW, hT);
        ctx.fillStyle = MR_GREEN;
        ctx.fillRect(cx - barW / 2, padT + plotH - hT - hN, barW, hN);

        const topY = padT + plotH - hT - hN;
        const volStr = g.checked >= 1000 ? `${(g.checked / 1000).toFixed(1)} тыс.л.` : `${Math.round(g.checked)} л.`;
        ctx.fillStyle = '#111827'; ctx.font = '13px "Segoe UI", Arial, sans-serif';
        ctx.textAlign = 'center'; ctx.textBaseline = 'bottom';
        ctx.fillText(`${g.totalRate.toFixed(1)}%`, cx, topY - 18);
        ctx.fillText(volStr, cx, topY - 3);

        ctx.fillStyle = '#1F2937'; ctx.font = 'bold 15px "Segoe UI", Arial, sans-serif';
        ctx.textBaseline = 'top';
        ctx.fillText(mrFormatThickness(g.thickness), cx, padT + plotH + 10);
    });

    ctx.strokeStyle = '#374151'; ctx.lineWidth = 1.5;
    ctx.strokeRect(padL, padT, plotW, plotH);

    // Легенда сверху справа
    const legW = 310, legH = 34;
    const legX = padL + plotW - legW - 12, legY = padT + 10;
    ctx.fillStyle = 'rgba(255,255,255,0.92)'; ctx.fillRect(legX, legY, legW, legH);
    ctx.strokeStyle = '#D1D5DB'; ctx.strokeRect(legX, legY, legW, legH);
    ctx.textAlign = 'left'; ctx.textBaseline = 'middle'; ctx.font = '14px "Segoe UI", Arial, sans-serif';
    ctx.fillStyle = MR_BLUE; ctx.fillRect(legX + 10, legY + 10, 20, 14);
    ctx.fillStyle = '#111827'; ctx.fillText('технологические', legX + 36, legY + 17);
    ctx.fillStyle = MR_GREEN; ctx.fillRect(legX + 175, legY + 10, 20, 14);
    ctx.fillStyle = '#111827'; ctx.fillText('природные', legX + 201, legY + 17);

    return canvas.toDataURL('image/png');
}

// График 6: ТОП-12 проблемных позиций
function mrDrawChart6TopBad(posFiltered, avgDefRate, minChecked = 1) {
    const top12 = posFiltered.slice(0, 12);
    if (top12.length === 0) return null;

    const W = 1220, H = Math.max(400, top12.length * 42 + 130);
    const { canvas, ctx } = plyCreateCanvas(W, H);
    const padL = 215, padR = 55, padT = 66, padB = 68;
    const plotW = W - padL - padR, plotH = H - padT - padB;

    const maxRate = Math.max(...top12.map(r => r.defRate), avgDefRate, 1);
    const xMax = plyNiceMax(maxRate * 1.32, 5);

    const titleSuffix = minChecked > 1 ? ` (партии от ${minChecked} листов)` : '';
    ctx.fillStyle = '#111827';
    ctx.font = 'bold 20px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(`ТОП-${top12.length} проблемных позиций${titleSuffix}`, W / 2, 30);

    ctx.font = '14px "Segoe UI", Arial, sans-serif';
    for (let i = 0; i <= 5; i++) {
        const val = (xMax / 5) * i;
        const x = padL + (val / xMax) * plotW;
        ctx.save();
        ctx.strokeStyle = i === 0 ? '#374151' : 'rgba(156,163,175,0.45)';
        ctx.lineWidth = i === 0 ? 1.5 : 1;
        if (i > 0) ctx.setLineDash([3, 4]);
        ctx.beginPath(); ctx.moveTo(x, padT); ctx.lineTo(x, padT + plotH); ctx.stroke();
        ctx.restore();

        ctx.fillStyle = '#374151'; ctx.textAlign = 'center'; ctx.textBaseline = 'top';
        ctx.fillText(val.toFixed(1).replace(/\.0$/, ''), x, padT + plotH + 8);
    }

    ctx.fillStyle = '#111827'; ctx.font = 'bold 15px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center'; ctx.fillText('Доля дефектных листов, %', padL + plotW / 2, H - 20);

    const slotH = plotH / top12.length;
    const barH = Math.min(28, slotH * 0.64);

    top12.forEach((r, idx) => {
        const cy = padT + slotH * idx + slotH / 2;
        const bw = (r.defRate / xMax) * plotW;
        ctx.fillStyle = r.defRate >= 7.0 ? MR_RED : MR_ORANGE;
        ctx.fillRect(padL, cy - barH / 2, bw, barH);

        ctx.fillStyle = '#1F2937'; ctx.font = 'bold 14px "Segoe UI", Arial, sans-serif';
        ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
        ctx.fillText(r.posFull, padL - 12, cy);

        ctx.fillStyle = '#111827'; ctx.font = '14px "Segoe UI", Arial, sans-serif';
        ctx.textAlign = 'left';
        ctx.fillText(`${r.defRate.toFixed(2)} % (${Math.round(r.totalDef)} деф.)`, padL + bw + 8, cy);
    });

    // Линия среднего уровня
    const avgX = padL + (avgDefRate / xMax) * plotW;
    ctx.save();
    ctx.strokeStyle = '#333333'; ctx.lineWidth = 2; ctx.setLineDash([7, 5]);
    ctx.beginPath(); ctx.moveTo(avgX, padT); ctx.lineTo(avgX, padT + plotH); ctx.stroke();
    ctx.restore();

    ctx.fillStyle = '#333333'; ctx.font = 'bold 14px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'left'; ctx.textBaseline = 'bottom';
    ctx.fillText(` средний уровень ${avgDefRate.toFixed(2)} %`, avgX + 4, padT + plotH - 6);

    ctx.strokeStyle = '#374151'; ctx.lineWidth = 1.5;
    ctx.strokeRect(padL, padT, plotW, plotH);
    return canvas.toDataURL('image/png');
}

// Цветовая карта YlOrRd для тепловой карты
function mrYlOrRdColor(t) {
    const clamped = Math.max(0, Math.min(1, t));
    const stops = [
        [255, 255, 204],
        [254, 217, 118],
        [253, 141, 60],
        [227, 26, 28],
        [128, 0, 38]
    ];
    const scaled = clamped * (stops.length - 1);
    const i = Math.min(stops.length - 2, Math.floor(scaled));
    const f = scaled - i;
    const r = Math.round(stops[i][0] + (stops[i + 1][0] - stops[i][0]) * f);
    const g = Math.round(stops[i][1] + (stops[i + 1][1] - stops[i][1]) * f);
    const b = Math.round(stops[i][2] + (stops[i + 1][2] - stops[i][2]) * f);
    return `rgb(${r},${g},${b})`;
}

// График 7: Карта дефектности крупнейших партий × ТОП-8 дефектов
function mrDrawChart7Heatmap(records, paretoTotal) {
    const top12Vol = [...records].sort((a, b) => b.checked - a.checked).slice(0, 12);
    const top8Defs = paretoTotal.slice(0, 8).map(r => r.defect);
    if (top12Vol.length === 0 || top8Defs.length === 0) return null;

    const matrix = top12Vol.map(r =>
        top8Defs.map(d => (r.checked > 0 ? ((r.defects[d] || 0) / r.checked) * 1000 : 0))
    );
    const maxVal = Math.max(...matrix.flat(), 1);

    const W = 1220, H = Math.max(430, top12Vol.length * 42 + 210);
    const { canvas, ctx } = plyCreateCanvas(W, H);
    const padL = 195, padR = 135, padT = 66, padB = 135;
    const plotW = W - padL - padR, plotH = H - padT - padB;

    ctx.fillStyle = '#111827';
    ctx.font = 'bold 20px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(`Карта дефектности: ${top12Vol.length} крупнейших партий × ТОП-${top8Defs.length} дефектов`, W / 2, 30);

    const cellW = plotW / top8Defs.length;
    const cellH = plotH / top12Vol.length;

    top12Vol.forEach((r, i) => {
        const y = padT + i * cellH;
        ctx.fillStyle = '#1F2937'; ctx.font = 'bold 14px "Segoe UI", Arial, sans-serif';
        ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
        ctx.fillText(r.posShort, padL - 12, y + cellH / 2);

        top8Defs.forEach((_, j) => {
            const x = padL + j * cellW;
            const val = matrix[i][j];
            const ratio = val / maxVal;
            ctx.fillStyle = mrYlOrRdColor(ratio);
            ctx.fillRect(x, y, cellW, cellH);

            if (val > 0.05) {
                ctx.fillStyle = ratio > 0.55 ? '#FFFFFF' : '#222222';
                ctx.font = 'bold 13px "Segoe UI", Arial, sans-serif';
                ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
                ctx.fillText(val.toFixed(1), x + cellW / 2, y + cellH / 2);
            }
        });
    });

    top8Defs.forEach((dname, j) => {
        const cx = padL + j * cellW + cellW / 2;
        ctx.save();
        ctx.translate(cx, padT + plotH + 12);
        ctx.rotate(-30 * Math.PI / 180);
        ctx.fillStyle = '#1F2937'; ctx.font = '14px "Segoe UI", Arial, sans-serif';
        ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
        ctx.fillText(dname, 0, 0);
        ctx.restore();
    });

    ctx.strokeStyle = '#374151'; ctx.lineWidth = 1.5;
    ctx.strokeRect(padL, padT, plotW, plotH);

    // Шкала colorbar справа
    const cbX = padL + plotW + 24, cbW = 24;
    for (let py = 0; py < plotH; py++) {
        const t = 1 - py / plotH;
        ctx.fillStyle = mrYlOrRdColor(t);
        ctx.fillRect(cbX, padT + py, cbW, 1.5);
    }
    ctx.strokeStyle = '#374151'; ctx.lineWidth = 1;
    ctx.strokeRect(cbX, padT, cbW, plotH);

    ctx.fillStyle = '#374151'; ctx.font = '13px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'left'; ctx.textBaseline = 'middle';
    for (let i = 0; i <= 4; i++) {
        const v = (maxVal / 4) * i;
        const y = padT + plotH - (i / 4) * plotH;
        ctx.fillText(v.toFixed(0), cbX + cbW + 8, y);
    }

    ctx.save();
    ctx.translate(W - 24, padT + plotH / 2); ctx.rotate(Math.PI / 2);
    ctx.fillStyle = '#111827'; ctx.font = 'bold 14px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center'; ctx.fillText('дефектов на 1000 листов', 0, 0);
    ctx.restore();

    return canvas.toDataURL('image/png');
}

// График 8: Распределение позиций и объёмов по уровню качества (% годных)
function mrDrawChart8QualityDist(records) {
    const bins = [
        { label: '< 92 %',    min: -Infinity, max: 92 },
        { label: '92–94 %',   min: 92,        max: 94 },
        { label: '94–96 %',   min: 94,        max: 96 },
        { label: '96–98 %',   min: 96,        max: 98 },
        { label: '98–100 %',  min: 98,        max: Infinity }
    ];

    const dist = bins.map(b => {
        const sub = records.filter(r => r.qualRate >= b.min && r.qualRate < b.max);
        return {
            label: b.label,
            cnt: sub.length,
            vol: sub.reduce((a, r) => a + r.checked, 0)
        };
    });

    const maxVolRaw = Math.max(...dist.map(d => d.vol), 1);
    const useThousands = maxVolRaw >= 1000;
    const volValues = dist.map(d => useThousands ? d.vol / 1000 : d.vol);
    const maxCnt = Math.max(...dist.map(d => d.cnt), 1);
    const y1Max = plyNiceMax(maxCnt * 1.24, 4);
    const y2Max = plyNiceMax(Math.max(...volValues, 1) * 1.24, useThousands ? 1 : 50);

    const W = 1220, H = 470;
    const { canvas, ctx } = plyCreateCanvas(W, H);
    const padL = 84, padR = 95, padT = 66, padB = 64;
    const plotW = W - padL - padR, plotH = H - padT - padB;

    ctx.fillStyle = '#111827';
    ctx.font = 'bold 20px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText('Распределение позиций и объёмов по уровню качества (% годных)', W / 2, 30);

    ctx.font = '14px "Segoe UI", Arial, sans-serif';
    for (let i = 0; i <= 5; i++) {
        const v1 = (y1Max / 5) * i;
        const v2 = (y2Max / 5) * i;
        const y = padT + plotH - (i / 5) * plotH;
        ctx.save();
        ctx.strokeStyle = i === 0 ? '#374151' : 'rgba(156,163,175,0.45)';
        ctx.lineWidth = i === 0 ? 1.5 : 1;
        if (i > 0) ctx.setLineDash([3, 4]);
        ctx.beginPath(); ctx.moveTo(padL, y); ctx.lineTo(padL + plotW, y); ctx.stroke();
        ctx.restore();

        ctx.fillStyle = '#374151'; ctx.textAlign = 'right';
        ctx.fillText(v1.toFixed(1).replace(/\.0$/, ''), padL - 10, y);

        ctx.fillStyle = '#374151'; ctx.textAlign = 'left';
        ctx.fillText(useThousands ? v2.toFixed(1).replace(/\.0$/, '') : Math.round(v2).toString(), padL + plotW + 10, y);
    }

    ctx.save();
    ctx.translate(24, padT + plotH / 2); ctx.rotate(-Math.PI / 2);
    ctx.fillStyle = '#111827'; ctx.font = 'bold 15px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center'; ctx.fillText('Позиций, шт.', 0, 0);
    ctx.restore();

    ctx.save();
    ctx.translate(W - 22, padT + plotH / 2); ctx.rotate(Math.PI / 2);
    ctx.fillStyle = '#111827'; ctx.font = 'bold 15px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center'; ctx.fillText(useThousands ? 'Объём, тыс. листов' : 'Объём, листов', 0, 0);
    ctx.restore();

    const slotW = plotW / dist.length;
    const barW = Math.min(72, slotW * 0.36);

    dist.forEach((d, idx) => {
        const cx = padL + slotW * idx + slotW / 2;
        const h1 = (d.cnt / y1Max) * plotH;
        const vVal = volValues[idx];
        const h2 = (vVal / y2Max) * plotH;

        ctx.fillStyle = MR_BLUE;
        ctx.fillRect(cx - barW - 2, padT + plotH - h1, barW, h1);
        ctx.fillStyle = MR_ORANGE;
        ctx.fillRect(cx + 2, padT + plotH - h2, barW, h2);

        ctx.fillStyle = '#111827'; ctx.font = 'bold 14px "Segoe UI", Arial, sans-serif';
        ctx.textAlign = 'center'; ctx.textBaseline = 'bottom';
        ctx.fillText(String(d.cnt), cx - barW / 2 - 2, padT + plotH - h1 - 5);
        ctx.fillText(useThousands ? vVal.toFixed(1) : String(Math.round(vVal)), cx + barW / 2 + 2, padT + plotH - h2 - 5);

        ctx.fillStyle = '#1F2937'; ctx.font = 'bold 15px "Segoe UI", Arial, sans-serif';
        ctx.textBaseline = 'top';
        ctx.fillText(d.label, cx, padT + plotH + 12);
    });

    ctx.strokeStyle = '#374151'; ctx.lineWidth = 1.5;
    ctx.strokeRect(padL, padT, plotW, plotH);

    // Легенда
    const legW = 225, legH = 60;
    const legX = padL + 14, legY = padT + 12;
    ctx.fillStyle = 'rgba(255,255,255,0.92)'; ctx.fillRect(legX, legY, legW, legH);
    ctx.strokeStyle = '#D1D5DB'; ctx.strokeRect(legX, legY, legW, legH);
    ctx.textAlign = 'left'; ctx.textBaseline = 'middle'; ctx.font = '14px "Segoe UI", Arial, sans-serif';
    ctx.fillStyle = MR_BLUE; ctx.fillRect(legX + 12, legY + 12, 22, 13);
    ctx.fillStyle = '#111827'; ctx.fillText('число позиций', legX + 42, legY + 19);
    ctx.fillStyle = MR_ORANGE; ctx.fillRect(legX + 12, legY + 36, 22, 13);
    ctx.fillStyle = '#111827'; ctx.fillText(useThousands ? 'объём, тыс. листов' : 'объём, листов', legX + 42, legY + 42);

    return canvas.toDataURL('image/png');
}

// ------------------- 3. СБОРКА HTML-СТРАНИЦ И ГЕНЕРАЦИЯ PDF -------------------

async function generateAnalyticalPdfReport(cycles, options = {}) {
    if (!window.jspdf || !window.jspdf.jsPDF || typeof window.html2canvas !== 'function') {
        throw new Error('Не загружены библиотеки jsPDF / html2canvas');
    }

    const records = mrPrepareRecords(cycles);
    if (records.length === 0) return null;

    const natSet = mrGetNaturalDefectNames(cycles);
    const isMonthly = Boolean(options.isMonthly);
    const periodTitle = options.periodTitle || 'Смена';
    const operatorStr = options.operator || '';
    const sourceLabel = options.sourceLabel || '';
    const minThicknessChk = isMonthly ? 300 : 1;
    const minProbChk = isMonthly ? 500 : 1;

    const totalChecked = records.reduce((a, r) => a + r.checked, 0);
    const totalDef = records.reduce((a, r) => a + r.totalDef, 0);
    const totalTech = records.reduce((a, r) => a + r.techDef, 0);
    const totalNat = records.reduce((a, r) => a + r.natDef, 0);
    const defRate = totalChecked > 0 ? (totalDef / totalChecked) * 100 : 0;
    const qualRate = 100 - defRate;
    const natShare = totalDef > 0 ? (totalNat / totalDef) * 100 : 0;
    const techShare = totalDef > 0 ? (totalTech / totalDef) * 100 : 0;

    // Группы
    const groupsOrder = ['0,9 мм ТУ-Р', '0,9 мм ТУ-Э', '1,41 мм ТУ-Р', '1,41 мм ТУ-Э'];
    const grpRows = groupsOrder.map(g => {
        const sub = records.filter(r => r.group === g);
        const c = sub.reduce((a, r) => a + r.checked, 0);
        const d = sub.reduce((a, r) => a + r.totalDef, 0);
        const t = sub.reduce((a, r) => a + r.techDef, 0);
        const n = sub.reduce((a, r) => a + r.natDef, 0);
        const rPct = c > 0 ? (d / c) * 100 : 0;
        return { group: g, posCnt: sub.length, checked: c, totalDef: d, techDef: t, natDef: n, defRate: rPct, qualRate: c > 0 ? 100 - rPct : 0 };
    });

    // Парето (общий + 0,9 + 1,41)
    const paretoTotal = mrBuildParetoData(records, natSet);
    const sub09 = records.filter(r => r.veneer === '0,9');
    const sub141 = records.filter(r => r.veneer === '1,41');
    const pareto09 = mrBuildParetoData(sub09, natSet);
    const pareto141 = mrBuildParetoData(sub141, natSet);

    // Фильтр для проблемных и лучших
    let posFiltered = records.filter(r => r.checked >= minProbChk);
    let actualMinProb = minProbChk;
    if (posFiltered.length === 0) {
        posFiltered = [...records];
        actualMinProb = 1;
    }
    posFiltered.sort((a, b) => b.defRate - a.defRate);
    const worst8 = posFiltered.slice(0, 8);
    const best5 = [...posFiltered].sort((a, b) => a.defRate - b.defRate).slice(0, 5);

    // Отрисовка графиков
    const img1 = mrDrawChart1Groups(grpRows);
    const img2 = mrDrawChart2Donuts(records);
    const img3a = pareto09.length > 0 ? mrDrawChart3Pareto(pareto09, 'рубашка 0,9 мм') : null;
    const img3b = pareto141.length > 0 ? mrDrawChart3Pareto(pareto141, 'рубашка 1,41 мм') : null;
    const img4 = paretoTotal.length > 0 ? mrDrawChart4Profile(records, paretoTotal) : null;
    const img5 = mrDrawChart5Thickness(records, minThicknessChk);
    const img6 = mrDrawChart6TopBad(posFiltered, defRate, actualMinProb);
    const img7 = mrDrawChart7Heatmap(records, paretoTotal);
    const img8 = mrDrawChart8QualityDist(records);

    // Помощники для HTML-таблиц
    const tblStyle = 'width:100%;border-collapse:collapse;margin:10px 0;font-size:12px;';
    const thStyle = `background:${MR_BLUE};color:#fff;padding:6px 6px;text-align:center;border:1px solid #2A6288;`;
    const tdStyle = 'border-bottom:1px solid #DDD;padding:5px 6px;text-align:center;';

    const grpRowsHtml = grpRows.map(r => `
        <tr>
            <td style="${tdStyle}"><b>${r.group}</b></td>
            <td style="${tdStyle}">${r.posCnt}</td>
            <td style="${tdStyle}">${mrFmtInt(r.checked)}</td>
            <td style="${tdStyle}">${mrFmtInt(r.totalDef)}</td>
            <td style="${tdStyle}">${mrFmtInt(r.techDef)}</td>
            <td style="${tdStyle}">${mrFmtInt(r.natDef)}</td>
            <td style="${tdStyle}"><b>${mrFmtFlt(r.defRate)}</b></td>
            <td style="${tdStyle}">${mrFmtFlt(r.qualRate)}</td>
        </tr>
    `).join('');

    const makeParetoRowsHtml = (arr) => arr.slice(0, 10).map((r, i) => `
        <tr>
            <td style="${tdStyle}">${i + 1}</td>
            <td style="${tdStyle}">${r.defect}</td>
            <td style="${tdStyle}">${r.nature}</td>
            <td style="${tdStyle}"><b>${mrFmtInt(r.count)}</b></td>
            <td style="${tdStyle}">${r.share.toFixed(1)}</td>
            <td style="${tdStyle}">${r.cumShare.toFixed(1)}</td>
            <td style="${tdStyle}">${mrFmtFlt(r.per1000)}</td>
        </tr>
    `).join('');

    const worst8Html = worst8.map(r => {
        const entries = Object.entries(r.defects);
        const keyDef = entries.length > 0 ? entries.sort((a, b) => b[1] - a[1])[0] : ['—', 0];
        return `
            <tr>
                <td style="${tdStyle}"><b>${r.position}</b></td>
                <td style="${tdStyle}">${r.veneer} мм</td>
                <td style="${tdStyle}">${mrFmtInt(r.checked)}</td>
                <td style="${tdStyle}">${mrFmtInt(r.totalDef)}</td>
                <td style="${tdStyle}color:${MR_RED};font-weight:bold;">${mrFmtFlt(r.defRate)}</td>
                <td style="${tdStyle}">${keyDef[0]} (${keyDef[1]})</td>
            </tr>
        `;
    }).join('');

    const best5Html = best5.map(r => `
        <tr>
            <td style="${tdStyle}"><b>${r.position}</b></td>
            <td style="${tdStyle}">${r.veneer} мм</td>
            <td style="${tdStyle}">${mrFmtInt(r.checked)}</td>
            <td style="${tdStyle}">${mrFmtInt(r.totalDef)}</td>
            <td style="${tdStyle}color:#27AE60;font-weight:bold;">${mrFmtFlt(r.defRate)}</td>
        </tr>
    `).join('');

    const top5Names = paretoTotal.slice(0, 5).map(r => r.defect).join(', ') || 'нет дефектов';
    const top5Share = paretoTotal.slice(0, 5).reduce((a, r) => a + r.share, 0);

    // Сборка отдельных страниц A4
    const pagesHtml = [];
    const h2Style = `color:#1F4E79;border-bottom:2px solid ${MR_BLUE};padding-bottom:4px;margin:14px 0 8px 0;font-size:18px;font-weight:bold;`;
    const h3Style = 'color:#2C3E50;margin:14px 0 6px 0;font-size:15px;font-weight:bold;';

    // СТРАНИЦА 1: Шапка + KPI + Главное + Раздел 1
    pagesHtml.push(`
        <h1 style="color:#1F4E79;margin:0 0 4px 0;font-size:24px;font-weight:800;">Отчёт по дефектовке продукции — ${periodTitle}</h1>
        <p style="margin:0 0 12px 0;font-size:13px;color:#444;">
            Анализ результатов входного контроля шлифованной плиты: рубашка 0,9 мм и 1,41 мм, спецификации ТУ-Э и ТУ-Р${operatorStr ? ' • Оператор: <b>' + operatorStr + '</b>' : ''}
        </p>
        <div style="display:flex;gap:10px;margin:12px 0;">
            <div style="flex:1;background:#F4F8FB;border-top:4px solid ${MR_BLUE};padding:10px;text-align:center;border-radius:4px;">
                <div style="font-size:22px;font-weight:bold;color:#1F4E79;">${mrFmtInt(totalChecked)}</div>
                <div style="font-size:11px;color:#555;">Проверено листов</div>
            </div>
            <div style="flex:1;background:#F4F8FB;border-top:4px solid ${MR_BLUE};padding:10px;text-align:center;border-radius:4px;">
                <div style="font-size:22px;font-weight:bold;color:${MR_RED};">${mrFmtInt(totalDef)}</div>
                <div style="font-size:11px;color:#555;">Выявлено дефектов</div>
            </div>
            <div style="flex:1;background:#F4F8FB;border-top:4px solid ${MR_BLUE};padding:10px;text-align:center;border-radius:4px;">
                <div style="font-size:22px;font-weight:bold;color:${MR_RED};">${mrFmtFlt(defRate)} %</div>
                <div style="font-size:11px;color:#555;">Доля дефектных</div>
            </div>
            <div style="flex:1;background:#F4F8FB;border-top:4px solid ${MR_BLUE};padding:10px;text-align:center;border-radius:4px;">
                <div style="font-size:22px;font-weight:bold;color:#27AE60;">${mrFmtFlt(qualRate)} %</div>
                <div style="font-size:11px;color:#555;">Уровень качества</div>
            </div>
            <div style="flex:1;background:#F4F8FB;border-top:4px solid ${MR_BLUE};padding:10px;text-align:center;border-radius:4px;">
                <div style="font-size:22px;font-weight:bold;color:#1F4E79;">${records.length}</div>
                <div style="font-size:11px;color:#555;">Позиций в отчёте</div>
            </div>
        </div>
        <div style="background:#EFEFEF;border-left:5px solid ${MR_BLUE};padding:10px 14px;margin:12px 0;font-size:12.5px;line-height:1.45;">
            <b>Главное:</b> Проверено <b>${mrFmtInt(totalChecked)}</b> листов, выявлено <b>${mrFmtInt(totalDef)}</b> дефектов — <b>${mrFmtFlt(defRate)}%</b> продукции имеет отклонения, уровень качества — <b>${mrFmtFlt(qualRate)}%</b>. Природные (шпоновые) дефекты составляют <b>${natShare.toFixed(1)}%</b> всех отклонений, технологические — <b>${techShare.toFixed(1)}%</b>. Пять видов дефектов дают <b>${top5Share.toFixed(1)}%</b> всех отклонений: ${top5Names}.
        </div>
        <h2 style="${h2Style}">1. Объём контроля и уровень дефектности по группам</h2>
        <img src="${img1}" style="width:100%;display:block;margin:8px auto;">
        <table style="${tblStyle}">
            <tr>
                <th style="${thStyle}">Группа</th><th style="${thStyle}">Позиций</th><th style="${thStyle}">Проверено, л.</th>
                <th style="${thStyle}">Дефектов</th><th style="${thStyle}">в т.ч. техн.</th><th style="${thStyle}">в т.ч. природн.</th>
                <th style="${thStyle}">Дефектность, %</th><th style="${thStyle}">Качество, %</th>
            </tr>
            ${grpRowsHtml}
            <tr style="background:#EAECEE;font-weight:bold;">
                <td style="${tdStyle}">Итого</td><td style="${tdStyle}">${records.length}</td><td style="${tdStyle}">${mrFmtInt(totalChecked)}</td>
                <td style="${tdStyle}">${mrFmtInt(totalDef)}</td><td style="${tdStyle}">${mrFmtInt(totalTech)}</td><td style="${tdStyle}">${mrFmtInt(totalNat)}</td>
                <td style="${tdStyle}">${mrFmtFlt(defRate)}</td><td style="${tdStyle}">${mrFmtFlt(qualRate)}</td>
            </tr>
        </table>
    `);

    // СТРАНИЦА 2: Раздел 2 (Бублики) + Раздел 3а (Парето 0,9 мм, если есть)
    let page2 = `
        <h2 style="${h2Style}">2. Структура дефектов: природа возникновения</h2>
        <img src="${img2}" style="width:100%;display:block;margin:8px auto;">
    `;
    if (img3a) {
        page2 += `
            <h2 style="${h2Style}">3а. Парето-анализ видов дефектов: рубашка 0,9 мм</h2>
            <img src="${img3a}" style="width:100%;display:block;margin:6px auto;">
            <table style="${tblStyle}">
                <tr>
                    <th style="${thStyle}">№</th><th style="${thStyle}">Вид дефекта (0,9 мм)</th><th style="${thStyle}">Природа</th>
                    <th style="${thStyle}">Кол-во, шт.</th><th style="${thStyle}">Доля, %</th><th style="${thStyle}">Накопл., %</th><th style="${thStyle}">На 1000 л.</th>
                </tr>
                ${makeParetoRowsHtml(pareto09)}
            </table>
        `;
    }
    pagesHtml.push(page2);

    // СТРАНИЦА 3: Раздел 3б (Парето 1,41 мм, если есть)
    if (img3b) {
        pagesHtml.push(`
            <h2 style="${h2Style}">3б. Парето-анализ видов дефектов: рубашка 1,41 мм</h2>
            <img src="${img3b}" style="width:100%;display:block;margin:8px auto;">
            <table style="${tblStyle}">
                <tr>
                    <th style="${thStyle}">№</th><th style="${thStyle}">Вид дефекта (1,41 мм)</th><th style="${thStyle}">Природа</th>
                    <th style="${thStyle}">Кол-во, шт.</th><th style="${thStyle}">Доля, %</th><th style="${thStyle}">Накопл., %</th><th style="${thStyle}">На 1000 л.</th>
                </tr>
                ${makeParetoRowsHtml(pareto141)}
            </table>
        `);
    }

    // СТРАНИЦА 4: Раздел 4 (Профиль) + Раздел 5 (Толщина)
    if (img4 || img5) {
        let page4 = '';
        if (img4) {
            page4 += `
                <h2 style="${h2Style}">4. Сравнение профилей дефектности 0,9 и 1,41 мм</h2>
                <img src="${img4}" style="width:100%;display:block;margin:8px auto;">
            `;
        }
        if (img5) {
            page4 += `
                <h2 style="${h2Style}">5. Влияние толщины плиты</h2>
                <img src="${img5}" style="width:100%;display:block;margin:8px auto;">
            `;
        }
        pagesHtml.push(page4);
    }

    // СТРАНИЦА 5: Раздел 6 (Проблемные позиции + Худшие/Лучшие)
    if (img6) {
        const subNote = actualMinProb > 1 ? ` (партии от ${actualMinProb} листов)` : '';
        pagesHtml.push(`
            <h2 style="${h2Style}">6. Проблемные позиции</h2>
            <img src="${img6}" style="width:100%;display:block;margin:6px auto;">
            <h3 style="${h3Style}">Худшие ${worst8.length} позиций${subNote}</h3>
            <table style="${tblStyle}">
                <tr>
                    <th style="${thStyle}">Позиция</th><th style="${thStyle}">Толщина рубашки</th><th style="${thStyle}">Проверено, л.</th>
                    <th style="${thStyle}">Дефектов</th><th style="${thStyle}">Дефектность, %</th><th style="${thStyle}">Ключевой дефект</th>
                </tr>
                ${worst8Html}
            </table>
            <h3 style="${h3Style}">Лучшие ${best5.length} позиций${subNote} — эталон процесса</h3>
            <table style="${tblStyle}">
                <tr>
                    <th style="${thStyle}">Позиция</th><th style="${thStyle}">Толщина рубашки</th><th style="${thStyle}">Проверено, л.</th>
                    <th style="${thStyle}">Дефектов</th><th style="${thStyle}">Дефектность, %</th>
                </tr>
                ${best5Html}
            </table>
        `);
    }

    // СТРАНИЦА 6: Раздел 7 (Карта дефектности) + Раздел 8 (Распределение по качеству)
    if (img7 || img8) {
        let page6 = '';
        if (img7) {
            page6 += `
                <h2 style="${h2Style}">7. Карта дефектности крупнейших партий</h2>
                <img src="${img7}" style="width:100%;display:block;margin:8px auto;">
            `;
        }
        if (img8) {
            page6 += `
                <h2 style="${h2Style}">8. Распределение позиций по уровню качества</h2>
                <img src="${img8}" style="width:100%;display:block;margin:8px auto;">
            `;
        }
        pagesHtml.push(page6);
    }

    // ПРИЛОЖЕНИЕ: Сводная таблица по всем позициям (разбивка по 26 строк на лист)
    const sortedApp = [...records].sort((a, b) =>
        a.veneer.localeCompare(b.veneer) || a.tu.localeCompare(b.tu) || (b.defRate - a.defRate)
    );
    const chunkSize = 26;
    for (let start = 0; start < sortedApp.length; start += chunkSize) {
        const chunk = sortedApp.slice(start, start + chunkSize);
        const isLast = start + chunkSize >= sortedApp.length;
        const rowsHtml = chunk.map(r => {
            const bg = r.defRate >= 7.0 ? 'background:#FADBD8;' : (r.defRate <= 3.0 ? 'background:#D4EFDF;' : '');
            return `
                <tr style="${bg}">
                    <td style="${tdStyle}">${r.veneer}</td>
                    <td style="${tdStyle}"><b>${r.position}</b></td>
                    <td style="${tdStyle}">${r.tu}</td>
                    <td style="${tdStyle}">${mrFmtInt(r.checked)}</td>
                    <td style="${tdStyle}">${mrFmtInt(r.techDef)}</td>
                    <td style="${tdStyle}">${mrFmtInt(r.natDef)}</td>
                    <td style="${tdStyle}"><b>${mrFmtInt(r.totalDef)}</b></td>
                    <td style="${tdStyle}"><b>${mrFmtFlt(r.defRate)}</b></td>
                    <td style="${tdStyle}">${mrFmtFlt(r.qualRate)}</td>
                </tr>
            `;
        }).join('');

        const totalRowHtml = isLast ? `
            <tr style="background:#EAECEE;font-weight:bold;">
                <td colspan="3" style="${tdStyle}">Итого</td>
                <td style="${tdStyle}">${mrFmtInt(totalChecked)}</td>
                <td style="${tdStyle}">${mrFmtInt(totalTech)}</td>
                <td style="${tdStyle}">${mrFmtInt(totalNat)}</td>
                <td style="${tdStyle}">${mrFmtInt(totalDef)}</td>
                <td style="${tdStyle}">${mrFmtFlt(defRate)}</td>
                <td style="${tdStyle}">${mrFmtFlt(qualRate)}</td>
            </tr>
        ` : '';

        const footerNote = isLast && sourceLabel
            ? `<p style="font-size:11px;color:#666;margin-top:8px;">Источник: ${sourceLabel}${operatorStr ? ' • Оператор: ' + operatorStr : ''}</p>`
            : '';

        pagesHtml.push(`
            <h2 style="${h2Style}">Приложение. Сводная таблица по всем позициям${start > 0 ? ' (продолжение)' : ''}</h2>
            <table style="${tblStyle}">
                <tr>
                    <th style="${thStyle}">Рубашка</th><th style="${thStyle}">Позиция</th><th style="${thStyle}">ТУ</th>
                    <th style="${thStyle}">Проверено</th><th style="${thStyle}">Техн.</th><th style="${thStyle}">Природн.</th>
                    <th style="${thStyle}">Всего деф.</th><th style="${thStyle}">Дефектность, %</th><th style="${thStyle}">Качество, %</th>
                </tr>
                ${rowsHtml}
                ${totalRowHtml}
            </table>
            ${footerNote}
        `);
    }

    // Рендерим каждую страницу в Canvas -> собираем портретный A4 PDF
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
    const pageW = doc.internal.pageSize.getWidth();
    const pageH = doc.internal.pageSize.getHeight();
    const margin = 8;
    const maxW = pageW - margin * 2;
    const maxH = pageH - margin * 2 - 6;

    const host = document.createElement('div');
    host.style.cssText = 'position:fixed;left:-12000px;top:0;width:880px;background:#ffffff;color:#222222;font-family:"Segoe UI",Arial,sans-serif;z-index:-1;';
    document.body.appendChild(host);

    try {
        for (let i = 0; i < pagesHtml.length; i++) {
            host.innerHTML = `<div style="padding:18px 22px;background:#ffffff;color:#222222;">${pagesHtml[i]}</div>`;
            const target = host.firstElementChild;

            // Ждем загрузки всех dataURL-картинок внутри блока
            const imgs = Array.from(target.querySelectorAll('img'));
            await Promise.all(imgs.map(img => img.complete ? Promise.resolve() : new Promise(res => { img.onload = res; img.onerror = res; })));

            const canvas = await window.html2canvas(target, {
                scale: 1.75,
                backgroundColor: '#ffffff',
                logging: false,
                useCORS: true
            });

            if (i > 0) doc.addPage('a4', 'portrait');
            const imgData = canvas.toDataURL('image/jpeg', 0.92);
            const ratio = Math.min(maxW / canvas.width, maxH / canvas.height);
            const drawW = canvas.width * ratio;
            const drawH = canvas.height * ratio;
            const x = (pageW - drawW) / 2;

            doc.addImage(imgData, 'JPEG', x, margin, drawW, drawH, undefined, 'FAST');
            doc.setFontSize(8);
            doc.setTextColor(130);
            doc.text(`${i + 1} / ${pagesHtml.length}`, pageW - margin, pageH - 4, { align: 'right' });
        }
    } finally {
        document.body.removeChild(host);
    }

    const blob = doc.output('blob');
    const dataUri = doc.output('datauristring');
    const base64 = dataUri.split(',')[1] || '';
    return { blob, base64, pagesCount: pagesHtml.length };
}