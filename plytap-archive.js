// =========================================================================
// PLYTAP-ARCHIVE.JS — Закрытие смены, папки Год/Месяц/Дата, накопление
// =========================================================================

let activeArchiveFolderView = null;

// ------------------- 1. ИНДИКАТОР ЗАГРУЗКИ (ОВЕРЛЕЙ) -------------------

function showPlyBusyOverlay(title, subtitle = '') {
    let el = document.getElementById('plyBusyOverlay');
    if (!el) {
        el = document.createElement('div');
        el.id = 'plyBusyOverlay';
        el.className = 'fixed inset-0 z-[70] hidden items-center justify-center bg-slate-950/85 backdrop-blur-sm p-4';
        el.innerHTML = `
            <div class="bg-slate-900 border-2 border-amber-500/70 rounded-2xl max-w-md w-full p-6 text-center shadow-2xl">
                <div class="w-12 h-12 mx-auto mb-3 rounded-full border-4 border-amber-500/30 border-t-amber-400 animate-spin"></div>
                <h4 id="plyBusyTitle" class="text-base sm:text-lg font-extrabold text-white"></h4>
                <p id="plyBusySub" class="text-xs sm:text-sm text-slate-400 mt-1.5"></p>
            </div>
        `;
        document.body.appendChild(el);
    }
    document.getElementById('plyBusyTitle').textContent = title || 'Формирование отчётов...';
    document.getElementById('plyBusySub').textContent = subtitle || 'Пожалуйста, подождите несколько секунд';
    el.classList.remove('hidden');
    el.classList.add('flex');
}

function hidePlyBusyOverlay() {
    const el = document.getElementById('plyBusyOverlay');
    if (el) {
        el.classList.add('hidden');
        el.classList.remove('flex');
    }
}

function downloadPdfFromBase64(b64, filename) {
    if (!b64) return;
    const bytes = base64ToUint8Array(b64);
    const blob = new Blob([bytes], { type: 'application/pdf' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1500);
}

// ------------------- 2. ПАРСИНГ .XLSX (ЕСЛИ НЕТ СЫРЫХ ЦИКЛОВ) -------------------

function parseMetaFromExcelFilename(filename, fallbackDate, fallbackShift) {
    let dateISO = fallbackDate || new Date().toISOString().slice(0, 10);
    let shift = String(fallbackShift || '1');

    const m = String(filename || '').match(/(\d{1,2})[.\-_](\d{1,2})[.\-_](\d{2,4})/);
    if (m) {
        const dd = m[1].padStart(2, '0');
        const mm = m[2].padStart(2, '0');
        let yy = m[3];
        if (yy.length === 2) yy = '20' + yy;
        dateISO = `${yy}-${mm}-${dd}`;
    }
    const s = String(filename || '').match(/(\d+)\s*см/i);
    if (s) shift = s[1];

    return { dateISO, shift };
}

function extractExcelCellNumeric(cellVal) {
    if (cellVal === null || cellVal === undefined) return 0;
    if (typeof cellVal === 'number') return Number.isFinite(cellVal) ? cellVal : 0;
    if (typeof cellVal === 'object' && cellVal.result !== undefined) {
        return Number(cellVal.result) || 0;
    }
    const num = parseFloat(String(cellVal).replace(',', '.').trim());
    return Number.isFinite(num) ? num : 0;
}

async function extractCyclesFromExcelFileRecord(fileRecord) {
    if (Array.isArray(fileRecord.cycles) && fileRecord.cycles.length > 0) {
        return fileRecord.cycles;
    }
    if (!fileRecord.base64 || typeof ExcelJS === 'undefined') return [];

    try {
        const bytes = base64ToUint8Array(fileRecord.base64);
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(bytes);

        const parsedMeta = parseMetaFromExcelFilename(fileRecord.filename, fileRecord.date, fileRecord.shift);
        const cycles = [];

        wb.eachSheet((ws) => {
            const sName = String(ws.name || '').trim();
            let veneer = null;
            if (sName.includes('0,9') || sName.includes('0.9')) veneer = '0,9';
            else if (sName.includes('1,41') || sName.includes('1.41')) veneer = '1,41';
            if (!veneer) return;

            const customTitles = {};
            const headerRow = ws.getRow(1);
            EXCEL_COLUMNS.forEach((colDef, idx) => {
                if (colDef.type === 'defect_cn' || colDef.type === 'defect_py') {
                    const txt = String(headerRow.getCell(idx + 1).text || '').trim();
                    if (txt && txt !== colDef.title) {
                        customTitles[colDef.key] = txt;
                    }
                }
            });

            let currentTu = 'ТУ-Э';
            ws.eachRow((row, rowNumber) => {
                if (rowNumber === 1) return;
                const colA = String(row.getCell(1).text || '').trim();
                if (!colA) return;
                const lowA = colA.toLowerCase();

                if (lowA.includes('ту-э')) { currentTu = 'ТУ-Э'; return; }
                if (lowA.includes('ту-р')) { currentTu = 'ТУ-Р'; return; }
                if (lowA.startsWith('итого') || lowA.startsWith('всего') || lowA === '%' || lowA.startsWith('оператор')) return;

                const checked = extractExcelCellNumeric(row.getCell(2).value);
                if (!checked || checked <= 0) return;

                const counts = createEmptyCounts();
                EXCEL_COLUMNS.forEach((colDef, idx) => {
                    if (colDef.type === 'defect_cn' || colDef.type === 'defect_py') {
                        counts[colDef.key] = extractExcelCellNumeric(row.getCell(idx + 1).value);
                    }
                });

                const metrics = calculateRowMetrics(counts, checked);
                cycles.push({
                    id: `imp_${fileRecord.id}_${veneer}_${rowNumber}`,
                    date: parsedMeta.dateISO,
                    shift: parsedMeta.shift,
                    operator: fileRecord.operator || '',
                    veneerThickness: veneer,
                    tuStandard: currentTu,
                    section: `Рубашка ${veneer} ${currentTu}`,
                    sort: colA,
                    checked,
                    counts,
                    customTitles,
                    sumO: metrics.sumO,
                    sumZ: metrics.sumZ,
                    sumAA: metrics.sumAA,
                    yieldAB: metrics.yieldAB
                });
            });
        });

        // Сохраняем восстановленные циклы в запись Excel, чтобы больше не парсить повторно
        if (cycles.length > 0) {
            fileRecord.cycles = cycles;
            fileRecord.date = parsedMeta.dateISO;
            fileRecord.shift = parsedMeta.shift;
            if (fileRecord.cyclesCount === '-' || !fileRecord.cyclesCount) {
                fileRecord.cyclesCount = cycles.length;
                fileRecord.totalSheets = cycles.reduce((a, c) => a + c.checked, 0);
                fileRecord.totalDefects = cycles.reduce((a, c) => a + c.sumAA, 0);
            }
            await saveFileRecordToIndexedDB(fileRecord);
        }

        return cycles;
    } catch (err) {
        console.error('Ошибка разбора Excel-файла:', err);
        return [];
    }
}

// ------------------- 3. НАКОПЛЕНИЕ ЗА МЕСЯЦ (00_Итого_за_месяц_графики) -------------------

async function collectAllCyclesForMonth(monthKey, extraShiftCycles = [], excludeFileId = null) {
    const allCycles = [];
    const usedFileIds = new Set();

    // 1. Берем циклы из уже сохранённых отчётов смен
    for (const rep of savedShiftReports) {
        if (rep.monthKey === monthKey && rep.fileId !== excludeFileId && Array.isArray(rep.cycles) && rep.cycles.length > 0) {
            allCycles.push(...rep.cycles);
            if (rep.fileId) usedFileIds.add(rep.fileId);
        }
    }

    // 2. Подхватываем смены из savedFiles, если по ним ещё нет записи в shift_reports
    for (const f of savedFiles) {
        if (!f || f.id === excludeFileId || usedFileIds.has(f.id)) continue;
        const meta = parseMetaFromExcelFilename(f.filename, f.date, f.shift);
        if (meta.dateISO.slice(0, 7) !== monthKey) continue;
        const cycs = await extractCyclesFromExcelFileRecord(f);
        if (cycs.length > 0) {
            allCycles.push(...cycs);
            usedFileIds.add(f.id);
        }
    }

    // 3. Добавляем текущую закрываемую смену
    if (Array.isArray(extraShiftCycles) && extraShiftCycles.length > 0) {
        allCycles.push(...extraShiftCycles);
    }

    return allCycles;
}

async function rebuildMonthlyReportForKey(monthKey, silent = false) {
    if (!monthKey || !/^\d{4}-\d{2}$/.test(monthKey)) return null;

    const [year, month] = monthKey.split('-');
    const monthFolder = PLYTAP_MONTH_FOLDERS[month] || `${month}_Месяц`;
    const monthDisplayLabel = `${year} / ${monthFolder}`;

    const allMonthCycles = await collectAllCyclesForMonth(monthKey);

    if (allMonthCycles.length === 0) {
        await idbDelete(IDB_STORE_MONTHLY_REPORTS, monthKey);
        savedMonthlyReports = savedMonthlyReports.filter(m => m.monthKey !== monthKey);
        return null;
    }

    const shiftsSet = new Set(allMonthCycles.map(c => `${c.date || ''}_см${c.shift || '1'}`));
    const totalSheets = allMonthCycles.reduce((a, c) => a + (Number(c.checked) || 0), 0);
    const totalDefects = allMonthCycles.reduce((a, c) => a + (Number(c.sumAA) || 0), 0);
    const goodPct = totalSheets > 0 ? ((totalSheets - totalDefects) / totalSheets) * 100 : 0;

    const monthCharts = generateMonthSummaryChartsBundle(allMonthCycles, monthDisplayLabel);
    const chartsPdf = monthCharts.length > 0 ? combineChartsToPdf(monthCharts) : null;

    const monthRecord = {
        monthKey,
        year,
        month,
        monthFolder,
        folderPath: `${year} / ${monthFolder} / 00_Итого_за_месяц_графики`,
        updatedAt: new Date().toLocaleString('ru-RU'),
        shiftsCount: shiftsSet.size,
        cyclesCount: allMonthCycles.length,
        totalSheets,
        totalDefects,
        goodPct,
        charts: monthCharts,
        chartsPdfBase64: chartsPdf ? chartsPdf.base64 : '',
        chartsPdfFilename: `Итого_графики_${monthFolder}_${year}.pdf`
    };

    await saveMonthlyReportToDB(monthRecord);
    if (!silent) {
        showToast(`Накопительные графики за ${monthFolder} ${year} обновлены (${shiftsSet.size} см.)`, 'emerald');
    }
    return monthRecord;
}

// ------------------- 4. ЗАКРЫТИЕ СМЕНЫ: ПАПКА PNG + АНАЛИТ. PDF + МЕСЯЦ -------------------

async function buildAndStoreShiftPackage(cycles, metaOptions = {}, autoDownloadPdf = false) {
    const cleanCycles = JSON.parse(JSON.stringify(cycles || [])).filter(c => Number(c.checked) > 0);
    if (cleanCycles.length === 0) return null;

    const dateISO = metaOptions.dateISO || (cleanCycles[0] && cleanCycles[0].date) || appState.date;
    const shiftNum = String(metaOptions.shiftNum || (cleanCycles[0] && cleanCycles[0].shift) || appState.shift || '1');
    const operatorStr = metaOptions.operator || (cleanCycles[0] && cleanCycles[0].operator) || appState.operator || '';
    const fileId = metaOptions.fileId || appState.currentShiftFileId || ('shift_' + Date.now());
    const excelFileName = metaOptions.excelFileName || `${plyFormatDateRuShort(dateISO)}г.${shiftNum}см.xlsx`;

    // Гарантируем, что в каждом цикле прописаны дата и смена этой смены
    cleanCycles.forEach(c => {
        c.date = dateISO;
        c.shift = shiftNum;
        if (!c.operator && operatorStr) c.operator = operatorStr;
    });

    const folderMeta = buildFolderMeta(dateISO, shiftNum);
    const periodLabel = `${folderMeta.dateRu} см.${shiftNum}`;

    // 1. Все циклы месяца для графика динамики между сменами
    const monthAllCycles = await collectAllCyclesForMonth(folderMeta.monthKey, cleanCycles, fileId);

    // 2. Графики смены (PNG) и сборный PDF из графиков
    const shiftCharts = generateShiftChartsBundle(cleanCycles, periodLabel, monthAllCycles);
    const chartsPdf = shiftCharts.length > 0 ? combineChartsToPdf(shiftCharts) : null;

    // 3. Аналитический PDF-отчёт за смену (порт monthly_report.py)
    const analyticalPdfFilename = `Отчет_по_дефектовке_${folderMeta.dateRu}_см${shiftNum}.pdf`;
    const analyticalPdf = await generateAnalyticalPdfReport(cleanCycles, {
        isMonthly: false,
        periodTitle: `${folderMeta.dateRu} (Смена ${shiftNum})`,
        operator: operatorStr,
        sourceLabel: excelFileName
    });

    const totalSheets = cleanCycles.reduce((a, c) => a + (Number(c.checked) || 0), 0);
    const totalDefects = cleanCycles.reduce((a, c) => a + (Number(c.sumAA) || 0), 0);
    const goodPct = totalSheets > 0 ? ((totalSheets - totalDefects) / totalSheets) * 100 : 0;

    const reportRecord = {
        id: 'rep_' + fileId,
        fileId,
        excelFileName,
        operator: operatorStr,
        date: dateISO,
        shift: shiftNum,
        year: folderMeta.year,
        month: folderMeta.month,
        monthKey: folderMeta.monthKey,
        monthFolder: folderMeta.monthFolder,
        shiftFolder: folderMeta.shiftFolder,
        fullPath: folderMeta.fullPath,
        createdAt: new Date().toLocaleString('ru-RU'),
        cyclesCount: cleanCycles.length,
        totalSheets,
        totalDefects,
        goodPct,
        cycles: cleanCycles,
        charts: shiftCharts,
        chartsPdfBase64: chartsPdf ? chartsPdf.base64 : '',
        chartsPdfFilename: `Графики_${folderMeta.dateRu}_см${shiftNum}.pdf`,
        analyticalPdfBase64: analyticalPdf ? analyticalPdf.base64 : '',
        analyticalPdfFilename
    };

    await saveShiftReportToDB(reportRecord);

    // 4. Обновляем накопительные графики за весь месяц
    await rebuildMonthlyReportForKey(folderMeta.monthKey, true);

    // 5. Скачиваем аналитический PDF смены с небольшой задержкой после .xlsx
    if (autoDownloadPdf && reportRecord.analyticalPdfBase64) {
        setTimeout(() => {
            downloadPdfFromBase64(reportRecord.analyticalPdfBase64, reportRecord.analyticalPdfFilename);
        }, 750);
    }

    return reportRecord;
}

// Вызывается автоматически внутри finishShiftAndGoHome() и handleStartShiftFromMenu()
async function finalizeShiftReportsPackage() {
    const cycles = Array.isArray(appState.accumulatedCycles) ? appState.accumulatedCycles : [];
    if (cycles.length === 0) return;

    const firstCyc = cycles[0] || {};
    const dateISO = appState.shiftStartDate || firstCyc.date || appState.date;
    const shiftNum = appState.shift || firstCyc.shift || '1';
    const fileId = appState.currentShiftFileId || ('file_' + Date.now());
    const savedRec = savedFiles.find(f => f.id === fileId);

    showPlyBusyOverlay(
        'Формирование отчётов смены и месяца...',
        'Собираем PNG-графики в папку, создаём аналитический PDF смены и обновляем сводку месяца'
    );

    try {
        const rep = await buildAndStoreShiftPackage(
            cycles,
            {
                dateISO,
                shiftNum,
                operator: appState.operator || firstCyc.operator || '',
                fileId,
                excelFileName: savedRec ? savedRec.filename : `${plyFormatDateRuShort(dateISO)}г.${shiftNum}см.xlsx`
            },
            true
        );

        renderSavedFilesList();
        if (rep) {
            showToast(`Отчёты сохранены в папку: ${rep.fullPath}`, 'emerald');
        }
    } catch (err) {
        console.error('Ошибка при формировании пакета отчётов смены:', err);
        showToast('Excel сохранён, но возникла ошибка при генерации PDF/графиков', 'amber');
    } finally {
        hidePlyBusyOverlay();
    }
}

// Сформировать графики и аналитический PDF для уже существующего/импортированного .xlsx файла
async function generateReportsForExistingFile(fileId) {
    const f = savedFiles.find(x => x.id === fileId);
    if (!f) return;

    showPlyBusyOverlay(
        `Обработка файла ${f.filename}...`,
        'Извлекаем данные из Excel, строим PNG-графики, аналитический PDF и обновляем месяц'
    );

    try {
        const cycles = await extractCyclesFromExcelFileRecord(f);
        if (!cycles || cycles.length === 0) {
            showToast('В этом Excel-файле не найдено строк с проверенными листами!', 'amber');
            return;
        }
        const meta = parseMetaFromExcelFilename(f.filename, f.date, f.shift);
        const rep = await buildAndStoreShiftPackage(
            cycles,
            {
                dateISO: meta.dateISO,
                shiftNum: meta.shift,
                operator: f.operator || '',
                fileId: f.id,
                excelFileName: f.filename
            },
            false
        );
        renderSavedFilesList();
        if (rep) {
            showToast(`Сформированы графики (${rep.charts.length} шт.) и PDF в папке ${rep.fullPath}`, 'emerald');
        }
    } catch (err) {
        console.error(err);
        showToast('Ошибка при построении отчётов из файла Excel', 'amber');
    } finally {
        hidePlyBusyOverlay();
    }
}

// ------------------- 5. ДЕЙСТВИЯ С ОТЧЁТАМИ В АРХИВЕ -------------------

function findShiftReportForFile(fileRec) {
    return savedShiftReports.find(r => r.fileId === fileRec.id) ||
        savedShiftReports.find(r => r.excelFileName === fileRec.filename) ||
        null;
}

function downloadShiftAnalyticalPdf(reportId) {
    const rep = savedShiftReports.find(r => r.id === reportId);
    if (!rep || !rep.analyticalPdfBase64) {
        showToast('Аналитический PDF для этой смены не найден', 'amber');
        return;
    }
    downloadPdfFromBase64(rep.analyticalPdfBase64, rep.analyticalPdfFilename);
    showToast(`Скачивание ${rep.analyticalPdfFilename}`, 'emerald');
}

async function shareShiftAnalyticalPdfToMax(reportId) {
    const rep = savedShiftReports.find(r => r.id === reportId);
    if (!rep || !rep.analyticalPdfBase64) {
        showToast('Аналитический PDF для этой смены не найден', 'amber');
        return;
    }
    const text = `Аналитический отчёт смены PlyTap (${rep.fullPath})\nОператор: ${rep.operator || '-'}\nПроверено: ${rep.totalSheets} л., Брак: ${rep.totalDefects} л., Годных: ${rep.goodPct.toFixed(2)}%`;
    await sharePdfToMax(rep.analyticalPdfBase64, rep.analyticalPdfFilename, text);
}

async function shareShiftChartsPdfToMax(reportId) {
    const rep = savedShiftReports.find(r => r.id === reportId);
    if (!rep || !rep.chartsPdfBase64) {
        showToast('PDF с графиками не найден', 'amber');
        return;
    }
    const text = `Графики смены PlyTap (${rep.fullPath})\nПроверено: ${rep.totalSheets} л., Брак: ${rep.totalDefects} л., Годных: ${rep.goodPct.toFixed(2)}%`;
    await sharePdfToMax(rep.chartsPdfBase64, rep.chartsPdfFilename, text);
}

async function shareMonthChartsPdfToMax(monthKey) {
    const m = savedMonthlyReports.find(x => x.monthKey === monthKey);
    if (!m || !m.chartsPdfBase64) {
        showToast('Сначала нажмите «Обновить» для расчёта графиков месяца', 'amber');
        return;
    }
    const text = `Итоговые графики за месяц (${m.monthFolder} ${m.year})\nСмен: ${m.shiftsCount}, Проверено: ${m.totalSheets} л., Брак: ${m.totalDefects} л., Годных: ${m.goodPct.toFixed(2)}%`;
    await sharePdfToMax(m.chartsPdfBase64, m.chartsPdfFilename, text);
}

async function rebuildMonthlyReportInteractive(monthKey) {
    showPlyBusyOverlay('Пересчёт накопительного отчёта месяца...', 'Суммируем все смены выбранного месяца');
    try {
        await rebuildMonthlyReportForKey(monthKey, false);
        renderSavedFilesList();
    } catch (e) {
        console.error(e);
        showToast('Ошибка пересчёта месяца', 'amber');
    } finally {
        hidePlyBusyOverlay();
    }
}

async function generateAndActionMonthlyAnalyticalPdf(monthKey, sendToMax = false) {
    const [year, month] = monthKey.split('-');
    const monthFolder = PLYTAP_MONTH_FOLDERS[month] || `${month}_Месяц`;
    const monthNameClean = monthFolder.replace(/^\d+_/, '').toLowerCase();

    showPlyBusyOverlay(
        `Аналитический PDF за ${monthNameClean} ${year}...`,
        'Собираем все смены месяца, строим 8 графиков и сводные таблицы'
    );

    try {
        const allMonthCycles = await collectAllCyclesForMonth(monthKey);
        if (allMonthCycles.length === 0) {
            showToast('Нет данных за этот месяц для построения PDF!', 'amber');
            return;
        }
        const pdf = await generateAnalyticalPdfReport(allMonthCycles, {
            isMonthly: true,
            periodTitle: `${monthNameClean} ${year}`,
            sourceLabel: `Архив PlyTap (${monthFolder} ${year})`
        });
        if (!pdf) return;

        const filename = `Отчет_по_дефектовке_${monthNameClean}_${year}.pdf`;
        if (sendToMax) {
            await sharePdfToMax(pdf.blob, filename, `Месячный аналитический отчёт PlyTap — ${monthNameClean} ${year}`);
        } else {
            downloadPdfFromBase64(pdf.base64, filename);
            showToast(`Скачан месячный отчёт: ${filename}`, 'emerald');
        }
    } catch (err) {
        console.error(err);
        showToast('Ошибка при создании месячного аналитического PDF', 'amber');
    } finally {
        hidePlyBusyOverlay();
    }
}

// ------------------- 6. МОДАЛЬНОЕ ОКНО ПРОСМОТРА ПАПКИ ГРАФИКОВ -------------------

function openArchiveFolderModal(type, keyId) {
    let folderData = null;
    if (type === 'shift') {
        const rep = savedShiftReports.find(r => r.id === keyId);
        if (!rep) return;
        folderData = {
            type: 'shift',
            id: rep.id,
            path: rep.fullPath,
            subtitle: `Оператор: ${rep.operator || '—'} • Проверено: ${rep.totalSheets} л. • Брак: ${rep.totalDefects} л. • Годных: ${rep.goodPct.toFixed(2)}%`,
            charts: rep.charts || [],
            chartsPdfBase64: rep.chartsPdfBase64,
            chartsPdfFilename: rep.chartsPdfFilename,
            analyticalPdfBase64: rep.analyticalPdfBase64,
            analyticalPdfFilename: rep.analyticalPdfFilename
        };
    } else if (type === 'month') {
        const m = savedMonthlyReports.find(x => x.monthKey === keyId);
        if (!m) return;
        folderData = {
            type: 'month',
            id: m.monthKey,
            path: m.folderPath,
            subtitle: `Накопление за месяц • Смен: ${m.shiftsCount} • Проверено: ${m.totalSheets} л. • Брак: ${m.totalDefects} л. • Годных: ${m.goodPct.toFixed(2)}%`,
            charts: m.charts || [],
            chartsPdfBase64: m.chartsPdfBase64,
            chartsPdfFilename: m.chartsPdfFilename,
            analyticalPdfBase64: '',
            analyticalPdfFilename: ''
        };
    }

    if (!folderData) return;
    activeArchiveFolderView = folderData;

    let modal = document.getElementById('plyArchiveFolderModal');
    if (!modal) {
        modal = document.createElement('div');
        modal.id = 'plyArchiveFolderModal';
        modal.className = 'fixed inset-0 z-50 hidden items-center justify-center bg-slate-950/90 backdrop-blur-sm p-3 sm:p-5';
        modal.innerHTML = `
            <div class="bg-slate-900 border-2 border-slate-700 rounded-2xl max-w-6xl w-full max-h-[92vh] flex flex-col overflow-hidden shadow-2xl">
                <div class="px-4 py-3 border-b border-slate-800 flex flex-wrap items-center justify-between gap-3">
                    <div class="flex items-center gap-2.5 min-w-0">
                        <div class="w-10 h-10 rounded-xl bg-amber-500/20 border border-amber-500/40 text-amber-400 flex items-center justify-center shrink-0">
                            <i class="ph-fill ph-folder-open text-2xl"></i>
                        </div>
                        <div class="min-w-0">
                            <h4 id="plyFolderModalPath" class="text-sm sm:text-base font-extrabold text-white mono truncate"></h4>
                            <p id="plyFolderModalSub" class="text-xs text-slate-400 truncate"></p>
                        </div>
                    </div>
                    <button type="button" onclick="closeArchiveFolderModal()" class="p-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white">
                        <i class="ph-bold ph-x text-lg"></i>
                    </button>
                </div>

                <div id="plyFolderModalActions" class="px-4 py-2.5 bg-slate-950 border-b border-slate-800 flex flex-wrap items-center gap-2"></div>

                <div id="plyFolderModalGrid" class="p-4 overflow-y-auto grid grid-cols-1 lg:grid-cols-2 gap-4 flex-1"></div>
            </div>
        `;
        modal.addEventListener('click', (e) => {
            if (e.target === modal) closeArchiveFolderModal();
        });
        document.body.appendChild(modal);
    }

    document.getElementById('plyFolderModalPath').textContent = folderData.path;
    document.getElementById('plyFolderModalSub').textContent = folderData.subtitle;

    const actionsEl = document.getElementById('plyFolderModalActions');
    let actionsHtml = `
        <button type="button" onclick="shareActiveFolderChartsPdf()" class="px-3.5 py-2 rounded-xl bg-gradient-to-r from-indigo-500 to-violet-600 hover:from-indigo-400 hover:to-violet-500 text-white text-xs font-extrabold flex items-center gap-1.5 shadow transition">
            <i class="ph-bold ph-paper-plane-tilt text-base"></i>
            <span>Отправить графики в MAX (PDF)</span>
        </button>
        <button type="button" onclick="downloadActiveFolderChartsPdf()" class="px-3 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 text-xs font-bold flex items-center gap-1.5 transition">
            <i class="ph-bold ph-file-pdf text-rose-400 text-base"></i>
            <span>Скачать графики PDF</span>
        </button>
        <button type="button" onclick="downloadAllPngsInActiveFolder()" class="px-3 py-2 rounded-xl bg-emerald-500/15 hover:bg-emerald-500/25 text-emerald-300 border border-emerald-500/40 text-xs font-bold flex items-center gap-1.5 transition">
            <i class="ph-bold ph-images text-base"></i>
            <span>Скачать все PNG (${folderData.charts.length})</span>
        </button>
    `;

    if (folderData.type === 'shift' && folderData.analyticalPdfBase64) {
        actionsHtml += `
            <button type="button" onclick="downloadShiftAnalyticalPdf('${folderData.id}')" class="px-3 py-2 rounded-xl bg-amber-500/20 hover:bg-amber-500/30 text-amber-300 border border-amber-500/40 text-xs font-extrabold flex items-center gap-1.5 transition">
                <i class="ph-bold ph-file-text text-base"></i>
                <span>Аналит. PDF смены</span>
            </button>
            <button type="button" onclick="shareShiftAnalyticalPdfToMax('${folderData.id}')" class="px-3 py-2 rounded-xl bg-sky-500/20 hover:bg-sky-500/30 text-sky-300 border border-sky-500/40 text-xs font-extrabold flex items-center gap-1.5 transition">
                <i class="ph-bold ph-paper-plane-tilt text-base"></i>
                <span>Аналит. PDF в MAX</span>
            </button>
        `;
    } else if (folderData.type === 'month') {
        actionsHtml += `
            <button type="button" onclick="generateAndActionMonthlyAnalyticalPdf('${folderData.id}', false)" class="px-3 py-2 rounded-xl bg-amber-500/20 hover:bg-amber-500/30 text-amber-300 border border-amber-500/40 text-xs font-extrabold flex items-center gap-1.5 transition">
                <i class="ph-bold ph-file-text text-base"></i>
                <span>Скачать аналит. PDF месяца</span>
            </button>
            <button type="button" onclick="generateAndActionMonthlyAnalyticalPdf('${folderData.id}', true)" class="px-3 py-2 rounded-xl bg-sky-500/20 hover:bg-sky-500/30 text-sky-300 border border-sky-500/40 text-xs font-extrabold flex items-center gap-1.5 transition">
                <i class="ph-bold ph-paper-plane-tilt text-base"></i>
                <span>Аналит. PDF месяца в MAX</span>
            </button>
        `;
    }
    actionsEl.innerHTML = actionsHtml;

    const gridEl = document.getElementById('plyFolderModalGrid');
    gridEl.innerHTML = folderData.charts.map((ch, idx) => `
        <div class="bg-slate-950 border border-slate-800 rounded-xl overflow-hidden flex flex-col">
            <div class="px-3 py-2 border-b border-slate-800 flex items-center justify-between gap-2">
                <div class="truncate">
                    <span class="text-xs font-extrabold text-white block truncate">${ch.title}</span>
                    <span class="text-[10px] text-slate-400 mono">${ch.filename}</span>
                </div>
                <div class="flex items-center gap-1.5 shrink-0">
                    <button type="button" onclick="openFolderChartZoom(${idx})" class="px-2 py-1 rounded-lg bg-slate-800 hover:bg-slate-700 text-sky-300 text-xs font-bold">
                        <i class="ph-bold ph-magnifying-glass-plus"></i>
                    </button>
                    <button type="button" onclick="downloadFolderSinglePng(${idx})" class="px-2.5 py-1 rounded-lg bg-emerald-500/15 hover:bg-emerald-500 hover:text-slate-950 text-emerald-300 border border-emerald-500/30 text-xs font-bold flex items-center gap-1 transition">
                        <i class="ph-bold ph-download-simple"></i>
                        <span>PNG</span>
                    </button>
                </div>
            </div>
            <div onclick="openFolderChartZoom(${idx})" class="bg-white p-1.5 cursor-pointer">
                <img src="${ch.dataUrl}" alt="${ch.title}" class="w-full h-auto block rounded">
            </div>
        </div>
    `).join('');

    modal.classList.remove('hidden');
    modal.classList.add('flex');
}

function closeArchiveFolderModal() {
    const modal = document.getElementById('plyArchiveFolderModal');
    if (modal) {
        modal.classList.add('hidden');
        modal.classList.remove('flex');
    }
}

function openFolderChartZoom(idx) {
    if (!activeArchiveFolderView || !activeArchiveFolderView.charts[idx]) return;
    const ch = activeArchiveFolderView.charts[idx];
    openChartLightbox(ch.dataUrl, ch.title, ch.filename);
}

function downloadFolderSinglePng(idx) {
    if (!activeArchiveFolderView || !activeArchiveFolderView.charts[idx]) return;
    const ch = activeArchiveFolderView.charts[idx];
    const a = document.createElement('a');
    a.href = ch.dataUrl;
    a.download = ch.filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
}

function downloadAllPngsInActiveFolder() {
    if (!activeArchiveFolderView || !activeArchiveFolderView.charts.length) return;
    activeArchiveFolderView.charts.forEach((ch, idx) => {
        setTimeout(() => downloadFolderSinglePng(idx), idx * 350);
    });
    showToast(`Скачивание ${activeArchiveFolderView.charts.length} картинок PNG...`, 'emerald');
}

function downloadActiveFolderChartsPdf() {
    if (!activeArchiveFolderView || !activeArchiveFolderView.chartsPdfBase64) return;
    downloadPdfFromBase64(activeArchiveFolderView.chartsPdfBase64, activeArchiveFolderView.chartsPdfFilename);
    showToast(`Скачан файл: ${activeArchiveFolderView.chartsPdfFilename}`, 'emerald');
}

async function shareActiveFolderChartsPdf() {
    if (!activeArchiveFolderView || !activeArchiveFolderView.chartsPdfBase64) return;
    await sharePdfToMax(
        activeArchiveFolderView.chartsPdfBase64,
        activeArchiveFolderView.chartsPdfFilename,
        `Графики PlyTap: ${activeArchiveFolderView.path}`
    );
}

// ------------------- 7. ОБНОВЛЁННЫЙ АРХИВ (ГОД -> МЕСЯЦ -> ПАПКИ СМЕН) -------------------

renderSavedFilesList = function() {
    const container = document.getElementById('savedFilesList');
    if (!container) return;

    if (savedFiles.length === 0 && savedShiftReports.length === 0) {
        container.innerHTML = `
            <div class="col-span-full bg-slate-950 border border-slate-800 rounded-2xl p-8 text-center">
                <i class="ph-duotone ph-file-xls text-4xl text-slate-600 mb-2"></i>
                <p class="text-sm font-bold text-slate-300">Нет сформированных файлов и отчётов</p>
                <p class="text-xs text-slate-500 mt-1">Завершите смену или загрузите .xlsx с устройства</p>
            </div>
        `;
        return;
    }

    const monthsRU = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
    const yearsMap = {};

    savedFiles.forEach(f => {
        const meta = parseMetaFromExcelFilename(f.filename, f.date, f.shift);
        const [year, month] = meta.dateISO.split('-');
        const monthKey = `${year}-${month}`;
        const mIdx = Math.max(0, Math.min(11, parseInt(month, 10) - 1));
        const monthFolder = PLYTAP_MONTH_FOLDERS[month] || `${month}_Месяц`;

        if (!yearsMap[year]) yearsMap[year] = {};
        if (!yearsMap[year][monthKey]) {
            yearsMap[year][monthKey] = {
                year,
                month,
                monthKey,
                monthFolder,
                label: `${monthsRU[mIdx]} ${year}`,
                files: []
            };
        }
        yearsMap[year][monthKey].files.push(f);
    });

    const sortedYears = Object.keys(yearsMap).sort((a, b) => b.localeCompare(a));
    let html = '';

    sortedYears.forEach(year => {
        html += `
            <div class="col-span-full flex items-center gap-2.5 mt-2 pt-2 border-t border-slate-800 first:mt-0 first:pt-0 first:border-t-0">
                <div class="w-9 h-9 rounded-xl bg-sky-500/20 border border-sky-500/40 text-sky-400 flex items-center justify-center">
                    <i class="ph-fill ph-folder text-xl"></i>
                </div>
                <h3 class="text-lg sm:text-xl font-black text-white mono">${year} год</h3>
            </div>
        `;

        const sortedMonthKeys = Object.keys(yearsMap[year]).sort((a, b) => b.localeCompare(a));

        sortedMonthKeys.forEach(monthKey => {
            const group = yearsMap[year][monthKey];
            const monthlyRep = savedMonthlyReports.find(m => m.monthKey === monthKey) || null;

            // Шапка папки месяца
            html += `
                <div class="col-span-full flex flex-wrap items-center justify-between gap-2 mt-1 mb-1 border-b border-slate-800 pb-2 pl-2">
                    <div class="flex items-center gap-2.5">
                        <div class="w-8 h-8 rounded-lg bg-amber-500/20 border border-amber-500/40 text-amber-400 flex items-center justify-center">
                            <i class="ph-bold ph-calendar-blank text-lg"></i>
                        </div>
                        <div>
                            <h4 class="text-base sm:text-lg font-black text-white">${group.label}</h4>
                            <span class="text-[11px] text-slate-400 mono">Папка: ${year} / ${group.monthFolder}</span>
                        </div>
                    </div>
                    <span class="px-2.5 py-1 rounded-lg bg-slate-800 border border-slate-700 text-slate-300 text-[11px] font-bold mono">
                        Смен: ${group.files.length}
                    </span>
                </div>
            `;

            // Карточка накопительной папки месяца: 00_Итого_за_месяц_графики
            html += `
                <div class="col-span-full bg-gradient-to-r from-slate-950 via-slate-900 to-slate-950 border-2 border-emerald-500/40 rounded-2xl p-4 flex flex-col lg:flex-row lg:items-center justify-between gap-3 shadow-lg">
                    <div class="flex items-start gap-3">
                        <div class="w-11 h-11 rounded-xl bg-emerald-500/20 border border-emerald-500/40 text-emerald-400 flex items-center justify-center shrink-0">
                            <i class="ph-fill ph-folder-star text-2xl"></i>
                        </div>
                        <div>
                            <div class="flex flex-wrap items-center gap-2">
                                <span class="text-sm sm:text-base font-black text-emerald-300 mono">00_Итого_за_месяц_графики</span>
                                <span class="px-2 py-0.5 rounded-full bg-emerald-500/15 border border-emerald-500/30 text-emerald-300 text-[10px] font-bold">Накопление</span>
                            </div>
                            ${monthlyRep ? `
                                <p class="text-xs text-slate-300 mt-1">
                                    Смен: <strong class="text-white mono">${monthlyRep.shiftsCount}</strong> •
                                    Проверено: <strong class="text-emerald-400 mono">${monthlyRep.totalSheets} л.</strong> •
                                    Брак: <strong class="text-amber-400 mono">${monthlyRep.totalDefects}</strong> •
                                    Годных: <strong class="text-emerald-300 mono">${monthlyRep.goodPct.toFixed(2)}%</strong>
                                </p>
                                <p class="text-[11px] text-slate-500 mt-0.5">Обновлено: ${monthlyRep.updatedAt}</p>
                            ` : `
                                <p class="text-xs text-slate-400 mt-1">Нажмите «Обновить», чтобы собрать общие накопительные графики по всем сменам за ${group.label}</p>
                            `}
                        </div>
                    </div>
                    <div class="flex flex-wrap items-center gap-2">
                        ${monthlyRep ? `
                            <button type="button" onclick="openArchiveFolderModal('month', '${monthKey}')" class="px-3 py-2 rounded-xl bg-amber-500/20 hover:bg-amber-500/30 text-amber-300 border border-amber-500/40 text-xs font-extrabold flex items-center gap-1.5 transition">
                                <i class="ph-fill ph-folder-open text-base"></i>
                                <span>Графики PNG (${(monthlyRep.charts || []).length})</span>
                            </button>
                            <button type="button" onclick="shareMonthChartsPdfToMax('${monthKey}')" class="px-3 py-2 rounded-xl bg-indigo-500/25 hover:bg-indigo-500/35 text-indigo-200 border border-indigo-500/40 text-xs font-extrabold flex items-center gap-1.5 transition">
                                <i class="ph-bold ph-paper-plane-tilt text-base"></i>
                                <span>Графики в MAX (PDF)</span>
                            </button>
                        ` : ''}
                        <button type="button" onclick="generateAndActionMonthlyAnalyticalPdf('${monthKey}', false)" class="px-3 py-2 rounded-xl bg-rose-500/15 hover:bg-rose-500/25 text-rose-300 border border-rose-500/40 text-xs font-extrabold flex items-center gap-1.5 transition">
                            <i class="ph-bold ph-file-pdf text-base"></i>
                            <span>Аналит. PDF месяца</span>
                        </button>
                        <button type="button" onclick="rebuildMonthlyReportInteractive('${monthKey}')" class="px-3 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 text-xs font-bold flex items-center gap-1.5 transition" title="Пересчитать сводку месяца по всем файлам">
                            <i class="ph-bold ph-arrows-clockwise text-base text-emerald-400"></i>
                            <span>Обновить</span>
                        </button>
                    </div>
                </div>
            `;

            // Карточки отдельных смен месяца
            html += group.files.map(f => {
                const rep = findShiftReportForFile(f);
                const meta = parseMetaFromExcelFilename(f.filename, f.date, f.shift);
                const folderMeta = buildFolderMeta(meta.dateISO, meta.shift);
                const goodPctStr = (Number(f.totalSheets) > 0 && f.totalDefects !== '-')
                    ? (((Number(f.totalSheets) - Number(f.totalDefects)) / Number(f.totalSheets)) * 100).toFixed(1) + '%'
                    : '—';

                return `
                    <div class="bg-slate-950 border border-slate-800 hover:border-slate-700 rounded-2xl p-4 flex flex-col justify-between gap-3 shadow transition">
                        <div class="flex items-start justify-between gap-3">
                            <div class="flex items-center gap-3 min-w-0">
                                <div class="w-11 h-11 rounded-xl bg-emerald-500/15 border border-emerald-500/30 flex items-center justify-center text-emerald-400 shrink-0">
                                    <i class="ph-bold ph-file-xls text-2xl"></i>
                                </div>
                                <div class="min-w-0">
                                    <h4 class="font-extrabold text-white text-sm sm:text-base mono break-all">${f.filename}</h4>
                                    <p class="text-[11px] text-amber-300/90 mono truncate mt-0.5">
                                        <i class="ph-fill ph-folder text-xs"></i> ${rep ? rep.fullPath : folderMeta.fullPath}
                                    </p>
                                    <p class="text-xs text-slate-400 mt-0.5">Создан: ${f.createdAt} • Смена ${f.shift}${f.operator ? ' • ' + f.operator : ''}</p>
                                </div>
                            </div>
                            <button type="button" onclick="deleteSavedFile('${f.id}')" class="text-slate-500 hover:text-rose-400 hover:bg-rose-950/50 rounded-lg p-1.5 transition shrink-0" title="Удалить смену и её отчёты из базы">
                                <i class="ph-bold ph-trash text-lg"></i>
                            </button>
                        </div>

                        <div class="grid grid-cols-4 gap-1.5 bg-slate-900 p-2.5 rounded-xl border border-slate-800 text-center text-xs">
                            <div>
                                <span class="text-slate-400 block text-[10px]">Циклов</span>
                                <strong class="text-white mono">${f.cyclesCount || 0}</strong>
                            </div>
                            <div class="border-l border-slate-800">
                                <span class="text-slate-400 block text-[10px]">Проверено</span>
                                <strong class="text-emerald-400 mono">${f.totalSheets || 0}</strong>
                            </div>
                            <div class="border-l border-slate-800">
                                <span class="text-slate-400 block text-[10px]">Брак</span>
                                <strong class="text-amber-400 mono">${f.totalDefects || 0}</strong>
                            </div>
                            <div class="border-l border-slate-800">
                                <span class="text-slate-400 block text-[10px]">Годных</span>
                                <strong class="text-sky-300 mono">${rep ? rep.goodPct.toFixed(1) + '%' : goodPctStr}</strong>
                            </div>
                        </div>

                        <!-- Кнопки Excel -->
                        <div class="flex items-center gap-2">
                            <button type="button" onclick="downloadSavedFile('${f.id}')" class="flex-1 py-2 px-3 rounded-xl bg-emerald-500 hover:bg-emerald-400 text-slate-950 font-extrabold text-xs flex items-center justify-center gap-1.5 shadow transition">
                                <i class="ph-bold ph-download-simple text-base"></i>
                                <span>Excel .xlsx</span>
                            </button>
                            <button type="button" onclick="openShareModal('${f.id}')" class="flex-1 py-2 px-3 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 font-bold text-xs flex items-center justify-center gap-1.5 transition">
                                <i class="ph-bold ph-paper-plane-tilt text-base text-indigo-400"></i>
                                <span>Excel в MAX</span>
                            </button>
                        </div>

                        <!-- Кнопки папки графиков PNG и аналитического PDF -->
                        ${rep ? `
                            <div class="grid grid-cols-3 gap-1.5 pt-1 border-t border-slate-800/80">
                                <button type="button" onclick="openArchiveFolderModal('shift', '${rep.id}')" class="py-2 px-2 rounded-xl bg-amber-500/15 hover:bg-amber-500/25 text-amber-300 border border-amber-500/40 font-extrabold text-[11px] flex items-center justify-center gap-1 transition" title="Открыть папку с PNG-графиками смены">
                                    <i class="ph-fill ph-folder-open text-sm"></i>
                                    <span>PNG (${(rep.charts || []).length})</span>
                                </button>
                                <button type="button" onclick="downloadShiftAnalyticalPdf('${rep.id}')" class="py-2 px-2 rounded-xl bg-rose-500/15 hover:bg-rose-500/25 text-rose-300 border border-rose-500/40 font-extrabold text-[11px] flex items-center justify-center gap-1 transition" title="Скачать аналитический PDF-отчёт смены">
                                    <i class="ph-bold ph-file-pdf text-sm"></i>
                                    <span>Отчёт PDF</span>
                                </button>
                                <button type="button" onclick="shareShiftAnalyticalPdfToMax('${rep.id}')" class="py-2 px-2 rounded-xl bg-indigo-500/20 hover:bg-indigo-500/35 text-indigo-200 border border-indigo-500/40 font-extrabold text-[11px] flex items-center justify-center gap-1 transition" title="Отправить аналитический PDF смены в MAX">
                                    <i class="ph-bold ph-paper-plane-tilt text-sm"></i>
                                    <span>PDF в MAX</span>
                                </button>
                            </div>
                        ` : `
                            <div class="pt-1 border-t border-slate-800/80">
                                <button type="button" onclick="generateReportsForExistingFile('${f.id}')" class="w-full py-2 px-3 rounded-xl bg-amber-500/15 hover:bg-amber-500/25 text-amber-300 border border-amber-500/40 font-extrabold text-xs flex items-center justify-center gap-1.5 transition">
                                    <i class="ph-bold ph-chart-bar text-base"></i>
                                    <span>Сформировать графики PNG и PDF</span>
                                </button>
                            </div>
                        `}
                    </div>
                `;
            }).join('');
        });
    });

    container.innerHTML = html;
};

// Расширенное удаление: удаляет и Excel, и связанные отчёты смены, и обновляет месяц
deleteSavedFile = function(fileId) {
    openCustomDialog('Удаление смены', 'Удалить Excel-файл и все связанные с ним графики/PDF из базы IndexedDB?', async () => {
        const targetFile = savedFiles.find(x => x.id === fileId);
        const linkedRep = targetFile ? findShiftReportForFile(targetFile) : null;
        const monthKey = linkedRep
            ? linkedRep.monthKey
            : (targetFile ? parseMetaFromExcelFilename(targetFile.filename, targetFile.date, targetFile.shift).dateISO.slice(0, 7) : null);

        savedFiles = savedFiles.filter(x => x.id !== fileId);
        await deleteFileRecordFromIndexedDB(fileId);

        if (linkedRep) {
            savedShiftReports = savedShiftReports.filter(r => r.id !== linkedRep.id);
            await idbDelete(IDB_STORE_SHIFT_REPORTS, linkedRep.id);
        }

        if (monthKey) {
            await rebuildMonthlyReportForKey(monthKey, true);
        }

        saveToStorage();
        renderSavedFilesList();
        refreshWorkspaceUI();
        showToast('Смена удалена, сводка месяца пересчитана', 'slate');
    });
};

// После полной загрузки базы отчётов обновляем интерфейс архива и живые графики
window.addEventListener('load', () => {
    setTimeout(async () => {
        if (typeof loadAnalyticsReportsFromDB === 'function') {
            await loadAnalyticsReportsFromDB();
            renderSavedFilesList();
            if (typeof renderLiveShiftCharts === 'function') renderLiveShiftCharts(true);
        }
    }, 450);
});