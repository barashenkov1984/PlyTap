// =========================================================================
// PLYTAP-DB.JS — Хранилище отчетов (Год / Месяц / Дата), JSON-бэкап и MAX
// =========================================================================

const PLYTAP_MONTH_FOLDERS = {
    '01': '01_Январь', '02': '02_Февраль', '03': '03_Март', '04': '04_Апрель',
    '05': '05_Май',    '06': '06_Июнь',    '07': '07_Июль', '08': '08_Август',
    '09': '09_Сентябрь','10': '10_Октябрь','11': '11_Ноябрь','12': '12_Декабрь'
};

let savedShiftReports = [];
let savedMonthlyReports = [];

// 1. Вспомогательные функции работы с таблицами IndexedDB
async function idbGetAll(storeName) {
    const db = await openArchiveDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(storeName, 'readonly');
        const req = tx.objectStore(storeName).getAll();
        req.onsuccess = () => resolve(req.result || []);
        req.onerror = () => reject(req.error);
    });
}

async function idbPut(storeName, record) {
    const db = await openArchiveDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(storeName, 'readwrite');
        const req = tx.objectStore(storeName).put(record);
        req.onsuccess = () => resolve(true);
        req.onerror = () => reject(req.error);
    });
}

async function idbDelete(storeName, key) {
    const db = await openArchiveDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(storeName, 'readwrite');
        const req = tx.objectStore(storeName).delete(key);
        req.onsuccess = () => resolve(true);
        req.onerror = () => reject(req.error);
    });
}

// 2. Построение пути виртуальной папки: Год / 11_Ноябрь / Отчет_22.11.2026_см1
function buildFolderMeta(dateISO, shiftNum) {
    const safeDate = (dateISO && /^\d{4}-\d{2}-\d{2}$/.test(dateISO))
        ? dateISO
        : new Date().toISOString().slice(0, 10);
    const [year, month, day] = safeDate.split('-');
    const monthFolder = PLYTAP_MONTH_FOLDERS[month] || `${month}_Месяц`;
    const monthKey = `${year}-${month}`;
    const dateRu = `${day}.${month}.${year}`;
    const shiftFolder = `Отчет_${dateRu}_см${shiftNum || '1'}`;
    return {
        year,
        month,
        monthKey,
        monthFolder,
        dateRu,
        shiftFolder,
        fullPath: `${year} / ${monthFolder} / ${shiftFolder}`
    };
}

// 3. Загрузка всех отчетов из IndexedDB при старте
async function loadAnalyticsReportsFromDB() {
    try {
        const [shifts, months] = await Promise.all([
            idbGetAll(IDB_STORE_SHIFT_REPORTS),
            idbGetAll(IDB_STORE_MONTHLY_REPORTS)
        ]);
        savedShiftReports = shifts.sort((a, b) =>
            String(b.date || '').localeCompare(String(a.date || '')) ||
            String(b.shift || '').localeCompare(String(a.shift || ''))
        );
        savedMonthlyReports = months.sort((a, b) =>
            String(b.monthKey || '').localeCompare(String(a.monthKey || ''))
        );
    } catch (e) {
        console.error('Ошибка загрузки отчетов из IndexedDB:', e);
    }
}

async function saveShiftReportToDB(reportObj) {
    await idbPut(IDB_STORE_SHIFT_REPORTS, reportObj);
    const idx = savedShiftReports.findIndex(r => r.id === reportObj.id);
    if (idx !== -1) savedShiftReports[idx] = reportObj;
    else savedShiftReports.unshift(reportObj);
}

async function saveMonthlyReportToDB(monthObj) {
    await idbPut(IDB_STORE_MONTHLY_REPORTS, monthObj);
    const idx = savedMonthlyReports.findIndex(r => r.monthKey === monthObj.monthKey);
    if (idx !== -1) savedMonthlyReports[idx] = monthObj;
    else savedMonthlyReports.unshift(monthObj);
}

// 4. Экспорт всей базы в JSON-файл
async function exportFullDatabaseJSON() {
    try {
        showToast('Подготовка резервной копии базы (JSON)...', 'sky');
        const [excelFiles, shiftReports, monthlyReports] = await Promise.all([
            idbGetAll(IDB_STORE_FILES),
            idbGetAll(IDB_STORE_SHIFT_REPORTS),
            idbGetAll(IDB_STORE_MONTHLY_REPORTS)
        ]);

        const backupPayload = {
            app: 'PlyTap',
            schemaVersion: 2,
            exportedAt: new Date().toISOString(),
            operators: authorizedOperators || [],
            activeState: appState || {},
            stores: {
                saved_excel_files: excelFiles,
                shift_reports: shiftReports,
                monthly_reports: monthlyReports
            }
        };

        const jsonStr = JSON.stringify(backupPayload);
        const blob = new Blob([jsonStr], { type: 'application/json;charset=utf-8' });
        const nowStr = new Date().toISOString().slice(0, 10).split('-').reverse().join('.');
        const fileName = `PlyTap_Backup_${nowStr}.json`;

        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 1500);

        showToast(
            `Бэкап сохранен: ${fileName} (Excel: ${excelFiles.length}, Смен: ${shiftReports.length}, Месяцев: ${monthlyReports.length})`,
            'emerald'
        );
    } catch (err) {
        console.error('Ошибка экспорта JSON:', err);
        showToast('Ошибка при экспорте базы JSON', 'amber');
    }
}

// 5. Импорт базы из JSON-файла
function handleImportDatabaseJSON(event) {
    const file = event.target.files ? event.target.files[0] : null;
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (e) => {
        try {
            const parsed = JSON.parse(String(e.target.result || '{}'));
            if (!parsed || parsed.app !== 'PlyTap' || !parsed.stores) {
                showToast('Выбранный файл не является резервной копией PlyTap!', 'amber');
                return;
            }

            // Восстанавливаем операторов
            if (Array.isArray(parsed.operators)) {
                parsed.operators.forEach(op => {
                    if (op && !authorizedOperators.includes(op)) {
                        authorizedOperators.push(op);
                    }
                });
            }

            // Восстанавливаем записи во все 3 таблицы IndexedDB
            const filesArr = parsed.stores.saved_excel_files || [];
            const shiftsArr = parsed.stores.shift_reports || [];
            const monthsArr = parsed.stores.monthly_reports || [];

            for (const f of filesArr) {
                if (f && f.id) await idbPut(IDB_STORE_FILES, f);
            }
            for (const s of shiftsArr) {
                if (s && s.id) await idbPut(IDB_STORE_SHIFT_REPORTS, s);
            }
            for (const m of monthsArr) {
                if (m && m.monthKey) await idbPut(IDB_STORE_MONTHLY_REPORTS, m);
            }

            saveToStorage();
            await loadFilesFromIndexedDB();
            await loadAnalyticsReportsFromDB();
            if (typeof renderSavedFilesList === 'function') renderSavedFilesList();
            if (typeof refreshWorkspaceUI === 'function') refreshWorkspaceUI();

            showToast(
                `База восстановлена! Файлов: ${filesArr.length}, отчетов смен: ${shiftsArr.length}, месяцев: ${monthsArr.length}`,
                'emerald'
            );
        } catch (err) {
            console.error('Ошибка импорта JSON:', err);
            showToast('Не удалось прочитать JSON-файл резервной копии', 'amber');
        } finally {
            event.target.value = '';
        }
    };
    reader.readAsText(file, 'utf-8');
}

// 6. Универсальная отправка PDF в MAX (через Web Share API на планшете или скачивание + MAX Web на ПК)
async function sharePdfToMax(pdfBase64OrBlob, filename, summaryText = '') {
    try {
        let blob;
        if (pdfBase64OrBlob instanceof Blob) {
            blob = pdfBase64OrBlob;
        } else {
            const bytes = base64ToUint8Array(pdfBase64OrBlob);
            blob = new Blob([bytes], { type: 'application/pdf' });
        }

        const fileObj = new File([blob], filename, { type: 'application/pdf' });

        if (navigator.canShare && navigator.canShare({ files: [fileObj] })) {
            try {
                await navigator.share({
                    title: filename,
                    text: summaryText || `Отчет PlyTap: ${filename}`,
                    files: [fileObj]
                });
                return;
            } catch (shareErr) {
                if (shareErr.name === 'AbortError') return;
            }
        }

        // Резервный способ: скачиваем PDF и открываем MAX
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 1500);

        try {
            if (summaryText && navigator.clipboard && navigator.clipboard.writeText) {
                await navigator.clipboard.writeText(summaryText);
            }
        } catch (_) {}

        const isMobile = /Android|webOS|iPhone|iPad|iPod/i.test(navigator.userAgent);
        if (isMobile) {
            showToast(`PDF «${filename}» сохранен в Загрузки! Прикрепите его в приложении MAX.`, 'amber');
        } else {
            window.open('https://web.max.ru/', '_blank', 'noopener');
            showToast(`PDF «${filename}» скачан, открыт MAX.`, 'emerald');
        }
    } catch (err) {
        console.error('Ошибка отправки PDF в MAX:', err);
        showToast('Ошибка при подготовке PDF к отправке', 'amber');
    }
}

// Автозагрузка отчетов после старта страницы
window.addEventListener('load', () => {
    setTimeout(() => {
        loadAnalyticsReportsFromDB();
    }, 300);
});