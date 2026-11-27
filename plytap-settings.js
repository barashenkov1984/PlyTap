// PLYTAP — Справочники технолога и пороги подсветки
const PLYTAP_SETTINGS_KEY = 'plytap_dictionary_settings_v1';

const PLYTAP_FACTORY_SETTINGS = {
    defectTitles: Object.fromEntries(
        EXCEL_COLUMNS
            .filter(c => ALL_DEFECT_KEYS.includes(c.key))
            .map(c => [c.key, c.title])
    ),
    sorts: {
        'ТУ-Э': [...SORTS_BY_TU['ТУ-Э']],
        'ТУ-Р': [...SORTS_BY_TU['ТУ-Р']]
    },
    thresholds: {
        yellowFrom: 6,
        pinkFrom: 11,
        redFrom: 21,
        percentPinkFrom: 5,
        percentRedFrom: 10
    }
};

let plytapSettings = JSON.parse(JSON.stringify(PLYTAP_FACTORY_SETTINGS));

function getPlytapSettings() {
    return JSON.parse(JSON.stringify(plytapSettings));
}

function validatePlytapSettings(settings) {
    if (!settings || !settings.defectTitles || !settings.sorts || !settings.thresholds) {
        throw new Error('Некорректный формат настроек.');
    }

    const defectTitles = {};
    const usedTitles = new Set();
    for (const key of ALL_DEFECT_KEYS) {
        const rawTitle = settings.defectTitles[key];
        const title = typeof rawTitle === 'string' ? rawTitle.trim() : '';
        if (!title || title.length > 80 || /[<>&"'\\\u0000-\u001f\u007f]/.test(title)) {
            throw new Error('Проверьте название дефекта в столбце ' + key + '.');
        }
        const normalized = title.toLocaleLowerCase('ru-RU').replace(/\s+/g, ' ');
        if (usedTitles.has(normalized)) {
            throw new Error('Названия дефектов должны быть различными: ' + title);
        }
        usedTitles.add(normalized);
        defectTitles[key] = title;
    }

    const sorts = {};
    for (const tu of ['ТУ-Э', 'ТУ-Р']) {
        const values = settings.sorts[tu];
        if (!Array.isArray(values) || values.some(value => typeof value !== 'string')) {
            throw new Error('Проверьте список сортов ' + tu + '.');
        }
        const list = values.map(value => value.trim().toUpperCase()).filter(Boolean);
        if (!list.length || list.some(value => !/^[\p{L}\p{N}/._+-]{1,30}$/u.test(value))) {
            throw new Error('Проверьте сорта ' + tu + '. Вводите по одному обозначению на строку, без пробелов.');
        }
        if (new Set(list).size !== list.length) {
            throw new Error('В списке ' + tu + ' есть повторяющиеся сорта.');
        }
        sorts[tu] = list;
    }

    const thresholds = {};
    for (const name of Object.keys(PLYTAP_FACTORY_SETTINGS.thresholds)) {
        const value = settings.thresholds[name];
        if ((typeof value !== 'number' && typeof value !== 'string') || String(value).trim() === '') {
            throw new Error('Заполните все пороги подсветки.');
        }
        thresholds[name] = Number(value);
    }

    const { yellowFrom, pinkFrom, redFrom, percentPinkFrom, percentRedFrom } = thresholds;
    if (![yellowFrom, pinkFrom, redFrom].every(Number.isSafeInteger) ||
        yellowFrom < 2 || pinkFrom <= yellowFrom || redFrom <= pinkFrom) {
        throw new Error('Пороги количества должны быть целыми: 2 ≤ жёлтый < розовый < красный.');
    }
    const validPercent = value => Number.isFinite(value) && value >= 0.1 && value <= 100 &&
        Math.abs(value * 10 - Math.round(value * 10)) < 1e-7;
    if (!validPercent(percentPinkFrom) || !validPercent(percentRedFrom) ||
        percentRedFrom <= percentPinkFrom) {
        throw new Error('Процентные пороги должны возрастать и находиться в диапазоне 0,1–100% (до одного знака после запятой).');
    }

    return { defectTitles, sorts, thresholds };
}

function applyPlytapSettings(settings) {
    plytapSettings = settings;
    for (const column of EXCEL_COLUMNS) {
        if (ALL_DEFECT_KEYS.includes(column.key)) {
            column.title = settings.defectTitles[column.key];
        }
    }
    for (const tu of ['ТУ-Э', 'ТУ-Р']) {
        SORTS_BY_TU[tu] = [...settings.sorts[tu]];
    }
}

function loadPlytapSettings() {
    try {
        const raw = localStorage.getItem(PLYTAP_SETTINGS_KEY);
        if (raw) applyPlytapSettings(validatePlytapSettings(JSON.parse(raw)));
    } catch (error) {
        console.warn('Не удалось загрузить настройки PlyTap:', error);
    }
}

function savePlytapSettings(settings) {
    const validated = validatePlytapSettings(settings);
    localStorage.setItem(PLYTAP_SETTINGS_KEY, JSON.stringify(validated));
    applyPlytapSettings(validated);
}

function getPlytapRenamedDefects() {
    return Object.fromEntries(
        ALL_DEFECT_KEYS
            .filter(key => plytapSettings.defectTitles[key] !== PLYTAP_FACTORY_SETTINGS.defectTitles[key])
            .map(key => [key, plytapSettings.defectTitles[key]])
    );
}

// Общие правила подсветки для предпросмотра и Excel.
function getPlytapCountBand(value) {
    const n = Number(value) || 0;
    const { yellowFrom, pinkFrom, redFrom } = plytapSettings.thresholds;
    if (n < 1) return -1;
    if (n < yellowFrom) return 0;
    if (n < pinkFrom) return 1;
    if (n < redFrom) return 2;
    return 3;
}

function getPlytapPercentBand(value) {
    const rounded = Math.round((Number(value) || 0) * 10) / 10;
    const { percentPinkFrom, percentRedFrom } = plytapSettings.thresholds;
    if (rounded >= percentRedFrom) return 2;
    if (rounded >= percentPinkFrom) return 1;
    return 0;
}

function getPlytapDefectCFRules() {
    const { yellowFrom, pinkFrom, redFrom } = plytapSettings.thresholds;
    const fill = argb => ({ fill: { type: 'pattern', pattern: 'solid', bgColor: { argb } } });
    return [
        { type: 'cellIs', operator: 'between', formulae: [1, yellowFrom - 1], style: fill('FFC6EFCE') },
        { type: 'cellIs', operator: 'between', formulae: [yellowFrom, pinkFrom - 1], style: fill('FFFFEB9C') },
        { type: 'cellIs', operator: 'between', formulae: [pinkFrom, redFrom - 1], style: fill('FFFFC7CE') },
        { type: 'cellIs', operator: 'greaterThanOrEqual', formulae: [redFrom], style: fill('FFFF4D4D') }
    ];
}

function getPlytapPercentCFRules(row) {
    const { percentPinkFrom, percentRedFrom } = plytapSettings.thresholds;
    const cell = 'C' + row;
    const font = argb => ({ font: { bold: true, italic: true, color: { argb } } });
    return [
        {
            type: 'expression',
            formulae: ['AND(ROUND(' + cell + ',1)>=' + percentPinkFrom + ',ROUND(' + cell + ',1)<' + percentRedFrom + ')'],
            style: font('FFFF6699')
        },
        {
            type: 'expression',
            formulae: ['ROUND(' + cell + ',1)>=' + percentRedFrom],
            style: font('FFFF0000')
        }
    ];
}

function refreshPlytapSettingsViews() {
    refreshWorkspaceUI();
    if (typeof renderLiveShiftCharts === 'function') renderLiveShiftCharts(true);
    renderMasterAccumulatedTable();
}

function renderPlytapSettings() {
    const screen = document.getElementById('tab-settings');
    if (!screen) return;

    const inputClass = 'w-full bg-slate-950 border border-slate-700 rounded-xl px-3 py-2.5 text-sm text-white focus:outline-none focus:border-amber-400 select-text';
    screen.innerHTML = `
        <form id="plytapSettingsForm" class="flex flex-col gap-4 pb-4">
            <div class="bg-slate-900 border border-slate-800 rounded-2xl p-4 sm:p-5">
                <h2 class="text-lg sm:text-xl font-extrabold text-white flex items-center gap-2">
                    <i class="ph-duotone ph-gear-six text-amber-400 text-2xl"></i>
                    Настройки справочников
                </h2>
                <p class="mt-1 text-xs sm:text-sm text-slate-400">
                    Изменения сохраняются на этом устройстве. Названия привязаны к столбцам Excel.
                </p>
            </div>

            <div class="bg-slate-900 border border-slate-800 rounded-2xl p-4 sm:p-5">
                <h3 class="text-base font-extrabold text-white mb-3">1. Дефекты</h3>
                <div class="grid grid-cols-1 lg:grid-cols-2 gap-4">
                    <div>
                        <h4 class="text-sm font-bold text-amber-400 mb-2">Технологические дефекты (C–N)</h4>
                        <div id="plytapSettingsDefectsCN" class="grid gap-2"></div>
                    </div>
                    <div>
                        <h4 class="text-sm font-bold text-teal-400 mb-2">Пороки древесины (P–Y)</h4>
                        <div id="plytapSettingsDefectsPY" class="grid gap-2"></div>
                    </div>
                </div>
            </div>

            <div class="bg-slate-900 border border-slate-800 rounded-2xl p-4 sm:p-5">
                <h3 class="text-base font-extrabold text-white mb-1">2. Сорта фанеры</h3>
                <p class="text-xs text-slate-400 mb-3">Один сорт на строку. Добавьте, измените или удалите нужное обозначение.</p>
                <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <label class="text-sm font-bold text-sky-400">
                        ТУ-Э
                        <textarea id="plytapSettingsSortE" class="${inputClass} mt-1 mono" rows="10" required spellcheck="false"></textarea>
                    </label>
                    <label class="text-sm font-bold text-sky-400">
                        ТУ-Р
                        <textarea id="plytapSettingsSortR" class="${inputClass} mt-1 mono" rows="10" required spellcheck="false"></textarea>
                    </label>
                </div>
            </div>

            <div class="bg-slate-900 border border-slate-800 rounded-2xl p-4 sm:p-5">
                <h3 class="text-base font-extrabold text-white mb-1">3. Пороги подсветки</h3>
                <p class="text-xs text-slate-400 mb-3">Укажите минимальное количество дефектов для каждого из следующих цветов.</p>
                <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
                    <label class="text-sm font-bold text-amber-300">
                        <span class="inline-block w-3 h-3 rounded-sm bg-[#FFEB9C] mr-1"></span> Жёлтый, от
                        <input id="plytapSettingsYellow" type="number" min="2" max="9007199254740991" step="1" required class="${inputClass} mt-1 mono">
                    </label>
                    <label class="text-sm font-bold text-pink-300">
                        <span class="inline-block w-3 h-3 rounded-sm bg-[#FFC7CE] mr-1"></span> Розовый, от
                        <input id="plytapSettingsPink" type="number" min="3" max="9007199254740991" step="1" required class="${inputClass} mt-1 mono">
                    </label>
                    <label class="text-sm font-bold text-rose-300">
                        <span class="inline-block w-3 h-3 rounded-sm bg-[#FF4D4D] mr-1"></span> Красный, от
                        <input id="plytapSettingsRed" type="number" min="4" max="9007199254740991" step="1" required class="${inputClass} mt-1 mono">
                    </label>
                </div>
                <p class="text-xs text-emerald-400 mt-2" id="plytapSettingsGreenHint"></p>
                <div class="mt-4 border-t border-slate-800 pt-4">
                    <h4 class="text-sm font-bold text-white mb-1">Строка «%» в сводной таблице</h4>
                    <p class="text-xs text-slate-400 mb-3">Доля дефекта от общего числа дефектов, в процентах.</p>
                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        <label class="text-sm font-bold text-pink-300">
                            Розовый текст, от %
                            <input id="plytapSettingsPctPink" type="number" min="0.1" max="100" step="0.1" required class="${inputClass} mt-1 mono">
                        </label>
                        <label class="text-sm font-bold text-rose-300">
                            Красный текст, от %
                            <input id="plytapSettingsPctRed" type="number" min="0.1" max="100" step="0.1" required class="${inputClass} mt-1 mono">
                        </label>
                    </div>
                </div>
            </div>

            <div class="flex flex-wrap gap-2 items-center justify-between">
                <button type="button" id="plytapSettingsReset" class="px-4 py-3 rounded-xl bg-slate-900 hover:bg-slate-800 border border-slate-700 text-slate-200 font-bold text-sm">
                    <i class="ph-bold ph-arrow-counter-clockwise"></i> Заводские значения
                </button>
                <button type="submit" class="px-5 py-3 rounded-xl bg-emerald-500 hover:bg-emerald-400 text-slate-950 font-extrabold text-sm">
                    <i class="ph-bold ph-check"></i> Сохранить настройки
                </button>
            </div>
        </form>
    `;

    const settings = getPlytapSettings();
    for (const key of ALL_DEFECT_KEYS) {
        const groupId = DEFECT_KEYS_CN.includes(key) ? 'plytapSettingsDefectsCN' : 'plytapSettingsDefectsPY';
        const label = document.createElement('label');
        label.className = 'flex items-center gap-2';
        const badge = document.createElement('span');
        badge.className = 'w-9 shrink-0 text-center rounded-lg bg-slate-800 text-slate-300 py-2 text-xs font-extrabold mono';
        badge.textContent = key;
        const input = document.createElement('input');
        input.className = inputClass;
        input.id = 'plytapSettingsDefect_' + key;
        input.value = settings.defectTitles[key];
        input.type = 'text';
        input.required = true;
        input.maxLength = 80;
        input.setAttribute('aria-label', 'Название дефекта, столбец ' + key);
        label.append(badge, input);
        document.getElementById(groupId).appendChild(label);
    }

    document.getElementById('plytapSettingsSortE').value = settings.sorts['ТУ-Э'].join('\n');
    document.getElementById('plytapSettingsSortR').value = settings.sorts['ТУ-Р'].join('\n');

    const fields = {
        yellowFrom: 'plytapSettingsYellow',
        pinkFrom: 'plytapSettingsPink',
        redFrom: 'plytapSettingsRed',
        percentPinkFrom: 'plytapSettingsPctPink',
        percentRedFrom: 'plytapSettingsPctRed'
    };
    for (const [key, id] of Object.entries(fields)) {
        document.getElementById(id).value = settings.thresholds[key];
    }

    const yellow = document.getElementById('plytapSettingsYellow');
    const updateGreenHint = () => {
        const upper = Number(yellow.value) - 1;
        document.getElementById('plytapSettingsGreenHint').textContent =
            Number.isSafeInteger(upper) && upper >= 1 ? 'Зелёная подсветка: 1–' + upper : '';
    };
    yellow.addEventListener('input', updateGreenHint);
    updateGreenHint();

    document.getElementById('plytapSettingsForm').addEventListener('submit', event => {
        event.preventDefault();
        const draft = {
            defectTitles: Object.fromEntries(ALL_DEFECT_KEYS.map(key => [
                key, document.getElementById('plytapSettingsDefect_' + key).value
            ])),
            sorts: {
                'ТУ-Э': document.getElementById('plytapSettingsSortE').value.split(/\r?\n/),
                'ТУ-Р': document.getElementById('plytapSettingsSortR').value.split(/\r?\n/)
            },
            thresholds: Object.fromEntries(Object.entries(fields).map(([key, id]) => [
                key, document.getElementById(id).value
            ]))
        };

        try {
            savePlytapSettings(draft);
        } catch (error) {
            showToast(error.message || 'Не удалось сохранить настройки.', 'amber');
            return;
        }
        renderPlytapSettings();
        refreshPlytapSettingsViews();
        showToast('Настройки сохранены.', 'emerald');
    });

    document.getElementById('plytapSettingsReset').addEventListener('click', () => {
        openCustomDialog(
            'Восстановить заводские настройки?',
            'Названия дефектов, списки сортов и пороги подсветки вернутся к исходным значениям.',
            () => {
                try {
                    savePlytapSettings(PLYTAP_FACTORY_SETTINGS);
                } catch (error) {
                    showToast(error.message || 'Не удалось восстановить настройки.', 'amber');
                    return;
                }
                renderPlytapSettings();
                refreshPlytapSettingsViews();
                showToast('Заводские настройки восстановлены.', 'emerald');
            }
        );
    });
}
