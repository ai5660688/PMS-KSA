// Daily Fit-up JavaScript functionality
(function () {
    'use strict';

    const dailyFitupConfig = window.dailyFitupConfig || {};
    const getEndpoint = (name) => (dailyFitupConfig.endpoints && dailyFitupConfig.endpoints[name]) || '';
    const resolvedEndpoints = {
        rfiDateTime: getEndpoint('rfiDateTime'),
        sheetsForLayout: getEndpoint('sheetsForLayout'),
        addRow: getEndpoint('addRow'),
        deleteRow: getEndpoint('deleteRow'),
        bulkUpdate: getEndpoint('bulkUpdate'),
        scheduleThickness: getEndpoint('scheduleThickness'),
        schedulesForDiameter: getEndpoint('schedulesForDiameter')
    };
    const {
        rfiDateTime: rfiDateTimeEndpoint,
        sheetsForLayout: sheetsForLayoutEndpoint,
        addRow: addRowEndpoint,
        deleteRow: deleteRowEndpoint,
        bulkUpdate: bulkUpdateEndpoint,
        scheduleThickness: scheduleThicknessEndpoint,
        schedulesForDiameter: schedulesForDiameterEndpoint
    } = resolvedEndpoints;
    const configLineClass = (dailyFitupConfig.lineClass || '').trim();

    // Global variables
    let pendingDeleteRow = null;
    let __supersedeResolve = null;
    const __rfiCache = new Map();
    const wpsState = new WeakMap();
    let frozenMeta = [];

    // Initialize on DOM ready
    document.addEventListener('DOMContentLoaded', function () {
        initializeUCButtons();
        initializeWeldTypeFilters();
        initializeAutoSubmit();
        initializeSheetLoading();
        initializePendingNoData();
        initializeSearchableSelects();
        initializeMainFunctionality();
        initializeConfirmAllModal();
        initializeScheduleMenus();
        initializeFilterToggle();

        // Apply weld type filter
        if (window.applyWeldTypeFilter) window.applyWeldTypeFilter();

        // Populate full RFI lists for all rows on initial load
        if (typeof window.reloadRfiListsForAllRows === 'function') window.reloadRfiListsForAllRows();
    });

    // ==================== UC Button Management ====================
    function initializeUCButtons() {
        // Update dynamic button labels when Fit-up Date changes
        const fitDate = document.getElementById('fitupDateInput');
        let t = null;
        let baselineTime = '';

        function extractTime(val) {
            const m = typeof val === 'string' ? val.match(/T(\d{2}:\d{2}(?::\d{2})?)/) : null;
            return m ? m[1] : '';
        }

        async function refreshButtons() {
            const st = await uc.checkExists();
            uc.setButtonsText(st);
        }

        function schedule() {
            if (t) clearTimeout(t);
            t = setTimeout(refreshButtons, 250);
        }

        function isComplete(val) {
            return typeof val === 'string' && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(?::\d{2}(?:\.\d{1,3})?)?$/.test(val);
        }

        function tryClose(prevTime) {
            if (fitDate && isComplete(fitDate.value)) {
                const nowTime = extractTime(fitDate.value);
                if (nowTime && nowTime !== (prevTime ?? '')) {
                    try { fitDate.blur(); } catch { }
                }
            }
        }

        if (fitDate) {
            fitDate.addEventListener('focus', () => { baselineTime = extractTime(fitDate.value); });
            fitDate.addEventListener('change', () => { schedule(); tryClose(baselineTime); if (typeof window.reloadRfiListsForAllRows === 'function') window.reloadRfiListsForAllRows(); });
            fitDate.addEventListener('input', () => { schedule(); tryClose(baselineTime); });
            fitDate.addEventListener('keydown', (e) => {
                if (e.key === 'Enter' || e.key === 'Escape') { tryClose(baselineTime); }
            });
        }

        // Sync date filter and header datetime for Date view
        const dateFilter = document.querySelector('input[name="fitupDateFilter"][type="date"]');
        if (dateFilter) {
            function setDateFilterFromHeader() {
                const hdr = document.getElementById('fitupDateInput');
                if (!hdr) return;
                const v = hdr.value || '';
                const datePart = v.split('T')[0] || '';
                if (datePart && dateFilter.value !== datePart) {
                    try { dateFilter.value = datePart; } catch { }
                }
            }

            function setHeaderFromDateFilter() {
                const hdr = document.getElementById('fitupDateInput');
                if (!hdr) return;
                const d = dateFilter.value || '';
                if (!d) return;
                const timeMatch = (hdr.value || '').match(/T(\d{2}:\d{2}(?::\d{2})?)/);
                const timePart = timeMatch ? timeMatch[1] : '00:00';
                const newVal = d + 'T' + timePart;
                if (hdr.value !== newVal) {
                    hdr.value = newVal;
                    try { hdr.dispatchEvent(new Event('change', { bubbles: true })); } catch { }
                }
            }

            try { setDateFilterFromHeader(); } catch { }

            dateFilter.addEventListener('change', () => {
                try { setHeaderFromDateFilter(); dateFilter.blur(); } catch { }
            });

            dateFilter.addEventListener('keydown', (e) => {
                if (e.key === 'Enter' || e.key === 'Escape') {
                    try { setHeaderFromDateFilter(); dateFilter.blur(); } catch { }
                }
            });

            const hdrInit = document.getElementById('fitupDateInput');
            if (hdrInit) {
                hdrInit.addEventListener('change', () => { try { setDateFilterFromHeader(); } catch { } });
            }
        }

        // Initial button state
        refreshButtons();

        // Set date from RFI selection
        const rfiSel = document.getElementById('rfiSelect');
        async function setDateFromRfi() {
            if (!rfiSel) return;
            const id = parseInt(rfiSel.value, 10);
            if (!id || isNaN(id)) return;
            if (!rfiDateTimeEndpoint) return;
            try {
                const url = `${rfiDateTimeEndpoint}?id=${encodeURIComponent(id)}`;
                const res = await fetch(url, { headers: { 'Accept': 'application/json' } });
                if (!res.ok) return;
                const data = await res.json();
                if (data && data.iso) {
                    const input = document.getElementById('fitupDateInput');
                    if (input) {
                        input.value = data.iso;
                        input.dispatchEvent(new Event('change', { bubbles: true }));
                    }
                }
            } catch { }
        }

        rfiSel?.addEventListener('change', setDateFromRfi);

        // Complete button handler
        document.getElementById('completeBtn')?.addEventListener('click', () =>
            uc.post(uc.endpoints.complete, 'completed')
        );

        // Confirm button handler with validation
        document.getElementById('confirmBtn')?.addEventListener('click', async function () {
            try {
                const rows = Array.from(document.querySelectorAll('#fitupTable tbody tr'));
                if (!rows || rows.length === 0) {
                    showStatus('No rows loaded to confirm.', false);
                    return;
                }

                const unconfirmedRows = rows.filter(r => {
                    const confirmedCb = r.querySelector('input[data-name="FitupConfirmed"]');
                    return !(confirmedCb && confirmedCb.checked === true);
                });

                if (unconfirmedRows.length > 0) {
                    const count = unconfirmedRows.length;
                    const msg = `There ${count === 1 ? 'is' : 'are'} ${count} joint${count === 1 ? '' : 's'} that ${count === 1 ? 'is' : 'are'} not marked as Confirmed.\n\nChoose "Confirm All" to mark every joint as Confirmed before using 'Daily Fit-up Confirmed', or select "Cancel" to review the table.`;
                    const proceed = await openConfirmAllModal(msg);
                    if (!proceed) {
                        showStatus('Confirmation aborted. Confirm every joint first.', false);
                        return;
                    }
                    const changed = markRowsConfirmed(unconfirmedRows);
                    if (changed > 0) {
                        if (typeof window.saveAll === 'function') {
                            await window.saveAll();
                        }
                    }
                }

                // Ensure header Fit-up Date aligns with displayed rows (needed for Report view so
                // CompleteDailyFitup uses the correct day and inserts the UC record / email).
                const fitHdr = document.getElementById('fitupDateInput');
                if (fitHdr && (!fitHdr.value || (typeof window.headerView === 'string' && window.headerView.toUpperCase() === 'REPORT'))) {
                    const firstRowDate = [...document.querySelectorAll('#fitupTable tbody input[data-name="FitupDate"]')]
                        .map(inp => (inp.value || '').trim())
                        .find(v => v);
                    if (firstRowDate) {
                        fitHdr.value = firstRowDate;
                        try { fitHdr.dispatchEvent(new Event('change', { bubbles: true })); } catch { }
                    }
                }

                await uc.post(uc.endpoints.confirm, 'confirmed');
            } catch (e) {
                console.error('Confirm pre-check error', e);
                showStatus('Error during confirm validation', false);
            }
        });

        // Re-evaluate buttons on project or location changes
        const proj = document.getElementById('projectSelect');
        const loc = document.getElementById('locationSelect');
        proj?.addEventListener('change', schedule);
        loc?.addEventListener('change', schedule);
    }

    // ==================== Weld Type Filters ====================
    function initializeWeldTypeFilters() {
        const allBtn = document.getElementById('wfAll');
        const noneBtn = document.getElementById('wfNone');
        const defaultBtn = document.getElementById('wfDefault');

        function storageKey() {
            const pid = document.getElementById('projectSelect')?.value || '0';
            const view = window.headerView || 'DWG';
            return `PMS:DailyFitup:WeldType:${pid}:${view}`;
        }

        function selectedValues() {
            return Array.from(document.querySelectorAll('#weldFilters input.wt-filter:checked'))
                .map(cb => (cb.value || '').trim());
        }

        function setSelections(values) {
            const set = new Set((values || []).map(v => String(v)));
            const all = document.querySelectorAll('#weldFilters input.wt-filter');
            all.forEach(cb => { cb.checked = set.has(cb.value); });
            return all;
        }

        function save() {
            try {
                localStorage.setItem(storageKey(), JSON.stringify(selectedValues()));
            } catch { }
        }

        function restore() {
            try {
                const raw = localStorage.getItem(storageKey());
                if (!raw) return false;
                const vals = JSON.parse(raw);
                if (!Array.isArray(vals)) return false;
                return setSelections(vals);
            } catch { return false; }
        }

        function apply() {
            const selected = Array.from(document.querySelectorAll('#weldFilters input.wt-filter:checked'))
                .map(cb => (cb.value || '').toUpperCase().trim());
            const rows = document.querySelectorAll('#fitupTable tbody tr');
            rows.forEach(r => {
                const wtSel = r.querySelector('select[data-name="WeldType"]');
                const wt = ((wtSel?.value) || '').toUpperCase().trim();
                r.style.display = (selected.length === 0 || selected.includes(wt)) ? '' : 'none';
            });
        }

        window.applyWeldTypeFilter = apply;
        window.saveWeldTypeFilterSelection = save;

        // Immediate restore before content flashes default checked
        const hadSavedEarly = restore();
        if (hadSavedEarly) { apply(); }

        // Update on checkbox change + persist
        document.addEventListener('change', e => {
            if (e.target && e.target.matches('#weldFilters input.wt-filter')) {
                save();
                apply();
            }
        });

        // Button handlers
        allBtn?.addEventListener('click', () => {
            document.querySelectorAll('#weldFilters input.wt-filter').forEach(c => c.checked = true);
            save();
            apply();
        });

        noneBtn?.addEventListener('click', () => {
            document.querySelectorAll('#weldFilters input.wt-filter').forEach(c => c.checked = false);
            save();
            apply();
        });

        defaultBtn?.addEventListener('click', () => {
            setSelections(defaultWeldTypes);
            save();
            apply();
        });
    }

    // ==================== Auto Submit ====================
    function initializeAutoSubmit() {
        const form = document.getElementById('filterForm');
        const doLoadInput = document.getElementById('doLoadInput');
        const loadBtn = document.getElementById('loadBtn');

        function submitForm() {
            if (!form) return;
            if (typeof form.requestSubmit === 'function') {
                form.requestSubmit();
            } else {
                form.submit();
            }
        }

        function submitWithLoadFlag() {
            if (!form) return;
            if (doLoadInput) doLoadInput.value = '1';
            submitForm();
        }

        function submitForHeaderRefresh() {
            if (!form) return;
            if (doLoadInput) doLoadInput.value = '0';
            submitForm();
        }

        if (loadBtn) {
            loadBtn.addEventListener('click', (e) => {
                e.preventDefault();
                submitWithLoadFlag();
            });
        }

        const headerControlIds = ['projectSelect', 'headerViewSelect', 'locationSelect'];
        headerControlIds.forEach((id) => {
            const ctrl = document.getElementById(id);
            if (!ctrl) return;
            ctrl.addEventListener('change', submitForHeaderRefresh);
        });

        if (form) {
            form.addEventListener('submit', () => {
                if (!doLoadInput) return;
                if (doLoadInput.value !== '0') {
                    doLoadInput.value = '1';
                }
            });
        }
    }

    // ==================== Sheet Loading ====================
    function initializeSheetLoading() {
        const layoutSel = document.getElementById('layoutSelect');
        const sheetSel = document.getElementById('sheetSelect');
        const projectSel = document.getElementById('projectSelect');
        if (!layoutSel || !sheetSel || !projectSel) return;

        function getSheetEndpoint() {
            return sheetsForLayoutEndpoint
                || (window.dailyFitupConfig && window.dailyFitupConfig.endpoints && window.dailyFitupConfig.endpoints.sheetsForLayout)
                || '';
        }

        async function loadSheets() {
            const endpoint = getSheetEndpoint();
            if (!endpoint) return;

            const layout = (layoutSel.value || '').trim();
            const prevSelected = sheetSel.value || '';

            if (!layout) {
                while (sheetSel.options.length > 0) sheetSel.remove(0);
                return;
            }

            const projectId = (projectSel.value || '').trim();
            if (!projectId) {
                sheetSel.disabled = false;
                return;
            }
            sheetSel.disabled = true;

            try {
                const url = `${endpoint}?projectId=${encodeURIComponent(projectId)}&layout=${encodeURIComponent(layout)}`;
                const res = await fetch(url, { headers: { 'Accept': 'application/json' } });
                if (!res.ok) throw new Error(`HTTP ${res.status}`);
                const data = await res.json();

                if (Array.isArray(data)) {
                    // Clear existing options only after successful fetch
                    while (sheetSel.options.length > 0) sheetSel.remove(0);

                    data.forEach(s => {
                        const opt = document.createElement('option');
                        opt.value = s;
                        opt.textContent = s;
                        sheetSel.appendChild(opt);
                    });

                    if (prevSelected && Array.from(sheetSel.options).some(o => o.value === prevSelected)) {
                        sheetSel.value = prevSelected;
                    } else if (data.length > 0) {
                        // Auto-select the first real sheet when the previous selection is no longer available
                        sheetSel.value = data[0];
                    } else {
                        sheetSel.selectedIndex = 0;
                    }

                    try { sheetSel.dispatchEvent(new Event('change', { bubbles: true })); } catch { /* ignore */ }
                }
            } catch {
                // Handle error silently - preserve existing options
            } finally {
                sheetSel.disabled = false;
            }
        }

        layoutSel.addEventListener('change', loadSheets);
        projectSel.addEventListener('change', () => {
            while (sheetSel.options.length > 0) sheetSel.remove(0);
            // Reload sheets for the current layout after clearing
            if ((layoutSel.value || '').trim()) {
                loadSheets();
            }
        });

        // Initial population: if layout is selected, ensure sheets are loaded.
        // If the server already rendered sheet options, keep them; otherwise fetch via AJAX.
        if (!layoutSel.value && layoutSel.options.length > 0) {
            const first = Array.from(layoutSel.options).find(o => (o.value || '').trim());
            if (first) {
                layoutSel.value = first.value;
            }
        }
        const hasServerSheets = sheetSel.options.length > 0 && Array.from(sheetSel.options).some(o => (o.value || '').trim());
        if (layoutSel.value && !hasServerSheets) {
            loadSheets();
        }
    }

    // ==================== Pending No Data ====================
    function initializePendingNoData() {
        function showPendingNoData(resetDate) {
            const sc = document.getElementById('scrollSync');
            if (sc) { sc.style.display = 'none'; }

            const bs = document.getElementById('bottomScroll');
            if (bs) { bs.style.display = 'none'; }

            const msg = document.getElementById('statusMsg');
            if (msg) { msg.textContent = ''; }

            const nd = document.getElementById('clientNoData');
            if (nd) { nd.style.display = 'block'; }

            if (resetDate) {
                const fitHdr = document.getElementById('fitupDateInput');
                if (fitHdr && typeof window.__initialHeaderFitupDate !== 'undefined') {
                    fitHdr.value = window.__initialHeaderFitupDate;
                    fitHdr.dispatchEvent(new Event('change', { bubbles: true }));
                }
            }
        }

        // Capture initial date
        (function captureInitial() {
            const fitHdr = document.getElementById('fitupDateInput');
            if (fitHdr) {
                window.__initialHeaderFitupDate = fitHdr.value;
            }
        })();

        const dynamicFilterIds = ['projectSelect', 'headerViewSelect', 'locationSelect', 'layoutSelect', 'sheetSelect'];
        dynamicFilterIds.forEach(id => {
            const sel = document.getElementById(id);
            if (!sel) return;
            sel.addEventListener('change', () => showPendingNoData(true));
        });

        const dateFilter = document.querySelector('input[name="fitupDateFilter"][type="date"]');
        dateFilter?.addEventListener('change', () => showPendingNoData(true));

        if (window.headerView === 'Date') {
            const headerFitupDate = document.getElementById('fitupDateInput');
            if (headerFitupDate) {
                const handleHeaderDateEdit = () => showPendingNoData(false);
                headerFitupDate.addEventListener('input', handleHeaderDateEdit);
                headerFitupDate.addEventListener('change', handleHeaderDateEdit);
            }
        }

        const reportFilter = document.querySelector('input[name="fitupReportFilter"]');
        if (reportFilter) {
            const handleReportFilterEdit = () => showPendingNoData(true);
            reportFilter.addEventListener('input', handleReportFilterEdit);
            reportFilter.addEventListener('change', handleReportFilterEdit);
        }
    }

    // ==================== Searchable Selects ====================
    function initializeSearchableSelects() {
        function makeSearchableSelect(sel) {
            if (!sel || sel.dataset.searchable) return;
            let buffer = '';
            let last = 0;

            sel.addEventListener('keydown', (e) => {
                const now = Date.now();
                if (now - last > 800) buffer = '';
                last = now;

                if (e.key === 'Backspace') {
                    buffer = buffer.slice(0, -1);
                    e.preventDefault();
                } else if (e.key && e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
                    buffer += e.key.toLowerCase();
                } else {
                    return;
                }

                const opts = Array.from(sel.options);
                if (opts.length === 0) return;

                const start = sel.selectedIndex >= 0 ? sel.selectedIndex : -1;
                const findIdx = (from) => {
                    for (let i = from + 1; i < opts.length; i++) {
                        if ((opts[i].text || '').toLowerCase().includes(buffer)) return i;
                    }
                    for (let i = 0; i <= from; i++) {
                        if ((opts[i].text || '').toLowerCase().includes(buffer)) return i;
                    }
                    return -1;
                };

                const idx = findIdx(start);
                if (idx >= 0) {
                    sel.selectedIndex = idx;
                    sel.dispatchEvent(new Event('change', { bubbles: true }));
                }
            });

            sel.dataset.searchable = '1';
        }

        window.__makeSearchableSelect = makeSearchableSelect;

        ['layoutSelect', 'sheetSelect', 'tackerSelect', 'projectSelect'].forEach(id => {
            const el = document.getElementById(id);
            if (el) makeSearchableSelect(el);
        });
    }

    // ==================== Main Functionality ====================
    function initializeMainFunctionality() {
        // Status display
        window.showStatus = function (msg, ok) {
            const s = document.getElementById('statusMsg');
            s.textContent = msg;
            s.style.color = ok ? '#176d8a' : '#b40000';
            if (ok) {
                setTimeout(() => {
                    if (s.textContent === msg) s.textContent = '';
                }, 3000);
            }
        };

        // Mark row as dirty
        function markRowDirty(tr) {
            if (!tr.dataset.dirty) {
                tr.dataset.dirty = '1';
                tr.classList.add('dirty');
            }
        }

        // Initialize row baseline
        function initRowBaseline(tr, markDirtyInitially) {
            const fields = tr.querySelectorAll('input[data-name], select[data-name]');
            fields.forEach(f => {
                if (!f.dataset.orig) {
                    f.dataset.orig = f.type === 'checkbox' ? (f.checked ? '1' : '0') : (f.value ?? '');
                    const evt = (f.tagName === 'SELECT' ? 'change' : 'input');
                    f.addEventListener(evt, () => {
                        markRowDirty(tr);
                        if (f.matches && f.matches('.wps-select, input[data-name="Wps"], .rfi-select')) {
                            f.dataset.userSelected = '1';
                        }
                    });
                    if (f.type === 'checkbox') f.addEventListener('change', () => markRowDirty(tr));
                }
            });
            if (markDirtyInitially) markRowDirty(tr);
        }

        // Collect DTO from row
        window.collectDto = function (tr) {
            const id = tr.getAttribute('data-id');
            const dto = { JointId: id ? parseInt(id, 10) : 0 };
            const inputs = tr.querySelectorAll('input[data-name], select[data-name]');

            inputs.forEach(i => {
                const name = i.getAttribute('data-name');
                let val = i.tagName === 'SELECT' ? i.value : (i.type === 'checkbox' ? i.checked : i.value);
                if (i.type !== 'checkbox' && typeof val === 'string' && val.trim() === '') val = null;
                dto[name] = val;
            });

            // Parse numeric fields
            ['DiaIn', 'OlDia', 'OlThick'].forEach(n => {
                if (Object.prototype.hasOwnProperty.call(dto, n)) {
                    const raw = dto[n];
                    if (raw === null || raw === undefined) {
                        dto[n] = null;
                    } else if (typeof raw === 'string') {
                        const v = parseFloat(raw);
                        dto[n] = isNaN(v) ? null : v;
                    }
                }
            });

            // Parse integer fields
            ['RfiId'].forEach(n => {
                if (Object.prototype.hasOwnProperty.call(dto, n)) {
                    const raw = dto[n];
                    if (raw === null || raw === undefined || raw === '') {
                        dto[n] = null;
                    } else if (typeof raw === 'string') {
                        const v = parseInt(raw, 10);
                        dto[n] = isNaN(v) ? null : v;
                    }
                }
            });

            // Handle FitupDate
            ['FitupDate'].forEach(n => {
                if (Object.prototype.hasOwnProperty.call(dto, n)) {
                    const raw = dto[n];
                    dto[n] = typeof raw === 'string' && raw.trim().length > 0 ? raw : null;
                }
            });

            // WPS handling
            const sel = tr.querySelector('select[data-name="Wps"]');
            if (sel) {
                dto.Wps = sel.value || null;
                const opt = sel.options[sel.selectedIndex];
                if (opt) {
                    const idAttr = opt.getAttribute('data-wps-id');
                    if (idAttr) {
                        const idNum = parseInt(idAttr, 10);
                        if (!isNaN(idNum) && idNum > 0) dto.WpsId = idNum;
                    }
                }
            } else {
                const inp = tr.querySelector('input[data-name="Wps"]');
                if (inp) dto.Wps = inp.value || null;
            }

            if (dto.Wps && dto.Wps.includes(':')) dto.Wps = dto.Wps.split(':')[0];

            return dto;
        };

        // Check if row is LET weld type
        function isLetWeldType(tr) {
            const wtSel = tr.querySelector('select[data-name="WeldType"]');
            return wtSel && wtSel.value && wtSel.value.trim().toUpperCase() === 'LET';
        }

        // Check if row is TH weld type
        function isThWeldType(tr) {
            const wtSel = tr.querySelector('select[data-name="WeldType"]');
            return !!(wtSel && wtSel.value && wtSel.value.toUpperCase().includes("TH"));
        }

        // Clear WPS
        function clearWps(tr) {
            const cell = tr.querySelector('td[data-col="wps"]');
            if (!cell) return;
            const sel = cell.querySelector('select[data-name="Wps"]');
            if (sel) {
                if (sel.value !== '') {
                    sel.value = '';
                    sel.dataset.userCleared = '1';
                    sel.dispatchEvent(new Event('change', { bubbles: true }));
                }
            } else {
                const inp = cell.querySelector('input[data-name="Wps"]');
                if (inp && inp.value !== '') {
                    inp.value = '';
                    inp.dispatchEvent(new Event('input', { bubbles: true }));
                }
            }
        }

        // Clear tack welder
        function clearTackWelder(tr) {
            const sel = tr.querySelector('select[data-name="TackWelder"]');
            if (sel) {
                if (sel.value !== '') {
                    sel.value = '';
                    sel.dispatchEvent(new Event('change', { bubbles: true }));
                }
            } else {
                const inp = tr.querySelector('input[data-name="TackWelder"]');
                if (inp && inp.value !== '') {
                    inp.value = '';
                    inp.dispatchEvent(new Event('input', { bubbles: true }));
                }
            }
        }

        // Validate LET thicknesses
        function validateLetThicknesses(showMsg, dirtyOnly = true) {
            const rows = [...document.querySelectorAll('#fitupTable tbody tr')];
            let invalidCount = 0;

            rows.forEach(r => {
                const isDirty = r.dataset.dirty === '1';
                if (dirtyOnly && !isDirty) {
                    r.classList.remove('require-olthick');
                    return;
                }

                if (isLetWeldType(r)) {
                    const thick = r.querySelector('input[data-name="OlThick"]');
                    const missing = !thick || !thick.value || !thick.value.trim();
                    if (missing && isDirty) {
                        r.classList.add('require-olthick');
                        invalidCount++;
                    } else {
                        r.classList.remove('require-olthick');
                    }
                } else {
                    r.classList.remove('require-olthick');
                }
            });

            if (showMsg && invalidCount > 0) {
                showStatus('Enter Actual Thick. for all LET weld types.', false);
            }
            return invalidCount === 0;
        }

        // Update row state based on checkboxes
        function updateRowState(tr) {
            const deleted = tr.querySelector('input[data-name="Deleted"]')?.checked;
            const cancelled = tr.querySelector('input[data-name="Cancelled"]')?.checked;
            const confirmed = tr.querySelector('input[data-name="FitupConfirmed"]')?.checked;
            let allowed = null;

            if (deleted || cancelled) {
                allowed = new Set(['Deleted', 'Cancelled']);
            } else if (confirmed) {
                allowed = new Set(['FitupConfirmed']);
            }

            const inputs = tr.querySelectorAll('input[data-name], select[data-name]');
            inputs.forEach(el => {
                const nm = el.getAttribute('data-name');
                if (allowed) {
                    el.disabled = !allowed.has(nm);
                } else {
                    if (nm) {
                        if (el.matches('input[type="checkbox"]')) {
                            el.disabled = false;
                        } else {
                          el.disabled = false;
                        }
                    }
                }
            });

            // Row action buttons
            const delBtn = tr.querySelector('.delete-row');
            const clrBtn = tr.querySelector('.clear-row');
            if (delBtn) delBtn.disabled = !!allowed;
            if (clrBtn) clrBtn.disabled = !!allowed;

            // Visual locked class
            if (confirmed) {
                tr.classList.add('locked');
            } else {
                tr.classList.remove('locked');
            }
        }

        // Add status handlers to row
        function addStatusHandlers(tr) {
            const map = ['Deleted', 'Cancelled', 'FitupConfirmed'];
            map.forEach(name => {
                const cb = tr.querySelector(`input[data-name="${name}"]`);
                if (cb && !cb.dataset.stateBound) {
                    cb.addEventListener('change', () => {
                        markRowDirty(tr);
                        updateRowState(tr);
                    });
                    cb.dataset.stateBound = '1';
                }
            });

            // Enforce exclusivity between Deleted and Cancelled
            const del = tr.querySelector('input[data-name="Deleted"]');
            const can = tr.querySelector('input[data-name="Cancelled"]');
            if (del && can) {
                del.addEventListener('change', () => {
                    if (del.checked) {
                        can.checked = false;
                        markRowDirty(tr);
                        updateRowState(tr);
                    }
                });
                can.addEventListener('change', () => {
                    if (can.checked) {
                        del.checked = false;
                        markRowDirty(tr);
                        updateRowState(tr);
                    }
                });
            }
        }

        // Renumber rows
        function renumberRows() {
            const rows = [...document.querySelectorAll('#fitupTable tbody tr')];
            rows.forEach((r, i) => {
                const cell = r.querySelector('.col-sn');
                if (cell) cell.textContent = String(i + 1);
            });
        }

        // Escape HTML
        function escapeHtml(str) {
            return ((str ?? '') + "").replace(/[&<>"';]/g, s => ({
                '&': '&amp;', '<': '&lt;', '>': '&gt;',
                '"': '&quot;', "'": '&#39;', ';': '&#59;'
            }[s]));
        }

        // Build options HTML
        function buildOptions(arr, selected) {
            if (!Array.isArray(arr)) arr = [];
            let list = [...arr];
            if (selected && !list.includes(selected)) list = [...list, selected];
            if (list.length === 0) return '';
            return list.map(v =>
                `<option value="${escapeHtml(v)}"${v === selected ? ' selected' : ''}>${escapeHtml(v)}</option>`
            ).join('');
        }

        // Build tacker options HTML
        function buildTackerOptionsHtml(selected) {
            const headerSel = document.getElementById('tackerSelect');
            const vals = headerSel ? Array.from(headerSel.options).map(o => o.value) : [];
            let html = '';
            vals.forEach(v => {
                const esc = escapeHtml(v);
                html += `<option value="${esc}"${v === selected ? ' selected' : ''}>${esc}</option>`;
            });
            if (!selected) {
                html += '<option value="" selected></option>';
            } else {
                html += '<option value=""></option>';
            }
            return html;
        }

        // RFI dropdown helpers
        function extractBareRfi(txt) {
            let s = (txt || '').trim();
            ['WS | ', 'FW | ', 'TH | '].forEach(p => {
                if (s.startsWith(p)) s = s.substring(p.length);
            });
            const idx = s.indexOf(' | ');
            return idx >= 0 ? s.substring(0, idx) : s;
        }

        function determineLocationBucket(tr) {
            const wtSel = tr.querySelector('select[data-name="WeldType"]');
            const wt = (wtSel?.value || '').toUpperCase();
            if (wt.includes('TH')) return 'TH';

            const rowLocSel = tr.querySelector('select[data-name="Location"]');
            const raw = (rowLocSel?.value || '').trim();
            const rowLoc = raw.toUpperCase();

            if (rowLoc.startsWith('WS') || rowLoc.startsWith('SHOP') || rowLoc.startsWith('WORK')) return 'WS';
            if (rowLoc.startsWith('FW') || rowLoc.startsWith('FIELD')) return 'FW';
            if (rowLoc === 'SHOP') return 'WS';
            if (rowLoc === 'FIELD') return 'FW';

            return 'FW';
        }

        // Load RFI options for row
        async function loadRowRfiOptions(tr, preselect) {
            try {
                const sel = tr.querySelector('select[data-name="RfiId"]');
                if (!sel) return;

                const projectId = document.getElementById('projectSelect')?.value;
                if (!projectId) return;

                const bucket = determineLocationBucket(tr);
                let fitupIso = document.getElementById('fitupDateInput')?.value || '';
                if (!fitupIso) {
                    fitupIso = tr.querySelector('input[data-name="FitupDate"]')?.value || '';
                }

                const cacheKey = `${projectId}|${bucket}|${fitupIso}`;
                let items = __rfiCache.get(cacheKey);

                if (!items) {
                    const base = new URL(rfiOptionsEndpoint, window.location.origin);
                    base.searchParams.set('projectId', String(projectId));
                    base.searchParams.set('location', String(bucket));
                    if (fitupIso) base.searchParams.set('fitupDateIso', fitupIso);

                    const res = await fetch(base.toString(), { headers: { 'Accept': 'application/json' } });
                    if (!res.ok) return;
                    items = await res.json();
                    if (!Array.isArray(items)) items = [];
                    __rfiCache.set(cacheKey, items);
                }

                const prevVal = sel.value || '';
                const prevText = sel.options[sel.selectedIndex]?.text || '';
                const preId = preselect ? String(preselect.id || preselect.Id || preselect.RFI_ID || preselect.Rfi_ID || preselect.RfiId || preselect) : '';
                let preText = preselect ? (preselect.value || preselect.Value || preselect.SubCon_RFI_No || preselect.RFI_NO || preselect.Rfi_No || preselect.Display || '') : '';
                preText = preText ? extractBareRfi(preText) : '';

                // Rebuild options
                while (sel.options.length > 0) sel.remove(0);

                const addOpt = (val, text, selected) => {
                    const o = document.createElement('option');
                    o.value = String(val || '');
                    o.textContent = text || '';
                    if (selected) o.selected = true;
                    sel.appendChild(o);
                };

                addOpt('', '', !prevVal && !preId);

                let foundPrev = false, foundPre = false;
                items.forEach(it => {
                    const id = it.id ?? it.Id ?? it.RFI_ID ?? it.Rfi_ID ?? it.RfiId;
                    const raw = it.value ?? it.Value ?? it.SubCon_RFI_No ?? it.RFI_NO ?? it.Rfi_No ?? it.display ?? it.Display ?? String(id || '');
                    const bare = extractBareRfi(raw);

                    if (String(id) === prevVal) foundPrev = true;
                    if (preId && String(id) === preId) foundPre = true;
                    addOpt(id, bare, false);
                });

                if (preId) {
                    if (foundPre) {
                        [...sel.options].forEach(o => { if (o.value === preId) o.selected = true; });
                    } else if (preText) {
                        addOpt(preId, preText, true);
                    }
                } else if (prevVal) {
                    if (foundPrev) {
                        [...sel.options].forEach(o => { if (o.value === prevVal) o.selected = true; });
                    } else {
                        addOpt(prevVal, extractBareRfi(prevText) || prevVal, true);
                    }
                }

                if (window.__makeSearchableSelect) window.__makeSearchableSelect(sel);
                sel.dataset.loaded = '1';
            } catch {
                // Handle error silently
            }
        }

        // Expose reload function
        window.reloadRfiListsForAllRows = function () {
            const rows = Array.from(document.querySelectorAll('#fitupTable tbody tr'));
            rows.forEach(tr => {
                try { loadRowRfiOptions(tr); } catch { }
            });
        };

        // Attach revision insert handler
        function attachRevInsert(tr) {
            const fitInp = tr.querySelector('input[data-name="FitupReport"]');
            if (!fitInp || fitInp.dataset.revAttach) return;

            fitInp.addEventListener('dblclick', () => {
                // Insert LS Rev
                if (lsRev) {
                    const revInp = tr.querySelector('input[data-name="Rev"]');
                    if (revInp && !revInp.disabled) {
                        revInp.value = lsRev;
                        markRowDirty(tr);
                    }
                }

                // Copy header date to row
                const headerDate = document.getElementById('fitupDateInput');
                if (headerDate) {
                    const rowDate = tr.querySelector('input[data-name="FitupDate"]');
                    if (rowDate && !rowDate.disabled) {
                        rowDate.value = headerDate.value || '';
                        rowDate.dispatchEvent(new Event('change', { bubbles: true }));
                        markRowDirty(tr);
                    }
                }

                // Copy fit-up report based on location
                const rowLocSel = tr.querySelector('select[data-name="Location"]');
                const rowLocLower = ((rowLocSel?.value) || '').trim().toLowerCase();
                const combinedEl = document.getElementById('fitupReportCombinedInput');
                const singleEl = document.getElementById('fitupReportHeaderInput');
                const headerLocSel = document.getElementById('locationSelect');
                const headerLocLower = ((headerLocSel?.value) || '').trim().toLowerCase();
                const rowIsTh = (() => {
                    const wtSel = tr.querySelector('select[data-name="WeldType"]');
                    return !!(wtSel && wtSel.value && wtSel.value.toUpperCase().includes('TH'));
                })();

                let headerRep = '';
                if (combinedEl && headerLocLower === 'all') {
                    const raw = (combinedEl.value || '').trim();
                    const parts = raw.split('/').map(p => p.trim());
                    const wsPart = parts[0] || '';
                    const fieldPart = parts[1] || '';
                    const thPart = parts[2] || '';
                    headerRep = rowIsTh ? (thPart || fieldPart || wsPart) :
                        (['ws', 'shop', 'workshop'].includes(rowLocLower) ? wsPart : (fieldPart || wsPart));
                } else if (singleEl) {
                    headerRep = singleEl.value || '';
                } else if (combinedEl) {
                    headerRep = combinedEl.value || '';
                }

                if (headerRep && fitInp && !fitInp.disabled) {
                    fitInp.value = headerRep;
                    markRowDirty(tr);
                }

                // Copy tacker
                const headerTacker = document.getElementById('tackerSelect');
                const rowTackerSel = tr.querySelector('select[data-name="TackWelder"]');
                if (rowIsTh) {
                    clearTackWelder(tr);
                } else if (headerTacker && rowTackerSel && !rowTackerSel.disabled) {
                    rowTackerSel.value = headerTacker.value || '';
                    rowTackerSel.dispatchEvent(new Event('change', { bubbles: true }));
                    markRowDirty(tr);
                }

                // Copy RFI
                const headerRfi = document.getElementById('rfiSelect');
                const rowRfiSel = tr.querySelector('select[data-name="RfiId"]');
                if (headerRfi && rowRfiSel && !rowRfiSel.disabled) {
                    const hdrId = headerRfi.value || '';
                    const hdrText = headerRfi.options[headerRfi.selectedIndex]?.text || '';
                    const bare = extractBareRfi(hdrText);
                    if (hdrId) {
                        if (![...rowRfiSel.options].some(o => o.value === hdrId)) {
                            const opt = document.createElement('option');
                            opt.value = hdrId;
                            opt.textContent = bare || hdrId;
                            rowRfiSel.insertBefore(opt, rowRfiSel.firstChild);
                        }
                        rowRfiSel.value = hdrId;
                        rowRfiSel.dispatchEvent(new Event('change', { bubbles: true }));
                        markRowDirty(tr);
                        loadRowRfiOptions(tr, { id: hdrId, value: bare });
                    } else {
                        loadRowRfiOptions(tr);
                    }
                } else {
                    loadRowRfiOptions(tr);
                }

                // Add hold note to remarks if on hold
                const onHold = !!tr.querySelector('.col-jointno.hold-alert, .col-spno.hold-alert, .col-sheet.hold-alert');
                if (onHold) {
                    const remInp = tr.querySelector('input[data-name="Remarks"]');
                    if (remInp && !remInp.disabled) {
                        const note = 'Fabricated on hold';
                        const cur = (remInp.value || '').trim();
                        if (cur.length === 0) remInp.value = note;
                        else if (!cur.toLowerCase().includes(note.toLowerCase())) {
                            remInp.value = cur + ' | ' + note;
                        }
                        remInp.dispatchEvent(new Event('input', { bubbles: true }));
                        markRowDirty(tr);
                    }
                }
            });

            fitInp.dataset.revAttach = '1';
        }

        // Add row button handler
        document.getElementById('addRowBtn')?.addEventListener('click', async () => {
            const projectSel = document.getElementById('projectSelect');
            const layoutSel = document.getElementById('layoutSelect');
            const sheetSel = document.getElementById('sheetSelect');
            const locationSel = document.getElementById('locationSelect');
            const locationValue = locationSel?.value || '';

            if (!projectSel || !layoutSel || !sheetSel) {
                showStatus('Page is missing required inputs. Reload the page.', false);
                return;
            }

            let rowLocation = locationValue === 'Threaded' ? 'TH' : (locationValue === 'All' ? 'Field' : locationValue);
            if (!layoutSel.value || !sheetSel.value) {
                showStatus('Enter layout & sheet first', false);
                return;
            }

            const tbody = document.querySelector('#fitupTable tbody');
            if (!tbody) {
                showStatus('Load data first', false);
                return;
            }

            if (!addRowEndpoint) {
                showStatus('Add row endpoint unavailable. Reload the page.', false);
                return;
            }

            try {
                const formBody = new URLSearchParams({
                    projectId: projectSel.value,
                    location: rowLocation,
                    layout: layoutSel.value,
                    sheet: sheetSel.value,
                    __RequestVerificationToken: token
                });

                const res = await fetch(addRowEndpoint, {
                    method: 'POST',
                    credentials: 'same-origin',
                    headers: {
                        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                        'RequestVerificationToken': token
                    },
                    body: formBody
                });

                if (!res.ok) {
                    let msg = 'Add failed';
                    try {
                        const text = await res.text();
                        let jsonErr = null;
                        try { jsonErr = JSON.parse(text); } catch { }
                        if (jsonErr && jsonErr.message) {
                            msg = jsonErr.message;
                        } else if (text) {
                            msg = text;
                        }
                    } catch { }
                    showStatus(msg, false);
                    return;
                }

                let data = null;
                try { data = await res.json(); } catch {
                    showStatus('Add failed', false);
                    return;
                }

                if (!data || data.success !== true) {
                    const msg = (data && data.message) ? data.message : 'Add failed';
                    showStatus(msg, false);
                    return;
                }

                const headerSheet = (sheetSel.value || '').trim();
                const sn = tbody.querySelectorAll('tr').length + 1;
                const rawLocOptions = (locationOptions && locationOptions.length ? locationOptions : [rowLocation]);
                const filteredLocOptions = rawLocOptions.filter(v => String(v).toLowerCase() !== 'threaded');
                const locSelected = (locationValue === 'Threaded') ? 'TH' :
                    ((filteredLocOptions.includes(data.location)) ? data.location :
                        (filteredLocOptions.length ? filteredLocOptions[0] : (data.location || '')));

                let weldNumberRaw = (data.weldNumber || '').toString();
                let weldNumber3 = weldNumberRaw;
                const numMatch = weldNumberRaw.match(/^[0-9]+$/);
                if (numMatch) {
                    const num = parseInt(weldNumberRaw, 10);
                    if (!isNaN(num)) weldNumber3 = String(num).padStart(3, '0');
                }

                let jAddSource = (jAddOptions && jAddOptions.length ? [...jAddOptions] : []);
                jAddSource = jAddSource.filter(x => x.toUpperCase() !== 'NEW');
                jAddSource.unshift('New');
                const jAddOptsHtml = buildOptions(jAddSource, 'New');

                let weldTypeSource = (weldTypeOptions && weldTypeOptions.length ? [...weldTypeOptions] : []);
                const defaultWt = 'BW';
                const weldTypeOptsHtml = buildOptions(weldTypeSource, defaultWt);
                const locationOptsHtml = buildOptions(filteredLocOptions, locSelected);
                const tackerOptsHtml = buildTackerOptionsHtml('');
                const matDescOpts = buildOptions(materialDescriptions || [], '');
                const matGradeOpts = buildOptions(materialGrades || [], '');

                const tr = document.createElement('tr');
                tr.dataset.id = data.id ?? '0';
                tr.innerHTML = `
          <td class="col-sn">${sn}</td>
          <td class="col-location"><select data-name="Location" name="Location" id="Location_New" class="searchable-select">${locationOptsHtml}</select></td>
          <td class="col-jointno"><input data-name="WeldNumber" name="WeldNumber" id="WeldNumber_New" value="${escapeHtml(weldNumber3)}" /></td>
          <td class="col-jadd"><select data-name="JAdd" name="JAdd" id="JAdd_New">${jAddOptsHtml}</select></td>
          <td class="col-wtype"><select data-name="WeldType" name="WeldType" id="WeldType_New">${weldTypeOptsHtml}</select></td>
          <td class="col-sheet"><input data-name="Sheet" name="Sheet" id="Sheet_New" value="${escapeHtml((data.sheet && data.sheet.trim()) ? data.sheet : headerSheet)}" /></td>
          <td class="col-rev"><input data-name="Rev" name="Rev" id="Rev_New" /></td>
          <td class="col-spno"><input data-name="SpoolNo" name="SpoolNo" id="SpoolNo_New" /></td>
          <td class="col-dia"><input data-name="DiaIn" name="DiaIn" id="DiaIn_New" class="dia-input" /></td>
          <td class="col-sch"><input data-name="Sch" name="Sch" id="Sch_New" class="sch-input" /></td>
          <td class="col-matA"><select data-name="MaterialA" name="MaterialA" id="MaterialA_New" class="searchable-select">${matDescOpts}<option value="" selected></option></select></td>
          <td class="col-matB"><select data-name="MaterialB" name="MaterialB" id="MaterialB_New" class="searchable-select">${matDescOpts}<option value="" selected></option></select></td>
          <td class="col-gradeA"><select data-name="GradeA" name="GradeA" id="GradeA_New" class="searchable-select">${matGradeOpts}<option value="" selected></option></select></td>
          <td class="col-gradeB"><select data-name="GradeB" name="GradeB" id="GradeB_New" class="searchable-select">${matGradeOpts}<option value="" selected></option></select></td>
          <td class="col-heatA"><select data-name="HeatNumberA" name="HeatNumberA" id="HeatNumberA_New" class="searchable-select heat-select"><option value="" selected></option></select></td>
          <td class="col-heatB"><select data-name="HeatNumberB" name="HeatNumberB" id="HeatNumberB_New" class="searchable-select heat-select"><option value="" selected></option></select></td>
          <td class="col-frep"><input autocomplete="off" data-name="FitupReport" name="FitupReport" id="FitupReport_New" /></td>
          <td class="col-rfi"><select data-name="RfiId" name="RfiId" id="RfiId_New" class="searchable-select rfi-select"><option value="" selected></option></select></td>
          <td class="col-fdate"><input type="datetime-local" data-name="FitupDate" name="FitupDate" id="FitupDate_New" /></td>
          <td class="col-wps" data-col="wps"><select data-name="Wps" name="Wps" id="Wps_New" class="wps-select" style="width:160px"><option value="" selected></option></select></td>
          <td class="col-tacker"><select data-name="TackWelder" name="TackWelder" id="TackWelder_New" class="searchable-select w110">${tackerOptsHtml}</select></td>
          <td class="col-olDia"><input data-name="OlDia" name="OlDia" id="OlDia_New" class="ol-dia-input" /></td>
          <td class="col-olSch"><input data-name="OlSch" name="OlSch" id="OlSch_New" class="olsched-input" /></td>
          <td class="col-olThick"><input data-name="OlThick" name="OlThick" id="OlThick_New" /></td>
          <td class="col-del chk-cell"><input type="checkbox" data-name="Deleted" name="Deleted" id="Deleted_New" /></td>
          <td class="col-can chk-cell"><input type="checkbox" data-name="Cancelled" name="Cancelled" id="Cancelled_New" /></td>
          <td class="col-conf chk-cell"><input type="checkbox" data-name="FitupConfirmed" name="FitupConfirmed" id="FitupConfirmed_New" /></td>
          <td class="col-remarks"><input data-name="Remarks" name="Remarks" id="Remarks_New" /></td>
          <td class="readonly col-updby"></td>
          <td class="readonly col-upddate"></td>
          <td class="col-clear row-actions"><button type="button" class="clear-row btn btn-outline">Clear</button></td>
          <td class="col-delbtn row-actions"><button type="button" class="delete-row btn btn-danger">Del</button></td>
        `;

                tbody.appendChild(tr);
                renumberRows();

                // Apply frozen columns
                if (typeof markFrozen === 'function') markFrozen();
                if (typeof applyFrozenPositions === 'function') applyFrozenPositions();

                attachDelete(tr);
                attachClear(tr);
                addStatusHandlers(tr);
                tr.dataset.initializing = '1';
                const markDirtyInitially = !data.id || data.id === 0;
                initRowBaseline(tr, markDirtyInitially);

                tr.querySelectorAll('select[data-name="MaterialA"], select[data-name="MaterialB"], select[data-name="GradeA"], select[data-name="GradeB"], select[data-name="RfiId"]').forEach(sel => {
                    if (window.__makeSearchableSelect) window.__makeSearchableSelect(sel);
                    sel.addEventListener('change', () => { markRowDirty(tr); });
                });

                addDynamicListeners(tr);
                addActivationTriggers(tr);
                attachRevInsert(tr);
                persistScrollInit();
                autoFillOlThick(tr).finally(() => { delete tr.dataset.initializing; });
                resizeBottomScroll();
                initHeatSelects(tr);
                refreshHeatOptions(tr, 'A');
                refreshHeatOptions(tr, 'B');
                loadRowRfiOptions(tr);

                if (window.applyWeldTypeFilter) window.applyWeldTypeFilter();

                setTimeout(() => {
                    try { tr.scrollIntoView({ behavior: 'smooth', block: 'center' }); } catch { }
                }, 0);

                showStatus('Joint added', true);
            } catch {
                showStatus('Error adding', false);
            }
        });

        // Delete row modal
        function openDeleteRowModal(tr) {
            pendingDeleteRow = tr;
            const loc = tr.querySelector('[data-name="Location"]')?.value?.trim() || '';
            const weldNo = tr.querySelector('input[data-name="WeldNumber"]').value?.trim() || '';
            const jAddEl = tr.querySelector('select[data-name="JAdd"]');
            const jAddRaw = jAddEl ? (jAddEl.value || '').trim() : '';
            const jAddPart = (jAddRaw && jAddRaw.toUpperCase() !== 'NEW') ? jAddRaw : '';

            let jointDisplay = '';
            if (loc && weldNo) {
                jointDisplay = `${loc}-${weldNo}${jAddPart}`;
            } else if (loc) {
                jointDisplay = `${loc}${weldNo ? ('-' + weldNo) : ''}${jAddPart}`;
            } else if (weldNo) {
                jointDisplay = `${weldNo}${jAddPart}`;
            }

            const msg = jointDisplay ? `Are you sure you want to delete Joint "${jointDisplay}"?` : 'Are you sure you want to delete this row?';
            const el = document.getElementById('confirmDeleteRowMessage');
            if (el) el.textContent = msg;

            const modal = document.getElementById('confirmDeleteRowModal');
            if (modal) {
                modal.style.display = 'flex';
                modal.setAttribute('aria-hidden', 'false');
                modal.querySelector('.close-modal')?.focus();
            }
        }

        window.closeDeleteRowModal = function () {
            const m = document.getElementById('confirmDeleteRowModal');
            if (m) {
                m.style.display = 'none';
                m.setAttribute('aria-hidden', 'true');
            }
            pendingDeleteRow = null;
        };

        window.confirmDeleteRow = async function () {
            if (!pendingDeleteRow) {
                closeDeleteRowModal();
                return;
            }

            const tr = pendingDeleteRow;
            pendingDeleteRow = null;
            closeDeleteRowModal();

            const id = tr.getAttribute('data-id');
            if (!id || id === '0') {
                tr.remove();
                renumberRows();
                resizeBottomScroll();
                showStatus('Deleted', true);
                return;
            }

            if (!deleteRowEndpoint) {
                showStatus('Delete endpoint unavailable.', false);
                return;
            }

            try {
                const res = await fetch(deleteRowEndpoint, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/x-www-form-urlencoded',
                        'RequestVerificationToken': token
                    },
                    body: new URLSearchParams({ id })
                });

                if (res.ok) {
                    tr.remove();
                    renumberRows();
                    resizeBottomScroll();
                    showStatus('Deleted', true);
                } else {
                    showStatus('Delete failed', false);
                }
            } catch {
                showStatus('Error deleting', false);
            }
        };

        // Attach delete handler
        function attachDelete(tr) {
            const btn = tr.querySelector('.delete-row');
            if (!btn) return;
            btn.addEventListener('click', () => {
                if (btn.disabled) return;
                openDeleteRowModal(tr);
            });
        }

        // Clear row fields
        function clearRowFields(tr) {
            try {
                const rep = tr.querySelector('input[data-name="FitupReport"]');
                if (rep && !rep.disabled) {
                    rep.value = '';
                    rep.dispatchEvent(new Event('input', { bubbles: true }));
                }

                const rfiSel = tr.querySelector('select[data-name="RfiId"]');
                if (rfiSel && !rfiSel.disabled) {
                    rfiSel.value = '';
                    rfiSel.dispatchEvent(new Event('change', { bubbles: true }));
                }

                const fdt = tr.querySelector('input[data-name="FitupDate"]');
                if (fdt && !fdt.disabled) {
                    fdt.value = '';
                    fdt.dispatchEvent(new Event('change', { bubbles: true }));
                }

                // Clear REV. value when clearing row fields
                const revInp = tr.querySelector('input[data-name="Rev"]');
                if (revInp && !revInp.disabled) {
                    revInp.value = '';
                    revInp.dispatchEvent(new Event('input', { bubbles: true }));
                }

                clearWps(tr);
                clearTackWelder(tr);
                markRowDirty(tr);
                showStatus('Row fields cleared. Remember to Save.', true);
            } catch { }
        }

        // Attach clear handler
        function attachClear(tr) {
            const btn = tr.querySelector('.clear-row');
            if (!btn) return;
            btn.addEventListener('click', () => {
                if (btn.disabled) return;
                clearRowFields(tr);
            });
        }

        // Attach handlers to existing rows
        document.querySelectorAll('#fitupTable tbody tr').forEach(tr => {
            attachClear(tr);
            attachDelete(tr);
        });

        // Modal close handlers
        window.addEventListener('click', e => {
            const modal = document.getElementById('confirmDeleteRowModal');
            if (e.target === modal) {
                closeDeleteRowModal();
            }
        });

        window.addEventListener('keydown', e => {
            if (e.key === 'Escape') {
                closeDeleteRowModal();
            }
        });

        // Save all button handler
        document.getElementById('saveAllBtn')?.addEventListener('click', saveAll);

        // Bulk update response handler
        window.handleBulkUpdateResponse = function (resp) {
            try {
                // Clear previous skipped/highlight markers
                document.querySelectorAll('#fitupTable tbody tr.validation-skip').forEach(r => r.classList.remove('validation-skip'));

                // Normalize response
                const updated = Number(resp && resp.updated) || 0;
                const skipped = Number(resp && resp.skipped) || 0;
                const errors = Array.isArray(resp && resp.errors) ? resp.errors : [];

                // Mark rows referenced in errors
                errors.forEach(err => {
                    const id = err && (err.id || err.Id || err.JointId);
                    if (id !== undefined && id !== null) {
                        const tr = document.querySelector(`#fitupTable tbody tr[data-id='${id}']`);
                        if (tr) tr.classList.add('validation-skip');
                    }
                });

                // Compose main message
                let mainMsg = '';
                if (skipped > 0) {
                    if (updated > 0) mainMsg = `Saved ${updated} row(s). Skipped ${skipped} row(s).`;
                    else mainMsg = `Skipped ${skipped} row(s).`;
                } else {
                    if (updated > 0) mainMsg = `Saved ${updated} row(s) successfully.`;
                    else mainMsg = 'No changes saved.';
                }

                // If there are errors, append first error reason
                if (errors.length > 0 && errors[0].message) {
                    const firstMsg = String(errors[0].message).replace(/^Skipped:\s*/i, '');
                    showStatus(`${mainMsg} ${firstMsg}`, skipped === 0 && updated > 0);
                } else {
                    showStatus(mainMsg, skipped === 0 && updated > 0);
                }
            } catch (e) {
                console.error('handleBulkUpdateResponse error', e);
            }
        };

        // Save all function
        async function saveAll() {
            const allRows = [...document.querySelectorAll('#fitupTable tbody tr')];
            if (allRows.length === 0) {
                showStatus('Nothing to save', false);
                return;
            }

            if (!validateLetThicknesses(true, true)) return;

            const dirty = allRows.filter(r => r.dataset.dirty === '1' || !r.getAttribute('data-id') || r.getAttribute('data-id') === '0');
            if (dirty.length === 0) {
                showStatus('No changes to save', true);
                return;
            }

            const dtos = dirty.map(collectDto);
            const localValidationErrors = [];
            const validDtos = [];

            dtos.forEach(d => {
                const hasReport = typeof d.FitupReport === 'string' && d.FitupReport.trim() !== '';
                const hasDate = !!d.FitupDate;

                if (hasReport && !hasDate) {
                    localValidationErrors.push({ id: d.JointId, message: 'Skipped: Fit-up Date is required when Fit-up Report is provided.' });
                    return;
                }

                if (hasDate && !hasReport) {
                    localValidationErrors.push({ id: d.JointId, message: 'Skipped: Fit-up Report is required when Fit-up Date is provided.' });
                    return;
                }

                function parseIso(s) {
                    if (!s) return null;
                    try {
                        const d = new Date(s);
                        return isNaN(d.getTime()) ? null : d;
                    } catch {
                        return null;
                    }
                }

                // Bevel PT date validation
                const bevelKeys = ['Bevel_PT_DATE', 'Bevel_PT', 'BEVEL_PT_DATE', 'bevel_pt_date', 'BevelPtDate', 'BevelPt'];
                let bevelVal = null;
                for (const k of bevelKeys) {
                    if (Object.prototype.hasOwnProperty.call(d, k) && d[k]) {
                        bevelVal = parseIso(d[k]);
                        break;
                    }
                }

                if (bevelVal) {
                    if (!d.FitupDate) {
                        localValidationErrors.push({ id: d.JointId, message: 'Skipped: Fit-up Date is required when Bevel PT Date is present.' });
                        return;
                    }
                    const fit = parseIso(d.FitupDate);
                    if (!fit) {
                        localValidationErrors.push({ id: d.JointId, message: 'Skipped: Fit-up Date is invalid.' });
                        return;
                    }
                    if (fit.getTime() < bevelVal.getTime()) {
                        localValidationErrors.push({
                            id: d.JointId,
                            message: `Skipped: Fit-up Date must be on or after Bevel PT Date ${bevelVal.toLocaleString()}.`
                        });
                        return;
                    }
                }

                // Presence-only validation for other dates
                const presenceKeys = [
                    'Date_Welded', 'DATE_WELDED', 'DateWelded', 'dateWelded',
                    'ACTUAL_DATE_WELDED', 'Actual_Date_Welded', 'ActualDateWelded', 'ActualDate',
                    'Root_PT_DATE', 'Root_PT', 'root_pt_date', 'RootPtDate',
                    'OTHER_NDE_DATE', 'OTHER_NDE', 'other_ndE_date', 'OtherNdeDate', 'Other_Nde_Date',
                    'PMI_DATE', 'PMI', 'PmiDate', 'pmi_date',
                    'NDE_REQUEST', 'NdeRequest', 'nde_request', 'Nde_Request',
                    'BSR_NDE_REQUEST', 'BsrNdeRequest', 'bsr_nde_request'
                ];

                for (const k of presenceKeys) {
                    if (Object.prototype.hasOwnProperty.call(d, k) && d[k]) {
                        if (!d.FitupDate) {
                            localValidationErrors.push({ id: d.JointId, message: `Skipped: Fit-up Date is required when ${k} is present.` });
                            return;
                        }
                    }
                }

                validDtos.push(d);
            });

            if (localValidationErrors.length > 0) {
                localValidationErrors.forEach(err => {
                    const tr = document.querySelector(`#fitupTable tbody tr[data-id='${err.id}']`);
                    if (tr) tr.classList.add('validation-skip');
                });
            }

            if (validDtos.length === 0) {
                const msg = localValidationErrors.length > 0 ? localValidationErrors[0].message : 'Skipped.';
                handleBulkUpdateResponse({ updated: 0, skipped: localValidationErrors.length, errors: localValidationErrors });
                showStatus(msg, false);
                return;
            }

            try {
                const saveBtn = document.getElementById('saveAllBtn');
                saveBtn.disabled = true;
                saveBtn.dataset.prevText = saveBtn.textContent;
                saveBtn.textContent = 'Saving...';

                if (!bulkUpdateEndpoint) {
                    showStatus('Save endpoint unavailable. Reload the page.', false);
                    return;
                }

                async function post(payload) {
                    if (!token) {
                        showStatus('Missing session token. Please reload the page.', false);
                        return new Response(null, { status: 400 });
                    }
                    return await fetch(bulkUpdateEndpoint, {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json',
                            'RequestVerificationToken': token
                        },
                        body: JSON.stringify(payload)
                    });
                }

                let res = await post(validDtos);

                // Parse once for reuse
                let data = null;
                try { data = await res.clone().json(); } catch { data = null; }

                const isSupersedeRequired = (res.status === 409) || (data && data.code === 'requireSupersedeConfirm');

                if (isSupersedeRequired) {
                    const msg = (data && data.message) || 'You are making modification on released spool, are you sure to supersede the previous reports?';
                    const proceed = await openSupersedeModal(msg);
                    const ids = data && Array.isArray(data.ids) ?
                        data.ids.map(x => parseInt(x, 10)).filter(x => !isNaN(x)) :
                        (data && data.id ? [parseInt(data.id, 10)].filter(x => !isNaN(x)) : []);

                    if (!proceed) {
                        const remaining = validDtos.filter(d => !ids.includes(d.JointId));
                        if (remaining.length === 0) {
                            handleBulkUpdateResponse({
                                updated: 0,
                                skipped: localValidationErrors.length + ids.length,
                                errors: [...localValidationErrors, ...ids.map(i => ({ id: i, message: 'Supersede skipped' }))]
                            });
                            showStatus('Supersede cancelled. All pending changes required supersede and were skipped.', false);
                            return;
                        }

                        res = await post(remaining);
                        data = null;
                        try { data = await res.clone().json(); } catch { data = null; }

                        if (res.ok) {
                            let j = data;
                            if (!j) j = { updated: remaining.length, skipped: localValidationErrors.length };

                            if (localValidationErrors.length > 0) {
                                j.errors = (Array.isArray(j.errors) ? j.errors : []).concat(localValidationErrors);
                                j.skipped = j.errors.length;
                            }

                            const updated = (j && typeof j.updated === 'number') ? j.updated : remaining.length;
                            const skippedSupersede = validDtos.length - remaining.length;
                            const additionalSkipped = (j && typeof j.skipped === 'number' ? j.skipped : 0) - localValidationErrors.length;
                            const totalSkipped = skippedSupersede + additionalSkipped + localValidationErrors.length;
                            const savedIds = new Set(remaining.map(r => r.JointId));

                            dirty.forEach(r => {
                                const idVal = parseInt(r.getAttribute('data-id') || '0', 10);
                                if (savedIds.has(idVal)) {
                                    delete r.dataset.dirty;
                                    r.classList.remove('dirty');
                                }
                            });

                            showStatus(`Saved ${updated} row(s). Skipped ${totalSkipped} row(s).`, true);
                            if (j) handleBulkUpdateResponse(j);
                        } else {
                            let m = 'Save failed';
                            try { m = await res.text(); if (!m) m = 'Save failed'; } catch { }
                            showStatus(m, false);
                        }
                        return;
                    }

                    const next = validDtos.map(d =>
                        (ids.includes(d.JointId) ? Object.assign({}, d, { ConfirmSupersede: true }) : d)
                    );
                    res = await post(next);
                    try { data = await res.clone().json(); } catch { data = null; }
                }

                if (res.ok) {
                    let j = data;
                    if (!j) {
                        try { j = await res.json(); } catch { j = { updated: validDtos.length, skipped: localValidationErrors.length }; }
                    }

                    if (localValidationErrors.length > 0) {
                        j.errors = (Array.isArray(j.errors) ? j.errors : []).concat(localValidationErrors);
                        j.skipped = j.errors.length;
                    }

                    const postedIds = new Set(validDtos.map(d => d.JointId));
                    dirty.forEach(r => {
                        const idVal = parseInt(r.getAttribute('data-id') || '0', 10);
                        if (postedIds.has(idVal)) {
                            delete r.dataset.dirty;
                            r.classList.remove('dirty');
                        }
                    });

                    // Improved messaging for save/skip
                    if (j.skipped > 0 && j.updated > 0) {
                        showStatus(`Saved ${j.updated} row(s). Skipped ${j.skipped} row(s).`, true);
                        if (j.errors && j.errors.length > 0 && j.errors[0].message) {
                            showStatus(j.errors[0].message, false);
                        }
                    } else if (j.skipped > 0 && j.updated === 0) {
                        showStatus(`Skipped ${j.skipped} row(s).`, false);
                        if (j.errors && j.errors.length > 0 && j.errors[0].message) {
                            showStatus(j.errors[0].message, false);
                        }
                    } else if (j.skipped === 0 && j.updated > 0) {
                        showStatus(`Saved ${j.updated} row(s) successfully.`, true);
                    } else {
                        showStatus('No changes saved.', false);
                    }

                    if (j) handleBulkUpdateResponse(j);
                } else {
                    let msg = 'Save failed';
                    try { msg = await res.text(); if (!msg) msg = 'Save failed'; } catch { }
                    showStatus(msg, false);
                }
            } catch (e) {
                showStatus('Error saving', false);
            } finally {
                const saveBtn = document.getElementById('saveAllBtn');
                if (saveBtn) {
                    saveBtn.disabled = false;
                    if (saveBtn.dataset.prevText) saveBtn.textContent = saveBtn.dataset.prevText;
                }
            }
        }
        window.saveAll = saveAll;

        // Supersede modal functions
        window.openSupersedeModal = function (message) {
            const modal = document.getElementById('supersedeConfirmModal');
            const msgEl = document.getElementById('supersedeConfirmMessage');
            if (msgEl) msgEl.textContent = message || 'You are making modification on released spool, are you sure to supersede the previous reports?';
            if (modal) {
                modal.style.display = 'flex';
                modal.setAttribute('aria-hidden', 'false');
            }
            return new Promise((resolve) => { __supersedeResolve = resolve; });
        };

        window.closeSupersedeModal = function () {
            const m = document.getElementById('supersedeConfirmModal');
            if (m) {
                m.style.display = 'none';
                m.setAttribute('aria-hidden', 'true');
            }
        };

        window.cancelSupersede = function () {
            if (__supersedeResolve) {
                try { __supersedeResolve(false); } catch { }
                __supersedeResolve = null;
            }
            closeSupersedeModal();
        };

        window.confirmSupersede = function () {
            if (__supersedeResolve) {
                try { __supersedeResolve(true); } catch { }
                __supersedeResolve = null;
            }
            closeSupersedeModal();
        };

        // Supersede modal event handlers
        window.addEventListener('click', e => {
            const modal = document.getElementById('supersedeConfirmModal');
            if (e.target === modal) {
                cancelSupersede();
            }
        });

        window.addEventListener('keydown', e => {
            if (e.key === 'Escape') {
                const modal = document.getElementById('supersedeConfirmModal');
                if (modal && modal.style.display === 'flex') {
                    cancelSupersede();
                }
            }
        });

        // WPS refresh function
        function shouldAutoOverrideWps(tr) {
            const wtSel = tr.querySelector('select[data-name="WeldType"]');
            const sel = tr.querySelector('select[data-name="Wps"]') || tr.querySelector('input[data-name="Wps"]');
            const userSelected = !!(sel && sel.dataset.userSelected === '1');
            const userCleared = !!(sel && sel.dataset.userCleared === '1');
            return !userSelected && !userCleared;
        }

        function buildWpsTitle(o) {
            const t = o && (o.thicknessRange || o.ThicknessRange || '');
            const pw = (o && (o.pwht ?? o.Pwht)) ? 'Y' : 'N';
            return `t: ${t} PWHT: ${pw}`;
        }

        async function refreshWps(tr, activateDefault = false, forceFirst = false) {
            const cell = tr.querySelector('td[data-col="wps"]');
            if (!cell) return;

            let st = wpsState.get(tr);
            if (!st) {
                st = { seq: 0, controller: null };
                wpsState.set(tr, st);
            }
            st.seq++;
            const currentSeq = st.seq;

            if (st.controller) {
                try { st.controller.abort(); } catch { }
            }
            st.controller = new AbortController();
            const signal = st.controller.signal;

            const projectSel = document.getElementById('projectSelect');
            const projectId = projectSel?.value || '';
            let lineClass = configLineClass;
            if (!lineClass) {
                lineClass = (tr.getAttribute('data-lineclass') || '').trim();
            }

            const thkInput = tr.querySelector('input[data-name="OlThick"]');
            const thickness = thkInput && thkInput.value ? parseFloat(thkInput.value) : null;
            const pwht = (typeof window !== 'undefined' && Object.prototype.hasOwnProperty.call(window, 'dailyFitupPwht')) ? window.dailyFitupPwht : null;
            const sch = (tr.querySelector('input[data-name="Sch"]').value || '').trim();
            const diaRaw = (tr.querySelector('input[data-name="DiaIn"]')?.value || '').trim();
            const olThickRaw = ( tr.querySelector('input[data-name="OlThick"]').value || '').trim();
            const dia = diaRaw && !isNaN(parseFloat(diaRaw)) ? parseFloat(diaRaw) : null;
            const olThick = olThickRaw && !isNaN(parseFloat(olThickRaw)) ? parseFloat(olThickRaw) : null;

            try {
                const urlParams = new URLSearchParams({
                    projectId: String(projectId),
                    lineClass: String(lineClass || ''),
                    pwht: String(pwht)
                });

                if (thickness != null) urlParams.append('thickness', thickness.toString());
                if (sch) urlParams.append('sch', sch);
                if (dia != null) urlParams.append('dia', dia.toString());
                if (olThick != null) urlParams.append('olThick', olThick.toString());

                const res = await fetch(`${wpsEndpointBase}?${urlParams.toString()}`, {
                    headers: { 'Accept': 'application/json' },
                    signal
                });

                if (!res.ok) return;
                const payload = await res.json();

                if (wpsState.get(tr)?.seq !== currentSeq) return;

                const dataList = Array.isArray(payload) ? payload : (payload.items || []);
                const deletedChecked = !!(tr.querySelector('input[data-name="Deleted"]') &&
                    tr.querySelector('input[data-name="Deleted"]').checked);
                const cancelledChecked = !!(tr.querySelector('input[data-name="Cancelled"]') &&
                    tr.querySelector('input[data-name="Cancelled"]').checked);
                const disabled = tr.classList.contains('locked') || deletedChecked || cancelledChecked;

                let currentVal = '';
                let userCleared = false;
                let userSelected = false;
                const existingSel = cell.querySelector('select[data-name="Wps"]');

                if (existingSel) {
                    currentVal = existingSel.value || '';
                    userCleared = existingSel.dataset.userCleared === '1';
                    userSelected = existingSel.dataset.userSelected === '1';
                } else {
                    const existingInput = cell.querySelector('input[data-name="Wps"]');
                    if (existingInput) {
                        currentVal = existingInput.value || '';
                        userSelected = existingInput.dataset.userSelected === '1';
                    }
                }

                if (currentVal.includes(':')) currentVal = currentVal.split(':')[0];
                const thType = isThWeldType(tr);

                if (!Array.isArray(dataList) || dataList.length === 0) {
                    cell.innerHTML = `<input data-name="Wps" value="${thType ? '' : currentVal}" ${disabled ? 'disabled' : ''} style="width:160px" />`;
                    const inp = cell.querySelector('input[data-name="Wps"]');
                    if (inp) {
                        inp.addEventListener('input', () => {
                            inp.dataset.userSelected = '1';
                            markRowDirty(tr);
                        });
                    }
                    return;
                }

                let finalVal = currentVal;
                const existsInData = v => dataList.some(o => o.wps === v || o.Wps === v);
                let preserveFallback = false;

                if (thType) {
                    finalVal = '';
                    userCleared = true;
                } else if (forceFirst) {
                    finalVal = dataList[0].wps || dataList[0].Wps;
                } else if (userCleared) {
                    finalVal = '';
                } else if (activateDefault) {
                    if (!finalVal || !existsInData(finalVal)) finalVal = dataList[0].wps || dataList[0].Wps;
                } else {
                    if (finalVal && !existsInData(finalVal)) preserveFallback = true;
                }

                const optsCore = dataList.map(o => {
                    const v = o.wps || o.Wps;
                    const selected = (v === finalVal) ? ' selected' : '';
                    const id = o.id || o.Id;
                    const title = buildWpsTitle(o);
                    return `<option data-wps-id="${id}" value="${v}"${selected} title="${title}">${v}</option>`;
                }).join('');

                const fallbackOpt = (preserveFallback && finalVal) ?
                    `<option data-wps-id="" value="${finalVal}" selected title="Saved (not in candidates)">${finalVal}</option>` :
                    '';

                const blankSelected = finalVal === '' ? ' selected' : '';
                const userSelAttr = (userSelected && finalVal !== '') ? 'data-userSelected="1"' : '';
                const userClrAttr = userCleared ? 'data-userCleared="1"' : '';
                const isDisabled = disabled ? 'disabled' : '';
                const styleAttr = 'style="width:160px"';

                cell.innerHTML = `<select data-name="Wps" class="wps-select" ${styleAttr} ${isDisabled} ${userSelAttr} ${userClrAttr}>${fallbackOpt}${optsCore}<option value=""${blankSelected}></option></select>`;

                const selEl = cell.querySelector('select[data-name="Wps"]');
                if (selEl) {
                    selEl.addEventListener('change', () => {
                        if (selEl.value === '') {
                            selEl.dataset.userCleared = '1';
                            selEl.dataset.userSelected = '';
                        } else {
                            selEl.dataset.userCleared = '';
                            selEl.dataset.userSelected = '1';
                        }
                        markRowDirty(tr);
                    });
                }
            } catch (e) {
                if (e?.name !== 'AbortError') console.error('refreshWps error', e);
            }
        }

        // Auto-fill OL thickness
        async function autoFillOlThick(tr) {
            const thickInput = tr.querySelector('input[data-name="OlThick"]');
            if (!thickInput) return;

            const initializing = tr.dataset.initializing === '1';

            if (isLetWeldType(tr)) {
                if (thickInput.dataset.userEdited === '1') {
                    tr.classList.add('require-olthick');
                    refreshWps(tr, false, false);
                    return;
                }

                const dia = parseFloat((tr.querySelector('input[data-name="DiaIn"]')?.value || '').trim());
                const sch = (tr.querySelector('input[data-name="Sch"]').value || '').trim();

                if (isNaN(dia) || !sch) {
                    tr.classList.add('require-olthick');
                    return;
                }

                if (!scheduleThicknessEndpoint) {
                    tr.classList.add('require-olthick');
                    return;
                }

                const url = `${scheduleThicknessEndpoint}?sch=${encodeURIComponent(sch)}&dia=${encodeURIComponent(dia)}`;

                try {
                    const res = await fetch(url, { headers: { 'Accept': 'application/json' } });
                    if (!res.ok) return;
                    const data = await res.json();
                    if (data && !isNaN(parseFloat(data.thickness))) {
                        const prev = thickInput.value;
                        const next = data.thickness;
                        const changed = String(prev ?? '') !== String(next ?? '');
                        thickInput.value = next;
                        if (changed && !initializing) {
                            thickInput.dispatchEvent(new Event('input', { bubbles: true }));
                            markRowDirty(tr);
                        }
                    }
                } catch { }
            } else {
                tr.classList.remove('require-olthick');
            }

            if (isThWeldType(tr)) {
                clearWps(tr);
                clearTackWelder(tr);
                refreshWps(tr, false, false);
            }

            let allowOverride = shouldAutoOverrideWps(tr);
            if (initializing) allowOverride = false;

            if (thickInput.dataset.userEdited === '1') {
                refreshWps(tr, allowOverride, allowOverride);
                return;
            }

            if (!thickInput.value || thickInput.value.trim() === '') {
                const valOf = n => (tr.querySelector(`input[data-name="${n}"]`)?.value || '').trim();
                const pairs = [];
                const bSch = valOf('Sch');
                const bDia = valOf('DiaIn');

                if (bSch && bDia) pairs.push({ sch: bSch, dia: parseFloat(bDia) });

                if (scheduleThicknessEndpoint) {
                    for (const p of pairs) {
                        if (!p.sch || isNaN(p.dia)) continue;
                        try {
                            const url = `${scheduleThicknessEndpoint}?sch=${encodeURIComponent(p.sch)}&dia=${encodeURIComponent(p.dia)}${lineMaterial ? `&material=${encodeURIComponent(lineMaterial)}` : ''}`;
                            const res = await fetch(url, { headers: { 'Accept': 'application/json' } });
                            if (!res.ok) continue;
                            const j = await res.json();
                            if (j && j.thickness) {
                                thickInput.value = j.thickness;
                                thickInput.dataset.autoGenerated = '1';
                                if (!initializing) markRowDirty(tr);
                                refreshWps(tr, allowOverride, allowOverride);
                                break;
                            }
                        } catch { }
                    }
                }

                if (!thickInput.value) refreshWps(tr, false, false);
            } else {
                if (thickInput.dataset.autoGenerated === '1') {
                    const prev = thickInput.value;
                    thickInput.value = '';
                    await autoFillOlThick(tr);
                    if (!thickInput.value) thickInput.value = prev;
                } else {
                    refreshWps(tr, allowOverride, allowOverride);
                }
            }
        }

        // Heat number handling
        function initHeatSelects(tr) {
            ['HeatNumberA', 'HeatNumberB'].forEach(n => {
                const sel = tr.querySelector(`select[data-name="${n}"]`);
                if (sel) {
                    sel.addEventListener('change', () => { markRowDirty(tr); });
                    if (window.__makeSearchableSelect) window.__makeSearchableSelect(sel);
                }
            });
        }

        function valuesForHeat(tr, side) {
            const dia = parseFloat((tr.querySelector('input[data-name="DiaIn"]')?.value || '').trim());
            const olDiaParsed = parseFloat((tr.querySelector('input[data-name="OlDia"]')?.value || '').trim());
            const sch = (tr.querySelector('input[data-name="Sch"]').value || '').trim();
            const olSch = (tr.querySelector('input[data-name="OlSch"]')?.value || '').trim();
            const matEl = tr.querySelector(`select[data-name="Material${side}"]`) ||
                tr.querySelector(`input[data-name="Material${side}"]`);
            const gradeEl = tr.querySelector(`select[data-name="Grade${side}"]`) ||
                tr.querySelector(`input[data-name="Grade${side}"]`);
            const mat = (matEl?.value || '').trim();
            const grade = (gradeEl?.value || '').trim();
            const pid = document.getElementById('projectSelect')?.value;

            return {
                pid, dia: (!isNaN(dia) ? dia : null), olDia: (!isNaN(olDiaParsed) ? olDiaParsed : null),
                sch, olSch, mat, grade
            };
        }

        async function refreshHeatOptions(tr, side) {
            try {
                const { pid, dia, olDia, sch, olSch, mat, grade } = valuesForHeat(tr, side);
                const sel = tr.querySelector(`select[data-name="HeatNumber${side}"]`);

                if (!sel || !pid || !dia || !mat || !grade || !sch) return;

                const params = new URLSearchParams({
                    projectId: String(pid),
                    side: String(side),
                    diaIn: String(dia),
                    material: mat,
                    grade: grade,
                    sch: sch
                });

                if (olDia != null) params.append('olDiaIn', String(olDia));
                if (olSch) params.append('olSch', olSch);

                const res = await fetch(`${heatEndpoint}?${params.toString()}`, {
                    headers: { 'Accept': 'application/json' }
                });

                if (!res.ok) return;
                const list = await res.json();
                const cur = sel.value || '';

                while (sel.options.length > 0) sel.remove(0);

                const add = (val, text, selected, desc) => {
                    const o = document.createElement('option');
                    o.value = val;
                    o.textContent = text;
                    if (selected) o.selected = true;
                    if (desc) o.title = desc;
                    sel.appendChild(o);
                };

                add('', '', cur === '');

                if (Array.isArray(list)) {
                    list.forEach(item => {
                        if (item && typeof item === 'object') {
                          const h = item.heat || '';
                          const d = item.description || '';
                          add(h, h, h === cur, d);
                        } else if (typeof item === 'string') {
                          add(item, item, item === cur);
                        }
                    });

                    if (cur && !list.some(it =>
                        (typeof it === 'string' ? it === cur : (it && typeof it === 'object' && it.heat === cur))
                    )) {
                        add(cur, cur, true);
                    }
                }

                if (window.__makeSearchableSelect) window.__makeSearchableSelect(sel);
            } catch {
                // Handle error silently
            }
        }

        function bindHeatDependencies(tr) {
            const debA = debounce(() => refreshHeatOptions(tr, 'A'), 300);
            const debB = debounce(() => refreshHeatOptions(tr, 'B'), 300);

            const map = [
                { sel: 'input[data-name="DiaIn"]', f: () => { debA(); debB(); } },
                { sel: 'input[data-name="OlDia"]', f: () => { debA(); debB(); } },
                { sel: 'input[data-name="Sch"]', f: () => { debA(); debB(); } },
                { sel: 'input[data-name="OlSch"]', f: () => { debA(); debB(); } },
                { sel: 'select[data-name="MaterialA"], input[data-name="MaterialA"]', f: debA },
                { sel: 'select[data-name="MaterialB"], input[data-name="MaterialB"]', f: debB },
                { sel: 'select[data-name="GradeA"], input[data-name="GradeA"]', f: debA },
                { sel: 'select[data-name="GradeB"], input[data-name="GradeB"]', f: debB }
            ];

            map.forEach(x => {
                const els = tr.querySelectorAll(x.sel);
                els.forEach(el => {
                    if (!el.dataset.heatBind) {
                        el.addEventListener('input', x.f);
                        el.addEventListener('change', x.f);
                        el.dataset.heatBind = '1';
                    }
                });
            });
        }

        // Add dynamic listeners to row
        function addDynamicListeners(tr) {
            const relatedNames = ['Sch', 'DiaIn'];
            relatedNames.forEach(n => {
                const inp = tr.querySelector(`input[data-name="${n}"]`);
                if (inp && !inp.dataset.dyn) {
                    inp.addEventListener('input', debounce(async () => {
                        const thick = tr.querySelector('input[data-name="OlThick"]');
                        if (thick && thick.dataset.userEdited !== '1') {
                            thick.dataset.autoGenerated = '1';
                            thick.value = '';
                        }
                        await autoFillOlThick(tr);

                        if (tr && tr.dataset.wpsActivated !== '1') {
                            if (!isThWeldType(tr)) activateWpsDefault(tr);
                            else {
                                clearWps(tr);
                                clearTackWelder(tr);
                            }
                        }
                    }, 300));
                    inp.dataset.dyn = '1';
                }
            });

            const thick = tr.querySelector('input[data-name="OlThick"]');
            if (thick && !thick.dataset.dyn) {
                thick.addEventListener('change', () => {
                    const trParent = thick.closest('tr');
                    if (trParent) {
                        if (trParent.dataset.wpsActivated !== '1') {
                            if (!isThWeldType(trParent)) activateWpsDefault(trParent);
                            else {
                                clearWps(trParent);
                                clearTackWelder(trParent);
                            }
                        } else {
                            const allow = shouldAutoOverrideWps(trParent);
                            refreshWps(trParent, allow, allow);
                            if (isThWeldType(trParent)) {
                                clearWps(trParent);
                                clearTackWelder(trParent);
                            }
                        }
                    }
                });
                thick.dataset.dyn = '1';
            }

            const wtSel = tr.querySelector('select[data-name="WeldType"]');
            if (wtSel && !wtSel.dataset.letbind) {
                wtSel.addEventListener('change', () => {
                    if (isLetWeldType(tr)) {
                        const thick = tr.querySelector('input[data-name="OlThick"]');
                        if (thick && thick.dataset.autoGenerated === '1') {
                            thick.value = '';
                            delete thick.dataset.autoGenerated;
                        }
                    } else {
                        tr.classList.remove('require-olthick');
                        const thick = tr.querySelector('input[data-name="OlThick"]');
                        if (thick && (!thick.value || thick.value.trim())) autoFillOlThick(tr);
                    }

                    if (isThWeldType(tr)) {
                        clearWps(tr);
                        clearTackWelder(tr);
                    }

                    loadRowRfiOptions(tr);
                    refreshWps(tr, false, false);
                });
                wtSel.dataset.letbind = '1';
            }

            const locSel = tr.querySelector('select[data-name="Location"]');
            if (locSel && !locSel.dataset.rfibind) {
                locSel.addEventListener('change', () => { loadRowRfiOptions(tr); });
                locSel.dataset.rfibind = '1';
            }

            // FitupDate hold binding
            const fitDateInp = tr.querySelector('input[data-name="FitupDate"]');
            if (fitDateInp && !fitDateInp.dataset.holdBind) {
                function handleFitupDateHold() {
                    try {
                        const onHold = !!tr.querySelector('.col-jointno.hold-alert, .col-spno.hold-alert, .col-sheet.hold-alert');
                        if (!onHold) return;

                        const remInp = tr.querySelector('input[data-name="Remarks"]');
                        if (remInp && !remInp.disabled) {
                            const note = 'Fabricated on hold';
                            const cur = (remInp.value || '').trim();
                            if (cur.length === 0) remInp.value = note;
                            else if (!cur.toLowerCase().includes(note.toLowerCase())) {
                                remInp.value = cur + ' | ' + note;
                            }
                            remInp.dispatchEvent(new Event('input', { bubbles: true }));
                            markRowDirty(tr);
                        }
                    } catch (e) { }
                }

                fitDateInp.addEventListener('change', handleFitupDateHold);
                fitDateInp.addEventListener('input', handleFitupDateHold);
                fitDateInp.dataset.holdBind = '1';
            }

            bindHeatDependencies(tr);
            initHeatSelects(tr);

            tr.querySelectorAll('select[data-name="MaterialA"], select[data-name="MaterialB"], select[data-name="GradeA"], select[data-name="GradeB"], select[data-name="RfiId"]').forEach(sel => {
                if (window.__makeSearchableSelect) window.__makeSearchableSelect(sel);
                sel.addEventListener('change', () => { markRowDirty(tr); });
            });

            const rfiSel = tr.querySelector('select[data-name="RfiId"]');
            if (rfiSel && !rfiSel.dataset.bound) {
                rfiSel.addEventListener('focus', () => {
                    if (rfiSel.options.length <= 1 || !rfiSel.dataset.loaded) {
                        loadRowRfiOptions(tr);
                    }
                });
                rfiSel.dataset.bound = '1';
            }
        }

        function activateWpsDefault(tr) {
            if (isThWeldType(tr)) {
                clearWps(tr);
                clearTackWelder(tr);
            } else {
                tr.dataset.wpsActivated = '1';
                const allow = shouldAutoOverrideWps(tr);
                refreshWps(tr, allow, allow);
            }
        }

        function addActivationTriggers(tr) {
            const fitRep = tr.querySelector('input[data-name="FitupReport"]');
            if (fitRep && !fitRep.dataset.wpsAct) {
                fitRep.addEventListener('dblclick', () => {
                    if (isThWeldType(tr)) {
                        clearWps(tr);
                        clearTackWelder(tr);
                    } else {
                        activateWpsDefault(tr);
                    }
                });
                fitRep.dataset.wpsAct = '1';
            }

            const thick = tr.querySelector('input[data-name="OlThick"]');
            if (thick && !thick.dataset.wpsAct) {
                thick.addEventListener('input', () => {
                    if (thick.value.trim() !== '') {
                        if (isThWeldType(tr)) {
                            clearWps(tr);
                            clearTackWelder(tr);
                        } else {
                            activateWpsDefault(tr);
                        }
                    }
                });
                thick.dataset.wpsAct = '1';
            }
        }

        // Initialize all existing rows
        document.querySelectorAll('#fitupTable tbody tr').forEach(tr => {
            tr.dataset.initializing = '1';
            addStatusHandlers(tr);
            initRowBaseline(tr, false);
            addDynamicListeners(tr);
            autoFillOlThick(tr);
            refreshWps(tr, false, false);
            addActivationTriggers(tr);
            attachRevInsert(tr);
            initHeatSelects(tr);
            refreshHeatOptions(tr, 'A');
            refreshHeatOptions(tr, 'B');
            loadRowRfiOptions(tr);
            updateRowState(tr);
            delete tr.dataset.initializing;
        });

        validateLetThicknesses(false, true);

        // Non-DWG view specific handlers
        if (window.headerView !== 'DWG') {
            const confirmBtn = document.getElementById('confirmAllBtn');
            const unconfirmBtn = document.getElementById('unconfirmAllBtn');
            const clearRemarksBtn = document.getElementById('clearRemarksBtn');

            function setConfirmAll(flag, options) {
                const opts = options || {};
                const rows = [...document.querySelectorAll('#fitupTable tbody tr')].filter(r => r.style.display !== 'none');
                const total = rows.length;
                let changed = 0;

                rows.forEach(r => {
                    const cb = r.querySelector('input[data-name="FitupConfirmed"]');
                    if (cb && cb.checked !== flag) {
                        cb.checked = flag;
                        cb.dispatchEvent(new Event('change', { bubbles: true }));
                        changed++;
                    }
                });

                if (!opts.silent) {
                    if (total === 0) {
                        showStatus('No rows loaded.', false);
                    } else if (changed === 0) {
                        showStatus('All displayed rows already confirmed.', true);
                    } else {
                        const action = flag ? 'Confirmed' : 'Unconfirmed';
                        showStatus(`${action} ${changed} row(s). Remember to Save.`, true);
                    }
                }

                return { changed, total };
            }

            function clearRemarks() {
                const rows = [...document.querySelectorAll('#fitupTable tbody tr')].filter(r => r.style.display !== 'none');
                let cleared = 0;

                rows.forEach(r => {
                    const inp = r.querySelector('input[data-name="Remarks"]');
                    if (inp && inp.value.trim() !== '') {
                        inp.value = '';
                        inp.dispatchEvent(new Event('input', { bubbles: true }));
                    }
                });

                showStatus(cleared > 0 ? `Cleared remarks in ${cleared} row(s). Remember to Save.` : 'No remarks to clear.', cleared > 0);
            }

            confirmBtn?.addEventListener('click', () => {
                const { changed, total } = setConfirmAll(true, { silent: true });
                if (total === 0) {
                    showStatus('No rows loaded.', false);
                    return;
                }
                if (changed === 0) {
                    showStatus('All displayed rows already confirmed.', true);
                    return;
                }

                // Ensure header Fit-up Date aligns with displayed rows (needed for Report view so
                // CompleteDailyFitup uses the correct day and inserts the UC record / email).
                const fitHdr = document.getElementById('fitupDateInput');
                if (fitHdr && (!fitHdr.value || (typeof window.headerView === 'string' && window.headerView.toUpperCase() === 'REPORT'))) {
                    const firstRowDate = [...document.querySelectorAll('#fitupTable tbody input[data-name="FitupDate"]')]
                        .map(inp => (inp.value || '').trim())
                        .find(v => v);
                    if (firstRowDate) {
                        fitHdr.value = firstRowDate;
                        try { fitHdr.dispatchEvent(new Event('change', { bubbles: true })); } catch { }
                    }
                }

                showStatus(`Confirmed ${changed} row(s). Remember to Save.`, true);
            });

            unconfirmBtn?.addEventListener('click', () => setConfirmAll(false));
            clearRemarksBtn?.addEventListener('click', clearRemarks);
        }
    }

    // ==================== Confirm All Modal ====================
    function initializeConfirmAllModal() {
        window.openConfirmAllModal = function (message) {
            return new Promise((resolve) => {
                const modal = document.getElementById('confirmUnconfirmedRowsModal');
                const msgEl = document.getElementById('confirmUnconfirmedRowsMessage');
                const okBtn = document.getElementById('confirmUnconfirmedRowsOk');

                if (msgEl) msgEl.textContent = message || '';

                function cleanup(result) {
                    try {
                        modal.style.display = 'none';
                        modal.setAttribute('aria-hidden', 'true');
                    } catch { }

                    if (okBtn) okBtn.removeEventListener('click', onOk);
                    modal.removeEventListener('click', onBackdrop);
                    document.removeEventListener('keydown', onEsc);
                    resolve(result);
                }

                function onOk() { cleanup(true); }
                function onBackdrop(e) { if (e.target === modal) cleanup(false); }
                function onEsc(e) { if (e.key === 'Escape') cleanup(false); }

                if (okBtn) okBtn.addEventListener('click', onOk);
                modal.addEventListener('click', onBackdrop);
                document.addEventListener('keydown', onEsc);

                modal.style.display = 'flex';
                modal.setAttribute('aria-hidden', 'false');
                try {
                    (modal.querySelector('.close-modal') || okBtn)?.focus();
                } catch { }
            });
        };

        window.markRowsConfirmed = function (rows) {
            let changed = 0;
            rows.forEach(r => {
                const cb = r.querySelector('input[data-name="FitupConfirmed"]');
                if (cb && cb.checked !== true) {
                    cb.checked = true;
                    cb.dispatchEvent(new Event('change', { bubbles: true }));
                    changed++;
                }
            });
            return changed;
        };

        window.closeConfirmUnconfirmedRowsModal = function () {
            try {
                const modal = document.getElementById('confirmUnconfirmedRowsModal');
                if (!modal) return;
                const evt = new MouseEvent('click', { bubbles: true, cancelable: true });
                modal.dispatchEvent(evt);
                modal.style.display = 'none';
                modal.setAttribute('aria-hidden', 'true');
            } catch (e) {
                try {
                    const m = document.getElementById('confirmUnconfirmedRowsModal');
                    if (m) {
                        m.style.display = 'none';
                        m.setAttribute('aria-hidden', 'true');
                    }
                } catch { }
            }
        };
    }

    // ==================== Schedule Menus ====================
    function initializeScheduleMenus() {
        const schedulesEndpoint = schedulesForDiameterEndpoint || '';
        if (!schedulesEndpoint) return;

        const cache = new Map();
        const scheduleInputSelector = 'input[data-name="Sch"], input[data-name="OlSch"]';
        const diaSelector = 'input[data-name="DiaIn"], input[data-name="OlDia"]';

        async function fetchSchedules(dia) {
            if (!schedulesEndpoint || dia == null || isNaN(dia)) return [];
            const key = Number(dia).toString();
            if (cache.has(key)) return cache.get(key);
            try {
                const res = await fetch(`${schedulesEndpoint}?dia=${encodeURIComponent(key)}`, {
                    headers: { 'Accept': 'application/json' }
                });
                if (!res.ok) {
                    cache.set(key, []);
                    return [];
                }
                const payload = await res.json();
                const normalized = Array.isArray(payload)
                    ? payload.map(item => {
                        if (typeof item === 'string') return item;
                        if (item && typeof item === 'object') {
                            return item.sch ?? item.Sch ?? item.value ?? item.Value ?? '';
                        }
                        return '';
                    }).filter(v => typeof v === 'string' && v.trim().length > 0)
                    : [];
                cache.set(key, normalized);
                return normalized;
            } catch {
                cache.set(key, []);
                return [];
            }
        }

        function positionMenu(menu, input, container) {
            const inputRect = input.getBoundingClientRect();
            const containerRect = (container || document.body).getBoundingClientRect();
            const viewportWidth = window.innerWidth || document.documentElement.clientWidth;
            const viewportHeight = window.innerHeight || document.documentElement.clientHeight;
            const spaceBelow = viewportHeight - inputRect.bottom - 8;
            const spaceAbove = inputRect.top - 8;
            const defaultMax = 190;
            const openUp = spaceBelow < Math.min(defaultMax, 140) && spaceAbove > spaceBelow;
            const maxHeight = Math.max(80, Math.min(defaultMax, openUp ? spaceAbove : spaceBelow));
            menu.style.maxHeight = `${maxHeight}px`;

            let left = Math.max(0, inputRect.left - containerRect.left);
            const menuWidth = Math.max(inputRect.width, 120);
            const overflowRight = (left + menuWidth + containerRect.left) - viewportWidth + 8;
            if (overflowRight > 0) {
                left = Math.max(0, left - overflowRight);
            }
            menu.style.left = `${left}px`;
            menu.style.minWidth = `${inputRect.width}px`;

            if (openUp) {
                const bottom = containerRect.bottom - inputRect.top + 2;
                menu.style.top = 'auto';
                menu.style.bottom = `${bottom}px`;
            } else {
                const top = inputRect.bottom - containerRect.top + 2;
                menu.style.top = `${top}px`;
                menu.style.bottom = 'auto';
            }
        }

        function removeMenu(input) {
            if (!input) return;
            const td = input.closest('td');
            const menu = td?.querySelector('.sch-float-menu');
            if (menu) menu.remove();
            input.classList.remove('dropdown-open');
            delete input.dataset.menuOpen;
        }

        function removeAllMenus() {
            document.querySelectorAll('.sch-float-menu').forEach(menu => menu.remove());
            document.querySelectorAll(scheduleInputSelector).forEach(inp => {
                inp.classList.remove('dropdown-open');
                delete inp.dataset.menuOpen;
            });
        }

        function buildMenu(items, anchorInput, placeholder) {
            removeMenu(anchorInput);
            const menu = document.createElement('div');
            menu.className = 'sch-float-menu';

            if (items && items.length) {
                items.forEach(val => {
                    const entry = document.createElement('div');
                    entry.className = 'sch-float-item';
                    entry.textContent = val;
                    entry.addEventListener('mousedown', (e) => {
                        e.preventDefault();
                        anchorInput.value = val;
                        anchorInput.dispatchEvent(new Event('input', { bubbles: true }));
                        anchorInput.dispatchEvent(new Event('change', { bubbles: true }));
                        removeMenu(anchorInput);
                    });
                    menu.appendChild(entry);
                });
            } else {
                const empty = document.createElement('div');
                empty.className = 'sch-float-item disabled';
                empty.textContent = placeholder || 'No schedules';
                menu.appendChild(empty);
            }

            const td = anchorInput.closest('td');
            (td || document.body).appendChild(menu);
            positionMenu(menu, anchorInput, td);
            menu.style.display = 'block';
            anchorInput.classList.add('dropdown-open');
            anchorInput.dataset.menuOpen = '1';
        }

        async function showMenu(input) {
            if (!input || input.disabled) return;
            removeAllMenus();
            const tr = input.closest('tr');
            if (!tr) return;

            const field = input.getAttribute('data-name') === 'OlSch' ? 'OlDia' : 'DiaIn';
            let diaText = tr.querySelector(`input[data-name="${field}"]`)?.value || '';
            if (!diaText && field === 'OlDia') {
                diaText = tr.querySelector('input[data-name="DiaIn"]')?.value || '';
            }

            const dia = parseFloat((diaText || '').trim());
            if (isNaN(dia)) {
                buildMenu([], input, 'Enter diameter first');
                return;
            }

            const list = await fetchSchedules(dia);
            buildMenu(list, input, list.length === 0 ? 'No schedules' : '');
        }

        function toggleMenu(input) {
            if (!input) return;
            if (input.dataset.menuOpen === '1') {
                removeMenu(input);
            } else {
                showMenu(input);
            }
        }

        function ensureCaret(input) {
            const td = input.closest('td');
            if (!td || td.querySelector('.sch-caret')) return;
            td.style.position = 'relative';
            const caret = document.createElement('span');
            caret.className = 'sch-caret';
            td.appendChild(caret);
        }

        document.addEventListener('focusin', (e) => {
            if (e.target && e.target.matches(scheduleInputSelector)) {
                ensureCaret(e.target);
            } else {
                removeAllMenus();
            }
        });

        document.addEventListener('click', (e) => {
            if (e.target && e.target.classList && e.target.classList.contains('sch-caret')) {
                const input = e.target.closest('td')?.querySelector(scheduleInputSelector);
                if (input && !input.disabled) {
                    toggleMenu(input);
                }
                return;
            }
            if (!e.target.closest('.sch-float-menu') && !(e.target.matches && e.target.matches(scheduleInputSelector))) {
                removeAllMenus();
            }
        });

        document.addEventListener('dblclick', (e) => {
            if (!e.target || !e.target.matches || !e.target.matches(scheduleInputSelector)) return;
            if (e.target.disabled) return;
            toggleMenu(e.target);
        });

        document.addEventListener('keydown', (e) => {
            if (!e.target || !e.target.matches || !e.target.matches(scheduleInputSelector)) return;
            if (e.key === 'ArrowDown') {
                e.preventDefault();
                showMenu(e.target);
            } else if (e.key === 'Escape') {
                removeMenu(e.target);
            }
        });

        const handleDiaChange = (e) => {
            if (!e.target || !e.target.matches || !e.target.matches(diaSelector)) return;
            const row = e.target.closest('tr');
            if (!row) return;
            row.querySelectorAll(scheduleInputSelector).forEach(inp => {
                if (inp.dataset.menuOpen === '1') {
                    showMenu(inp);
                }
            });
        };

        document.addEventListener('input', handleDiaChange);
        document.addEventListener('change', handleDiaChange);
    }

    // ==================== Filter Toggle ====================
    function initializeFilterToggle() {
        const toggleBtn = document.getElementById('filterToggle');
        const header = document.querySelector('.fitup-header');
        if (!toggleBtn || !header) return;

        const headerForm = document.getElementById('filterForm');
        const weldFilters = document.getElementById('weldFilters');
        const sections = [headerForm, weldFilters].filter(Boolean);
        const storageKey = 'PMS:DailyFitup:FiltersCollapsed';

        const applyState = (collapsed) => {
            sections.forEach(section => {
                if (!section) return;
                section.hidden = collapsed;
            });
            header.classList.toggle('filters-collapsed', collapsed);
            toggleBtn.textContent = collapsed ? 'Show Filters' : 'Hide Filters';
            toggleBtn.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
        };

        let collapsed = false;
        try {
            collapsed = localStorage.getItem(storageKey) === '1';
        } catch { /* ignore */ }

        applyState(collapsed);

        toggleBtn.addEventListener('click', () => {
            collapsed = !collapsed;
            applyState(collapsed);
            try {
                localStorage.setItem(storageKey, collapsed ? '1' : '0');
            } catch { /* ignore */ }
            if (typeof resizeTableWrapper === 'function') resizeTableWrapper();
        });
    }
    // ==================== Table Layout Functions ====================
    function resizeTableWrapper() {
        const wrapper = document.getElementById('tableWrapper');
        if (!wrapper) return;
        var rect = wrapper.getBoundingClientRect();
        var footer = document.querySelector('.footer');
        var footerH = footer ? footer.offsetHeight : 42;
        var margin = 4;
        var available = window.innerHeight - rect.top - footerH - margin;
        if (available < 120) available = 120;
        wrapper.style.maxHeight = available + 'px';
    }

    function resizeBottomScroll() {
        const inner = document.getElementById('bottomScrollInner');
        const table = document.getElementById('fitupTable');
        const wrapper = document.getElementById('tableWrapper');
        const bottom = document.getElementById('bottomScroll');

        if (inner) {
            var contentWidth = (wrapper ? wrapper.scrollWidth : 0) || (table ? table.scrollWidth : 0);
            if (contentWidth > 0) {
                inner.style.width = contentWidth + 'px';
            }
        }

        if (bottom) {
            const max = bottom.scrollWidth - bottom.clientWidth;
            if (bottom.scrollLeft > max) {
                bottom.scrollLeft = Math.max(0, max);
            }
        }
    }

    function bindScrollSync() {
        const bottom = document.getElementById('bottomScroll');
        const wrapper = document.getElementById('tableWrapper');

        if (!bottom || bottom.dataset.bound) return;

        const apply = () => {
            if (wrapper) wrapper.scrollLeft = bottom.scrollLeft;
        };

        bottom.addEventListener('scroll', apply);

        if (wrapper) {
            wrapper.addEventListener('wheel', e => {
                if (Math.abs(e.deltaX) > 0) {
                    bottom.scrollLeft += e.deltaX;
                    e.preventDefault();
                }
            }, { passive: false });
        }

        if (window.ResizeObserver) {
            const table = document.getElementById('fitupTable');
            if (table) {
                const ro = new ResizeObserver(() => { resizeTableWrapper(); resizeBottomScroll(); });
                ro.observe(table);
            }
        }

        window.addEventListener('resize', () => { resizeTableWrapper(); resizeBottomScroll(); });
        apply();
        bottom.dataset.bound = '1';
    }

    function markFrozen() {
        const isDwg = window.headerView === 'DWG';
        const classes = isDwg ?
            ['col-sn', 'col-location', 'col-jointno', 'col-jadd', 'col-wtype'] :
            ['col-sn', 'col-layout', 'col-location', 'col-jointno', 'col-jadd', 'col-wtype'];

        const table = document.getElementById('fitupTable');
        if (!table) return;

        classes.forEach(cls => {
            table.querySelectorAll('th.' + cls + ', td.' + cls).forEach(cell => {
                cell.dataset.frozen = '1';
                cell.classList.add('frozen-col');
            });
        });

        computeFrozenLefts();
    }

    function computeFrozenLefts() {
        frozenMeta = [];
        const table = document.getElementById('fitupTable');
        if (!table) return;

        const row = table.querySelector('thead tr');
        if (!row) return;

        const frozenCells = [...row.querySelectorAll('[data-frozen="1"]')];
        let left = 0;

        frozenCells.forEach(c => {
            const w = c.getBoundingClientRect().width;
            frozenMeta.push({ cls: [...c.classList].find(x => x.startsWith('col-')), left });
            left += w;
        });

        applyFrozenPositions();
    }

    function applyFrozenPositions() {
        const table = document.getElementById('fitupTable');
        if (!table) return;

        frozenMeta.forEach(m => {
            if (!m.cls) return;
            table.querySelectorAll('th.' + m.cls + ', td.' + m.cls).forEach(cell => {
                cell.style.left = m.left + 'px';
                cell.style.position = 'sticky';
            });
        });
    }

    function persistScrollInit() {
        try {
            const wrapper = document.getElementById('tableWrapper');
            const bottom = document.getElementById('bottomScroll');
            window.pms_scrollSnapshot = {
                wrapperLeft: wrapper ? wrapper.scrollLeft : 0,
                bottomLeft: bottom ? bottom.scrollLeft : 0
            };

            setTimeout(() => {
                try {
                    if (typeof markFrozen === 'function') markFrozen();
                    if (typeof computeFrozenLefts === 'function') computeFrozenLefts();
                    if (typeof resizeTableWrapper === 'function') resizeTableWrapper();
                    if (typeof bindScrollSync === 'function') bindScrollSync();
                    if (typeof resizeBottomScroll === 'function') resizeBottomScroll();

                    const w = document.getElementById('tableWrapper');
                    const b = document.getElementById('bottomScroll');
                    if (window.pms_scrollSnapshot) {
                        if (w) {
                            try { w.scrollLeft = window.pms_scrollSnapshot.wrapperLeft; } catch { }
                        }
                        if (b) {
                            try { b.scrollLeft = window.pms_scrollSnapshot.bottomLeft; } catch { }
                        }
                    }
                } catch { }
            }, 50);
        } catch { }
    }

    // Expose layout helpers so Razor inline scripts (persistScrollInit, etc.) can reuse them
    window.markFrozen = markFrozen;
    window.computeFrozenLefts = computeFrozenLefts;
    window.applyFrozenPositions = applyFrozenPositions;
    window.bindScrollSync = bindScrollSync;
    window.resizeBottomScroll = resizeBottomScroll;
    window.resizeTableWrapper = resizeTableWrapper;
    window.persistScrollInit = persistScrollInit;

    // Initialize table layout – deferred because dailyfitup.js is loaded with defer;
    // the table may not have completed layout when the IIFE body runs.
    function initTableLayout() {
        markFrozen();
        resizeTableWrapper();
        bindScrollSync();
        resizeBottomScroll();
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', function () {
            requestAnimationFrame(initTableLayout);
        });
    } else {
        requestAnimationFrame(initTableLayout);
    }

    // Debounce utility
    function debounce(fn, ms) {
        let t;
        return function () {
            clearTimeout(t);
            const a = arguments;
            const th = this;
            t = setTimeout(() => fn.apply(th, a), ms);
        };
    }
})();