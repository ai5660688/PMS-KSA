// Ensure shared status auto-hide helper is available globally
(function(){
    if(typeof window.scheduleStatusAutoHide !== 'function'){
        window.scheduleStatusAutoHide = function(snapshot){
            try{
                var host = document.getElementById('statusMsg');
                if(!host) return;
                var token = Date.now().toString();
                host.dataset.statusTimerToken = token;
                setTimeout(function(){
                    try{ if(host.dataset.statusTimerToken === token) host.textContent = ''; }catch{}
                }, 3000);
            }catch{}
        };
    }
})();

// DailyWelding-specific UI tweaks
// Repurpose dailyfitup-dwg helpers and adapt behaviors for welding context
(function () {
    function log(msg) { try { console.debug('[DailyWelding] ' + msg); } catch { } }

    // RFI cache and helpers (copied/adjusted from Daily Fit-up behavior)
    const __rfiCache = new Map();

    function extractBareRfi(txt) {
        try {
            var s = (txt || '').toString();
            // remove leading location prefix like "WS | ", "FW | ", "TH | "
            s = s.replace(/^\s*(WS\s*\|\s*|FW\s*\|\s*|TH\s*\|\s*)/i, '');
            var idx = s.indexOf(' | ');
            if (idx >= 0) return s.substring(0, idx).trim();
            return s.trim();
        } catch (e) { return (txt || '').toString(); }
    }

    function determineLocationBucket(tr) {
        try {
            if (!tr) return 'FW';
            var wt = '';
            try { wt = (tr.querySelector('.col-wtype')?.textContent || tr.querySelector('[data-name="WeldType"]')?.value || '').toString().trim(); } catch { wt = ''; }
            if (wt && wt.toUpperCase().indexOf('TH') >= 0) return 'TH';

            var rowLoc = '';
            try {
                rowLoc = (tr.getAttribute('data-location-code') || '').toString().trim();
                if (!rowLoc) {
                    var locCell = tr.querySelector('.col-location');
                    rowLoc = (locCell ? (locCell.textContent || '') : '').trim();
                }
            } catch { rowLoc = ''; }
            var up = (rowLoc || '').toUpperCase();
            if (!up) return 'FW';
            if (up.startsWith('WS') || up.indexOf('SHOP') >= 0 || up.indexOf('WORK') >= 0) return 'WS';
            if (up.startsWith('FW') || up.indexOf('FIELD') >= 0) return 'FW';
            if (up.startsWith('TH')) return 'TH';
            return 'FW';
        } catch (e) { return 'FW'; }
    }

    async function loadRowRfiOptions(tr, preselect) {
        try {
            if (!tr) return;
            var sel = tr.querySelector('select[data-name="RfiId"], select[name="RfiId"], select.rfi-select, select.rfi-async-select, select[data-name="RFI_ID_DWR"]');
            if (!sel) return;

            if (sel.dataset.refreshing === '1') return;
            sel.dataset.refreshing = '1';

            // preserve previous options in case fetch fails
            var savedOptions = Array.from(sel.options || []).map(function (o) { return { value: o.value, text: o.text, display: o.getAttribute && o.getAttribute('data-display') }; });
            var previous = sel.value || '';
            var prevText = '';
            try { var idx = sel.selectedIndex; if (idx >= 0) prevText = (sel.options[idx]?.text || '').trim(); } catch { prevText = ''; }

            var projectId = '';
            try { projectId = (document.getElementById('projectSelect')?.value || document.getElementById('projectId')?.value || document.getElementById('projectSelect')?.getAttribute('data-value') || '').toString(); } catch { projectId = ''; }
            if (!projectId) {
                try { delete sel.dataset.refreshing; } catch (e) { }
                return;
            }

            var bucket = determineLocationBucket(tr);
            var fitupIso = (document.getElementById('fitupDateInput')?.value || '').toString();
            if (!fitupIso) {
                var rowFitInp = tr.querySelector('input[data-name="FitupDate"], input[name="FitupDate"]');
                if (rowFitInp) fitupIso = rowFitInp.value || '';
            }

            var cacheKey = projectId + '|' + bucket + '|' + fitupIso;
            var items = __rfiCache.get(cacheKey);
            if (!items) {
                var url = '/Home/GetWeldingRfiOptions?projectId=' + encodeURIComponent(projectId) + '&location=' + encodeURIComponent(bucket);
                if (fitupIso) url += '&fitupDateIso=' + encodeURIComponent(fitupIso);
                try {
                    const controller = typeof AbortController !== 'undefined' ? new AbortController() : null;
                    const timer = controller ? setTimeout(function () { try { controller.abort(); } catch { } }, 8000) : null;
                    const resp = await fetch(url, { headers: { 'Accept': 'application/json' }, signal: controller ? controller.signal : undefined });
                    if (timer) try { clearTimeout(timer); } catch { }
                    if (resp.ok) {
                        items = await resp.json();
                    } else {
                        items = [];
                    }
                } catch (e) {
                    items = [];
                }
                if (!Array.isArray(items)) items = [];
                try { __rfiCache.set(cacheKey, items); } catch (e) { }
            }

            // rebuild options
            try {
                sel.innerHTML = '';

                // Always create a visible blank first option (non-breaking space) so native select and Select2 show a blank choice
                var empty = document.createElement('option');
                empty.value = '';
                empty.text = '\u00A0';
                try { empty.setAttribute('data-display', ''); } catch (e) { }
                sel.appendChild(empty);

                var foundPrev = false;
                var preId = preselect ? String(preselect.id || preselect.Id || preselect.RfiId || preselect.Rfi_ID || preselect) : '';
                var preText = '';
                if (preselect) { preText = extractBareRfi(preselect.value || preselect.Value || preselect.Display || preselect.display || preselect.text || ''); }

                (items || []).forEach(function (it) {
                    try {
                        var id = (it.id !== undefined && it.id !== null) ? it.id : (it.Id !== undefined && it.Id !== null ? it.Id : (it.RFI_ID !== undefined && it.RFI_ID !== null ? it.RFI_ID : ''));
                        // Skip additional blanks; we already inserted one
                        if(id === '' || id === null || typeof id === 'undefined') return;
                        var raw = (it.Value ?? it.value ?? it.SubCon_RFI_No ?? it.RFI_NO ?? it.Rfi_No ?? it.Display ?? it.display ?? String(id || ''));
                        var text = extractBareRfi(raw);
                        var o = document.createElement('option'); o.value = String(id || ''); o.text = text;
                        if (it.Display || it.display) try { o.setAttribute('data-display', String(it.Display || it.display)); } catch (e) { }
                        sel.appendChild(o);
                        if (String(id) === String(previous)) foundPrev = true;
                    } catch (e) { }
                });

                // Ensure only one blank option exists
                try {
                    var blanks = Array.from(sel.options || []).filter(o => (o.value || '') === '');
                    blanks.slice(1).forEach(o => { try { o.remove(); } catch(_) {} });
                } catch (e) { }

                // preselect header-provided option if present
                if (preId) {
                    var hasPre = Array.from(sel.options || []).some(o => String(o.value) === String(preId));
                    if (hasPre) {
                        sel.value = String(preId);
                    } else if (preText) {
                        var op = document.createElement('option'); op.value = String(preId); op.text = preText; op.selected = true; sel.appendChild(op);
                    }
                } else if (previous) {
                    if (foundPrev) {
                        sel.value = previous;
                    } else {
                        var ap = document.createElement('option'); ap.value = previous; ap.text = prevText || previous; ap.selected = true; sel.appendChild(ap);
                    }
                } else {
                    // ensure the empty option remains selected
                    try { sel.selectedIndex = 0; } catch (e) { }
                }

                // Respect locked value
                try {
                    var lockVal = sel.dataset.lockRfiValue || '';
                    if (sel.dataset.lockRfi === '1' && lockVal) {
                        var lockFound = Array.from(sel.options || []).some(o => String(o.value) === String(lockVal));
                        if (!lockFound) {
                            var lop = document.createElement('option'); lop.value = lockVal; lop.text = sel.dataset.lockRFiText || lockVal; sel.appendChild(lop);
                        }
                        sel.value = lockVal;
                    }
                } catch (e) { }

                sel.dataset.loaded = '1';

                // If Select2 is in use for this select, notify it of the change so the placeholder/selection is shown correctly
                try {
                    if (window.jQuery && jQuery && jQuery(sel).data('select2')) {
                        try { jQuery(sel).trigger('change.select2'); } catch (e) { }
                    }
                } catch (e) { }
            } catch (e) {
                // restore previous options on error
                try { sel.innerHTML = ''; savedOptions.forEach(function (it) { var opt = document.createElement('option'); opt.value = it.value; opt.text = it.text || it.value; if (it.display) opt.setAttribute('data-display', it.display); sel.appendChild(opt); }); sel.value = previous; } catch (er) { }
            }

            try { delete sel.dataset.refreshing; } catch (e) { }
        } catch (e) {
            try { if (tr) { var s = tr.querySelector('select[data-name="RfiId"], select[name="RfiId"], select.rfi-select, select.rfi-async-select, select[data-name="RFI_ID_DWR"]'); if (s) delete s.dataset.refreshing; } } catch (er) { }
            console.debug('[DailyWelding] loadRowRfiOptions error:', e);
        }
    }

    function refreshRowRfiOptions(tr) { try { loadRowRfiOptions(tr); } catch (e) { console.debug('[DailyWelding] refreshRowRfiOptions delegate error:', e); } }

    window.reloadRfiListsForAllRows = function () { try { var rows = Array.from(document.querySelectorAll('#fitupTable tbody tr')); rows.forEach(function (tr) { try { loadRowRfiOptions(tr); } catch (e) { } }); } catch (e) { } };

    function ensureConfirmAllModalHelpers() {
        try {
            if (typeof window.openConfirmAllModal !== 'function') {
                window.openConfirmAllModal = function (message) {
                    return new Promise(function (resolve) {
                        try {
                            var modal = document.getElementById('confirmUnconfirmedRowsModal');
                            var msgEl = document.getElementById('confirmUnconfirmedRowsMessage');
                            var okBtn = document.getElementById('confirmUnconfirmedRowsOk');

                            if (!modal) {
                                if (typeof window.confirm === 'function') {
                                    resolve(window.confirm(message || ''));
                                } else {
                                    resolve(true);
                                }
                                return;
                            }

                            if (msgEl) msgEl.textContent = message || '';

                            function cleanup(result) {
                                try {
                                    modal.style.display = 'none';
                                    modal.setAttribute('aria-hidden', 'true');
                                } catch { }
                                try { if (okBtn) okBtn.removeEventListener('click', onOk); } catch { }
                                try { modal.removeEventListener('click', onBackdrop); } catch { }
                                try { document.removeEventListener('keydown', onEsc); } catch { }
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
                            try { (modal.querySelector('.close-modal') || okBtn)?.focus(); } catch { }
                        } catch (e) {
                            try {
                                if (typeof window.confirm === 'function') {
                                    resolve(window.confirm(message || ''));
                                } else { resolve(true); }
                            } catch { resolve(true); }
                        }
                    });
                };
            }

            if (typeof window.closeConfirmUnconfirmedRowsModal !== 'function') {
                window.closeConfirmUnconfirmedRowsModal = function () {
                    try {
                        var modal = document.getElementById('confirmUnconfirmedRowsModal');
                        if (!modal) return;
                        var evt = new MouseEvent('click', { bubbles: true, cancelable: true });
                        modal.dispatchEvent(evt);
                        modal.style.display = 'none';
                        modal.setAttribute('aria-hidden', 'true');
                    } catch (e) {
                        try {
                            var m = document.getElementById('confirmUnconfirmedRowsModal');
                            if (m) {
                                m.style.display = 'none';
                                m.setAttribute('aria-hidden', 'true');
                            }
                        } catch { }
                    }
                };
            }

            if (typeof window.markRowsConfirmed !== 'function') {
                window.markRowsConfirmed = function (rows) {
                    var changed = 0;
                    try {
                        (rows || []).forEach(function (r) {
                            try {
                                var cb = r.querySelector('input[data-name="FitupConfirmed"], input[name="FitupConfirmed"]');
                                if (cb && cb.checked !== true) {
                                    cb.checked = true;
                                    try { cb.dispatchEvent(new Event('change', { bubbles: true })); } catch { }
                                    changed++;
                                }
                            } catch { }
                        });
                    } catch { }
                    return changed;
                };
            }

            // Ensure rows being bulk-confirmed also receive header welding/actual dates when blank
            if (typeof window.applyHeaderDatesToRows !== 'function') {
                window.applyHeaderDatesToRows = function (rows) {
                    try {
                        var hdrFit = document.getElementById('fitupDateInput');
                        var hdrAct = document.getElementById('actualDateInput');
                        var fitVal = hdrFit ? (hdrFit.value || '') : '';
                        var actVal = hdrAct ? (hdrAct.value || fitVal) : fitVal;
                        (rows || []).forEach(function (r) {
                            try {
                                var rowFit = r.querySelector('input[data-name="DATE_WELDED"], input[name="DATE_WELDED"], input[data-name="FitupDate"], input[name="FitupDate"]');
                                var rowAct = r.querySelector('input[data-name="ACTUAL_DATE_WELDED"], input[name="ACTUAL_DATE_WELDED"], input[data-name="ActualDate"], input[name="ActualDate"], input[data-name="ActualDateWelded"], input[name="ActualDateWelded"]');
                                var touched = false;
                                if (rowFit && !rowFit.value && fitVal) {
                                    rowFit.value = fitVal;
                                    touched = true;
                                    try { rowFit.dispatchEvent(new Event('change', { bubbles: true })); } catch { }
                                }
                                if (rowAct && !rowAct.value && actVal) {
                                    rowAct.value = actVal;
                                    touched = true;
                                    try { rowAct.dispatchEvent(new Event('change', { bubbles: true })); } catch { }
                                }
                                if (touched) {
                                    try { markRowDirty(r); } catch { }
                                }
                            } catch { }
                        });
                    } catch (e) {
                        console.debug('[DailyWelding] applyHeaderDatesToRows error:', e);
                    }
                };
            }
        } catch (e) {
            console.debug('[DailyWelding] ensureConfirmAllModalHelpers error:', e);
        }
    }

    function replaceButtonTextByPattern(id, pattern, replacement) {
        try {
            var el = document.getElementById(id);
            if (!el) return;
            el.textContent = (el.textContent || '').replace(pattern, replacement);
        } catch (e) {/*ignore*/ }
    }

    function hideAddJoint() {
        try {
            var b = document.getElementById('addRowBtn');
            if (b) b.style.display = 'none';
        } catch (e) { }
    }

    function markRowDirty(tr){
        try {
            if(!tr) return;
            // set a consistent dirty flag
            tr.dataset.dirty = '1';
            // add class if not already present
            if (!tr.classList.contains('dirty')) tr.classList.add('dirty');
            // mark child editable controls as dirty for downstream logic
            var controls = tr.querySelectorAll('input, select, textarea');
            controls.forEach(function(ctrl){
                try { ctrl.dataset.dirty = '1'; } catch(e){}
            });
            // notify listeners without causing bubbles to re-trigger handlers unexpectedly
            try {
                var evt = new CustomEvent('rowdirty', { detail: { row: tr }, bubbles: false });
                tr.dispatchEvent(evt);
            } catch(e){}
        } catch(e){}
    }

    // Utility: insert Welded on hold note when row is on hold
    function isRowOnHold(row){
        try { return !!(row && (row.querySelector('.col-locjoint.hold-alert') || row.querySelector('.col-spno.hold-alert') || row.querySelector('.col-sheet.hold-alert'))); } catch { return false; }
    }
    function insertHoldRemark(row){
        try {
            if(!row || !isRowOnHold(row)) return;
            var rem = row.querySelector('input[data-name="DWR_REMARKS"], input[name="DWR_REMARKS"], input[data-name="DwrRemarks"], input[name="DwrRemarks"], input[data-name="Remarks"], input[name="Remarks"]');
            if(!rem || rem.disabled) return;
            var note = 'Welded on hold';
            var cur = (rem.value || '').trim();
            if(cur.length === 0){ rem.value = note; }
            else if(cur.toLowerCase().indexOf(note.toLowerCase()) < 0){ rem.value = cur + ' | ' + note; }
            try { rem.dispatchEvent(new Event('input',{bubbles:true})); } catch(e) {}
            markRowDirty(row);
        } catch(e) { }
    }

    function setRowWeldingDefaults(tr) {
        if (!tr) return;
        try {
            var selects = Array.from(tr.querySelectorAll('select[data-name], select[name]'));
            if (selects.length === 0) return;

            function find(nameRegex) {
                return selects.find(s => {
                    var n = (s.getAttribute('data-name') || s.getAttribute('name') || '').toString();
                    return nameRegex.test(n);
                });
            }

            // robust regexes for expected fields (prefer welding-specific names)
            var rootA = find(/^(ROOT_A|TackWelder)$/i);
            var rootB = find(/^(ROOT_B|TackWelderB)$/i);
            var fillA = find(/^(FILL_A|TackWelderFillA)$/i);
            var fillB = find(/^(FILL_B|TackWelderFillB)$/i);
            var capA = find(/^(CAP_A|TackWelderCapA)$/i);
            var capB = find(/^(CAP_B|TackWelderCapB)$/i);

            var prefer = null;
            if (rootA && rootA.value && rootA.value.trim() !== '') prefer = rootA.value;
            // If no rootA value, but there is an option literally "ROOT A", prefer that
            if (!prefer && rootA) {
                var opt = Array.from(rootA.options || []).map(o => o.value || o.text || '').find(v => /^\s*ROOT\s*A\s*$/i.test(v));
                if (opt) prefer = opt;
            }

            if (prefer) {
                [rootB, fillA, fillB, capA, capB].forEach(s => {
                    if (!s) return;
                    var cur = (s.value || '').toString();
                    if (cur && cur.trim() !== '') return; // don't override user value
                    // try to pick matching option
                    var matchOpt = Array.from(s.options || []).find(o => (o.value || o.text || '').toString() === prefer);
                    if (matchOpt) { s.value = matchOpt.value; }
                    else {
                        // if prefer is like a label present in options with case-insensitive match
                        var ci = Array.from(s.options || []).find(o => (o.value || o.text || '').toString().toLowerCase() === (prefer || '').toLowerCase());
                        if (ci) s.value = ci.value;
                        else {
                            // fallback: if option 'ROOT A' exists
                            var ra = Array.from(s.options || []).find(o => /^\s*ROOT\s*A\s*$/i.test((o.value || o.text || '').toString()));
                            if (ra) s.value = ra.value;
                        }
                    }
                });
            }
        } catch (e) { /* swallow */ }
    }

    // Format date strings to DD-MM-YYYY if possible; otherwise return original
    function formatDate(val) {
        try {
            if (!val) return '';
            // Support input types date/datetime-local; split on non-digits
            var parts = String(val).replace(/[T\s].*$/, '').split(/[-/]/);
            if (parts.length === 3) {
                var y = parts[0].length === 4 ? parts[0] : parts[2];
                var m = parts[0].length === 4 ? parts[1] : parts[1];
                var d = parts[0].length === 4 ? parts[2] : parts[0];
                // pad
                d = (d.length === 1 ? '0' + d : d);
                m = (m.length === 1 ? '0' + m : m);
                return d + '-' + m + '-' + y;
            }
            return val;
        } catch { return val; }
    }

    // Try to extract Joint label from row
    function getJointLabel(tr) {
        try {
            var txt = '';
            var cell = tr.querySelector('.col-locjoint');
            if (cell) txt = (cell.textContent || '').trim();
            if (!txt) {
                var inp = tr.querySelector('input[data-name="WeldNumber"], input[name="WeldNumber"]');
                if (inp) txt = (inp.value || '').trim();
            }
            if (!txt) {
                // fallback to dataset id
                txt = tr.dataset.id ? ('ID ' + tr.dataset.id) : '';
            }
            return txt || 'Unknown';
        } catch { return 'Unknown'; }
    }

    // Compute and set validation state/messages for a row
    function validateWeldingRow(tr) {
        try {
            if (!tr) return;
            var fr = tr.querySelector('input[data-name="POST_VISUAL_INSPECTION_QR_NO"], input[name="POST_VISUAL_INSPECTION_QR_NO"], input[data-name="FitupReport"], input[name="FitupReport"], input[data-name="WeldingReport"], input[name="WeldingReport"]');
            var rowFit = tr.querySelector('input[data-name="DATE_WELDED"], input[name="DATE_WELDED"], input[data-name="FitupDate"], input[name="FitupDate"], input[data-name="WeldingDate"], input[name="WeldingDate"]');
            var rowAct = tr.querySelector('input[data-name="ACTUAL_DATE_WELDED"], input[name="ACTUAL_DATE_WELDED"], input[data-name="ActualDate"], input[name="ActualDate"], input[data-name="ActualDateWelded"], input[name="ActualDateWelded"]');
            var fitupInp = tr.querySelector('input[data-name="FitupDate"], input[name="FitupDate"]');

            var reportVal = fr ? (fr.value || '').trim() : '';
            var weldDateVal = rowFit ? (rowFit.value || '').trim() : '';
            var actualDateVal = rowAct ? (rowAct.value || '').trim() : '';
            var fitupVal = fitupInp ? (fitupInp.value || '').trim() : '';

            var messages = [];
            // Welding Report vs Welding/Actual Date presence checks
            if (weldDateVal && !reportVal) messages.push('Welding Report is required when Date Welded is provided.');
            if (reportVal && !weldDateVal) messages.push('Date Welded is required when Welding Report is provided.');
            if (actualDateVal && !reportVal) messages.push('Welding Report is required when Actual Date is provided.');
            if (reportVal && !actualDateVal) messages.push('Actual Date is required when Welding Report is provided.');

            // Ensure Actual Date is not after Welding (Date Welded)
            try {
                if (weldDateVal && actualDateVal) {
                    var wdt = new Date(weldDateVal);
                    var adt = new Date(actualDateVal);
                    if (isFinite(wdt.getTime()) && isFinite(adt.getTime())) {
                        var wDay = new Date(wdt.getFullYear(), wdt.getMonth(), wdt.getDate()).getTime();
                        var aDay = new Date(adt.getFullYear(), adt.getMonth(), adt.getDate()).getTime();
                        if (aDay > wDay) {
                            messages.push('Actual Date must be on or before Welding Date');
                        }
                    }
                }
            } catch (e) { }

            // Ordering issues collection

            // Fit-up vs Date Welded ordering when both available (Fit-up must be on-or-before Date Welded)
            if (fitupVal && weldDateVal) {
                try {
                    var f = new Date(fitupVal);
                    var w = new Date(weldDateVal);
                    if (isFinite(f.getTime()) && isFinite(w.getTime())) {
                        var fDay = new Date(f.getFullYear(), f.getMonth(), f.getDate()).getTime();
                        var wDay = new Date(w.getFullYear(), w.getMonth(), w.getDate()).getTime();
                        if (fDay > wDay) {
                            orderingIssues.push({ label: 'Date Welded', date: weldDateVal });
                        }
                    }
                } catch (e) {}
            }

            // Additional: enforce presence and ordering against other related inspection dates
            try {
                var candidateKeys = [
                    { keys: ['Root_PT_DATE','ROOT_PT_DATE','RootPtDate','Root_PT'], label: 'Root PT Date' },
                    { keys: ['OTHER_NDE_DATE','Other_NDE_DATE','OtherNdeDate','OTHER_NDE'], label: 'Other NDE Date' },
                    { keys: ['PMI_DATE','PmiDate','PMI'], label: 'PMI Date' },
                    { keys: ['NDE_REQUEST','Date_NDE_Request','DATE_NDE_WAS_REQUESTED','NdeRequest'], label: 'NDE Request Date' },
                    { keys: ['BSR_NDE_REQUEST','BSR_DATE_NDE_WAS_REQ','BsrNdeRequest'], label: 'BSR NDE Request Date' },
                    // include Fit-up as a candidate for presence rule (already used above for ordering)
                    { keys: ['FitupDate','FITUP_DATE'], label: 'Fit-up Date' }
                ];

                var found = [];
                candidateKeys.forEach(function(cand){
                    try {
                        for (var i = 0; i < cand.keys.length; i++) {
                            var k = cand.keys[i];
                            var sel = tr.querySelector('input[data-name="' + k + '"]') || tr.querySelector('input[name="' + k + '"]') || tr.querySelector('td[data-name="' + k + '"]') || tr.querySelector('.' + k.toLowerCase());
                            if (sel) {
                                var v = '';
                                try { v = (sel.value || sel.textContent || '').trim(); } catch (e) { v = ''; }
                                if (v) { found.push({ label: cand.label, value: v }); break; }
                            }
                        }
                    } catch (e) {}
                });

                // For each found candidate, require weld date presence and ensure weld <= candidate
                found.forEach(function(c){
                    try {
                        if (!weldDateVal) {
                            orderingIssues.push({ label: c.label, date: c.value, reason: 'missing-weld-date' });
                        } else {
                            var w = new Date(weldDateVal);
                            var cdt = new Date(c.value);
                            if (isFinite(w.getTime()) && isFinite(cdt.getTime())) {
                                var wDay2 = new Date(w.getFullYear(), w.getMonth(), w.getDate()).getTime();
                                var cDay = new Date(cdt.getFullYear(), cdt.getMonth(), cdt.getDate()).getTime();
                                if (wDay2 > cDay) {
                                    orderingIssues.push({ label: c.label, date: c.value });
                                }
                            }
                        }
                    } catch (e) {}
                });
            } catch (e) { }

            // Build uniform message when ordering issues exist
            var msg = '';
            if (orderingIssues.length > 0) {
                var joint = getJointLabel(tr);
                var missing = orderingIssues.some(function(x){ return x.reason === 'missing-weld-date'; });
                if (missing) {
                    msg = 'Joint ' + joint + ': Date Welded is required when related inspection dates exist.';
                } else {
                    var parts = orderingIssues.map(function(x){ return x.label + ' ' + formatDate(x.date); });
                    msg = 'Joint ' + joint + ': Date Welded must be on or before ' + parts.join(' or ') + '.';
                }
            }

            // If no ordering message, consider presence messages
            if (!msg && messages.length > 0) {
                var joint2 = getJointLabel(tr);
                msg = 'Joint ' + joint2 + ': ' + messages.join(' ');
            }

            // reflect UI state
            tr.dataset.validationMessage = msg;
            if (msg) {
                tr.classList.add('invalid');
                tr.title = msg;
                try {
                    var statusCell = tr.querySelector('.col-status, td.status');
                    if (statusCell) {
                        statusCell.textContent = msg;
                        statusCell.classList.add('text-danger');
                    }
                } catch (e){}

                // Ensure the row is scrolled into view if the message is new
                try {
                    var existing = document.querySelector('.tbl-validate-scroll');
                    if (existing) existing.classList.remove('tbl-validate-scroll');
                    tr.classList.add('tbl-validate-scroll');
                    setTimeout(function() {
                        try { tr.scrollIntoView({ behavior: 'smooth', block: 'center' }); } catch (e) { }
                        setTimeout(function() {
                          try { tr.classList.remove('tbl-validate-scroll'); } catch (e) { }
                        }, 500);
                    }, 50);
                } catch (e){}
            } else {
                tr.classList.remove('invalid');
                if (tr.title) tr.title = '';
                try {
                    var statusCell2 = tr.querySelector('.col-status, td.status');
                    if (statusCell2) {
                        statusCell2.textContent = '';
                        statusCell2.classList.remove('text-danger');
                    }
                } catch (e){}
            }
        } catch (e) { console.debug('[DailyWelding] validateWeldingRow error:', e); }
    }

    function refreshRowWpsOptions(tr, force) {
        try {
            if (!tr) return;
            var wpsSel = tr.querySelector('select[data-name="WPS_ID_DWR"], select[name="WPS_ID_DWR"], select[data-name="Wps"], select[name="Wps"]');
            var wpsInp = tr.querySelector('input[data-name="WPS_ID_DWR"], input[name="WPS_ID_DWR"], input[data-name="Wps"], input[name="Wps"]');
            if (!wpsSel && !wpsInp) return;

            // If already refreshing and not forced, skip
            var target = wpsSel || wpsInp;
            if (!force && target.dataset.wpsRefreshing === '1') return;
            target.dataset.wpsRefreshing = '1';

            // find projectId from header controls
            var projEl = document.getElementById('projectSelect') || document.getElementById('projectId');
            var projectId = projEl ? (projEl.value || projEl.getAttribute('data-value') || '') : '';
            if (!projectId) {
                // fallback to meta tag
                var meta = document.querySelector('meta[name="project-id"]');
                if (meta) projectId = meta.getAttribute('content') || '';
            }
            if (!projectId) { try { delete target.dataset.wpsRefreshing; } catch (e) { } return; }

            // derive row params
            var lineClass = tr.dataset.lineClass || '';
            if (!lineClass) {
                var lc = tr.querySelector('.col-lineclass');
                lineClass = lc ? (lc.textContent || '').trim() : '';
            }
            var schEl = tr.querySelector('input[data-name="Sch"], input[name="Sch"], select[data-name="Sch"], select[name="Sch"], td.col-sch');
            var sch = '';
            if (schEl) sch = (schEl.value || schEl.textContent || '').trim();
            var diaEl = tr.querySelector('input[data-name="DiaIn"], input[name="DiaIn"], td.col-dia');
            var dia = '';
            if (diaEl) dia = (diaEl.value || diaEl.textContent || '').trim();
            var olThickEl = tr.querySelector('input[data-name="OlThick"], input[name="OlThick"], td.col-olthick');
            var olThick = '';
            if (olThickEl) olThick = (olThickEl.value || olThickEl.textContent || '').trim();

            var url = '/Home/GetWpsCandidates?projectId=' + encodeURIComponent(projectId);
            if (lineClass) url += '&lineClass=' + encodeURIComponent(lineClass);
            if (sch) url += '&sch=' + encodeURIComponent(sch);
            if (dia) url += '&dia=' + encodeURIComponent(dia);
            if (olThick) url += '&olThick=' + encodeURIComponent(olThick);

            var previous = '';
            try { previous = target.value || ''; } catch (e) { previous = ''; }

            // small debounce
            setTimeout(function () {
                fetch(url, { method: 'GET', headers: { 'Accept': 'application/json' }, signal: AbortSignal.timeout ? AbortSignal.timeout(5000) : undefined })
                    .then(function (r) { return r.ok ? r.json() : []; })
                    .then(function (list) {
                        try {
                            if (!Array.isArray(list) || list.length === 0) return;

                            // If select element, rebuild options
                            if (wpsSel) {
                                // preserve any existing empty/default option
                                var keepEmpty = Array.from(wpsSel.options || []).some(function (o) { return !(o.value || '').trim(); });
                                wpsSel.innerHTML = '';
                                if (keepEmpty) {
                                    var em = document.createElement('option'); em.value = ''; em.text = '(none)'; wpsSel.appendChild(em);
                                }
                                list.forEach(function (it) {
                                    var opt = document.createElement('option');
                                    opt.value = String(it.id ?? it.Id ?? it.id ?? '');
                                    opt.text = String(it.wps ?? it.Wps ?? it.wps ?? '');
                                    wpsSel.appendChild(opt);
                                });
                                // restore previous selection when possible, otherwise keep blank
                                if (previous) {
                                    var match = Array.from(wpsSel.options || []).find(function (o) { return String(o.value) === String(previous); });
                                    if (match) wpsSel.value = previous;
                                }
                            } else if (wpsInp) {
                                // If input with datalist, populate datalist
                                var listId = wpsInp.getAttribute('list');
                                var dl = listId ? document.getElementById(listId) : null;
                                if (!dl) {
                                    // create a datalist if missing
                                    listId = 'dwr-wps-dl-' + (tr.dataset.id || Math.random().toString(36).slice(2,8));
                                    dl = document.createElement('datalist'); dl.id = listId; document.body.appendChild(dl);
                                    wpsInp.setAttribute('list', listId);
                                }
                                dl.innerHTML = '';
                                list.forEach(function (it) {
                                    var opt = document.createElement('option'); opt.value = String(it.wps ?? it.Wps ?? it.wps ?? ''); dl.appendChild(opt);
                                });
                            }
                        } catch (e) { console.debug('[DailyWelding] Error applying WPS candidates', e); }
                    })
                    .catch(function (err) { console.debug('[DailyWelding] Fetch WPS candidates failed', err); })
                    .finally(function () { try { delete target.dataset.wpsRefreshing; } catch (e) { } });
            }, 50);
        } catch (e) { try { delete (tr && tr.dataset && tr.dataset.wpsRefreshing); } catch { } }
    }

    function attachRowHandlers(tr) {
        if (!tr) return;
        if (tr.dataset.dwInit === '1') return;

        try {
            // Ensure RFI options are populated for this row
            try { refreshRowRfiOptions(tr); } catch (e) { console.debug('[DailyWelding] refreshRowRfiOptions failed', e); }

            // Clear button
            try {
                var clearBtn = tr.querySelector('button.clear-row');
                if (clearBtn && !clearBtn.dataset.dwBind) {
                    clearBtn.addEventListener('click', function (e) {
                        try {
                            e.preventDefault();
                            if (tr.classList.contains('locked')) return;

                            var setEmpty = function (selector) {
                                try {
                                    var els = tr.querySelectorAll(selector);
                                    els.forEach(function (el) {
                                        try {
                                            if (!el || el.disabled) return;
                                            if (el.tagName === 'SELECT') el.value = '';
                                            else el.value = '';
                                            try { el.dispatchEvent(new Event('input', { bubbles: true })); } catch { }
                                            try { el.dispatchEvent(new Event('change', { bubbles: true })); } catch { }
                                        } catch (err) { }
                                    });
                                } catch (err) { }
                            };

                            // Clear the set of fields required by Daily Welding clear button
                            setEmpty('input[data-name="POST_VISUAL_INSPECTION_QR_NO"], input[name="POST_VISUAL_INSPECTION_QR_NO"], input[data-name="FitupReport"], input[name="FitupReport"]');

                            // Clear RFI selects (support multiple selectors if present)
                            var rfis = tr.querySelectorAll('select[data-name="RfiId"], select[name="RfiId"], select.rfi-select, select.rfi-async-select, select[data-name="RFI_ID_DWR"]');
                            rfis.forEach(function (rfi) {
                                try {
                                    if (!rfi || rfi.disabled) return;
                                    rfi.value = '';
                                    try { delete rfi.dataset.lockRfi; } catch { }
                                    try { delete rfi.dataset.lockRfiValue; } catch { }
                                    try { delete rfi.dataset.lockRfiText; } catch { }
                                    try { rfi.dispatchEvent(new Event('change', { bubbles: true })); } catch { }
                                } catch (e) { }
                            });

                            setEmpty('input[data-name="DATE_WELDED"], input[name="DATE_WELDED"], input[data-name="FitupDate"], input[name="FitupDate"], input[data-name="WeldingDate"], input[name="WeldingDate"]');
                            setEmpty('input[data-name="ACTUAL_DATE_WELDED"], input[name="ACTUAL_DATE_WELDED"], input[data-name="ActualDate"], input[name="ActualDate"], input[data-name="ActualDateWelded"], input[name="ActualDateWelded"]');

                            setEmpty('input[data-name="WPS_ID_DWR"], input[name="Wps"], input[name="Wps"]');
                            setEmpty('select[data-name="WPS_ID_DWR"], select[name="Wps"], select[data-name="Wps"]');

                            // Root/Fill/Cap selectors
                            setEmpty('select[data-name="ROOT_A"], select[name="ROOT_A"], select[data-name="TackWelder"], select[name="TackWelder"]');
                            setEmpty('select[data-name="ROOT_B"], select[name="ROOT_B"], select[data-name="TackWelderB"], select[name="TackWelderB"]');
                            setEmpty('select[data-name="FILL_A"], select[name="FILL_A"], select[data-name="TackWelderFillA"], select[name="TackWelderFillA"]');
                            setEmpty('select[data-name="FILL_B"], select[name="FILL_B"], select[data-name="TackWelderFillB"], select[name="TackWelderFillB"]');
                            setEmpty('select[data-name="CAP_A"], select[name="CAP_A"], select[data-name="TackWelderCapA"], select[name="TackWelderCapA"]');
                            setEmpty('select[data-name="CAP_B"], select[name="CAP_B"], select[data-name="TackWelderCapB"], select[name="TackWelderCapB"]');

                            setEmpty('input[data-name="PREHEAT_TEMP_C"], input[name="PreheatTempC"]');
                            setEmpty('select[data-name="IP_or_T"], select[name="IP_or_T"]');
                            setEmpty('select[data-name="Open_Closed"], select[name="Open_Closed"]');
                            setEmpty('input[data-name="DWR_REMARKS"], input[name="Remarks"]');

                            markRowDirty(tr);
                            validateWeldingRow(tr);
                        } catch (err) { console.debug('[DailyWelding] Clear row failed:', err); }
                    });
                    clearBtn.dataset.dwBind = '1';
                }
            } catch (e) { console.debug('[DailyWelding] bind clear failed', e); }

            // Double-click on Welding Report input: copy header values and defaults into row
            try {
                var fr = tr.querySelector('input[data-name="POST_VISUAL_INSPECTION_QR_NO"], input[name="POST_VISUAL_INSPECTION_QR_NO"], input[data-name="FitupReport"], input[name="FitupReport"]');
                if (fr && !fr.dataset.dwBind) {
                    fr.addEventListener('dblclick', function () {
                        try {
                            var prevReport = fr.value || '';
                            var headerFit = document.getElementById('fitupDateInput');
                            var headerAct = document.getElementById('actualDateInput');
                            var rowFit = tr.querySelector('input[data-name="DATE_WELDED"], input[name="DATE_WELDED"], input[data-name="FitupDate"], input[name="FitupDate"]');
                            var rowAct = tr.querySelector('input[data-name="ACTUAL_DATE_WELDED"], input[name="ACTUAL_DATE_WELDED"], input[data-name="ActualDate"], input[name="ActualDate"], input[data-name="ActualDateWelded"], input[name="ActualDateWelded"]');
                            var prevRowFit = rowFit ? (rowFit.value || '') : '';
                            var prevRowAct = rowAct ? (rowAct.value || '') : '';

                            // Determine header report value based on row location
                            var rowLoc = tr.getAttribute('data-location-code') || '';
                            if (!rowLoc) {
                                var locCell2 = tr.querySelector('.col-location');
                                rowLoc = (locCell2 ? (locCell2.textContent || '') : '').trim().toUpperCase();
                            } else { rowLoc = rowLoc.toUpperCase(); }
                            var headerLocSel = document.getElementById('locationSelect');
                            var headerIsAll = headerLocSel && headerLocSel.value === 'All';
                            var headerSingle = document.getElementById('weldingReportHeaderInput');
                            var headerCombined = document.getElementById('weldingReportCombinedInput');
                            var reportValue = '';
                            if (headerIsAll && headerCombined) {
                                var parts = (headerCombined.value || '').split('/').map(function (p) { return (p || '').trim(); });
                                var ws = parts[0] || ''; var fw = parts[1] || ''; var th = parts[2] || '';
                                if (rowLoc === 'WS') reportValue = ws || (headerSingle && headerSingle.value ? headerSingle.value : '');
                                else if (rowLoc === 'TH') reportValue = th || fw || (headerSingle && headerSingle.value ? headerSingle.value : '');
                                else reportValue = fw || (headerSingle && headerSingle.value ? headerSingle.value : '');
                            } else if (headerSingle) {
                                reportValue = headerSingle.value || '';
                            }
                            if (fr && reportValue) fr.value = reportValue; else fr.value = prevReport;

                            // Copy weld dates from header if blank
                            var fitVal = headerFit ? (headerFit.value || '') : '';
                            var actVal = headerAct ? (headerAct.value || fitVal) : fitupIso;
                            if (rowFit && !rowFit.value && fitVal) rowFit.value = fitVal;
                            if (rowAct && !rowAct.value && actVal) rowAct.value = actVal;

                            // Copy tacker header values where matching option exists
                            var copyFromHeader = function(id, selector) {
                                try {
                                    var src = document.getElementById(id);
                                    var dest = tr.querySelector(selector);
                                    if (!src || !dest) return;
                                    var v = src.value || '';
                                    if (!v) return;
                                    var match = Array.from(dest.options || []).find(function (o) { return (o.value || '').toString() === v; })
                                        || Array.from(dest.options || []).find(function (o) { return (o.text || '').trim().toLowerCase() === v.trim().toLowerCase(); });
                                    if (match) dest.value = match.value;
                                } catch (e) { }
                            };
                            copyFromHeader('tackerRootASelect', 'select[data-name="ROOT_A"], select[name="ROOT_A"], select[data-name="TackWelder"], select[name="TackWelder"]');
                            copyFromHeader('tackerRootBSelect', 'select[data-name="ROOT_B"], select[name="ROOT_B"], select[data-name="TackWelderB"], select[name="TackWelderB"]');
                            copyFromHeader('tackerFillASelect', 'select[data-name="FILL_A"], select[name="FILL_A"], select[data-name="TackWelderFillA"], select[name="TackWelderFillA"]');
                            copyFromHeader('tackerFillBSelect', 'select[data-name="FILL_B"], select[name="FILL_B"], select[data-name="TackWelderFillB"], select[name="TackWelderFillB"]');
                            copyFromHeader('tackerCapASelect', 'select[data-name="CAP_A"], select[name="CAP_A"], select[data-name="TackWelderCapA"], select[name="TackWelderCapA"]');
                            copyFromHeader('tackerCapBSelect', 'select[data-name="CAP_B"], select[name="CAP_B"], select[data-name="TackWelderCapB"], select[name="TackWelderCapB"]');

                            // Copy RFI from header selection
                            try {
                                var headerRfi = document.getElementById('rfiSelect');
                                var rowRfiSel = tr.querySelector('select[data-name="RfiId"], select[name="RfiId"], select.rfi-select, select.rfi-async-select, select[data-name="RFI_ID_DWR"]');
                                if (headerRfi && rowRfiSel && !rowRfiSel.disabled) {
                                    var hdrId = headerRfi.value || '';
                                    var hdrText = '';
                                    try { var hdrIdx = headerRfi.selectedIndex; if (hdrIdx >= 0) hdrText = (headerRfi.options[hdrIdx]?.text || '').trim(); } catch { hdrText = ''; }
                                    var bare = extractBareRfi(hdrText);
                                    if (hdrId) {
                                        var exists = Array.from(rowRfiSel.options || []).some(function (o) { return String(o.value) === String(hdrId); });
                                        if (!exists) {
                                            var opt = document.createElement('option');
                                            opt.value = hdrId;
                                            opt.text = bare || hdrId;
                                            rowRfiSel.insertBefore(opt, rowRfiSel.firstChild);
                                        }
                                        rowRfiSel.value = hdrId;
                                        try { rowRfiSel.dispatchEvent(new Event('change', { bubbles: true })); } catch { }
                                        markRowDirty(tr);
                                        loadRowRfiOptions(tr, { id: hdrId, value: bare });
                                    } else {
                                        loadRowRfiOptions(tr);
                                    }
                                } else {
                                    loadRowRfiOptions(tr);
                                }
                            } catch (er) { console.debug('[DailyWelding] Copy RFI on dblclick failed', er); }

                            setRowWeldingDefaults(tr);

                            // Set first non-empty WPS candidate
                            try {
                                var wpsSel = tr.querySelector('select[data-name="WPS_ID_DWR"], select[name="WPS_ID_DWR"], select[data-name="Wps"], select[name="Wps"]');
                                if (wpsSel && wpsSel.options && wpsSel.options.length) {
                                    for (var i = 0; i < wpsSel.options.length; i++) {
                                        if ((wpsSel.options[i].value || '').trim() !== '') { wpsSel.selectedIndex = i; break; }
                                    }
                                }
                            } catch (e) { }

                            // Set default preheat
                            try { var pre = tr.querySelector('input[data-name="PREHEAT_TEMP_C"], input[name="PreheatTempC"]'); if (pre && !pre.value) pre.value = '10'; } catch (e) { }

                            markRowDirty(tr);
                            insertHoldRemark(tr);
                            validateWeldingRow(tr);
                        } catch (err) { console.debug('[DailyWelding] dblclick handler failed:', err); }
                    });
                    fr.dataset.dwBind = '1';
                }
            } catch (e) { console.debug('[DailyWelding] bind dblclick failed', e); }

            // Lazy-load RFI on focus
            try {
                var rfiSel = tr.querySelector('select[data-name="RfiId"], select[name="RfiId"], select.rfi-select, select.rfi-async-select, select[data-name="RFI_ID_DWR"]');
                if (rfiSel && !rfiSel.dataset.dwBind) {
                    rfiSel.addEventListener('focus', function() { setTimeout(function(){ loadRowRfiOptions(tr); }, 50); });
                    rfiSel.dataset.dwBind = '1';
                }
            } catch (e) { console.debug('[DailyWelding] bind rfi focus failed', e); }

            // Lazy-load WPS on focus
            try {
                var wpsSel = tr.querySelector('select[data-name="WPS_ID_DWR"], select[name="WPS_ID_DWR"], select[data-name="Wps"], select[name="Wps"]');
                var wpsInp = tr.querySelector('input[data-name="WPS_ID_DWR"], input[name="Wps"], input[name="Wps"]');
                var target = wpsSel || wpsInp;
                if (target && !target.dataset.dwBind) {
                    target.addEventListener('focus', function() { setTimeout(function(){ refreshRowWpsOptions(tr, true); }, 50); });
                    target.dataset.dwBind = '1';
                }
            } catch (e) { console.debug('[DailyWelding] bind wps focus failed', e); }

            // Validate initially
            try { validateWeldingRow(tr); } catch (e) { console.debug('[DailyWelding] initial validate failed', e); }

            // Actual date hold note
            try {
                var actualInput = tr.querySelector('input[data-name="ACTUAL_DATE_WELDED"], input[name="ACTUAL_DATE_WELDED"], input[data-name="ActualDate"], input[name="ActualDate"]');
                if (actualInput && !actualInput.dataset.dwHoldBind) {
                    actualInput.addEventListener('change', function(){ insertHoldRemark(tr); });
                    actualInput.addEventListener('dblclick', function(){ insertHoldRemark(tr); });
                    actualInput.dataset.dwHoldBind = '1';
                }
            } catch (e) { console.debug('[DailyWelding] bind actual date failed', e); }

            // Ensure individual control edits mark the row dirty and revalidate
            try {
                var editableControls = tr.querySelectorAll('input, select, textarea');
                editableControls.forEach(function (ctrl) {
                    try {
                        if (!ctrl || ctrl.dataset.dwBindChange === '1') return;
                        var onUserEdit = function (evt) {
                            try {
                                // Only mark dirty for user-initiated events
                                if (evt && evt.isTrusted === false) return;
                                markRowDirty(tr);
                                try { validateWeldingRow(tr); } catch (e) { }
                            } catch (e) { }
                        };
                        ctrl.addEventListener('change', onUserEdit);
                        ctrl.addEventListener('input', onUserEdit);
                        // Support Select2 change event if present
                        try {
                            if (window.jQuery && jQuery && jQuery(ctrl).data && jQuery(ctrl).data('select2')) {
                                try { jQuery(ctrl).on('change.select2', onUserEdit); } catch (e) { }
                            }
                        } catch (e) { }
                        ctrl.dataset.dwBindChange = '1';
                    } catch (e) { }
                });
            } catch (e) { console.debug('[DailyWelding] bind per-control change handlers failed', e); }

            // mark initialized
            tr.dataset.dwInit = '1';
        } catch (e) {
            console.debug('[DailyWelding] Error in attachRowHandlers:', e);
        }
    }

    function softenMissingFitupLock(tr) {
        try {
            if (!tr) return;
            if (tr.classList.contains('missing-fitup') && tr.classList.contains('locked')) {
                tr.classList.remove('locked');
                tr.classList.add('soft-locked');
            }
        } catch { }
    }

    function processExistingRows(root) {
        try {
            var scope = root || document;
            var rows = scope.querySelectorAll('#fitupTable tbody tr');
            rows.forEach(r => {
                try {
                    softenMissingFitupLock(r);
                    // Only set defaults for rows not yet initialized
                    if (r.dataset.dwInit !== '1') {
                        setRowWeldingDefaults(r);
                        attachRowHandlers(r);
                    } else {
                        // ensure validation is up-to-date
                        validateWeldingRow(r);
                    }
                } catch (e) {
                    console.debug('[DailyWelding] Error processing row:', e);
                }
            });
        } catch (e) {
            console.debug('[DailyWelding] Error in processExistingRows:', e);
        }
    }

    function setupTbodyObserver(tbody) {
        try {
            if (!tbody) return;
            if (tbody.dataset.dwObserved === '1') return;
            var mo = new MutationObserver(function (muts) {
                try {
                    muts.forEach(m => {
                        if (m.type === 'childList' && m.addedNodes && m.addedNodes.length) {
                            m.addedNodes.forEach(n => {
                                if (n.nodeType === 1 && n.matches('tr')) {
                                    setTimeout(function () {
                                        try {
                                            softenMissingFitupLock(n);
                                            setRowWeldingDefaults(n);
                                            attachRowHandlers(n);
                                        } catch (e) {
                                            console.debug('[DailyWelding] Error processing new row:', e);
                                        }
                                    }, 100); // Delay to ensure DOM is fully ready
                                }
                            });
                        }
                    });
                } catch (e) {
                    console.debug('[DailyWelding] Error in mutation observer:', e);
                }
            });
            mo.observe(tbody, { childList: true, subtree: false });
            tbody.dataset.dwObserved = '1';
        } catch (e) {
            console.debug('[DailyWelding] Error in setupTbodyObserver:', e);
        }
    }

    function observeNewRows() {
        try {
            // Try immediate wiring if table already exists
            var tbody = document.querySelector('#fitupTable tbody');
            if (tbody) {
                processExistingRows();
                setupTbodyObserver(tbody);
                // Do not return; continue to set up a root observer to handle re-loads
            }

            // If table not yet in DOM (fragment injected later), observe root container
            var root = document.getElementById('dwr-table-root');
            if (!root) return;
            // Always observe root to handle multiple loads
            var mo = new MutationObserver(function () {
                try {
                    // On any added child, look for table tbody and initialize bindings
                    var tbodies = root.querySelectorAll('#fitupTable tbody');
                    if (tbodies && tbodies.length) {
                        tbodies.forEach(function(tb){
                            processExistingRows(root);
                            setupTbodyObserver(tb);
                        });
                    }
                } catch (e) {
                    console.debug('[DailyWelding] Error observing root:', e);
                }
            });
            mo.observe(root, { childList: true, subtree: true });
        } catch (e) {
            console.debug('[DailyWelding] Error in observeNewRows:', e);
        }
    }

    function renameUiLabels() {
        try {
            // Replace Fit-up -> Welding for primary buttons
            replaceButtonTextByPattern('completeBtn', /Fit-?up/gi, 'Welding');
            replaceButtonTextByPattern('confirmBtn', /Fit-?up/gi, 'Welding');
            // Update toolbar title
            var title = document.querySelector('.tb-title');
            if (title && title.textContent) {
                title.textContent = title.textContent.replace(/Fit-?up/gi, 'Welding');
            }
            // Also update status and any other labels that may contain the word
            var elems = document.querySelectorAll('button, label, span');
            elems.forEach(el => {
                if (el && el.textContent && /Fit-?up/i.test(el.textContent)) {
                    el.textContent = el.textContent.replace(/Fit-?up/gi, 'Welding');
                }
            });
        } catch (e) {
            console.debug('[DailyWelding] Error in renameUiLabels:', e);
        }
    }

    function patchUcSetButtons() {
        try {
            if (window.uc && typeof window.uc.setButtonsText === 'function' && !window.uc.__dwPatched) {
                var orig = window.uc.setButtonsText.bind(window.uc);
                window.uc.setButtonsText = function () {
                    try { orig.apply(window.uc, arguments); } catch (e) { }
                    try { renameUiLabels(); } catch (e) { }
                };
                window.uc.__dwPatched = true;
            }
        } catch (e) {
            console.debug('[DailyWelding] Error in patchUcSetButtons:', e);
        }
    }

    // Helper to forward top-level toolbar button clicks into the injected fragment (or call global functions)
    function forwardToolbarButtons() {
        function clickCandidates(cands, topEl) {
            try {
                // search inside fragment root first
                var root = document.getElementById('dwr-table-root');
                for (var i = 0; i < cands.length; i++) {
                    var id = cands[i];
                    if (!id) continue;
                    // prefer fragment-local element
                    var el = (root && root.querySelector('#' + id));
                    if (el) { try { el.click(); return true; } catch (e) { } }
                    // if global element exists, ensure we don't forward back to the top-level button itself
                    var globalEl = document.getElementById(id);
                    if (globalEl && globalEl !== topEl) { try { globalEl.click(); return true; } catch (e) { } }
                }
                return false;
            } catch (e) {
                console.debug('[DailyWelding] Error in clickCandidates:', e);
                return false;
            }
        }

        // mapping: top-level id -> candidate fragment ids (in order) and optional global fn name
        var map = {
            'saveAllBtn': { ids: ['btnSave'], fn: 'save' },
            // Buttons with welding-specific handlers (complete/confirm) are intentionally excluded here
            'confirmAllBtn': { ids: ['btnConfirmAll', 'confirmAllBtn'], fn: 'confirmAll' },
            'unconfirmAllBtn': { ids: ['btnUnconfirmAll', 'unconfirmAllBtn'], fn: 'unconfirmAll' },
            'clearRemarksBtn': { ids: ['btnClearRemarks', 'clearRemarksBtn'], fn: 'clearRemarks' },
            'openDwgBtn': { ids: ['btnOpenDwg', 'openDgwBtn'], fn: 'openDwg' },
            'downloadDwgBtn': { ids: ['btnDownloadDwg', 'downloadDwgBtn'], fn: 'downloadDwg' }
        };

        Object.keys(map).forEach(function (topId) {
            try {
                var top = document.getElementById(topId);
                if (!top) return;
                top.addEventListener('click', function (e) {
                    try {
                        e.preventDefault();
                        var entry = map[topId];
                        var handled = clickCandidates(entry.ids || [], top);
                        if (!handled && entry.fn && typeof window[entry.fn] === 'function') {
                            try { window[entry.fn](); handled = true; } catch (e) { }
                        }
                        if (!handled) {
                            // last resort: try dispatching a click on any element with matching data-action attribute inside fragment
                            var root = document.getElementById('dwr-table-root');
                            if (root) {
                                for (var i = 0; i < (entry.ids || []).length; i++) {
                                    var el = root.querySelector('[data-action="' + entry.ids[i] + '"]');
                                    if (el) {
                                        try { el.click(); handled = true; break; } catch (e) { }
                                    }
                                }
                            }
                        }
                        if (!handled) {
                          console.debug('[DailyWelding] toolbar button ' + topId + ' had no target to forward to');
                        }
                    } catch (e) {
                        console.debug('[DailyWelding] forward toolbar click failed', e);
                    }
                });
            } catch (e) {
                console.debug('[DailyWelding] Error setting up forward for ' + topId, e);
            }
        });
    }

    // Remove any legacy inline onclick confirmations that may block toolbar actions
    function cleanToolbarHandlers() {
        try {
            ['completeBtn', 'confirmBtn'].forEach(function (id) {
                var btn = document.getElementById(id);
                if (!btn) return;
                try { btn.onclick = null; } catch (e) { }
                try { btn.removeAttribute && btn.removeAttribute('onclick'); } catch (e) { }
            });
        } catch (e) {
            console.debug('[DailyWelding] Error cleaning toolbar handlers:', e);
        }
    }

    // Toggle the Weld Type filter panel when the Hide/Show Filters button is clicked
    function attachFilterToggle() {
        try {
            var btn = document.getElementById('filterToggle');
            if (!btn) return;
            var filters = document.getElementById('weldFilters') || document.querySelector('.weld-filters');
            if (!filters) return;
            // Initialize button text from aria-expanded if present
            function updateButton(expanded) {
                try {
                    btn.setAttribute('aria-expanded', expanded ? 'true' : 'false');
                    btn.textContent = expanded ? 'Hide Filters' : 'Show Filters';
                } catch (e) {
                    console.debug('[DailyWelding] Error updating button:', e);
                }
            }
            // Start state: visible unless collapsed class present
            var isCollapsed = filters.classList.contains('collapsed') || window.getComputedStyle(filters).display === 'none';
            updateButton(!isCollapsed);
            btn.addEventListener('click', function (e) {
                try {
                    e.preventDefault();
                    filters.classList.toggle('collapsed');
                    var nowCollapsed = filters.classList.contains('collapsed');
                    updateButton(!nowCollapsed);
                } catch (err) {
                    console.debug('[DailyWelding] filter toggle failed', err);
                }
            });
        } catch (e) {
            console.debug('[DailyWelding] attachFilterToggle error:', e);
        }
    }

    function syncActualDateHeaderAndFilter() {
        try {
            var header = document.getElementById('actualDateInput');
            var filter = document.querySelector('input[name="actualDateFilter"][type="date"]');
            if (!header || !filter) return;

            function setFilterFromHeader() {
                try {
                    var v = header.value || '';
                    var datePart = (v.split('T')[0] || '').trim();
                    filter.value = datePart;
                } catch { }
            }

            function setHeaderFromFilter() {
                try {
                    var d = filter.value || '';
                    if (!d) {
                        header.value = '';
                        try { header.dispatchEvent(new Event('change', { bubbles: true })); } catch { }
                        return;
                    }
                    var timeMatch = (header.value || '').match(/T(\d{2}:\d{2}(?::\d{2})?)/);
                    var timePart = timeMatch ? timeMatch[1] : '00:00';
                    var next = d + 'T' + timePart;
                    if (header.value !== next) {
                        header.value = next;
                        try { header.dispatchEvent(new Event('change', { bubbles: true })); } catch { }
                    }
                } catch { }
            }

            try { setFilterFromHeader(); } catch { }
            header.addEventListener('change', function () { setFilterFromHeader(); });
            header.addEventListener('input', function () { setFilterFromHeader(); });
            filter.addEventListener('change', function () { setHeaderFromFilter(); try { filter.blur(); } catch { } });
            filter.addEventListener('keydown', function (e) {
                if (e.key === 'Enter' || e.key === 'Escape') {
                    setHeaderFromFilter();
                    try { filter.blur(); } catch { }
                }
            });
        } catch (e) {
            console.debug('[DailyWelding] syncActualDateHeaderAndFilter error:', e);
        }
    }

    function bindBulkToolbarActions() {
        try {
            var confirmBtn = document.getElementById('confirmAllBtn');
            var unconfirmBtn = document.getElementById('unconfirmAllBtn');
            var clearRemarksBtn = document.getElementById('clearRemarksBtn');

            if (!confirmBtn && !unconfirmBtn && !clearRemarksBtn) return;

            function status(msg, ok) {
                try {
                    if (typeof window.showStatus === 'function') {
                        window.showStatus(msg, ok);
                        if (ok && window.scheduleStatusAutoHide) window.scheduleStatusAutoHide(msg || '');
                        return;
                    }
                } catch { }
                var host = document.getElementById('statusMsg');
                if (host) {
                    host.textContent = msg || '';
                    host.style.color = ok ? '#176d8a' : '#b40000';
                    if (ok && window.scheduleStatusAutoHide) {
                        window.scheduleStatusAutoHide(msg || '');
                    }
                } else if (msg) {
                    try { (ok ? console.log : console.warn)(msg); } catch { }
                }
            }

            function visibleRows() {
                return Array.from(document.querySelectorAll('#fitupTable tbody tr')).filter(function (r) {
                    if (!r) return false;
                    try {
                        var disp = (r.style && r.style.display) ? r.style.display : '';
                        if (disp && disp.toLowerCase() === 'none') return false;
                        var cs = window.getComputedStyle ? window.getComputedStyle(r) : null;
                        if (cs && cs.display && cs.display.toLowerCase() === 'none') return false;
                    } catch { }
                    return true;
                });
            }

            function setConfirmAll(flag, options) {
                var opts = options || {};
                var rows = visibleRows();
                var total = rows.length;
                var changed = 0;

                rows.forEach(function (r) {
                    var cb = r.querySelector('input[data-name="FitupConfirmed"], input[name="FitupConfirmed"]');
                    if (cb && cb.checked !== flag) {
                        cb.checked = true;
                        try { cb.dispatchEvent(new Event('change', { bubbles: true })); } catch { }
                        markRowDirty(r);
                        changed++;
                    }
                });

                if (!opts.silent) {
                    if (total === 0) status('No rows loaded.', false);
                    else if (changed === 0) status('All displayed rows already confirmed.', true);
                    else status((flag ? 'Confirmed ' : 'Unconfirmed ') + changed + ' row(s). Remember to Save.', true);
                }
                return { changed: changed, total: total };
            }

            function clearRemarks() {
                var rows = visibleRows();
                var cleared = 0;
                rows.forEach(function (r) {
                    var inp = r.querySelector('input[data-name="DWR_REMARKS"], input[name="DWR_REMARKS"], input[data-name="DwrRemarks"], input[name="DwrRemarks"], input[data-name="Remarks"], input[name="Remarks"]');
                    if (inp && (inp.value || '').trim() !== '') {
                        inp.value = '';
                        try { inp.dispatchEvent(new Event('input', { bubbles: true })); } catch { }
                        markRowDirty(r);
                        cleared++;
                    }
                });
                status(cleared > 0 ? ('Cleared remarks in ' + cleared + ' row(s). Remember to Save.') : 'No remarks to clear.', cleared > 0);
            }

            if (confirmBtn && confirmBtn.dataset.dwBulkBind !== '1') {
                confirmBtn.addEventListener('click', function (e) {
                    try { e && e.preventDefault && e.preventDefault(); } catch { }
                    setConfirmAll(true);
                });
                confirmBtn.dataset.dwBulkBind = '1';
            }
            if (unconfirmBtn && unconfirmBtn.dataset.dwBulkBind !== '1') {
                unconfirmBtn.addEventListener('click', function (e) {
                    try { e && e.preventDefault && e.preventDefault(); } catch { }
                    setConfirmAll(false);
                });
                unconfirmBtn.dataset.dwBulkBind = '1';
            }
            if (clearRemarksBtn && clearRemarksBtn.dataset.dwBulkBind !== '1') {
                clearRemarksBtn.addEventListener('click', function (e) {
                    try { e && e.preventDefault && e.preventDefault(); } catch { }
                    clearRemarks();
                });
                clearRemarksBtn.dataset.dwBulkBind = '1';
            }
        } catch (e) {
            console.debug('[DailyWelding] bindBulkToolbarActions error:', e);
        }
    }

    function init() {
        try {
            if (window.dailyWeldingInitDone) return;
            window.dailyWeldingInitDone = true;
            ensureStatusAutoClear();
            hideAddJoint();
            renameUiLabels();
            cleanToolbarHandlers();
            observeNewRows();
            processExistingRows();
            // populate RFI lists for rows to match Daily Fit-up initial behavior
            try { if (typeof window.reloadRfiListsForAllRows === 'function') window.reloadRfiListsForAllRows(); } catch (e) { }

            // Clear and reload RFI cache when header date or project changes
            try {
                var hdr = document.getElementById('fitupDateInput');
                if (hdr && !hdr.dataset.dwRfiBind) {
                    hdr.addEventListener('change', function () { try { __rfiCache.clear(); if (typeof window.reloadRfiListsForAllRows === 'function') window.reloadRfiListsForAllRows(); } catch (e) { } });
                    hdr.addEventListener('input', function () { try { __rfiCache.clear(); if (typeof window.reloadRfiListsForAllRows === 'function') window.reloadRfiListsForAllRows(); } catch (e) { } });
                    hdr.dataset.dwRfiBind = '1';
                }
                var proj = document.getElementById('projectSelect');
                if (proj && !proj.dataset.dwRfiBind) {
                    proj.addEventListener('change', function () { try { __rfiCache.clear(); if (typeof window.reloadRfiListsForAllRows === 'function') window.reloadRfiListsForAllRows(); } catch (e) { } });
                    proj.dataset.dwRfiBind = '1';
                }
            } catch (e) { }

            ensureConfirmAllModalHelpers();
            patchUcSetButtons();
            forwardToolbarButtons();
            attachFilterToggle();
            syncActualDateHeaderAndFilter();
            bindBulkToolbarActions();
            // Do not call unknown observeHeaderDateChanges twice or at all when missing
            try {
                if (typeof observeHeaderDateChanges === 'function') {
                    observeHeaderDateChanges();
                }
            } catch (e) { }
            log('initialized');
        } catch (e) {
            console.debug('[DailyWelding] Error during initialization:', e);
        }
    }

    // Use a single initialization to prevent multiple initializations
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', function () {
            // Add a small delay to ensure all other scripts have loaded
            setTimeout(init, 100);
        });
    } else {
        setTimeout(init, 100);
    }

    // Observe status message changes to auto-hide after 3s
    function ensureStatusAutoClear(){
        try{
            var host = document.getElementById('statusMsg');
            if(!host || host.dataset.dwStatusObs==='1') return;
            var obs = new MutationObserver(function(){
                try{
                    var txt = (host.textContent || '').trim();
                    if(txt && window.scheduleStatusAutoHide){
                        window.scheduleStatusAutoHide(txt);
                    }
                }catch{}
            });
            obs.observe(host,{childList:true,subtree:true,characterData:true});
            host.dataset.dwStatusObs='1';
            var initial = (host.textContent || '').trim();
            if(initial && window.scheduleStatusAutoHide){ window.scheduleStatusAutoHide(initial); }
        }catch{}
    }
    ensureStatusAutoClear();
})();


    
// New: table fragment initializer (existing code unchanged)
(function(){
    window.dailyWelding = window.dailyWelding || {};

    function loadScript(src){ return new Promise(function(resolve,reject){ var s=document.createElement('script'); s.src=src; s.onload=resolve; s.onerror=function(e){ reject(e); }; document.head.appendChild(s); }); }
    function loadCss(href){ return new Promise(function(resolve,reject){ var l=document.createElement('link'); l.rel='stylesheet'; l.href=href; l.onload=resolve; l.onerror=function(e){ reject(e); }; document.head.appendChild(l); }); }

    async function ensureSelect2(){
        try{
            var localJq = '/lib/jquery/dist/jquery.min.js';
            var localSelect2Css = '/lib/select2/css/select2.min.css';
            var localSelect2Js = '/lib/select2/js/select2.min.js';

            if(!window.jQuery){ await loadScript(localJq); }
            if(!window.jQuery) return false;
            var $ = window.jQuery;
            if(!jQuery.fn || !jQuery.fn.select2){
                await loadCss(localSelect2Css);
                await loadScript(localSelect2Js);
            }
            return !!(window.jQuery && window.jQuery.fn && window.jQuery.fn.select2);
        }catch(e){ console.warn('[DailyWelding] ensureSelect2 failed', e); return false; }
    }

    async function initRfiSelects(root){
        try{
            var ok = await ensureSelect2();
            if(!ok) return;
            var $ = window.jQuery;
            var endpoint = '/Home/GetWeldingRfiOptions';
            var scope = root || document;
            var selects = Array.from(scope.querySelectorAll('select.rfi-async-select'));
            selects.forEach(function(sel){
                try{
                    if(sel.dataset.dwRfiInit==='1') return;
                    // Ensure at most one visible blank option for native/select2
                    try {
                        var blanks = Array.from(sel.options || []).filter(o => (o.value || '') === '');
                        blanks.slice(1).forEach(o => { try { o.remove(); } catch(_) {} });
                        if (blanks.length === 0) {
                            var blank = document.createElement('option'); blank.value = ''; blank.text = '\u00A0'; sel.insertBefore(blank, sel.firstChild);
                        }
                    } catch(e) { /* ignore */ }

                    var $sel = $(sel);
                    var $tr = $sel.closest('tr');
                    var loc = ($tr && $tr.attr && $tr.attr('data-location-code')) || '';
                    var projectId = $sel.data('project') || '';
                    var fitupDate = $sel.data('fitupdate') || '';
                    $sel.select2({
                        // Visible placeholder so Select2 shows a clear empty option
                        placeholder: $sel.attr('data-placeholder') || '\u00A0',
                         allowClear: true,
                         width: 'resolve',
                         ajax: {
                            url: endpoint,
                            dataType: 'json',
                            delay: 250,
                            data: function(params){
                                return {
                                    location: loc || '',
                                    fitupDateIso: fitupDate || '',
                                    projectId: projectId || '',
                                    q: params.term || ''
                                };
                            },
                            processResults: function(data){
                                const items = (Array.isArray(data) ? data : []).map(function(i){ return { id: i.Id, text: i.Display || (i.Value || i.Id) }; });
                                return { results: items };
                            }
                        },
                        minimumInputLength: 0,
                        templateResult: function(item){ return item && item.text ? item.text : item; },
                        templateSelection: function(item){ return item && item.text ? item.text : item; }
                    });
                    // ensure existing selected option remains visible in Select2
                    var cur = $sel.val();
                    if(cur){
                        if($sel.find('option[value="' + cur + '"]').length === 0){
                            var txt = $sel.attr('data-current-display') || ('RFI ' + cur);
                            var opt = new Option(txt, cur, true, true);
                            $sel.append(opt).trigger('change');
                        }
                    }
                    sel.dataset.dwRfiInit='1';
                }catch(e){ console.warn('[DailyWelding] initRfiSelects per-select failed', e); }
            });
        }catch(e){ console.warn('[DailyWelding] initRfiSelects failed', e); }
    }

    window.dailyWelding.initTableFragment = async function(root){
        try{
            var scope = root || document;
            await initRfiSelects(scope);
        }catch(e){ console.debug('[DailyWelding] initTableFragment error', e); }
    };

})();

(function () {
    const config = window.dailyWeldingUcConfig || {};
    const endpoints = config.endpoints || {};
    if (!endpoints.exists || !endpoints.complete || !endpoints.confirm) {
        return;
    }

    const defaultState = { updatedExists: false, confirmedExists: false };
    let toolbarStatusTimer = null;
    const uc = window.uc = {
        token: '',
        endpoints,
        headerView: config.headerView || '',
        getProjectId() { return document.getElementById('projectSelect')?.value; },
        getLocation() { return document.getElementById('locationSelect')?.value; },
        getFitupIso() {
            const actual = document.getElementById('actualDateInput')?.value;
            if (actual && actual.trim()) return actual;
            return document.getElementById('fitupDateInput')?.value;
        },
        mapHeaderLocation(raw) {
            const s = (raw || '').trim().toUpperCase();
            if (s === 'ALL') return 'ALL';
            if (s.startsWith('WS') || s.includes('SHOP') || s.includes('WORK')) return 'WS';
            if (s.startsWith('FW') || s.includes('FIELD')) return 'FW';
            if (s.startsWith('TH') || s.includes('THREAD')) return 'TH';
            return s;
        },
        getAllLocationCodes() { return ['WS', 'FW', 'TH']; }
    };

    function updateToolbarStatus(message, ok) {
        try {
            const host = document.getElementById('statusMsg');
            if (!host) return false;
            host.textContent = message || '';
            host.style.color = ok ? '#176d8a' : '#b40000';
            if (toolbarStatusTimer) {
                clearTimeout(toolbarStatusTimer);
                toolbarStatusTimer = null;
            }
            if (message && ok) {
                toolbarStatusTimer = setTimeout(function () {
                    if (host.textContent === message) {
                        host.textContent = '';
                    }
                }, 3000);
            }
            return true;
        } catch {
            return false;
        }
    }

    function safeStatus(message, ok) {
        if (!message) return;
        try {
            if (typeof window.showStatus === 'function') {
                const handled = window.showStatus(message, ok);
                if (ok && window.scheduleStatusAutoHide) window.scheduleStatusAutoHide(message);
                if (handled !== false) {
                    return;
                }
            }
        } catch { /* ignore */ }
        if (updateToolbarStatus(message, ok)) {
            return;
        }
        if (ok) {
            console.log(message);
        } else {
            console.warn(message);
        }
    }

    const normalizeState = (raw) => ({
        updatedExists: !!(raw && raw.updatedExists),
        confirmedExists: !!(raw && raw.confirmedExists)
    });

    function buildTooltip(breakdown, key) {
        if (!breakdown) return '';
        const parts = [];
        Object.keys(breakdown).forEach(code => {
            const section = breakdown[code] || defaultState;
            parts.push(`${code}: ${section[key] ? 'Yes' : 'No'}`);
        });
        return parts.join(' | ');
    }

    function currentHeaderView() {
        const hv = document.getElementById('headerViewSelect')?.value || uc.headerView || '';
        return hv.trim().toLowerCase();
    }

    uc.checkExists = async function () {
        const endpoint = this.endpoints.exists;
        const projectId = this.getProjectId();
        const fitupIso = this.getFitupIso();
        const headerLoc = (this.getLocation() || '').trim();
        const headerLocLower = headerLoc.toLowerCase();
        const mappedHeaderLoc = this.mapHeaderLocation(headerLoc);

        if (!endpoint || !projectId || !fitupIso) {
            if (headerLocLower === 'all') {
                const breakdown = {};
                this.getAllLocationCodes().forEach(code => { breakdown[code] = { ...defaultState }; });
                return { ...defaultState, breakdown };
            }
            return { ...defaultState };
        }

        const fetchState = async (locOverride) => {
            const locationParam = locOverride ?? mappedHeaderLoc;
            if (!locationParam) return { ...defaultState };
            const params = new URLSearchParams({
                projectId: String(projectId),
                location: String(locationParam),
                fitupDateIso: String(fitupIso),
                actualDateIso: String(fitupIso)
             });
            try {
                const res = await fetch(`${endpoint}?${params.toString()}`, { headers: { 'Accept': 'application/json' } });
                if (!res.ok) return { ...defaultState };
                const data = await res.json();
                return normalizeState(data);
            } catch {
                return { ...defaultState };
            }
        };

        if (headerLocLower === 'all') {
            const codes = this.getAllLocationCodes();
            const breakdown = {};
            const aggregate = { ...defaultState, breakdown };
            await Promise.all(codes.map(async code => {
                const state = await fetchState(code);
                breakdown[code] = state;
                if (state.updatedExists) aggregate.updatedExists = true;
                if (state.confirmedExists) aggregate.confirmedExists = true;
            }));
            return aggregate;
        }

        return fetchState();
    };

    uc.setButtonsText = function (state) {
        const normalized = normalizeState(state);
        const location = (this.getLocation() || '').trim().toLowerCase();
        const applyTooltip = (btn, key) => {
            if (!btn) return;
            if (location === 'all' && state && state.breakdown) {
                btn.title = buildTooltip(state.breakdown, key);
            } else {
                btn.title = '';
            }
        };

        const completeBtn = document.getElementById('completeBtn');
        if (completeBtn) {
            completeBtn.textContent = normalized.updatedExists ? 'Update Daily Welding' : 'Daily Welding Completed';
            applyTooltip(completeBtn, 'updatedExists');
        }

        const confirmBtn = document.getElementById('confirmBtn');
        if (confirmBtn) {
            const isDateView = currentHeaderView() === 'date';
            confirmBtn.style.display = isDateView ? 'inline-block' : 'none';
            confirmBtn.textContent = normalized.confirmedExists ? 'Update Confirmed Welding' : 'Daily Welding Confirmed';
            applyTooltip(confirmBtn, 'confirmedExists');
        }
    };

    uc.post = async function (endpoint, actionType) {
        const projectId = this.getProjectId();
        const fitupIso = this.getFitupIso();
        const locRaw = this.getLocation() || '';
        const token = this.token || document.querySelector('input[name="__RequestVerificationToken"]')?.value || '';
        if (!endpoint) { safeStatus('Endpoint unavailable. Reload the page.', false); return; }
        if (!projectId) { safeStatus('Select Project first.', false); return; }
        if (!fitupIso) { safeStatus('Select Actual Date first.', false); return; }
        if (!token) { safeStatus('Missing security token. Reload the page.', false); return; }

        const locCode = this.mapHeaderLocation(locRaw);
        const isAll = locCode === 'ALL';
        const targetLocs = locCode ? (isAll ? ['ALL'] : [locCode]) : [];
        if (targetLocs.length === 0) {
            safeStatus('Select Location first.', false);
            return;
        }

        const targetBtn = actionType === 'confirmed' ? document.getElementById('confirmBtn') : document.getElementById('completeBtn');
        const originalText = targetBtn?.textContent || '';

        try {
            if (targetBtn) {
                targetBtn.disabled = true;
                targetBtn.dataset.prevText = originalText;
                targetBtn.textContent = 'Working...';
            }

            let okCount = 0;
            let failCount = 0;
            for (const loc of targetLocs) {
                const params = new URLSearchParams({
                    projectId: String(projectId),
                    location: String(loc),
                    fitupDateIso: String(fitupIso),
                    actualDateIso: String(fitupIso),
                    actionType: actionType || 'completed',
                    __RequestVerificationToken: token
                });
                try {
                    const res = await fetch(endpoint, {
                        method: 'POST',
                        credentials: 'same-origin',
                        headers: {
                            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                            'RequestVerificationToken': token
                        },
                        body: params.toString()
                    });
                    if (res.ok) okCount++; else failCount++;
                } catch {
                    failCount++;
                }
            }

            if (okCount > 0) {
                const scopeLabel = isAll ? 'All locations' : `${okCount} location${okCount > 1 ? 's' : ''}`;
                safeStatus(`Done (${scopeLabel})`, true);
                try {
                    const st = await this.checkExists();
                    this.setButtonsText(st);
                } catch { /* ignore */ }
                const kind = actionType === 'confirmed' ? 'confirmed' : 'completed';
                const label = originalText || (kind === 'confirmed' ? 'Daily Welding Confirmed' : 'Daily Welding Completed');
                await this.sendMail(kind, label);
            } else if (failCount > 0) {
                safeStatus('Operation failed.', false);
            }
        } catch (err) {
            console.error('DailyWelding uc.post failed', err);
            safeStatus('Request failed.', false);
        } finally {
            if (targetBtn) {
                targetBtn.disabled = false;
                if (targetBtn.dataset.prevText) {
                    targetBtn.textContent = targetBtn.dataset.prevText;
                    delete targetBtn.dataset.prevText;
                }
            }
        }
    };

    uc.sendMail = async function (kind, label) {
        try {
            const endpoint = this.endpoints.mail;
            const token = this.token || document.querySelector('input[name="__RequestVerificationToken"]')?.value || '';
            if (!endpoint || !token) return;
            const projectId = this.getProjectId();
            const fitupIso = this.getFitupIso();
            const locRaw = this.getLocation() || '';
            if (!projectId || !fitupIso) return;

            const postMail = async (locCode) => {
                if (!locCode) return;
                const form = new URLSearchParams({
                    projectId: String(projectId),
                    location: String(locCode),
                    fitupDateIso: String(fitupIso),
                    actualDateIso: String(fitupIso),
                    actionLabel: String(label || ''),
                    kind: String(kind || '')
                });
                try {
                    const res = await fetch(endpoint, {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                            'RequestVerificationToken': token
                        },
                        body: form.toString()
                    });
                    if (!res.ok) {
                        console.warn('DailyWelding email send failed', locCode, await res.text().catch(() => ''));
                    }
                } catch (err) {
                    console.warn('DailyWelding email send error', locCode, err);
                }
            };

            if ((locRaw || '').trim().toUpperCase() === 'ALL') {
                await postMail('ALL');
                return;
            }

            const mapped = this.mapHeaderLocation(locRaw);
            if (mapped) await postMail(mapped);
        } catch (err) {
            console.warn('DailyWelding email send error', err);
        }
    };

    let refreshTimer = null;
    function refreshButtons() {
        uc.checkExists().then(state => uc.setButtonsText(state)).catch(() => { /* ignore */ });
    }
    function scheduleRefresh() {
        if (refreshTimer) clearTimeout(refreshTimer);
        refreshTimer = setTimeout(refreshButtons, 250);
    }

    async function handleConfirmClick() {
        try {
            const rows = Array.from(document.querySelectorAll('#fitupTable tbody tr'));
            if (!rows || rows.length === 0) {
                safeStatus('No rows loaded to confirm.', false);
                return;
            }

            const unconfirmedRows = rows.filter(function (r) {
                const cb = r.querySelector('input[data-name="FitupConfirmed"], input[name="FitupConfirmed"]');
                return !(cb && cb.checked === true);
            });

            if (unconfirmedRows.length > 0) {
                const count = unconfirmedRows.length;
                const msg = 'There ' + (count === 1 ? 'is' : 'are') + ' ' + count + ' joint' + (count === 1 ? '' : 's') + ' that ' + (count === 1 ? 'is' : 'are') + ' not marked as Confirmed.\n\nChoose "Confirm All" to mark every joint as Confirmed before using \'Daily Welding Confirmed\', or select "Cancel" to review the table.';
                let proceed = false;
                if (typeof window.openConfirmAllModal === 'function') {
                    proceed = await window.openConfirmAllModal(msg);
                } else if (typeof window.confirm === 'function') {
                    proceed = window.confirm(msg);
                }
                if (!proceed) {
                    safeStatus('Confirmation aborted. Confirm every joint first.', false);
                    return;
                }

                // Populate missing welding/actual dates from header before bulk confirming so rows can be saved
                try { if (typeof window.applyHeaderDatesToRows === 'function') { window.applyHeaderDatesToRows(unconfirmedRows); } } catch { }

                const changed = typeof window.markRowsConfirmed === 'function'
                    ? window.markRowsConfirmed(unconfirmedRows)
                    : (function () {
                        let cnt = 0;
                        unconfirmedRows.forEach(function (r) {
                            const cb = r.querySelector('input[data-name="FitupConfirmed"], input[name="FitupConfirmed"]');
                            if (cb && cb.checked !== true) {
                                cb.checked = true;
                                try { cb.dispatchEvent(new Event('change', { bubbles: true })); } catch { }
                                try { markRowDirty(r); } catch { }
                                cnt++;
                            }
                        });
                        return cnt;
                    })();

                if (changed > 0) {
                    safeStatus('Confirmed ' + changed + ' joint(s). Remember to Save.', true);
                    try { if (typeof window.save === 'function') { await window.save(); } } catch { }
                }
            }

            await uc.post(uc.endpoints.confirm, 'confirmed');
        } catch (err) {
            console.error('Confirm pre-check error', err);
            safeStatus('Error during confirm validation', false);
        }
    }

    document.addEventListener('DOMContentLoaded', function () {
        uc.token = document.querySelector('input[name="__RequestVerificationToken"]')?.value || uc.token || '';
        refreshButtons();

        const fitDate = document.getElementById('fitupDateInput');
        fitDate?.addEventListener('change', scheduleRefresh);
        fitDate?.addEventListener('input', scheduleRefresh);

        const projectSel = document.getElementById('projectSelect');
        projectSel?.addEventListener('change', scheduleRefresh);

        const locSel = document.getElementById('locationSelect');
        locSel?.addEventListener('change', scheduleRefresh);

        const headerViewSel = document.getElementById('headerViewSelect');
        headerViewSel?.addEventListener('change', function () {
            uc.headerView = headerViewSel.value || uc.headerView;
            scheduleRefresh();
        });

        document.getElementById('completeBtn')?.addEventListener('click', function (e) {
            try {
                e?.preventDefault?.();
                e?.stopPropagation?.();
                e?.stopImmediatePropagation?.();
            } catch { }
            uc.post(uc.endpoints.complete, 'completed');
        });
        document.getElementById('confirmBtn')?.addEventListener('click', function (e) {
            try {
                e?.preventDefault?.();
                e?.stopPropagation?.();
                e?.stopImmediatePropagation?.();
            } catch { }
            handleConfirmClick();
        });
    });
})();