(function () {
    function getAntiForgeryToken() {
        try {
            const el = document.querySelector('input[name="__RequestVerificationToken"]');
            return el ? el.value : '';
        } catch {
            return '';
        }
    }

    function init() {
        const cfg = window.__dfrFormConfig || {};
        const token = getAntiForgeryToken();
        const locationOptions = Array.isArray(cfg.locationOptions) ? cfg.locationOptions.slice() : [];
        const jAddOptions = Array.isArray(cfg.jAddOptions) ? cfg.jAddOptions.slice() : [];
        const weldTypeOptions = Array.isArray(cfg.weldTypeOptions) ? cfg.weldTypeOptions.slice() : [];
        const materialDescriptions = Array.isArray(cfg.materialDescriptions) ? cfg.materialDescriptions.slice() : [];
        const materialGrades = Array.isArray(cfg.materialGrades) ? cfg.materialGrades.slice() : [];
        const fitupDateHeader = typeof cfg.fitupDateHeader === 'string' ? cfg.fitupDateHeader : '';
        const bulkUpdateUrl = cfg.bulkUpdateUrl || '';
        const addRowUrl = cfg.addRowUrl || '';

        function status(msg, ok) {
            const e = document.getElementById('dfrMsg');
            if (!e) return;
            e.textContent = msg;
            e.style.color = ok ? '#176d8a' : '#b40000';
            if (ok) {
                setTimeout(() => {
                    if (e.textContent === msg) e.textContent = '';
                }, 2500);
            }
        }

        function markDirty(tr) {
            if (!tr || tr.dataset.dirty === '1') return;
            tr.dataset.dirty = '1';
            tr.classList.add('dirty');
        }

        function syncDeletedCancelled(tr) {
            if (!tr) return;
            const fitupInput = tr.querySelector('input[data-name="FitupDate"]');
            const deletedCb = tr.querySelector('input[data-name="Deleted"]');
            const cancelledCb = tr.querySelector('input[data-name="Cancelled"]');
            const lockedRow = tr.classList.contains('locked');
            if (!fitupInput || !deletedCb || !cancelledCb) return;
            const hasFitup = (fitupInput.value || '').trim().length > 0;
            // Deleted: only when FITUP_DATE is not null
            deletedCb.disabled = lockedRow || !hasFitup;
            deletedCb.title = !hasFitup ? 'Requires Fit-up Date' : '';
            if (!hasFitup && deletedCb.checked) { deletedCb.checked = false; }
            // Cancelled: only when FITUP_DATE is null
            cancelledCb.disabled = lockedRow || hasFitup;
            cancelledCb.title = hasFitup ? 'Not allowed when Fit-up Date exists' : '';
            if (hasFitup && cancelledCb.checked) { cancelledCb.checked = false; }
        }

        function initRow(tr) {
            if (!tr) return;
            tr.querySelectorAll('input[data-name],select[data-name]').forEach(inp => {
                if (inp.dataset.wire) return;
                const handler = () => {
                    markDirty(tr);
                    if (inp.getAttribute('data-name') === 'FitupDate') {
                        syncDeletedCancelled(tr);
                    }
                };
                if (inp.type === 'checkbox') {
                    inp.addEventListener('change', handler);
                } else {
                    inp.addEventListener('input', handler);
                    if (inp.tagName === 'SELECT') inp.addEventListener('change', handler);
                }
                inp.dataset.wire = '1';
            });
            syncDeletedCancelled(tr);
        }

        document.querySelectorAll('#dfrTable tbody tr').forEach(initRow);

        function dto(tr) {
            const o = { JointId: parseInt(tr?.dataset?.id || '0', 10) };
            tr.querySelectorAll('input[data-name],select[data-name]').forEach(i => {
                const n = i.getAttribute('data-name');
                let v = i.type === 'checkbox' ? i.checked : i.value;
                if (i.type !== 'checkbox' && v != null) {
                    v = v.trim();
                    if (v === '') v = null;
                }
                o[n] = v;
            });
            ['DiaIn'].forEach(n => {
                if (o[n] != null) {
                    const f = parseFloat(o[n]);
                    o[n] = !isNaN(f) ? f : null;
                }
            });
            return o;
        }

        async function save() {
            const dirty = [...document.querySelectorAll('#dfrTable tbody tr')].filter(r => r.dataset.dirty === '1');
            if (dirty.length === 0) {
                status('No changes', true);
                return;
            }
            if (!bulkUpdateUrl) {
                status('Save endpoint missing', false);
                return;
            }
            const payload = dirty.map(dto);
            try {
                const res = await fetch(bulkUpdateUrl, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'RequestVerificationToken': token
                    },
                    body: JSON.stringify(payload)
                });
                if (res.ok) {
                    let j;
                    try { j = await res.json(); } catch { j = { updated: dirty.length }; }
                    dirty.forEach(r => { delete r.dataset.dirty; r.classList.remove('dirty'); });
                    const updated = j && j.updated != null ? j.updated : dirty.length;
                    status(`Saved ${updated} row(s).`, true);
                } else {
                    status('Save failed', false);
                }
            } catch {
                status('Save failed', false);
            }
        }

        const saveBtn = document.getElementById('btnSave');
        if (saveBtn) saveBtn.addEventListener('click', save);

        function escapeHtml(str) {
            return (str == null ? '' : str).replace(/[&<>"']/g, ch => {
                switch (ch) {
                    case '&': return '&amp;';
                    case '<': return '&lt;';
                    case '>': return '&gt;';
                    case '"': return '&quot;';
                    case '\'': return '&#39;';
                    default: return ch;
                }
            });
        }

        function buildOptions(arr, selected) {
            if (!Array.isArray(arr) || arr.length === 0) {
                return selected ? `<option value="${escapeHtml(selected)}" selected>${escapeHtml(selected)}</option>` : '';
            }
            return arr.map(v => `<option value="${escapeHtml(v)}"${v === selected ? ' selected' : ''}>${escapeHtml(v)}</option>`).join('');
        }

        const addBtn = document.getElementById('btnAdd');
        if (addBtn) {
            addBtn.addEventListener('click', async () => {
                const form = document.getElementById('dfrHeaderForm');
                if (!form) return;
                if (!addRowUrl) {
                    status('Add endpoint missing', false);
                    return;
                }
                const projectId = form.projectId.value;
                const location = form.location.value;
                const layout = form.layout.value;
                const sheet = form.sheet.value;
                if (location === 'All') {
                    status('Please select Shop or Field for Location before adding a row.', false);
                    return;
                }
                if (!layout || !sheet) {
                    status('Layout & Sheet required', false);
                    return;
                }
                try {
                    const res = await fetch(addRowUrl, {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/x-www-form-urlencoded',
                            'RequestVerificationToken': token
                        },
                        body: new URLSearchParams({ projectId, location, layout, sheet })
                    });
                    let data = null;
                    try { data = await res.json(); } catch { }
                    if (!res.ok || !data || data.success === false) {
                        const msg = data && data.message ? data.message : 'Add failed';
                        status(msg, false);
                        return;
                    }
                    let tbody = document.querySelector('#dfrTable tbody');
                    if (!tbody) {
                        location.reload();
                        return;
                    }
                    const sn = tbody.querySelectorAll('tr').length + 1;
                    const tr = document.createElement('tr');
                    tr.dataset.id = data.id;
                    const locOpts = buildOptions(locationOptions && locationOptions.length ? locationOptions : [location], location === 'Shop' ? 'WS' : 'FW');
                    const jAddOpts = buildOptions(jAddOptions && jAddOptions.length ? jAddOptions : ['New'], 'New');
                    const weldTypeOpts = buildOptions(weldTypeOptions && weldTypeOptions.length ? weldTypeOptions : [], '');
                    tr.innerHTML = `<td>${sn}</td>
        <td><select data-name='Location' name='Location' aria-labelledby='colLocation'>${locOpts}</select></td>
        <td><input data-name='WeldNumber' name='WeldNumber' aria-labelledby='colJointNo' value='${escapeHtml(data.weldNumber || '')}' /></td>
        <td><select data-name='JAdd' name='JAdd' aria-labelledby='colJAdd'>${jAddOpts}</select></td>
        <td>${weldTypeOpts ? `<select data-name='WeldType' name='WeldType' aria-labelledby='colWeldType'>${weldTypeOpts}<option value='' selected></option></select>` : `<input data-name='WeldType' name='WeldType' aria-labelledby='colWeldType' />`}</td>
        <td><input data-name='Rev' name='Rev' aria-labelledby='colRev' /></td>
        <td><input data-name='DiaIn' name='DiaIn' aria-labelledby='colDiaIn' /></td>
        <td><input data-name='Sch' name='Sch' aria-labelledby='colSch' /></td>
        <td><input data-name='MaterialA' name='MaterialA' aria-labelledby='colMatA' list='materialDescList' /></td>
        <td><input data-name='MaterialB' name='MaterialB' aria-labelledby='colMatB' list='materialDescList' /></td>
        <td><input data-name='GradeA' name='GradeA' aria-labelledby='colGradeA' list='materialGradeList' /></td>
        <td><input data-name='GradeB' name='GradeB' aria-labelledby='colGradeB' list='materialGradeList' /></td>
        <td><input data-name='Wps' name='Wps' aria-labelledby='colWps' /></td>
        <td><input data-name='FitupReport' name='FitupReport' aria-labelledby='colFitupReport' value='${escapeHtml(data.weldNumber || '')}' /></td>
        <td><input data-name='FitupDate' name='FitupDate' aria-labelledby='colFitupDate' value='${escapeHtml(fitupDateHeader)}' /></td>
        <td><input data-name='TackWelder' name='TackWelder' aria-labelledby='colTacker' /></td>
        <td class='chk-cell'><input type='checkbox' data-name='Deleted' name='Deleted' aria-labelledby='colDeleted' /></td>
        <td class='chk-cell'><input type='checkbox' data-name='Cancelled' name='Cancelled' aria-labelledby='colCancelled' /></td>
        <td class='chk-cell'><input type='checkbox' data-name='FitupConfirmed' name='FitupConfirmed' aria-labelledby='colConfirmed' /></td>
        <td><input data-name='Remarks' name='Remarks' aria-labelledby='colRemarks' /></td>
        <td></td>
        <td></td>`;
                    tbody.appendChild(tr);
                    initRow(tr);
                    markDirty(tr);
                    status('Row added', true);
                } catch {
                    status('Add failed', false);
                }
            });
        }

        (function () {
            const form = document.getElementById('dfrHeaderForm');
            if (!form) return;
            const layout = document.getElementById('layoutInput');
            const sheet = document.getElementById('sheetInput');
            const proj = form.querySelector('select[name="projectId"]');
            const locSel = document.getElementById('dfrLocationSelect');
            const wrap = document.getElementById('dfrTableWrap');
            if (!wrap) return;
            const initial = {
                projectId: proj ? proj.value : '',
                location: locSel ? locSel.value : '',
                layout: layout ? layout.value : '',
                sheet: sheet ? sheet.value : ''
            };
            function same(a, b) {
                const a1 = (a == null ? '' : String(a)).trim();
                const b1 = (b == null ? '' : String(b)).trim();
                return a1 === b1;
            }
            function show() { wrap.style.display = ''; }
            function hide() { wrap.style.display = 'none'; }
            function updateVisibility() {
                const isInitial = same(proj ? proj.value : null, initial.projectId)
                    && same(locSel ? locSel.value : null, initial.location)
                    && same(layout ? layout.value : null, initial.layout)
                    && same(sheet ? sheet.value : null, initial.sheet);
                if (isInitial) show(); else hide();
            }
            if (layout) layout.addEventListener('input', updateVisibility);
            if (sheet) sheet.addEventListener('input', updateVisibility);
            if (proj) proj.addEventListener('change', updateVisibility);
            if (locSel) locSel.addEventListener('change', updateVisibility);
            updateVisibility();
        })();
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
