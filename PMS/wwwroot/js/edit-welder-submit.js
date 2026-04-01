document.addEventListener('DOMContentLoaded', () => {
    const shared = window.editWelderShared;
    if (!shared) return;

    const {
        isNew,
        routes,
        state,
        set,
        setLoading,
        renderExistingQuals,
        loadQualOptions,
        applyDefaultWqt,
        suggestWqtAgency,
        suggestTestDate,
        wqtSel,
        jccEl
    } = shared;

    // Intercept Save for Add-Welder flow and perform AJAX submit to refresh qualifications only
    const welderFormEl = document.getElementById('welderForm');
    let ajaxSubmitEnabled = true;

    const fallbackToStandardPost = (form) => {
        ajaxSubmitEnabled = false;
        welderFormEl?.removeEventListener('submit', handleAjaxSubmit);
        form.submit();
    };

    const handleAjaxSubmit = async function (ev) {
        if (!isNew || !ajaxSubmitEnabled) return;
        ev.preventDefault();
        setLoading(true);
        const form = this;
        const url = form.action || routes.saveWelder;
        const fd = new FormData(form);
        try {
            const res = await fetch(url, {
                method: 'POST',
                body: fd,
                headers: {
                    'X-Requested-With': 'XMLHttpRequest',
                    'Accept': 'application/json'
                }
            });
            const contentType = res.headers.get('content-type') || '';
            if (!res.ok || !contentType.includes('application/json')) {
                fallbackToStandardPost(form);
                return;
            }
            const data = await res.json().catch(() => null);
            if (!data) {
                fallbackToStandardPost(form);
                return;
            }

            const toInt = (v) => {
                const n = parseInt(v, 10);
                return Number.isFinite(n) ? n : 0;
            };

            // Set flash message
            const setTopFlash = (msg) => {
                try {
                    if (!msg) return;
                    const container = document.querySelector('.container');
                    if (!container) return;
                    let el = document.getElementById('topFlashMsg');
                    // If server-rendered message exists, reuse it
                    if (!el) {
                        const serverMsg = container.querySelector('.msg.text-danger');
                        if (serverMsg) {
                            el = serverMsg;
                            el.id = 'topFlashMsg';
                        }
                    }
                    if (!el) {
                        el = document.createElement('div');
                        el.id = 'topFlashMsg';
                        el.className = 'msg';
                        el.setAttribute('role','status');
                        el.setAttribute('aria-live','polite');
                        const title = container.querySelector('.page-title');
                        if (title && title.parentNode) title.parentNode.insertBefore(el, title.nextSibling);
                        else container.insertBefore(el, container.firstChild);
                    }
                    // determine style: success vs error
                    const isSuccess = /successfully|saved|uploaded|updated|added/i.test(msg);
                    // clear other style classes
                    el.className = 'msg';
                    if (isSuccess) el.classList.add('msg-success');
                    else el.classList.add('text-danger');
                    el.textContent = msg;
                    el.style.display = '';
                    // make focusable so screen readers announce and keyboard users can reach it
                    el.tabIndex = -1;
                    el.focus();
                    // scroll to top of page so message is visible at the very top
                    try { window.scrollTo({ top: 0, behavior: 'smooth' }); } catch (e) { }
                } catch (e) { /* ignore */ }
            };

            setTopFlash(data?.msg || '');

            // Update existing qualifications table
            if (data?.qualifications) {
                renderExistingQuals(data.qualifications);
            }

            // Update Qualification Information fields without touching Welder Information
            const q = data?.qualification;
            if (q) {
                set('Qualification.JCC_No', q.JCC_No ?? '');
                set('Qualification.Test_Date', q.Test_Date ?? '');
                set('Qualification.Welding_Process', q.Welding_Process ?? '');
                set('Qualification.Material_P_No', q.Material_P_No ?? '');
                set('Qualification.Diameter_Range', q.Diameter_Range ?? '');
                set('Qualification.Max_Thickness', q.Max_Thickness ?? '');
                set('Qualification.Qualification_Cert_Ref_No', q.Qualification_Cert_Ref_No ?? '');
                set('Qualification.WQT_Agency', q.WQT_Agency ?? '');
                set('Qualification.Batch_No', q.Batch_No ?? '');
                set('Qualification.Date_Issued', q.Date_Issued ?? '');
                set('Qualification.Remarks', q.Remarks ?? '');
                set('Qualification.DATE_OF_LAST_CONTINUITY', q.DATE_OF_LAST_CONTINUITY ?? '');
                set('Qualification.RECORDING_THE_CONTINUITY_RECORD', q.RECORDING_THE_CONTINUITY_RECORD ?? '');

                const selHidden = document.getElementById('SelectedJcc');
                if (selHidden) selHidden.value = q.JCC_No ?? '';

                // Reset AddNewQualification flag after successful save to prevent
                // stale flag from routing subsequent saves through the add-new path
                const addNewFlagReset = document.getElementById('AddNewQualification');
                if (addNewFlagReset) addNewFlagReset.value = 'false';

                // Reapply sensible defaults after save for Add flow so qualification inputs return to starting defaults
                if (isNew) {
                    // Ensure Status stays as default Waiting for Approval
                    const statusEl = document.getElementById('Welder_Status');
                    if (statusEl && !statusEl.value) {
                        const opt = Array.from(statusEl.options).find(o => (o.value||'').toLowerCase() === 'waiting for approval'.toLowerCase());
                        if (opt) statusEl.value = opt.value;
                    }

                    // Sync Batch No defaults (value + data-max-batch) from server response
                    const batchEl = document.getElementById('Batch_No');
                    const batchFromResponse = toInt(q.Batch_No ?? q.batchNo ?? q.batch_No);
                    const maxBatchFromResponse = toInt(data?.maxBatchNo ?? data?.MaxBatchNo);
                    if (batchEl) {
                        const currentMax = toInt(batchEl.dataset.maxBatch);
                        const newMax = Math.max(currentMax, maxBatchFromResponse, batchFromResponse);
                        if (newMax > 0) batchEl.dataset.maxBatch = String(newMax);

                        const currentVal = toInt(batchEl.value);
                        const targetBatch = batchFromResponse || maxBatchFromResponse || currentVal;
                        if (targetBatch && targetBatch !== currentVal) batchEl.value = targetBatch;
                    }

                    // Refresh WQT Agency using returned value and latest defaults
                    const ensureWqtOption = (sel, val) => {
                        if (!sel || !val) return;
                        const existing = Array.from(sel.options).find(o => (o.value || '').toLowerCase() === String(val).toLowerCase());
                        if (existing) return;
                        const addNewOpt = Array.from(sel.options).find(o => o.value === '__new__');
                        const opt = document.createElement('option');
                        opt.value = val;
                        opt.textContent = val;
                        if (addNewOpt && addNewOpt.parentNode === sel) sel.insertBefore(opt, addNewOpt);
                        else sel.appendChild(opt);
                    };

                    const wqtValue = q.WQT_Agency ?? q.wqt_Agency ?? q.wqtAgency ?? '';
                    if (wqtSel) {
                        if (wqtValue) {
                            ensureWqtOption(wqtSel, wqtValue);
                            const match = Array.from(wqtSel.options).find(o => (o.value || '').toLowerCase() === String(wqtValue).toLowerCase());
                            if (match) wqtSel.value = match.value;
                            wqtSel.dataset.userChoice = wqtSel.value || '';
                            state.userChangedWqt = false;
                        }
                        if (!wqtSel.value) await suggestWqtAgency(true);
                    }

                    // Reapply default WQT agency (batch-specific or page default) and refresh option lists
                    await applyDefaultWqt();
                    loadQualOptions();
                    // Ensure Test Date suggestion runs after qualifications refreshed so blank Test Date gets a sensible default
                    await suggestTestDate(true);
                }
            }

            if (jccEl) jccEl.focus();
        } catch (e) {
            console.error('AJAX submit error:', e);
            fallbackToStandardPost(form);
        } finally {
            if (ajaxSubmitEnabled) setLoading(false);
        }
    };

    if (welderFormEl && isNew) {
        welderFormEl.addEventListener('submit', handleAjaxSubmit);
    }
});
