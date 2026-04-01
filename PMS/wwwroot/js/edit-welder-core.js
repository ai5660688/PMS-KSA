document.addEventListener('DOMContentLoaded', () => {
    const configEl = document.getElementById('editWelderConfig');
    let config = {};
    if (configEl && configEl.textContent) {
        try {
            config = JSON.parse(configEl.textContent);
        } catch (e) {
            config = {};
        }
    }
    const routes = config.routes || {};
    const defaultWqt = typeof config.defaultWqtAgency === 'string' ? config.defaultWqtAgency : '';
    document.body.setAttribute('data-default-wqt', defaultWqt);

    // Ensure any server-rendered flash (.msg or .msg-success) is visible on page load
    try {
        const container = document.querySelector('.container');
        if (container) {
            const top = document.getElementById('topFlashMsg') || container.querySelector('.msg-success') || container.querySelector('.msg.text-danger') || container.querySelector('.msg');
            if (top) {
                try { top.tabIndex = -1; } catch (e) { }
                try { top.focus(); } catch (e) { }
                try { window.scrollTo({ top: 0, behavior: 'smooth' }); } catch (e) { }
            }
        }
    } catch (e) { /* ignore */ }

    const isNew = (document.getElementById("IsAddFlow")?.value === "true");
    const state = { addQualificationMode: false, userChangedWqt: false };

    // Initialize variables
    const addQualFlagHidden = document.getElementById('AddNewQualification');
    const qualSection = document.getElementById('qualSection');
    const jccInput = document.getElementById('JCC_No');

    function setQualRequired(on){
        const requiredFields = [jccInput, 'Qualification_Welding_Process', 'Qualification_Material_P_No'];
        requiredFields.forEach(el => {
            const node = typeof el === 'string' ? document.getElementById(el) : el;
            if (!node) return;
            if (on) node.setAttribute('required','required');
            else node.removeAttribute('required');
        });
    }

    // Initialize qualification section visibility and requirements
    if (isNew) {
        const shouldShowQualSection = addQualFlagHidden?.value === 'true';
        if (!shouldShowQualSection && qualSection) {
            qualSection.style.display='none';
        }
        setQualRequired(shouldShowQualSection);
    } else {
        setQualRequired(true); // edit flow always requires
    }

    const qs = (sel, root=document) => root.querySelector(sel);
    const $ = (name) => {
        const id = name.replace(/\./g, "_");
        return document.getElementById(id) || qs(`[name='${name}']`);
    };
    const set = (name, v) => {
        const el = $(name);
        if (el) el.value = v ?? "";
    };
    const get = (name) => {
        const el = $(name);
        return el ? el.value.trim() : "";
    };
    const getOrInit = (name) => {
        const el = $(name);
        if (!el) return "";
        const val = (el.value || "").trim();
        if (val) return val;
        const init = (el.getAttribute("data-init-value") || "").trim();
        return init;
    };
    const saveBtn = document.getElementById("saveBtn");
    let lastReqId = 0;

    const setLoading = (on) => {
        if (saveBtn) saveBtn.disabled = !!on;
        const sec = document.getElementById('qualSection');
        if (sec) sec.classList.toggle('loading', !!on);
    };

    const pick = (o, ...keys) => keys.map(k => o?.[k]).find(v => v != null);
    const dateOnly = (v) => v ? String(v).substring(0, 10) : "";
    const symbolInput = $("Welder.Welder_Symbol");
    const projectSelect = document.getElementById('Project_Welder');

    // Searchable select widget (same as Welder Register)
    function makeSearchableSelect(sel){
        if(!sel || sel.dataset.searchable) return;
        sel.dataset.searchable='1';
        sel.style.display='none';
        var wrap = document.createElement('div');
        wrap.className = 'ss-wrap';
        sel.parentNode.insertBefore(wrap, sel);
        wrap.appendChild(sel);

        var input = document.createElement('input');
        input.type = 'text';
        input.className = 'ss-input';
        input.setAttribute('autocomplete','off');
        input.setAttribute('spellcheck','false');
        input.placeholder = 'Search project...';
        wrap.appendChild(input);

        var list = document.createElement('div');
        list.className = 'ss-list';
        wrap.appendChild(list);

        var items = [];
        var highlighted = -1;
        var open = false;

        function buildItems(){
            items = [];
            Array.from(sel.options).forEach(function(opt){
                items.push({ value: opt.value, text: opt.text || opt.value });
            });
        }
        function syncInputText(){
            var idx = sel.selectedIndex;
            input.value = idx >= 0 && sel.options[idx] ? (sel.options[idx].text || sel.options[idx].value) : '';
        }
        function render(filter){
            list.innerHTML = '';
            highlighted = -1;
            var lc = (filter || '').toLowerCase();
            var count = 0;
            items.forEach(function(item, i){
                if(lc && !(item.text||'').toLowerCase().includes(lc)) return;
                var div = document.createElement('div');
                div.className = 'ss-item';
                div.textContent = item.text;
                div.dataset.idx = String(i);
                if(item.value === sel.value) div.classList.add('ss-active');
                div.addEventListener('mousedown', function(e){
                    e.preventDefault();
                    pick(i);
                });
                list.appendChild(div);
                count++;
            });
            if(count === 0){
                var empty = document.createElement('div');
                empty.className = 'ss-empty';
                empty.textContent = 'No match';
                list.appendChild(empty);
            }
        }
        function pick(idx){
            if(idx < 0 || idx >= items.length) return;
            sel.value = items[idx].value;
            syncInputText();
            closeList();
            sel.dispatchEvent(new Event('change', { bubbles: true }));
        }
        function showList(){
            if(open) return;
            open = true;
            render(input.value !== (items[sel.selectedIndex]?.text || '') ? input.value : '');
            list.style.display = 'block';
            wrap.classList.add('ss-open');
        }
        function closeList(){
            open = false;
            list.style.display = 'none';
            wrap.classList.remove('ss-open');
            highlighted = -1;
        }
        function highlightItem(dir){
            var visible = list.querySelectorAll('.ss-item');
            if(!visible.length) return;
            if(dir === 'down') highlighted = highlighted < visible.length - 1 ? highlighted + 1 : 0;
            else highlighted = highlighted > 0 ? highlighted - 1 : visible.length - 1;
            visible.forEach(function(el, i){ el.classList.toggle('ss-highlight', i === highlighted); });
            if(visible[highlighted]) visible[highlighted].scrollIntoView({ block:'nearest' });
        }
        input.addEventListener('focus', function(){
            buildItems();
            input.select();
            showList();
        });
        input.addEventListener('input', function(){
            buildItems();
            showList();
            render(input.value);
        });
        input.addEventListener('keydown', function(e){
            if(e.key === 'ArrowDown'){
                e.preventDefault();
                if(!open){ showList(); render(''); }
                highlightItem('down');
            } else if(e.key === 'ArrowUp'){
                e.preventDefault();
                highlightItem('up');
            } else if(e.key === 'Enter'){
                e.preventDefault();
                if(open && highlighted >= 0){
                    var visible = list.querySelectorAll('.ss-item');
                    if(visible[highlighted]){
                        var idx = parseInt(visible[highlighted].dataset.idx, 10);
                        pick(idx);
                    }
                }
            } else if(e.key === 'Escape'){
                syncInputText();
                closeList();
                input.blur();
            } else if(e.key === 'Tab'){
                syncInputText();
                closeList();
            }
        });
        input.addEventListener('blur', function(){
            setTimeout(function(){
                if(!wrap.contains(document.activeElement)){
                    syncInputText();
                    closeList();
                }
            }, 150);
        });
        buildItems();
        syncInputText();
    }

    if (projectSelect) makeSearchableSelect(projectSelect);

    const fetchJson = async (url) => {
        const r = await fetch(url, { headers: { "Accept": "application/json" }});
        if (!r.ok) return { ok:false, status:r.status };
        return { ok:true, data: await r.json() };
    };

    const debounce = (fn, ms=300) => {
        let h;
        return (...a) => {
            clearTimeout(h);
            h=setTimeout(()=>fn(...a),ms);
        };
    };

    const existingSection = document.getElementById('existingQualsAdd');
    const existingBody = document.getElementById('existingQualsAddBody');
    const esc = (s='') => s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');

    // File upload UX
    const fileInput = document.getElementById('UploadQualFile');
    const fileNameEl = document.getElementById('UploadQualFileName');
    const uploadBtn = document.getElementById('UploadQualSubmit');
    const uploadForm = document.getElementById('uploadQualForm');

    // Track user-changes on Welder Location
    const locationSel = document.getElementById('Welder_Location');
    let userChangedLocation = false;
    locationSel?.addEventListener('change', () => { userChangedLocation = true; });

    // Suggest location based on Batch No when adding
    const batchInput = document.getElementById('Batch_No');
    async function suggestLocation(force=false){
        try {
            const raw = batchInput?.value?.trim();
            const url = raw ? `${routes.suggestLocation}?batchNo=${encodeURIComponent(raw)}` : routes.suggestLocation;
            const res = await fetch(url, { headers:{ 'Accept':'application/json' } });
            if (!res.ok) {
                if (force && locationSel) locationSel.value = 'Shop';
                return;
            }
            const data = await res.json();
            const s = (data && (data.suggested||data.Suggested)) || '';
            if (!s) {
                if (force && locationSel) locationSel.value = 'Shop';
                return;
            }
            if (force || (!userChangedLocation && isNew && locationSel)){
                locationSel.value = s;
                if (force) userChangedLocation = false;
            }
        } catch {
            if (force && locationSel) locationSel.value = 'Shop';
        }
    }

    batchInput?.addEventListener('change', () => {
        // Only suggest location/WQT/test date in add mode or when adding new qualification
        if (isNew || state.addQualificationMode) {
            suggestLocation(false);
            suggestWqtAgency(false);
            suggestTestDate(true);
        }
    });

    // Only suggest location in add mode on page load
    if (isNew) suggestLocation(false);

    // Suggest WQT Agency
    const wqtSel = document.getElementById('Qualification_WQT_Agency');

    async function suggestWqtAgency(applyIfEmptyOnly=true){
        try {
            if (!wqtSel) return;
            const raw = batchInput?.value?.trim();
            const url = raw ? `${routes.suggestWqt}?batchNo=${encodeURIComponent(raw)}` : routes.suggestWqt;
            const res = await fetch(url, { headers:{ 'Accept':'application/json' } });
            if (!res.ok) return;
            const data = await res.json();
            const agency = (data && (data.agency||data.Agency)) || '';
            if (!agency) return;
            const opt = Array.from(wqtSel.options).find(o => (o.value||'').toLowerCase() === agency.toLowerCase());
            const shouldApply = (!state.userChangedWqt && (applyIfEmptyOnly ? !wqtSel.value : true));
            if (opt && shouldApply){
                wqtSel.value = opt.value;
                wqtSel.dataset.userChoice = opt.value;
            }
        } catch {}
    }

    // Only suggest WQT in add mode or when adding new qualification
    if (isNew || state.addQualificationMode) {
        suggestWqtAgency(false);
    }

    // Suggest Test Date
    const testDateInput = document.getElementById('Test_Date');
    async function suggestTestDate(applyIfEmptyOnly=true){
        try {
            if (!testDateInput) return;
            if (applyIfEmptyOnly && testDateInput.value) return;
            const raw = batchInput?.value?.trim();
            const url = raw ? `${routes.suggestTestDate}?batchNo=${encodeURIComponent(raw)}` : routes.suggestTestDate;
            const res = await fetch(url, { headers:{ 'Accept': 'application/json' } });
            if (!res.ok) return;
            const data = await res.json();
            const date = (data && (data.date||data.Date)) || '';
            if (date) testDateInput.value = date;
        } catch {}
    }

    // Only suggest test date in add mode or when adding new qualification
    if (isNew || state.addQualificationMode) {
        suggestTestDate(true);
    }

    // Existing quals rendering
    const renderExistingQuals = (items) => {
        if (!existingSection || !existingBody) return;
        if (!items || !items.length) {
            existingBody.innerHTML = '<tr><td colspan="5" style="text-align:center;color:#7a9aa8;">None</td></tr>';
            return;
        }

        const normalizeKey = (k) => String(k||'').replace(/[^a-z0-9]/gi, '').toLowerCase();
        const getVal = (obj, base) => {
            if (!obj) return undefined;
            const target = normalizeKey(base);
            for (const key of Object.keys(obj)) {
                if (normalizeKey(key) === target) return obj[key];
            }
            return undefined;
        };

        existingBody.innerHTML = items.map(q => {
            const jcc = esc(getVal(q, 'JCC_No') || getVal(q, 'JCC') || getVal(q, 'jcc') || getVal(q, 'jccNo') || '');
            const wp  = esc(getVal(q, 'Welding_Process') || getVal(q, 'weldingProcess') || '');
            const mat = esc(getVal(q, 'Material_P_No') || getVal(q, 'materialPNo') || '');
            const dia = esc(getVal(q, 'Diameter_Range') || getVal(q, 'diameterRange') || '');
            const thk = esc(getVal(q, 'Max_Thickness') || getVal(q, 'maxThickness') || '');
            return `<tr><td>${jcc}</td><td>${wp}</td><td>${mat}</td><td>${dia}</td><td>${thk}</td></tr>`;
        }).join('');
    };

    const fillWelder = (w) => {
        set("Welder.Welder_ID", pick(w, "Welder_ID","welder_ID","welderId"));
        set("Welder.Welder_Symbol", pick(w, "Welder_Symbol","welder_Symbol","welderSymbol"));
        set("Welder.Name", pick(w, "Name","name"));
        set("Welder.Iqama_No", pick(w, "Iqama_No","iqama_No","iqamaNo"));
        set("Welder.Passport", pick(w, "Passport","passport"));
        set("Welder.Welder_Location", pick(w, "Welder_Location","welder_Location","welderLocation"));
        set("Welder.Mobile_No", pick(w, "Mobile_No","mobile_No","mobileNo"));
        set("Welder.Email", pick(w, "Email","email"));
        set("Welder.Mobilization_Date", dateOnly(pick(w, "Mobilization_Date","mobilization_Date","mobilizationDate")));
        set("Welder.Demobilization_Date", dateOnly(pick(w, "Demobilization_Date","demobilization_Date","demobilizationDate")));
        set("Welder.Status", pick(w, "Status","status"));
        set("Welder.Project_Welder", pick(w, "Project_Welder","project_Welder","projectWelder"));
    };

    const clearWelder = () => {
        ["Welder.Welder_ID","Welder.Name","Welder.Iqama_No","Welder.Passport","Welder.Mobile_No","Welder.Email","Welder.Mobilization_Date","Welder.Demobilization_Date"].forEach(n => set(n, ""));
        // Preserve defaults for new-welder flow (e.g. Status) so user-provided page defaults remain
        if (!isNew) {
            // only clear status when editing flow (existing welder was cleared)
            set("Welder.Status", "");
        }
        if (existingSection) existingSection.style.display='none';
        if (existingBody) existingBody.innerHTML='<tr><td colspan="5" style="text-align:center;color:#7a9aa8;">None</td></tr>';
    };

    const loadExistingQuals = async (welderId) => {
        if (!welderId || !existingSection) return;
        const res = await fetchJson(`${routes.getQualifications}?welderId=${welderId}`);
        if (!res.ok) {
            renderExistingQuals([]);
            return;
        }
        const items = res.data?.items || [];
        renderExistingQuals(items);
        existingSection.style.display = 'block';
    };

    const onSymbolChanged = debounce(async () => {
        const sym = get("Welder.Welder_Symbol").replace(/\s+/g,"");
        if (!sym) {
            clearWelder();
            await suggestLocation(true);
            return;
        }
        const projectVal = projectSelect?.value?.trim();
        const query = new URLSearchParams({ symbol: sym });
        if (projectVal) query.set('projectId', projectVal);
        const res = await fetchJson(`${routes.findWelder}?${query.toString()}`);
        if (!res.ok) {
            clearWelder();
            await suggestLocation(true);
            await ensureJcc(true);
            return;
        }
        fillWelder(res.data);
        await ensureJcc(true);
        const welderId = res.data?.Welder_ID || res.data?.welder_ID || res.data?.welderId;
        if (isNew && welderId > 0) loadExistingQuals(welderId);
    }, 400);

    symbolInput?.addEventListener("input", onSymbolChanged);
    symbolInput?.addEventListener("blur", onSymbolChanged);
    projectSelect?.addEventListener('change', () => { onSymbolChanged(); });

    const toArr = (v) => Array.isArray(v) ? v : (v ? [String(v)] : []);
    const matchCI = (items, value) => {
        if (!value) return "";
        const lc = value.toLowerCase();
        return (items || []).find(x => String(x).toLowerCase() === lc) || "";
    };

    const buildOptions = (items, keep) => {
        const seenLC = new Set();
        const opts = [];
        for (const it of toArr(items)) {
            const raw = String(it);
            const key = raw.toLowerCase();
            if (seenLC.has(key)) continue;
            seenLC.add(key);
            const val = esc(raw);
            opts.push(`<option value="${val}">${val}</option>`);
        }
        const k = typeof keep === "string" ? keep : "";
        const kKey = k.toLowerCase();
        if (k && k !== "__new__" && !seenLC.has(kKey)) {
            const val = esc(k);
            seenLC.add(kKey);
            opts.push(`<option value="${val}">${val}</option>`);
        }
        opts.push(`<option value="__new__">+ Add new...</option>`);
        return opts.join("");
    };

    const fields = {
        batch: "Qualification.Batch_No",
        proc: "Qualification.Welding_Process",
        mat: "Qualification.Material_P_No",
        codeRef: "Qualification.Code_Reference",
        rootF: "Qualification.Consumable_Root_F_No",
        rootSpec: "Qualification.Consumable_Root_Spec",
        capF: "Qualification.Consumable_Filling_Cap_F_No",
        capSpec: "Qualification.Consumable_Filling_Cap_Spec",
        pos: "Qualification.Position_Progression",
        dia: "Qualification.Diameter_Range",
        thick: "Qualification.Max_Thickness",
        jcc: "Qualification.JCC_No",
        aramco: "Qualification.Received_from_Aramco"
    };

    const cascadeOrder = [fields.proc, fields.mat, fields.codeRef, fields.rootF, fields.rootSpec, fields.capF, fields.capSpec, fields.pos, fields.dia, fields.thick];
    const indexOfField = new Map(cascadeOrder.map((n, i) => [n, i]));

    const toggleAddNew = (sel) => {
        if (!sel) return;
        const addNew = document.getElementById(sel.id + "_New");
        if (!addNew) return;
        if (sel.value === "__new__") {
            addNew.style.display = "block";
            if (!addNew.name) {
                addNew.name = sel.dataset.origName || sel.name || "";
                sel.dataset.origName = sel.dataset.origName || sel.name || "";
                sel.name = "";
            }
        } else {
            addNew.style.display = "none";
            if (addNew.name) {
                sel.name = sel.dataset.origName || sel.name || addNew.name;
                addNew.name = "";
            }
        }
    };

    const initAddNewSelect = (sel, options = {}) => {
        if (!sel) return;
        const { onChange, onInput, trackUserChoice } = options;
        toggleAddNew(sel);
        sel.addEventListener('change', () => {
            if (trackUserChoice) sel.dataset.userChoice = sel.value || '';
            toggleAddNew(sel);
            if (onChange) onChange();
        });
        const addNew = document.getElementById(`${sel.id}_New`);
        if (addNew) {
            addNew.addEventListener('input', () => {
                if (trackUserChoice) sel.dataset.userChoice = '__new__';
                if (onInput) onInput();
            });
        }
    };

    const wqtAgencySel = document.getElementById('Qualification_WQT_Agency');
    initAddNewSelect(wqtAgencySel, {
        trackUserChoice: true,
        onChange: () => { state.userChangedWqt = true; },
        onInput: () => { state.userChangedWqt = true; }
    });

    // Initialize Status dropdown with same add-new logic as Welding Process
    const statusSel = document.getElementById('Welder_Status');
    if (statusSel) {
        // Apply initial toggle and respect preselected value
        const init = (statusSel.getAttribute('data-init-value') || '').trim();
        if (init && !statusSel.value) {
            const opt = Array.from(statusSel.options).find(o => (o.value || '').toLowerCase() === init.toLowerCase());
            if (opt) statusSel.value = opt.value;
        }
        initAddNewSelect(statusSel, { trackUserChoice: true });
    }

    const updateSelect = (name, items, effective) => {
        const el = $(name);
        if (!el) return;
        const prev = el.value || "";
        const userChoice = el.dataset.userChoice || "";
        let prefer = userChoice || prev || "";
        const list = toArr(items);
        const hasPrefer = matchCI(list, prefer) !== "";
        const shouldUseEffective = !userChoice && (!prefer || !hasPrefer);
        const canonicalPrefer = matchCI(list, prefer) || prefer;
        const canonicalEff = matchCI(list, effective) || effective || "";
        const keep = canonicalPrefer || canonicalEff || "";
        el.innerHTML = buildOptions(list, keep);
        let finalValue = canonicalPrefer;
        if (shouldUseEffective && canonicalEff) finalValue = canonicalEff;
        if (!finalValue) {
            for (const opt of el.options) {
                const val = (opt.value || "").trim();
                if (val && val !== "__new__") {
                    finalValue = opt.value;
                    break;
                }
            }
        }
        el.value = finalValue || "";
        toggleAddNew(el);
    };

    const clearDownstream = (changedFieldName) => {
        const idx = indexOfField.get(changedFieldName);
        if (idx == null) return;
        for (let i = idx + 1; i < cascadeOrder.length; i++) {
            const n = cascadeOrder[i];
            const sel = $(n);
            if (!sel) continue;
            sel.dataset.userChoice = "";
            sel.value = "";
            const addNew = document.getElementById((sel.id || n.replace(/\./g, "_")) + "_New");
            if (addNew) {
                addNew.value = "";
                if (!sel.name) sel.name = sel.dataset.origName || sel.name || n;
                addNew.name = "";
            }
        }
    };

    const loadQualOptions = debounce(async () => {
        const reqId = ++lastReqId;
        setLoading(true);
        try {
            const p = new URLSearchParams();
            const vProc = getOrInit("Qualification.Welding_Process");
            const vMat = getOrInit("Qualification.Material_P_No");
            const vCodeRef = getOrInit("Qualification.Code_Reference");
            const vRootF = getOrInit("Qualification.Consumable_Root_F_No");
            const vRootSpec = getOrInit("Qualification.Consumable_Root_Spec");
            const vCapF = getOrInit("Qualification.Consumable_Filling_Cap_F_No");
            const vCapSpec = getOrInit("Qualification.Consumable_Filling_Cap_Spec");
            const vPos = getOrInit("Qualification.Position_Progression");
            const vDia = getOrInit("Qualification.Diameter_Range");
            const vThk = getOrInit("Qualification.Max_Thickness");

            if (vProc) p.set("weldingProcess", vProc);
            if (vMat) p.set("materialPNo", vMat);
            if (vCodeRef) p.set("codeReference", vCodeRef);
            if (vRootF) p.set("rootFNo", vRootF);
            if (vRootSpec) p.set("rootSpec", vRootSpec);
            if (vCapF) p.set("fillCapFNo", vCapF);
            if (vCapSpec) p.set("fillCapSpec", vCapSpec);
            if (vPos) p.set("positionProgression", vPos);

            const fetchAndApply = async(urlParams) => {
                const res = await fetchJson(`${routes.getQualOptions}?${urlParams.toString()}`);
                if (!res.ok || reqId !== lastReqId) return { applied:false };
                const d = res.data || {};
                const eff = (pick(d, "effective", "Effective") || {});
                const toList = (...keys) => toArr(pick(d, ...keys));
                const effVal = (...keys) => pick(eff, ...keys);

                // For edit mode, preserve existing values by using getOrInit instead of effective values
                const isEditing = !isNew && !state.addQualificationMode;

                updateSelect(fields.proc, toList("weldingProcesses","WeldingProcesses"),
                    isEditing ? getOrInit("Qualification.Welding_Process") : effVal("weldingProcess"));
                updateSelect(fields.mat, toList("materialPNos","MaterialPNos"),
                    isEditing ? getOrInit("Qualification.Material_P_No") : effVal("materialPNo"));
                updateSelect(fields.codeRef, toList("codeReferences","CodeReferences"),
                    isEditing ? getOrInit("Qualification.Code_Reference") : effVal("codeReference"));
                updateSelect(fields.rootF, toList("rootFNos","RootFNos"),
                    isEditing ? getOrInit("Qualification.Consumable_Root_F_No") : effVal("rootFNo"));
                updateSelect(fields.rootSpec, toList("rootSpecs","RootSpecs"),
                    isEditing ? getOrInit("Qualification.Consumable_Root_Spec") : effVal("rootSpec"));
                updateSelect(fields.capF, toList("fillCapFNos","FillCapFNos"),
                    isEditing ? getOrInit("Qualification.Consumable_Filling_Cap_F_No") : effVal("fillCapFNo"));
                updateSelect(fields.capSpec, toList("fillCapSpecs","FillCapSpecs"),
                    isEditing ? getOrInit("Qualification.Consumable_Filling_Cap_Spec") : effVal("fillCapSpec"));
                updateSelect(fields.pos, toList("positionProgressions","PositionProgressions"),
                    isEditing ? getOrInit("Qualification.Position_Progression") : effVal("positionProgression"));
                updateSelect(fields.dia, toList("diameterRanges","DiameterRanges"),
                    isEditing ? getOrInit("Qualification.Diameter_Range") : effVal("diameterRange"));
                updateSelect(fields.thick, toList("maxThicknesses","MaxThicknesses"),
                    isEditing ? getOrInit("Qualification.Max_Thickness") : effVal("maxThickness"));

                if (!state.addQualificationMode) {
                    cascadeOrder.forEach(n => {
                        const el = $(n);
                        if (!el) return;
                        const desired = (el.dataset.userChoice || el.getAttribute("data-init-value") || "").trim();
                        if (desired) {
                            const opt = Array.from(el.options).find(o => (o.value || "").toLowerCase() === desired.toLowerCase());
                            if (opt) el.value = opt.value;
                            toggleAddNew(el);
                        }
                    });
                }

                if (state.addQualificationMode) state.addQualificationMode = false;
                return { applied:true, effective: eff };
            };

            await fetchAndApply(p);
        } finally {
            if (reqId === lastReqId) setLoading(false);
        }
    }, 200);

    function activateSelects() {
        cascadeOrder.forEach(function (n) {
            const el = $(n);
            if (el) {
                el.disabled = false;
                el.removeAttribute('aria-disabled');
            }
        });
    }

    [...cascadeOrder, fields.aramco].forEach(function (n) {
        const sel = $(n);
        if (!sel) return;
        if (!sel.dataset.userChoice) {
            const baseId = sel.id || n.replace(/\./g, "_");
            const addNew = document.getElementById(baseId + "_New");
            const newVal = (addNew && addNew.value && addNew.value.trim()) || "";
            if (newVal) sel.dataset.userChoice = "__new__";
            else if (sel.value) sel.dataset.userChoice = sel.value;
            else {
                const init = (sel.getAttribute("data-init-value") || "").trim();
                if (init) sel.dataset.userChoice = init;
            }
        }
        toggleAddNew(sel);
        sel.addEventListener("change", function () {
            sel.dataset.userChoice = sel.value || "";
            toggleAddNew(sel);
            clearDownstream(n);
            loadQualOptions();
        });
        const baseId2 = sel.id || n.replace(/\./g, "_");
        const addNew2 = document.getElementById(baseId2 + "_New");
        if (addNew2) addNew2.addEventListener('input', function () {
            sel.dataset.userChoice = "__new__";
            clearDownstream(n);
            loadQualOptions();
        });
    });

    const jccEl = document.getElementById("JCC_No") || document.querySelector("[name='Qualification.JCC_No']");
    const batchEl = document.getElementById("Batch_No");

    const setAutoJcc = (val) => {
        if (jccEl) jccEl.dataset.autoJcc = val || "";
    };

    const isManualJcc = () => {
        if (!jccEl) return false;
        const autoVal = (jccEl.dataset.autoJcc || "").trim();
        const cur = (jccEl.value || "").trim();
        return !!cur && (!autoVal || cur !== autoVal);
    };

    jccEl?.addEventListener('input', () => {
        if (!jccEl) return;
        const autoVal = (jccEl.dataset.autoJcc || "").trim();
        const cur = (jccEl.value || "").trim();
        if (autoVal && cur !== autoVal) setAutoJcc("");
    });

    const buildNextJccUrl = () => {
        const p = new URLSearchParams();
        const projectVal = projectSelect?.value?.trim();
        const locVal = locationSel?.value?.trim();
        const batchVal = batchEl?.value?.trim();
        if (projectVal) p.set("projectId", projectVal);
        if (locVal) p.set("welderLocation", locVal);
        if (batchVal) p.set("batchNo", batchVal);
        const query = p.toString();
        return query ? `${routes.nextJcc}?${query}` : routes.nextJcc;
    };

    const updateJccPlaceholder = () => {
        if (!jccEl) return;
        const batch = batchEl?.value?.trim();
        jccEl.placeholder = batch ? `Click to generate next JCC No (Batch ${batch})` : "Click to generate next JCC No";
    };

    const ensureJcc = debounce(async (force = false) => {
        if (!jccEl) return;
        const currentVal = (jccEl.value || "").trim();
        if (!force && currentVal) return;
        if (force && isManualJcc()) return;
        const url = buildNextJccUrl();
        const res = await fetch(url, { headers: { "Accept": "application/json" } });
        if (!res.ok) return;
        const data = await res.json();
        const next = data?.jcc ?? data?.JCC ?? data?.next ?? null;
        if (!next) return;
        const allowReplace = !isManualJcc() || force;
        if (allowReplace) {
            jccEl.value = next;
            setAutoJcc(next);
        }
    }, 150);

    ["focus","click"].forEach(evt => jccEl?.addEventListener(evt, () => ensureJcc(false)));
    batchEl?.addEventListener("change", () => {
        updateJccPlaceholder();
        ensureJcc(true);
    });
    locationSel?.addEventListener("change", () => {
        updateJccPlaceholder();
        ensureJcc(true);
    });
    projectSelect?.addEventListener("change", () => {
        updateJccPlaceholder();
        ensureJcc(true);
    });

    // Initialize for edit mode without AJAX call
    activateSelects();
    updateJccPlaceholder();
    ensureJcc(false);

    // Handle qualification options based on mode
    if (isNew) {
        // In add mode, fetch options
        loadQualOptions();
    } else {
        // In edit mode, preserve existing values without AJAX call
        // First, ensure all dropdowns show their current values
        const fieldsToInitialize = [...cascadeOrder, fields.aramco, 'Qualification.WQT_Agency'];
        fieldsToInitialize.forEach(fieldName => {
            const sel = $(fieldName);
            if (!sel) return;

            // Get the saved value from data-init-value
            const savedValue = (sel.getAttribute("data-init-value") || "").trim();
            if (savedValue) {
                // Find the option matching the saved value
                const matchingOption = Array.from(sel.options).find(option =>
                    option.value && option.value.toLowerCase() === savedValue.toLowerCase());

                if (matchingOption) {
                    sel.value = matchingOption.value;
                    sel.dataset.userChoice = matchingOption.value;
                } else {
                    // If not in options, mark for "Add new"
                    sel.dataset.userChoice = "__new__";
                    const newInput = document.getElementById(sel.id + "_New");
                    if (newInput) {
                        newInput.value = savedValue;
                        newInput.style.display = "block";
                        sel.name = "";
                    }
                }
                toggleAddNew(sel);
            }
        });

        // Only fetch options if user interacts (cascading will still work)
        loadQualOptions();
    }

    // Apply default WQT only for add flows
    const applyDefaultWqt = async () => {
        try {
            // Only apply defaults in add mode or when adding new qualification
            if ((isNew || state.addQualificationMode)) {
                if (typeof wqtSel !== 'undefined' && wqtSel) {
                    await suggestWqtAgency(false);
                    if (!wqtSel.value) applyDefaultWqtFromBody();
                }
            }
        } catch (e) { /* ignore */ }
    };

    const applyDefaultWqtFromBody = () => {
        const def = (document.body.getAttribute('data-default-wqt')||'').trim();
        if (!def) return;
        const candidates = [];
        if (typeof wqtSel !== 'undefined' && wqtSel) candidates.push(wqtSel);
        const maybe = document.getElementById('Qualification_WQT_Agency');
        if (maybe && candidates.indexOf(maybe) === -1) candidates.push(maybe);
        candidates.forEach(sel => {
            try {
                const opt = Array.from(sel.options).find(o => (o.value||'').toLowerCase() === def.toLowerCase());
                if (opt) {
                    sel.value = opt.value;
                    sel.dataset.userChoice = opt.value;
                }
            } catch (e) { /* ignore */ }
        });
    };

    applyDefaultWqt();

    const addBtn = document.getElementById('addQualificationBtn');
    addBtn?.addEventListener('click', async () => {
        const addNewFlag = document.getElementById('AddNewQualification');
        if (addNewFlag) addNewFlag.value='true';
        state.addQualificationMode = true;
        if (jccEl) jccEl.value='';
        const selJccHidden = document.getElementById('SelectedJcc');
        if (selJccHidden) selJccHidden.value='';

        // Clear fields
        ['Test_Date','Qualification_Cert_Ref_No','Date_Issued','Remarks','DATE_OF_LAST_CONTINUITY','RECORDING_THE_CONTINUITY_RECORD'].forEach(id => {
            const el = document.getElementById(id);
            if (el) el.value='';
        });

        cascadeOrder.concat([fields.aramco]).forEach(n => {
            const sel = $(n);
            if (!sel) return;
            sel.dataset.userChoice='';
            sel.setAttribute('data-init-value','');
            sel.value='';
            toggleAddNew(sel);
        });

        if (batchInput) {
            const maxBatch = parseInt(batchInput.dataset.maxBatch||'0',10);
            if (maxBatch>0) batchInput.value = maxBatch;
        }

        await applyDefaultWqt();
        loadQualOptions();
        ensureJcc(true);
        await suggestTestDate(true);
        document.getElementById('qualSection')?.scrollIntoView({behavior:'smooth', block:'start'});
    });

    if (fileInput && fileNameEl && uploadBtn && uploadForm) {
        fileInput.addEventListener('change', () => {
            const file = fileInput.files && fileInput.files[0];
            fileNameEl.textContent = file ? file.name : 'No file chosen';
            uploadBtn.disabled = !file;
        });
    }

    window.editWelderShared = {
        isNew,
        routes,
        state,
        set,
        get,
        getOrInit,
        setLoading,
        renderExistingQuals,
        loadQualOptions,
        applyDefaultWqt,
        suggestWqtAgency,
        suggestTestDate,
        wqtSel,
        jccEl
    };
});
