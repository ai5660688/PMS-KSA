(function(){
    const dataEl = document.getElementById('drawingsData');
    const config = {
        mode: dataEl?.dataset.mode || '',
        selectedProjectId: dataEl?.dataset.selectedProjectId || '',
        drawingsUrl: dataEl?.dataset.drawingsUrl || '',
        getSheetRevisionUrl: dataEl?.dataset.getSheetRevisionUrl || '',
        updateSheetRevisionUrl: dataEl?.dataset.updateSheetRevisionUrl || '',
        suggestRevisionTagUrl: dataEl?.dataset.suggestRevisionTagUrl || ''
    };

    function readKeys(){
        const keysEl = document.getElementById('drawingKeysJson');
        if(!keysEl) return [];
        try{
            return JSON.parse(keysEl.textContent || '[]');
        }catch{
            return [];
        }
    }

    function syncFilterFormValues(projectId, layout, sheet){
        var forms = [
            document.getElementById('singleUploadForm'),
            document.getElementById('bulkUploadForm'),
            document.getElementById('reorderForm')
        ];
        forms.forEach(function(form){
            if(!form) return;
            var pidInput = form.querySelector('input[name="projectId"]');
            var layoutInput = form.querySelector('input[name="layout"]');
            var sheetInput = form.querySelector('input[name="sheet"]');
            if(pidInput && projectId !== undefined) pidInput.value = projectId;
            if(layoutInput && layout !== undefined) layoutInput.value = layout || '';
            if(sheetInput && sheet !== undefined) sheetInput.value = sheet || '';
        });
    }

    function refreshRevisions(layout, sheet){
        var projectSel = document.getElementById('projectSelect');
        var projectId = (projectSel && projectSel.value ? projectSel.value : config.selectedProjectId) || '';
        var baseUrl = config.drawingsUrl || '';
        if(!baseUrl) return;
        var url = new URL(baseUrl, window.location.origin);
        if(projectId) url.searchParams.set('projectId', projectId);
        if(layout) url.searchParams.set('layout', layout);
        if(sheet) url.searchParams.set('sheet', sheet);

        syncFilterFormValues(projectId, layout, sheet);
        config.selectedProjectId = projectId;

        fetch(url.toString(), { credentials: 'same-origin' })
            .then(r => r.ok ? r.text() : '')
            .then(html => {
                if(!html) return;
                var doc = new DOMParser().parseFromString(html, 'text/html');
                var newBlock = doc.getElementById('revisionsBlock');
                var newPlaceholder = doc.getElementById('revisionsPlaceholder');
                var currentBlock = document.getElementById('revisionsBlock');
                var currentPlaceholder = document.getElementById('revisionsPlaceholder');

                if(currentBlock && newBlock){
                    currentBlock.innerHTML = newBlock.innerHTML;
                    currentBlock.style.display = newBlock.style.display || '';
                }
                if(currentPlaceholder && newPlaceholder){
                    currentPlaceholder.innerHTML = newPlaceholder.innerHTML;
                    currentPlaceholder.style.display = newPlaceholder.style.display || '';
                }

                var newData = doc.getElementById('drawingsData');
                if(newData && newData.dataset){
                    if(newData.dataset.mode) config.mode = newData.dataset.mode;
                    if(newData.dataset.selectedProjectId) config.selectedProjectId = newData.dataset.selectedProjectId;
                }

                var newKeys = doc.getElementById('drawingKeysJson');
                var currentKeys = document.getElementById('drawingKeysJson');
                if(currentKeys && newKeys){
                    currentKeys.textContent = newKeys.textContent || '[]';
                    if(typeof window.refreshDrawingSheetMap === 'function'){
                        window.refreshDrawingSheetMap();
                    }
                }

                var newLayoutSel = doc.getElementById('layoutSelect');
                var newSheetSel = doc.getElementById('sheetSelect');
                var layoutSel = document.getElementById('layoutSelect');
                var sheetSel = document.getElementById('sheetSelect');

                if(layoutSel && newLayoutSel){
                    layoutSel.innerHTML = newLayoutSel.innerHTML;
                    layoutSel.disabled = newLayoutSel.disabled;
                    layoutSel.value = newLayoutSel.value;
                }
                if(sheetSel && newSheetSel){
                    sheetSel.innerHTML = newSheetSel.innerHTML;
                    sheetSel.disabled = newSheetSel.disabled;
                    sheetSel.value = newSheetSel.value;
                }

                var newRevInput = doc.getElementById('lsRevInput');
                var revInput = document.getElementById('lsRevInput');
                if(revInput && newRevInput){
                    revInput.value = newRevInput.value || '';
                }

                if(projectSel && newData?.dataset?.selectedProjectId){
                    projectSel.value = newData.dataset.selectedProjectId;
                }
            })
            .catch(()=>{ /* ignore refresh errors */ });
    }

    // Searchable dropdown widget: wraps a native <select> with a text input + filterable list
    (function(){
        function makeSearchableSelect(sel){
            if(!sel || sel.dataset.searchable) return;
            sel.dataset.searchable='1';

            // Hide the original select and build the wrapper
            sel.style.display='none';
            var wrap = document.createElement('div');
            wrap.className = 'ss-wrap';
            // Copy width classes from the original select
            if(sel.classList.contains('w350')) wrap.classList.add('w350');
            else if(sel.classList.contains('w130')) wrap.classList.add('w130');
            else if(sel.classList.contains('w90')) wrap.classList.add('w90');
            sel.parentNode.insertBefore(wrap, sel);
            wrap.appendChild(sel);

            var input = document.createElement('input');
            input.type = 'text';
            input.className = 'ss-input form-control';
            input.setAttribute('autocomplete','off');
            input.setAttribute('spellcheck','false');
            input.placeholder = '';
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
                close();
                sel.dispatchEvent(new Event('change', { bubbles: true }));
            }

            function show(){
                if(open) return;
                open = true;
                render(input.value !== (items[sel.selectedIndex]?.text || '') ? input.value : '');
                list.style.display = 'block';
                wrap.classList.add('ss-open');
            }

            function close(){
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
                show();
            });

            input.addEventListener('input', function(){
                buildItems();
                show();
                render(input.value);
            });

            input.addEventListener('keydown', function(e){
                if(e.key === 'ArrowDown'){
                    e.preventDefault();
                    if(!open){ show(); render(''); }
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
                    close();
                    input.blur();
                } else if(e.key === 'Tab'){
                    syncInputText();
                    close();
                }
            });

            input.addEventListener('blur', function(){
                setTimeout(function(){
                    if(!wrap.contains(document.activeElement)){
                        syncInputText();
                        close();
                    }
                }, 150);
            });

            // Observe option changes (e.g., refillSheets dynamically replaces options)
            var mo = new MutationObserver(function(){
                buildItems();
                syncInputText();
                if(open) render(input.value);
            });
            mo.observe(sel, { childList: true, subtree: true });

            // Initial state
            buildItems();
            syncInputText();
        }
        ['projectSelect','layoutSelect','sheetSelect'].forEach(function(id){ var el=document.getElementById(id); if(el) makeSearchableSelect(el); });
    })();

    // When Layout changes in Sheet mode, submit the form using the currently selected Sheet if it exists for the new Layout; otherwise use the first sheet
    (function(){
        if(config.mode !== 'Sheet') return;
        var layoutSel = document.getElementById('layoutSelect');
        var sheetSel = document.getElementById('sheetSelect');
        var revInput = document.getElementById('lsRevInput');
        var revSaveBtn = document.getElementById('lsRevSaveBtn');
        var revStatusEl = document.getElementById('lsRevStatus');
        var revisionsBlock = document.getElementById('revisionsBlock');
        var revisionsPlaceholder = document.getElementById('revisionsPlaceholder');
        if(!layoutSel || !sheetSel || !revInput) return;

        var map = {};
        function buildSheetMap(){
            var pairs = readKeys();
            map = {};
            if(Array.isArray(pairs)){
                pairs.forEach(function(key){
                    if(typeof key !== 'string') return;
                    var parts = key.split('|');
                    if(parts.length<2) return;
                    var lay = parts[0];
                    var sh = parts[1];
                    if(!map[lay]) map[lay]=[];
                    if(map[lay].indexOf(sh)<0) map[lay].push(sh);
                });
            }
            Object.keys(map).forEach(function(l){ map[l].sort(); });
        }
        buildSheetMap();
        window.refreshDrawingSheetMap = buildSheetMap;

        function refillSheets(layout){
            var list = map[layout] || [];
            while(sheetSel.options.length>0) sheetSel.remove(0);
            if(list.length===0){
                var op=document.createElement('option');
                op.value='';
                op.textContent='';
                sheetSel.appendChild(op);
                return;
            }
            list.forEach(function(s){
                var op=document.createElement('option');
                op.value=s;
                op.textContent=s;
                sheetSel.appendChild(op);
            });
        }
        function getAntiForgery(){
            var t = document.querySelector('input[name="__RequestVerificationToken"]');
            return t ? t.value : '';
        }
        function setStatus(msg, isError){
            if(!revStatusEl) return;
            revStatusEl.textContent = msg || '';
            revStatusEl.classList.toggle('error', !!isError);
            if(msg){
                clearTimeout(setStatus._t);
                setStatus._t = setTimeout(()=>{ revStatusEl.textContent=''; },3000);
            }
        }
        function fetchRev(layout, sheet, callback){
            if(!layout || !sheet){ revInput.value=''; if(callback) callback(); return; }
            fetch(config.getSheetRevisionUrl + `?projectId=${encodeURIComponent(config.selectedProjectId)}&layout=${encodeURIComponent(layout)}&sheet=${encodeURIComponent(sheet)}`, { credentials:'same-origin' })
                .then(r=>r.ok?r.json():null)
                .then(d=>{ if(d) {
                    // strip any outer parentheses for the editable input
                    var rev = d.rev || '';
                    rev = rev.replace(/^[()]+|[()]+$/g, '');
                    revInput.value = rev;
                }
                if(callback) callback(); })
                .catch(()=>{ if(callback) callback(); });
        }

        async function saveRev(){
            if(!layoutSel || !sheetSel){ return; }
            var layout = (layoutSel.value || '').trim();
            var sheet = (sheetSel.value || '').trim();
            if(!layout || !sheet){
                setStatus('Select layout and sheet first.', true);
                return;
            }
            var token = getAntiForgery();
            if(!token){
                setStatus('Missing session token. Please reload.', true);
                return;
            }
            var value = (revInput.value || '').trim();
            var prevText = revSaveBtn ? revSaveBtn.textContent : '';
            if(revSaveBtn){
                revSaveBtn.disabled = true;
                revSaveBtn.textContent = 'Saving...';
            }
            try{
                var form = new URLSearchParams();
                form.append('projectId', config.selectedProjectId);
                form.append('layout', layout);
                form.append('sheet', sheet);
                form.append('lsRev', value);
                form.append('__RequestVerificationToken', token);
                var response = await fetch(config.updateSheetRevisionUrl, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                        'RequestVerificationToken': token,
                        'X-Requested-With': 'XMLHttpRequest'
                    },
                    body: form.toString()
                });
                if(!response.ok){
                    let err = '';
                    try{ err = (await response.text()).trim(); }catch{}
                    setStatus(err || 'Save failed.', true);
                    return;
                }
                var data = await response.json().catch(()=>null);
                if(data && data.ok){
                    if(data.rev){ revInput.value = String(data.rev); }
                    setStatus('Revision saved.', false);
                    fetchRev(layout, sheet);
                } else {
                    var msg = data && data.message ? data.message : '';
                    setStatus(msg || 'Save failed.', true);
                }
            } catch(e){
                setStatus('Save failed.', true);
            } finally {
                if(revSaveBtn){
                    revSaveBtn.disabled = false;
                    revSaveBtn.textContent = prevText || 'Save';
                }
            }
        }

        // Hide revisions if there is no sheet on load (defensive) and show placeholder
        if(sheetSel && !sheetSel.value){ if(revisionsBlock) revisionsBlock.style.display='none'; if(revisionsPlaceholder) revisionsPlaceholder.style.display=''; }

        // On layout change: refill sheets, try to retain current sheet; then submit the form to refresh revision list
        layoutSel.addEventListener('change', function(){
            var layout = layoutSel.value;
            var prevSheet = sheetSel.value;
            refillSheets(layout);
            var hasPrev = Array.from(sheetSel.options).some(function(o){ return o.value === prevSheet; });
            if(hasPrev){
                sheetSel.value = prevSheet;
            } else if(sheetSel.options.length>0){
                sheetSel.selectedIndex = 0;
            }
            // Fetch current rev for selected sheet then refresh revisions area via AJAX
            var selected = sheetSel.value || '';
            fetchRev(layoutSel.value, selected, function(){
                refreshRevisions(layoutSel.value, selected);
            });
        });
        // On sheet change: hide placeholder, fetch rev then refresh revisions area via AJAX
        sheetSel.addEventListener('change', function(){
            if(revisionsPlaceholder) revisionsPlaceholder.style.display='none';
            // fetch current rev (for the editable input) then refresh revisions list
            fetchRev(layoutSel.value, sheetSel.value, function(){ refreshRevisions(layoutSel.value, sheetSel.value); });
        });

        if(revSaveBtn){
            revSaveBtn.addEventListener('click', saveRev);
        }

    })();

    // Bulk upload in manageable batches to avoid gigantic single request
    (function(){
        var bulkForm = document.getElementById('bulkUploadForm');
        var filesInput = document.getElementById('bulkFilesInput');
        var uploadBtn = document.getElementById('bulkUploadBtn');
        var chooseInfo = document.getElementById('bulkChooseInfo');
        var totalWrap = document.getElementById('bulkProgressWrap');
        var totalBar = document.getElementById('bulkProgressBar');
        var chooseLbl = document.getElementById('bulkChooseBtn');
        var statusEl = document.getElementById('bulkUploadStatus');
        var uploadList = document.getElementById('bulkUploadList');
        var closeBtn = document.getElementById('bulkUploadCloseBtn');
        var listMap = new Map();

        window.prepareBulkList = function(){
            if(!filesInput) return;
            var files = filesInput.files || [];
            listMap.clear();
            if(uploadList){ uploadList.innerHTML=''; uploadList.style.display = files.length ? '' : 'none'; }
            if(closeBtn) closeBtn.style.display = 'none';
            if(!files.length){ if(chooseInfo) chooseInfo.textContent=''; if(uploadBtn) uploadBtn.disabled = true; return; }
            if(chooseInfo) chooseInfo.textContent = 'Selected ' + files.length + ' file' + (files.length>1?'s':'');
            if(uploadBtn) uploadBtn.disabled = false;
            Array.from(files).forEach(function(f){
                if(!uploadList) return;
                var li = document.createElement('li');
                var nameSpan = document.createElement('span');
                nameSpan.textContent = f.name;
                var statusSpan = document.createElement('span');
                statusSpan.className = 'status';
                statusSpan.textContent = 'Pending';
                li.appendChild(nameSpan);
                li.appendChild(statusSpan);
                uploadList.appendChild(li);
                listMap.set(f.name, li);
            });
        };

        function setBulkStatus(files, text, statusClass){
            Array.from(files || []).forEach(function(f){
                var li = listMap.get(f.name);
                if(!li) return;
                var statusEl = li.querySelector('span.status');
                if(statusEl){
                    statusEl.textContent = text;
                    statusEl.className = 'status ' + statusClass;
                }
            });
        }

        function setBulkStatusByNames(names, text, statusClass){
            (names || []).forEach(function(name){
                var li = listMap.get(name);
                if(!li) return;
                var statusEl = li.querySelector('span.status');
                if(statusEl){
                    statusEl.textContent = text;
                    statusEl.className = 'status ' + statusClass;
                }
            });
        }

        function getToken(){
            var t = bulkForm ? bulkForm.querySelector('input[name="__RequestVerificationToken"]') : null;
            return t ? t.value : '';
        }

        function sumSizes(arr){ return arr.reduce(function(s,f){ return s + (f && f.size ? f.size : 0); }, 0); }

        if(bulkForm){
            bulkForm.addEventListener('submit', function(e){
                e.preventDefault();
                var files = Array.from((filesInput && filesInput.files) ? filesInput.files : []);
                if(!files.length) return;
                var token = getToken();
                if(uploadBtn) uploadBtn.disabled = true;
                if(filesInput) filesInput.disabled = true;
                if(chooseLbl){ chooseLbl.setAttribute('aria-disabled','true'); chooseLbl.style.pointerEvents='none'; }

                var totalBytes = sumSizes(files) || 1;
                var uploadedTotal = 0;
                if(totalWrap) totalWrap.style.display='block';
                if(totalBar) totalBar.style.width='0%';
                if(statusEl){ statusEl.textContent='Starting upload...'; statusEl.classList.remove('error'); }

                // Batch settings
                var BATCH_SIZE = 50; // tune as needed
                var batches = [];
                for(var i=0;i<files.length;i+=BATCH_SIZE){ batches.push(files.slice(i, i+BATCH_SIZE)); }

                var aggregate = { added:0, skipped:0, failed:0 };

                function uploadBatch(idx){
                    if(idx >= batches.length){
                        if(statusEl){ statusEl.textContent = 'Upload complete. ' + `Added ${aggregate.added}, Skipped ${aggregate.skipped}, Failed ${aggregate.failed}.`; }
                        if(totalBar) totalBar.style.width='100%';
                        if(uploadList) uploadList.style.display = '';
                        if(closeBtn) closeBtn.style.display = '';
                        return;
                    }
                    var chunk = batches[idx];
                    var fd = new FormData();
                    fd.append('projectId', config.selectedProjectId);
                    if(token) fd.append('__RequestVerificationToken', token);
                    chunk.forEach(function(f){ fd.append('files', f, f.name); });

                    var chunkBytes = sumSizes(chunk);
                    var baseUploaded = uploadedTotal;
                    if(statusEl){ statusEl.textContent = `Uploading batch ${idx+1}/${batches.length} (${chunk.length} files)...`; }

                    var xhr = new XMLHttpRequest();
                    xhr.upload.onprogress = function(ev){
                        if(ev.lengthComputable){
                            var pct = ((baseUploaded + ev.loaded) / totalBytes) * 100;
                            if(totalBar) totalBar.style.width = pct.toFixed(2) + '%';
                        }
                    };
                    xhr.onreadystatechange = function(){
                        if(xhr.readyState===4){
                            if(xhr.status>=200 && xhr.status<400){
                                uploadedTotal += chunkBytes;
                                    try{
                                        var resp = JSON.parse(xhr.responseText);
                                        if(resp){
                                            aggregate.added += resp.added||0;
                                            aggregate.skipped += resp.skipped||0;
                                            aggregate.failed += resp.failed||0;
                                            if(resp.addedFiles || resp.skippedFiles || resp.failedFiles){
                                                setBulkStatusByNames(resp.addedFiles, 'Uploaded', 'ok');
                                                setBulkStatusByNames(resp.skippedFiles, 'Skipped', 'skip');
                                                setBulkStatusByNames(resp.failedFiles, 'Failed', 'fail');
                                            } else if(resp.failed > 0 && resp.added === 0){
                                                setBulkStatus(chunk, 'Failed', 'fail');
                                            } else {
                                                setBulkStatus(chunk, 'Uploaded', 'ok');
                                            }
                                        } else {
                                            setBulkStatus(chunk, 'Uploaded', 'ok');
                                        }
                                    }catch{
                                        setBulkStatus(chunk, 'Uploaded', 'ok');
                                    }
                                uploadBatch(idx+1);
                            } else {
                                // Mark whole chunk as failed and continue
                                uploadedTotal += chunkBytes;
                                aggregate.failed += chunk.length;
                                setBulkStatus(chunk, 'Failed', 'fail');
                                if(statusEl){ statusEl.textContent = `A batch failed. Continuing (${idx+1}/${batches.length})...`; statusEl.classList.add('error'); }
                                uploadBatch(idx+1);
                            }
                        }
                    };
                    xhr.open('POST', bulkForm.getAttribute('action'), true);
                    xhr.setRequestHeader('X-Requested-With','XMLHttpRequest');
                    xhr.send(fd);
                }

                uploadBatch(0);
            });
            if(closeBtn){
                closeBtn.addEventListener('click', function(){
                    if(uploadList){ uploadList.innerHTML = ''; uploadList.style.display = 'none'; }
                    if(totalWrap) totalWrap.style.display = 'none';
                    if(totalBar) totalBar.style.width = '0%';
                    if(statusEl) statusEl.textContent = '';
                    if(chooseInfo) chooseInfo.textContent = '';
                    if(filesInput){ filesInput.value = ''; filesInput.disabled = false; }
                    if(uploadBtn){ uploadBtn.disabled = true; }
                    if(chooseLbl){ chooseLbl.removeAttribute('aria-disabled'); chooseLbl.style.pointerEvents = ''; }
                    closeBtn.style.display = 'none';
                    listMap.clear();
                });
            }
        }
    })();

    // Edit modal and delete confirm logic
    (function(){
        var editModal = document.getElementById('editDrawingModal');
        var deleteModal = document.getElementById('confirmDeleteModal');
        var deleteMsg = document.getElementById('confirmDeleteMessage');
        var pendingDeleteForm = null;
        var lastFocused = new Map();

        function getFocusable(modal){
            if(!modal) return [];
            return Array.from(modal.querySelectorAll('button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])'))
                .filter(el => !el.disabled && el.offsetParent !== null);
        }
        function showModal(modal, focusSelector){
            if(!modal) return;
            lastFocused.set(modal, document.activeElement);
            modal.style.display = 'flex';
            modal.setAttribute('aria-hidden', 'false');
            var focusEl = focusSelector ? modal.querySelector(focusSelector) : null;
            if(!focusEl){
                var focusables = getFocusable(modal);
                focusEl = focusables[0];
            }
            if(focusEl) focusEl.focus();
        }
        function hideModal(modal){
            if(!modal) return;
            modal.style.display = 'none';
            modal.setAttribute('aria-hidden', 'true');
            var prev = lastFocused.get(modal);
            if(prev && typeof prev.focus === 'function') prev.focus();
            lastFocused.delete(modal);
        }
        function trapFocus(modal, closeFn){
            if(!modal) return;
            modal.addEventListener('keydown', function(e){
                if(e.key === 'Escape'){
                    e.preventDefault();
                    if(closeFn) closeFn();
                    return;
                }
                if(e.key !== 'Tab') return;
                var focusables = getFocusable(modal);
                if(!focusables.length) return;
                var first = focusables[0];
                var last = focusables[focusables.length - 1];
                if(e.shiftKey && document.activeElement === first){
                    e.preventDefault();
                    last.focus();
                } else if(!e.shiftKey && document.activeElement === last){
                    e.preventDefault();
                    first.focus();
                }
            });
        }

        window.openEditDrawingModal = function(btn){
            var tr = btn.closest('tr');
            if(!tr) return;
            var id = tr.getAttribute('data-id');
            var file = tr.getAttribute('data-file') || '';
            var rev = tr.getAttribute('data-rev') || '';
            document.getElementById('editDrawingId').value = id || '';
            document.getElementById('editFileName').value = file;
            document.getElementById('editRevisionTag').value = rev;
            showModal(editModal, '#editFileName');
        };
        window.closeEditDrawingModal = function(){ hideModal(editModal); };

        // Intercept delete forms
        document.addEventListener('submit', function(e){
            var form = e.target;
            if(form && form.classList && form.classList.contains('delete-form')){
                e.preventDefault();
                var tr = form.closest('tr');
                var rev = tr ? (tr.getAttribute('data-rev') || '') : '';
                // Compose: Are you sure you want to delete revision "Layout - Revision Tag"?
                var layoutSel = document.getElementById('layoutSelect');
                var layoutVal = layoutSel ? (layoutSel.value || '') : '';
                var label = layoutVal && rev ? (layoutVal + ' - ' + rev) : (layoutVal || rev || '');
                if(deleteMsg){ deleteMsg.textContent = 'Are you sure you want to delete revision "' + label + '"?'; }
                pendingDeleteForm = form;
                showModal(deleteModal, '.btn-danger');
            }
        }, true);
        window.closeDeleteModal = function(){ pendingDeleteForm = null; hideModal(deleteModal); };
        window.confirmDelete = function(){ if(pendingDeleteForm){ hideModal(deleteModal); pendingDeleteForm.submit(); pendingDeleteForm=null; } };

        // close on backdrop click
        [editModal, deleteModal].forEach(function(m){ if(!m) return; m.addEventListener('click', function(e){ if(e.target===m) hideModal(m); }); });
        trapFocus(editModal, window.closeEditDrawingModal);
        trapFocus(deleteModal, window.closeDeleteModal);
    })();

    (function(){
        const fileInput = document.getElementById('singleFileInput');
        const revInput = document.getElementById('revisionTagInput');
        const statusEl = document.getElementById('singleUploadStatus');
        const progressWrap = document.getElementById('singleProgressWrap');
        const progressBar = document.getElementById('singleProgressBar');

        function resetSingleUploadStatus(){
            if(statusEl){ statusEl.textContent=''; statusEl.classList.remove('error'); }
            if(progressWrap){ progressWrap.style.display='none'; }
            if(progressBar){ progressBar.style.width='0%'; }
        }
        window.resetSingleUploadStatus = resetSingleUploadStatus;

        function applySuggestedTag(tag){
            if(!revInput) return;
            if(tag){
                revInput.value = tag;
                revInput.dataset.autoFilled = '1';
            } else if(revInput.dataset.autoFilled === '1'){
                revInput.value = '';
                delete revInput.dataset.autoFilled;
            }
        }

        async function requestSuggestedTag(fileName){
            if(!revInput) return;
            if(!fileName){
                applySuggestedTag('');
                return;
            }
            try{
                const url = config.suggestRevisionTagUrl + '?fileName=' + encodeURIComponent(fileName);
                const res = await fetch(url, { credentials: 'same-origin' });
                if(!res.ok) return;
                const data = await res.json();
                if(data && typeof data.revisionTag === 'string'){
                    applySuggestedTag(data.revisionTag);
                }
            }catch{ /* ignore suggestion errors */ }
        }

        if(fileInput){
            fileInput.addEventListener('change', ()=>{
                const name = (fileInput.files && fileInput.files[0]) ? fileInput.files[0].name : '';
                if(revInput) delete revInput.dataset.autoFilled;
                requestSuggestedTag(name);
            });
            if(fileInput.files && fileInput.files.length > 0){
                requestSuggestedTag(fileInput.files[0].name);
            }
        }

        revInput?.addEventListener('input', ()=>{
            if(revInput.dataset.autoFilled === '1'){
                delete revInput.dataset.autoFilled;
            }
        });
    })();

    // Ensure revisions refresh when filter header changes (Project/Layout/Sheet)
    (function(){
        var projectSel = document.getElementById('projectSelect');
        var layoutSel = document.getElementById('layoutSelect');
        var sheetSel = document.getElementById('sheetSelect');

        function currentLayout(){ return layoutSel ? (layoutSel.value || '') : ''; }
        function currentSheet(){ return config.mode === 'Sheet' && sheetSel ? (sheetSel.value || '') : ''; }
        function doRefresh(){ refreshRevisions(currentLayout(), currentSheet()); }

        if(projectSel){
            projectSel.addEventListener('change', function(){
                var projectId = projectSel.value || '';
                var baseUrl = config.drawingsUrl || window.location.pathname;
                var url = new URL(baseUrl, window.location.origin);
                if(projectId){
                    url.searchParams.set('projectId', projectId);
                }
                url.searchParams.delete('layout');
                url.searchParams.delete('sheet');
                window.location.href = url.toString();
            });
        }
        if(config.mode !== 'Sheet' && layoutSel){
            layoutSel.addEventListener('change', function(){
                doRefresh();
            });
        }
    })();

    window.collectOrder = function(){
        var ids = Array.from(document.querySelectorAll('#rev-tbody tr')).map(tr => tr.getAttribute('data-id'));
        var orderedIds = document.getElementById('orderedIds');
        if(orderedIds) orderedIds.value = ids.join(',');
    };

    window.moveUp = function(btn){ var tr = btn.closest('tr'); if (tr && tr.previousElementSibling) tr.parentNode.insertBefore(tr, tr.previousElementSibling); };
    window.moveDown = function(btn){ var tr = btn.closest('tr'); if (tr && tr.nextElementSibling) tr.parentNode.insertBefore(tr.nextElementSibling, tr); };
})();
