// DWG open/download helpers with logging and popup-blocker fallback
(function(){
  let currentConfig = null;
  let delegatedBound = false;
  function buildQueryUrl(base, params){
    const u = new URL(base, window.location.origin);
    if(params && typeof params === 'object'){
      Object.keys(params).forEach(k => { if(params[k] !== undefined && params[k] !== null) u.searchParams.set(k, String(params[k])); });
    }
    return u.toString();
  }
  function safeStatus(msg, ok){ try { if(typeof window.showStatus === 'function') window.showStatus(msg, ok); } catch{} }
  function tryOpenInNewTab(url){
    try{
      const w = window.open(url, '_blank');
      if(!w){ window.location.href = url; }
    }catch{ window.location.href = url; }
  }
  function sanitizeFilename(name){
    if(!name) return name;
    return name.replace(/[\\\/:\*?"<>|\r\n\t]+/g,'').trim();
  }
  function getExtFromDisposition(filename){
    if(!filename) return '';
    const idx = filename.lastIndexOf('.', filename.length);
    if(idx<0) return '';
    return filename.substring(idx);
  }
  function getAntiForgeryToken(){
    try{ const el = document.querySelector('input[name="__RequestVerificationToken"]'); return el? el.value : ''; }catch{ return ''; }
  }
  function setBusy(btn, on){
    if(!btn) return;
    if(on){ btn.disabled=true; btn.dataset.prevText = btn.textContent; btn.textContent='Working...'; }
    else { btn.disabled=false; if(btn.dataset.prevText){ btn.textContent=btn.dataset.prevText; delete btn.dataset.prevText; }}
  }

  // Delegated row click handler so newly injected tables (Daily Welding fragment) work without re-init
  function bindDelegatedRowOpen(){
    if(delegatedBound) return;
    delegatedBound = true;
    document.addEventListener('click', function(e){
      try{
        const btn = e.target?.closest('.open-dwg-row');
        if(!btn) return;
        if(!currentConfig || !currentConfig.openUrlBase) return;
        e.preventDefault();
        const pid = (document.getElementById('projectSelect')?.value || '').trim();
        if(!pid){ safeStatus('Select Project first', false); return; }
        const tr = btn.closest('tr');
        const layoutCell = tr?.querySelector('.col-layout');
        const layout = (layoutCell?.textContent || btn.dataset.layout || '').trim();
        const sheetInput = tr?.querySelector('input[data-name="Sheet"], select[data-name="Sheet"]');
        let sheet = sheetInput ? (sheetInput.value || sheetInput.textContent || '') : '';
        if(!sheet){
          const sheetCell = tr?.querySelector('.col-sheet');
          sheet = sheetCell ? (sheetCell.textContent || '') : '';
        }
        if(!layout){ safeStatus('Layout missing for this row', false); return; }
        const url = buildQueryUrl(currentConfig.openUrlBase, { projectId: pid, layout: layout, sheet: sheet || undefined });
        tryOpenInNewTab(url);
      }catch(err){ console.debug('[dailyfitup-dwg] delegated open error', err); }
    });
  }

  function init(openUrlBase, downloadUrlBase, downloadZipBase, headerView){
    currentConfig = { openUrlBase, downloadUrlBase, downloadZipBase, headerView };
    bindDelegatedRowOpen();
    function attachHandlers(){
      const openBtn = document.getElementById('openDwgBtn');
      const dlBtn = document.getElementById('downloadDwgBtn');
      function currentProjectId(){ return document.getElementById('projectSelect')?.value || ''; }
      function currentLayout(){ return document.getElementById('layoutSelect')?.value || ''; }
      function currentSheet(){ return document.getElementById('sheetSelect')?.value || ''; }
      function effectiveHeaderView(){ return document.getElementById('headerViewSelect')?.value || headerView || 'DWG'; }

      openBtn?.addEventListener('click', function(){
        try{
          if(effectiveHeaderView() !== 'DWG'){ safeStatus('Open DWG only available in DWG view', false); return; }
          const pid = currentProjectId(); const layout=currentLayout();
          if(!pid || !layout){ safeStatus('Select Project and Layout first', false); return; }
          const url = buildQueryUrl(openUrlBase, { projectId: pid, layout: layout, sheet: currentSheet() });
          tryOpenInNewTab(url);
        }catch{ safeStatus('Failed to open drawing', false); }
      });

      dlBtn?.addEventListener('click', function(){
        try{
          const pid=currentProjectId(); if(!pid){ safeStatus('Select Project first', false); return; }
          const view=effectiveHeaderView();
          if(view==='DWG'){
            const layout=currentLayout(); if(!layout){ safeStatus('Select Layout first', false); return; }
            const sheet=currentSheet();
            const url=buildQueryUrl(downloadUrlBase,{ projectId: pid, layout: layout, sheet: sheet });
            (async()=>{
              setBusy(dlBtn,true);
              try{
                const res=await fetch(url,{ credentials:'same-origin', headers:{ 'RequestVerificationToken': getAntiForgeryToken(), 'Accept':'application/pdf' }});
                if(!res.ok){ tryOpenInNewTab(url); return; }
                const blob=await res.blob();
                let filename='';
                const cd=res.headers.get('content-disposition');
                if(cd){ const m=/filename\*=UTF-8''([^;\n\r]+)/i.exec(cd)||/filename="?([^";]+)"?/i.exec(cd); if(m) filename=decodeURIComponent(m[1].replace(/\\u0022/g,'').trim()); }
                if(!filename){ filename=`Drawing_${pid}_${new Date().toISOString().replace(/[:.]/g,'')}.pdf`; }
                filename=sanitizeFilename(filename);
                const obj=URL.createObjectURL(blob); const a=document.createElement('a'); a.href=obj; a.download=filename; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(obj);
                safeStatus('Download started', true);
              }catch{ tryOpenInNewTab(url); }
              finally{ setBusy(dlBtn,false); }
            })();
          } else {
            if(!downloadZipBase){ safeStatus('ZIP endpoint not configured', false); return; }
            const rows=Array.from(document.querySelectorAll('#fitupTable tbody tr'));
            const pairs=[]; const seen=new Set();
            rows.forEach(tr=>{
              if(tr.style.display==='none') return;
              const layoutCell=tr.querySelector('.col-layout');
              let layoutVal=layoutCell? (layoutCell.textContent||'').trim() : (currentLayout()||'');
              const sheetInput=tr.querySelector('input[data-name="Sheet"]')||tr.querySelector('select[data-name="Sheet"]');
              let sheetVal=sheetInput? (sheetInput.value||'').trim() : (tr.querySelector('.col-sheet')?.textContent||'').trim();
              if(!layoutVal && !sheetVal) return;
              const key=layoutVal.toUpperCase()+'|'+sheetVal.toUpperCase();
              if(!seen.has(key)){ seen.add(key); pairs.push({ Layout: layoutVal||null, Sheet: sheetVal||null }); }
            });
            if(pairs.length===0){ safeStatus('No drawings to download', false); return; }
            const payload={ ProjectId: Number(pid), Items: pairs.map(p=>({ Layout: p.Layout, Sheet: p.Sheet })) };
            setBusy(dlBtn,true);
            fetch(downloadZipBase,{ method:'POST', credentials:'same-origin', headers:{ 'Content-Type':'application/json','RequestVerificationToken': getAntiForgeryToken(), 'Accept':'application/zip' }, body: JSON.stringify(payload) })
              .then(async res=>{
                if(!res.ok){ const t=await res.text(); safeStatus(t||'Bulk download failed', false); return; }
                const blob=await res.blob(); let filename=''; const cd=res.headers.get('content-disposition');
                if(cd){ const m=/filename\*=UTF-8''([^;\n\r]+)/i.exec(cd)||/filename="?([^";]+)"?/i.exec(cd); if(m) filename=decodeURIComponent(m[1].replace(/\\u0022/g,'').trim()); }
                if(!filename) filename=`Drawings_${pid}_${new Date().toISOString().replace(/[:.]/g,'')}.zip`;
                filename=sanitizeFilename(filename);
                const obj=URL.createObjectURL(blob); const a=document.createElement('a'); a.href=obj; a.download=filename; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(obj);
                safeStatus('Download started', true);
              })
              .catch(()=> safeStatus('Failed to download ZIP', false))
              .finally(()=> setBusy(dlBtn,false));
          }
        }catch{ safeStatus('Failed to download drawing(s)', false); }
      });
      
      function resolveRowLayout(tr, btn){
        const cellText = tr?.querySelector('.col-layout')?.textContent || '';
        const dataVal = btn?.dataset?.layout || '';
        return (cellText || dataVal || '').trim();
      }
      function resolveRowSheet(tr, btn){
        const sheetInput = tr?.querySelector('input[data-name="Sheet"], select[data-name="Sheet"]');
        let val = sheetInput ? (sheetInput.value || sheetInput.textContent || '') : '';
        if(!val){
          const cellText = tr?.querySelector('.col-sheet')?.textContent || '';
          val = cellText;
        }
        if(!val && btn?.dataset?.sheet){ val = btn.dataset.sheet; }
        return (val || '').trim();
      }

      const table = document.getElementById('fitupTable');
      if(table){
        table.addEventListener('click', function(e){
          const btn = e.target?.closest('.open-dwg-row');
          if(!btn) return;
          e.preventDefault();
          const pid = currentProjectId();
          if(!pid){ safeStatus('Select Project first', false); return; }
          const tr = btn.closest('tr');
          const layout = resolveRowLayout(tr, btn);
          const sheet = resolveRowSheet(tr, btn);
          if(!layout){ safeStatus('Layout missing for this row', false); return; }
          const url = buildQueryUrl(openUrlBase, { projectId: pid, layout: layout, sheet: sheet || undefined });
          tryOpenInNewTab(url);
        });
      }
    }
    if(document.readyState==='loading') document.addEventListener('DOMContentLoaded', attachHandlers); else attachHandlers();
  }
  window.__dailyfitupDwgInit = init;
})();