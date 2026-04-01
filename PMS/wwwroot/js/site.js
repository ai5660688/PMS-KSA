// Site.js - Complete implementation for the Contact Us form

// Safe wrappers to avoid TS deprecation diagnostics while keeping legacy behavior
function execCommandSafe(command, value) {
    const doc = /** @type {any} */ (document);
    try {
        return typeof doc.execCommand === 'function'
            ? doc.execCommand(command, false, value ?? null)
            : false;
    } catch (error) {
        console.error('Error executing command:', error);
        return false;
    }
}
function queryCommandStateSafe(command) {
    const doc = /** @type {any} */ (document);
    try {
        return typeof doc.queryCommandState === 'function'
            ? doc.queryCommandState(command)
            : false;
    } catch (_e) {
        return false;
    }
}
function queryCommandValueSafe(command) {
    const doc = /** @type {any} */ (document);
    try {
        return typeof doc.queryCommandValue === 'function'
            ? doc.queryCommandValue(command)
            : null;
    } catch {
        return null;
    }
}
function rgbToHex(rgb) {
    const m = /^rgb\((\d+),\s*(\d+),\s*(\d+)\)$/i.exec(rgb || '');
    if (!m) return null;
    const toHex = (n) => ('0' + parseInt(n, 10).toString(16)).slice(-2);
    return `#${toHex(m[1])}${toHex(m[2])}${toHex(m[3])}`;
}
function getEditor() {
    return document.getElementById('contents');
}
function isSelectionInEditor() {
    const sel = window.getSelection();
    const editor = getEditor();
    return !!(sel && sel.rangeCount > 0 && editor && editor.contains(sel.anchorNode));
}

// Keep/restore selection when toolbar controls steal focus
let savedSelection = null;
function saveSelection() {
    if (!isSelectionInEditor()) return;
    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0) {
        savedSelection = sel.getRangeAt(0).cloneRange();
    }
}
function restoreSelection() {
    if (!savedSelection) return;
    const sel = window.getSelection();
    if (!sel) return;
    sel.removeAllRanges();
    sel.addRange(savedSelection);
}
function ensureEditorSelection() {
    const editor = getEditor();
    const sel = window.getSelection();
    if (!editor) return;
    if (!sel || sel.rangeCount === 0 || !editor.contains(sel.anchorNode)) {
        const range = document.createRange();
        range.selectNodeContents(editor);
        range.collapse(false);
        sel.removeAllRanges();
        sel.addRange(range);
        saveSelection();
    }
}

// Utilities to target the current list item and list container, and map sizes
function getEnclosingListItem() {
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return null;
    const editor = getEditor();
    let node = sel.anchorNode;
    while (node && node !== editor) {
        if (node.nodeType === 1 && node.tagName === 'LI') return node;
        node = node.parentNode;
    }
    return null;
}
function getEnclosingList() {
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return null;
    const editor = getEditor();
    let node = sel.anchorNode;
    while (node && node !== editor) {
        if (node.nodeType === 1 && (node.tagName === 'UL' || node.tagName === 'OL')) return node;
        node = node.parentNode;
    }
    return null;
}
function mapFontSizeToCss(value) {
    const map = { '1': '8pt', '2': '10pt', '3': '12pt', '4': '14pt', '5': '18pt', '6': '24pt', '7': '36pt' };
    return map[String(value)] || null;
}

// Ensure lists visually align markers too
function ensureListAlignmentSupport(list) {
    if (!list) return;
    // Make markers part of the inline flow so text-align affects them
    list.style.listStylePosition = 'inside';
    // Remove default left padding that would fight centering/right alignment
    list.style.paddingLeft = '0';
    list.style.marginLeft = '0';
}

// Color list markers (bullets/numbers) by styling the LI too
function colorEnclosingListItem(value) {
    const li = getEnclosingListItem();
    if (li) li.style.color = value; // ::marker adopts color from LI
}

// Apply text color and ensure list markers match
function applyForeColor(value) {
    execCommandSafe('styleWithCSS', true);
    execCommandSafe('foreColor', value);
    colorEnclosingListItem(value);
}

// Apply toolbar state to the LI and UL/OL so list reflects color/font/size/align/bold/italic
function applyListSideEffects(command, value) {
    const li = getEnclosingListItem();
    const list = getEnclosingList();

    switch (command) {
        case 'foreColor':
            if (li) li.style.color = value;
            break;
        case 'fontName':
            if (li) li.style.fontFamily = value;
            break;
        case 'fontSize': {
            const css = mapFontSizeToCss(value);
            if (li && css) li.style.fontSize = css;
            break;
        }
        case 'justifyLeft':
        case 'justifyCenter':
        case 'justifyRight': {
            const align = command === 'justifyLeft' ? 'left' : command === 'justifyCenter' ? 'center' : 'right';
            if (list) {
                ensureListAlignmentSupport(list);
                list.style.textAlign = align;
            }
            if (li) li.style.textAlign = align;
            break;
        }
        case 'bold': {
            const on = !!queryCommandStateSafe('bold');
            if (li) li.style.fontWeight = on ? 'bold' : 'normal';
            break;
        }
        case 'italic': {
            const on = !!queryCommandStateSafe('italic');
            if (li) li.style.fontStyle = on ? 'italic' : 'normal';
            break;
        }
        case 'insertUnorderedList':
        case 'insertOrderedList': {
            if (list) ensureListAlignmentSupport(list);

            // When a list is created/toggled, sync LI with current toolbar selections
            if (li) {
                const fontSelect = document.getElementById('fontFamily');
                if (fontSelect && fontSelect.value) li.style.fontFamily = fontSelect.value;

                const sizeSelect = document.getElementById('fontSize');
                if (sizeSelect && sizeSelect.value) {
                    const css = mapFontSizeToCss(sizeSelect.value);
                    if (css) li.style.fontSize = css;
                }

                const colorPicker = document.getElementById('fontColor');
                if (colorPicker && colorPicker.value) li.style.color = colorPicker.value;

                li.style.fontWeight = queryCommandStateSafe('bold') ? 'bold' : 'normal';
                li.style.fontStyle = queryCommandStateSafe('italic') ? 'italic' : 'normal';
            }
            // Also honor current alignment state for the new list
            if (list) {
                const left = queryCommandStateSafe('justifyLeft');
                const center = queryCommandStateSafe('justifyCenter');
                const right = queryCommandStateSafe('justifyRight');
                const align = center ? 'center' : right ? 'right' : 'left';
                list.style.textAlign = align;
                if (li) li.style.textAlign = align;
            }
            break;
        }
    }
}

// Format text using document.execCommand
function formatText(command, value = null) {
    const editor = getEditor();
    if (!editor) return;

    editor.focus();
    restoreSelection();
    ensureEditorSelection();

    try {
        const cssCommands = new Set(['bold','italic','underline','fontName','fontSize','foreColor']);
        if (cssCommands.has(command)) execCommandSafe('styleWithCSS', true);

        if (command === 'foreColor' && value) {
            applyForeColor(value);
        } else {
            execCommandSafe(command, value ?? null);
        }

        applyListSideEffects(command, value);
    } catch (error) {
        console.error('Error executing command:', error);
    }

    updateActiveUIFromSelection();
    saveSelection();
}

// Helpers to set toolbar button active state
function setAlignActive(which, active) {
    const btn = document.querySelector(`button[data-command="${which}"]`);
    if (!btn) return;
    btn.classList.toggle('active', !!active);
    btn.setAttribute('aria-pressed', active ? 'true' : 'false');
}

// Sync toolbar active states and current values from selection
function updateActiveUIFromSelection() {
    const commands = ['bold', 'italic', 'underline', 'justifyLeft', 'justifyCenter', 'justifyRight'];
    commands.forEach(command => {
        const button = document.querySelector(`button[data-command="${command}"]`);
        if (!button) return;
        const isActive = !!queryCommandStateSafe(command);
        button.classList.toggle('active', isActive);
        button.setAttribute('aria-pressed', isActive ? 'true' : 'false');
    });

    // Lists toggle
    const ulBtn = document.querySelector('button[data-command="insertUnorderedList"]');
    if (ulBtn) ulBtn.classList.toggle('active', !!queryCommandStateSafe('insertUnorderedList'));
    const olBtn = document.querySelector('button[data-command="insertOrderedList"]');
    if (olBtn) olBtn.classList.toggle('active', !!queryCommandStateSafe('insertOrderedList'));

    // If selection is inside a UL/OL, reflect its text-align on the toolbar
    const li = getEnclosingListItem();
    const list = getEnclosingList();
    let align = null;
    if (li) align = getComputedStyle(li).textAlign;
    if (!align && list) align = getComputedStyle(list).textAlign;
    if (align) {
        setAlignActive('justifyLeft', align === 'left' || align === 'start');
        setAlignActive('justifyCenter', align === 'center');
        setAlignActive('justifyRight', align === 'right' || align === 'end');
    }

    // Sync color
    const colorPicker = document.getElementById('fontColor');
    if (colorPicker) {
        const val = queryCommandValueSafe('foreColor');
        const hex = (val && val.startsWith('#')) ? val : (val && val.startsWith('rgb')) ? rgbToHex(val) : null;
        if (hex) colorPicker.value = hex;
        colorPicker.style.background = colorPicker.value;
    }

    // Sync font family
    const fontSelect = document.getElementById('fontFamily');
    if (fontSelect) {
        let name = queryCommandValueSafe('fontName');
        if (typeof name === 'string') name = name.replace(/^['"]|['"]$/g, '');
        const option = Array.from(fontSelect.options).find(o => o.value.toLowerCase() === String(name || '').toLowerCase());
        if (option) fontSelect.value = option.value;
    }

    // Sync font size (legacy 1..7 only)
    const sizeSelect = document.getElementById('fontSize');
    if (sizeSelect) {
        const v = queryCommandValueSafe('fontSize');
        if (v && /^[1-7]$/.test(String(v))) sizeSelect.value = String(v);
    }
}

// Store HTML content before form submission
function setupFormHandlers() {
    const form = document.querySelector('form');
    if (form) {
        form.addEventListener('submit', function () {
            // let the browser submit to the server
            document.getElementById('contentsHtml').value =
                document.getElementById('contents').innerHTML;
        });
    }
}

// Set up toolbar event handlers
function setupToolbar() {
    // Buttons: prevent pointerdown so editor selection is preserved (mouse/touch)
    document.querySelectorAll('.format-toolbar button[data-command]').forEach(button => {
        button.addEventListener('pointerdown', e => e.preventDefault());
        const command = button.getAttribute('data-command');
        button.addEventListener('click', () => formatText(command));
        button.setAttribute('aria-pressed', 'false');
    });

    // Font family select
    const fontSelect = document.getElementById('fontFamily');
    if (fontSelect) {
        fontSelect.addEventListener('change', function () {
            restoreSelection(); ensureEditorSelection();
            formatText('fontName', this.value);
        });
    }

    // Font size select (expects 1..7)
    const sizeSelect = document.getElementById('fontSize');
    if (sizeSelect) {
        sizeSelect.addEventListener('change', function () {
            restoreSelection(); ensureEditorSelection();
            formatText('fontSize', this.value);
        });
    }

    // Color picker
    const colorPicker = document.getElementById('fontColor');
    if (colorPicker) {
        const paintPicker = (el) => { el.style.background = el.value; };
        paintPicker(colorPicker);
        colorPicker.addEventListener('input', function () {
            paintPicker(this);
            restoreSelection(); ensureEditorSelection();
            formatText('foreColor', this.value);
        });
        colorPicker.addEventListener('change', function () {
            paintPicker(this);
            restoreSelection(); ensureEditorSelection();
            formatText('foreColor', this.value);
        });
    }
}

// Initialize the editor and event handlers
function initEditor() {
    const editor = getEditor();
    if (editor) {
        editor.setAttribute('contenteditable', 'true');
        editor.setAttribute('tabindex', '0');
        editor.addEventListener('input', () => { updateActiveUIFromSelection(); saveSelection(); });
        editor.addEventListener('click', () => { updateActiveUIFromSelection(); saveSelection(); });
        editor.addEventListener('keyup', () => { updateActiveUIFromSelection(); saveSelection(); });
        editor.addEventListener('mouseup', () => { updateActiveUIFromSelection(); saveSelection(); });
        document.addEventListener('selectionchange', () => {
            if (isSelectionInEditor()) { updateActiveUIFromSelection(); saveSelection(); }
        });

        // Removed default placeholder text
        ensureEditorSelection();
    }
    updateActiveUIFromSelection();
}

// Dashboard menu setup
function setupDashboardMenus() {
    const dropdowns = Array.from(document.querySelectorAll('.dropdown'));

    // Toggle top-level dropdowns on click (mobile/touch)
    dropdowns.forEach(dd => {
        const toggle = dd.querySelector('.dropdown-toggle');
        if (!toggle) return;
        toggle.addEventListener('click', e => {
            e.preventDefault();
            // close other dropdowns
            dropdowns.forEach(other => { if (other !== dd) other.classList.remove('active'); });
            dd.classList.toggle('active');
        });
    });

    // Toggle submenus on click
    document.querySelectorAll('.submenu-parent > .has-submenu').forEach(link => {
        link.addEventListener('click', e => {
            e.preventDefault();
            const parent = link.closest('.submenu-parent');
            if (!parent) return;
            // close sibling submenus
            parent.parentElement.querySelectorAll('.submenu-parent').forEach(sib => {
                if (sib !== parent) sib.classList.remove('active');
            });
            parent.classList.toggle('active');
        });
    });

    // Close menus when clicking outside
    document.addEventListener('click', e => {
        if (!(e.target instanceof Element)) return;
        if (!e.target.closest('.dropdown')) {
            dropdowns.forEach(d => d.classList.remove('active'));
            document.querySelectorAll('.submenu-parent').forEach(s => s.classList.remove('active'));
        }
    });

    // Escape to close
    document.addEventListener('keydown', e => {
        if (e.key === 'Escape') {
            dropdowns.forEach(d => d.classList.remove('active'));
            document.querySelectorAll('.submenu-parent').forEach(s => s.classList.remove('active'));
        }
    });
}

// Initialize everything when DOM is loaded
document.addEventListener('DOMContentLoaded', function () {
    setupFormHandlers();
    setupToolbar();
    initEditor();
    setupDashboardMenus();
});

// Utility helpers
function clearEditor() {
    getEditor().innerHTML = '';
    updateActiveUIFromSelection();
}
function getEditorContent() {
    return getEditor().innerHTML;
}
function setEditorContent(html) {
    getEditor().innerHTML = html;
    updateActiveUIFromSelection();
}