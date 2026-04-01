document.addEventListener('DOMContentLoaded', () => {
    let pendingQual = null;
    let pendingQualFile = null;
    const modal = document.getElementById('confirmDeleteQualModal');
    let lastActiveElement = null;
    const modalFocusableSelector = 'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])';

    const getModalFocusables = () => {
        if (!modal) return [];
        return Array.from(modal.querySelectorAll(modalFocusableSelector)).filter(el => !el.disabled && el.offsetParent !== null);
    };

    const handleModalKeydown = (event) => {
        if (!modal || modal.getAttribute('aria-hidden') === 'true') return;
        if (event.key === 'Escape') {
            event.preventDefault();
            window.closeQualDeleteModal();
            return;
        }
        if (event.key !== 'Tab') return;
        const focusables = getModalFocusables();
        if (!focusables.length) {
            event.preventDefault();
            return;
        }
        const first = focusables[0];
        const last = focusables[focusables.length - 1];
        if (event.shiftKey && document.activeElement === first) {
            event.preventDefault();
            last.focus();
        } else if (!event.shiftKey && document.activeElement === last) {
            event.preventDefault();
            first.focus();
        }
    };

    const openModal = () => {
        if (!modal) return;
        lastActiveElement = document.activeElement;
        modal.style.display = 'flex';
        modal.setAttribute('aria-hidden', 'false');
        const focusables = getModalFocusables();
        const target = focusables[0] || modal;
        if (typeof target.focus === 'function') target.focus();
        document.addEventListener('keydown', handleModalKeydown);
    };

    const closeModal = () => {
        if (!modal) return;
        modal.style.display = 'none';
        modal.setAttribute('aria-hidden', 'true');
        document.removeEventListener('keydown', handleModalKeydown);
        if (lastActiveElement && typeof lastActiveElement.focus === 'function') {
            lastActiveElement.focus();
        }
        lastActiveElement = null;
    };

    window.closeQualDeleteModal = function(){
        closeModal();
        pendingQual = null;
        pendingQualFile = null;
    };

    window.confirmQualDelete = function(){
        // If file-delete pending, prioritize that
        if (pendingQualFile) {
            const formF = document.getElementById('deleteQualFileForm');
            if (!formF) return;
            document.getElementById('deleteFileWelderId').value = pendingQualFile.welderId;
            document.getElementById('deleteFileJcc').value = pendingQualFile.jcc;
            closeQualDeleteModal();
            formF.submit();
            return;
        }
        if (!pendingQual) return;
        const form = document.getElementById('deleteQualForm');
        if (!form) return;
        document.getElementById('deleteWelderId').value = pendingQual.welderId;
        document.getElementById('deleteJcc').value = pendingQual.jcc;
        closeQualDeleteModal();
        form.submit();
    };

    const openQualDeleteModal = (welderId, jcc) => {
        pendingQual = { welderId, jcc };
        pendingQualFile = null;
        const msgEl = document.getElementById('confirmDeleteQualMessage');
        if (msgEl) msgEl.textContent = `Are you sure you want to delete qualification "${jcc}"?`;
        openModal();
    };

    const openQualFileDeleteModal = (welderId, jcc) => {
        pendingQualFile = { welderId, jcc };
        pendingQual = null;
        const msgEl = document.getElementById('confirmDeleteQualMessage');
        if (msgEl) msgEl.textContent = `Are you sure you want to delete the qualification file for "${jcc}"?`;
        openModal();
    };

    document.querySelectorAll('.delete-qual').forEach(btn => {
        btn.addEventListener('click', () => {
            const welderId = btn.getAttribute('data-welder');
            const jcc = btn.getAttribute('data-jcc');
            if (!welderId || !jcc) return;
            openQualDeleteModal(welderId, jcc);
        });
    });

    // Use delegated click handler so dynamically inserted buttons also work
    document.addEventListener('click', function (ev) {
        try {
            const t = ev.target;
            if (!t || !t.closest) return;
            const fileBtn = t.closest('.delete-qual-file');
            if (!fileBtn) return;
            const j = fileBtn.getAttribute('data-jcc') || fileBtn.dataset.jcc;
            const w = fileBtn.getAttribute('data-welder') || fileBtn.dataset.welder;
            if (!j || !w) return;
            openQualFileDeleteModal(w, j);
        } catch (e) { /* ignore */ }
    });

    window.addEventListener('click', function (event) {
        if (event.target === modal) closeQualDeleteModal();
    });
});
