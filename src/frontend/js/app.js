// Main application controller

function setPageTitle(page) {
    const titles = {
        dashboard: 'Dashboard',
        workers: 'Ouvriers',
        presence: 'Présences',
        documents: 'Documents',
        settings: 'Paramètres'
    };
    const label = titles[page] || 'Dashboard';
    const titleEl = document.getElementById('page-title');
    if (titleEl) titleEl.textContent = label;
}

function toggleSidebar() {
    document.body.classList.toggle('sidebar-collapsed');
}

function initTheme() {
    document.body.classList.remove('dark');
    localStorage.removeItem('theme');
}

function initDashboardCounters() {
    const docsCount = Number(localStorage.getItem('docs-generated') || 0);
    const dashDocs = document.getElementById('dash-docs-count');
    if (dashDocs) dashDocs.textContent = docsCount;
}

// Page navigation
function setupNavigation() {
    document.querySelectorAll('.nav-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const page = btn.dataset.page;

            document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');

            document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
            const pageEl = document.getElementById(`${page}-page`);
            if (pageEl) pageEl.classList.add('active');

            setPageTitle(page);

            if (page === 'presence' && typeof loadWorkersForPresence === 'function') {
                loadWorkersForPresence();
            } else if (page === 'documents' && typeof loadWorkersForDocuments === 'function') {
                loadWorkersForDocuments();
            }
        });
    });
}

function setupHeaderActions() {
    const toggleBtn = document.getElementById('sidebar-toggle');
    if (toggleBtn) toggleBtn.addEventListener('click', toggleSidebar);

    const quickDocs = document.getElementById('quick-generate-docs');
    if (quickDocs) {
        quickDocs.addEventListener('click', () => {
            document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
            const docsBtn = document.querySelector('.nav-btn[data-page="documents"]');
            if (docsBtn) docsBtn.classList.add('active');
            document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
            const docsPage = document.getElementById('documents-page');
            if (docsPage) docsPage.classList.add('active');
            setPageTitle('documents');
            if (typeof loadWorkersForDocuments === 'function') {
                loadWorkersForDocuments();
            }
        });
    }

    const headerSearch = document.getElementById('header-search');
    if (headerSearch) {
        headerSearch.addEventListener('keyup', (e) => {
            if (e.key === 'Enter') {
                const workerSearch = document.getElementById('worker-search');
                if (workerSearch) {
                    workerSearch.value = headerSearch.value;
                    const workersBtn = document.querySelector('.nav-btn[data-page=\"workers\"]');
                    if (workersBtn) workersBtn.click();
                    if (typeof searchWorkers === 'function') {
                        searchWorkers();
                    }
                }
            }
        });
    }
}

function setupShortcuts() {
    document.addEventListener('keydown', (e) => {
        if (e.ctrlKey && e.key.toLowerCase() === 'k') {
            e.preventDefault();
            const headerSearch = document.getElementById('header-search');
            if (headerSearch) headerSearch.focus();
        }
        if (e.ctrlKey && e.key.toLowerCase() === 'n') {
            e.preventDefault();
            const addBtn = document.getElementById('add-worker-btn');
            if (addBtn) addBtn.click();
        }
    });
}

function setupResponsiveSidebar() {
    // Keep sidebar open by default; user can still toggle manually.
    document.body.classList.remove('sidebar-collapsed');
}

// Loading overlay utilities
function showLoading() {
    const overlay = document.getElementById('loading-overlay');
    if (overlay) overlay.style.display = 'flex';
}

function hideLoading() {
    const overlay = document.getElementById('loading-overlay');
    if (overlay) overlay.style.display = 'none';
}

// Notification utilities
function showNotification(message, type = 'success') {
    const notification = document.createElement('div');
    notification.className = `toast ${type}`;
    notification.textContent = message;

    const container = document.getElementById('toast-container') || document.body;
    container.appendChild(notification);

    setTimeout(() => {
        notification.classList.add('hide');
        setTimeout(() => notification.remove(), 300);
    }, 3000);
}

// Initialize
initTheme();
initDashboardCounters();
setPageTitle('dashboard');
setupNavigation();
setupHeaderActions();
setupShortcuts();
setupResponsiveSidebar();

// Export utilities
window.appUtils = {
    showLoading,
    hideLoading,
    showNotification
};
