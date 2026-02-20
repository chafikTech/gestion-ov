// Presence management page

let currentPresenceWorkerId = null;
let currentPresenceYear = null;
let currentPresenceMonth = null;
let selectedDays = new Set();

// Initialize presence page
// Since scripts are loaded at the end of body, DOM is already ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', async () => {
        setupPresenceEventListeners();
        setCurrentMonth();
        await loadWorkersForPresence();
    });
} else {
    // DOM is already loaded, execute immediately
    setupPresenceEventListeners();
    setCurrentMonth();
    loadWorkersForPresence();
}

function getSelectedPeriod() {
    const year = parseInt(document.getElementById('presence-year').value, 10);
    const month = parseInt(document.getElementById('presence-month').value, 10);
    return { year, month };
}

function normalizeAttachmentValue(value) {
    if (value === null || value === undefined || value === '') return null;
    const n = Number(value);
    if (!Number.isFinite(n)) return null;
    const asInt = Math.trunc(n);
    return asInt >= 0 ? asInt : null;
}

// Setup event listeners
function setupPresenceEventListeners() {
    document.getElementById('load-presence-btn').addEventListener('click', loadPresence);
    document.getElementById('save-presence-btn').addEventListener('click', savePresence);

    // Update period-dependent values and worker ordering
    document.getElementById('presence-month').addEventListener('change', handlePresencePeriodChange);
    document.getElementById('presence-year').addEventListener('change', handlePresencePeriodChange);

    const firstInput = document.getElementById('presence-days-input-first');
    const secondInput = document.getElementById('presence-days-input-second');
    if (firstInput) {
        firstInput.addEventListener('input', syncDaysFromInputs);
        firstInput.addEventListener('blur', normalizeDaysInputs);
    }
    if (secondInput) {
        secondInput.addEventListener('input', syncDaysFromInputs);
        secondInput.addEventListener('blur', normalizeDaysInputs);
    }
}

function handlePresencePeriodChange() {
    currentPresenceWorkerId = null;
    currentPresenceYear = null;
    currentPresenceMonth = null;
    selectedDays = new Set();
    document.getElementById('worker-info').style.display = 'none';
    document.getElementById('calendar-container').style.display = 'none';
    document.getElementById('presence-stats').style.display = 'none';
    updateLastDay();
    loadWorkersForPresence();
}

function renderAttachOrderTable(workers) {
    const tbody = document.getElementById('presence-attach-tbody');
    if (!tbody) return;

    if (!Array.isArray(workers) || workers.length === 0) {
        tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;color:var(--text-secondary);">Aucun ouvrier</td></tr>';
        return;
    }

    tbody.innerHTML = workers.map(worker => {
        const attachmentValue = worker.attachmentNumber === null || worker.attachmentNumber === undefined
            ? ''
            : String(worker.attachmentNumber);

        return `
            <tr data-worker-id="${worker.id}">
                <td>${worker.nom_prenom || ''}</td>
                <td>${worker.cin || ''}</td>
                <td>${worker.type || ''}</td>
                <td>
                    <input
                        type="number"
                        min="0"
                        step="1"
                        class="form-input presence-attach-input"
                        data-worker-id="${worker.id}"
                        value="${attachmentValue}"
                        placeholder="-"
                    >
                </td>
            </tr>
        `;
    }).join('');

    tbody.querySelectorAll('.presence-attach-input').forEach((input) => {
        input.addEventListener('blur', () => {
            const normalized = normalizeAttachmentValue(input.value);
            input.value = normalized === null ? '' : String(normalized);
        });
    });
}

function collectMonthlyAttachmentEntries() {
    const inputs = document.querySelectorAll('.presence-attach-input');
    return Array.from(inputs).map((input) => ({
        workerId: parseInt(input.dataset.workerId, 10),
        attachmentNumber: normalizeAttachmentValue(input.value)
    }));
}

// Load workers dropdown (ordered by N° F. d'attach for selected month)
async function loadWorkersForPresence() {
    const select = document.getElementById('presence-worker');
    const previousSelected = currentPresenceWorkerId || parseInt(select.value, 10) || null;

    try {
        const { year, month } = getSelectedPeriod();
        const workers = await window.api.presence.getWorkersForMonth(year, month);

        select.innerHTML = '<option value="">Sélectionner un ouvrier...</option>';

        workers.forEach(worker => {
            const option = document.createElement('option');
            option.value = worker.id;
            const attach = worker.attachmentNumber === null || worker.attachmentNumber === undefined
                ? ''
                : ` • N° ${worker.attachmentNumber}`;
            option.textContent = `${worker.nom_prenom} (${worker.cin})${attach}`;
            option.dataset.type = worker.type;
            option.dataset.salary = worker.salaire_journalier;
            option.dataset.attachmentNumber = worker.attachmentNumber ?? '';
            select.appendChild(option);
        });

        if (previousSelected && workers.some(w => Number(w.id) === Number(previousSelected))) {
            select.value = String(previousSelected);
        }

        renderAttachOrderTable(workers);
    } catch (error) {
        console.error('Error loading workers:', error);
        renderAttachOrderTable([]);
    }
}

// Set current month and year
function setCurrentMonth() {
    const now = new Date();
    document.getElementById('presence-month').value = now.getMonth() + 1;
    document.getElementById('presence-year').value = now.getFullYear();
    updateLastDay();
}

// Update last day display
function updateLastDay() {
    const { year, month } = getSelectedPeriod();
    const daysInMonth = new Date(year, month, 0).getDate();
    const maxDaysEl = document.getElementById('presence-max-days');
    if (maxDaysEl) maxDaysEl.textContent = daysInMonth;
    const firstMaxEl = document.getElementById('presence-max-days-first');
    if (firstMaxEl) firstMaxEl.textContent = 15;
    const secondMaxEl = document.getElementById('presence-max-days-second');
    if (secondMaxEl) secondMaxEl.textContent = Math.max(0, daysInMonth - 15);
    normalizeDaysInputs();
}

// Load presence data
async function loadPresence() {
    const workerId = document.getElementById('presence-worker').value;
    const { year, month } = getSelectedPeriod();

    if (!workerId) {
        window.appUtils.showNotification('Veuillez sélectionner un ouvrier', 'error');
        return;
    }

    try {
        window.appUtils.showLoading();

        // Get worker info
        const worker = await window.api.workers.getById(parseInt(workerId, 10));
        displayWorkerInfo(worker);

        // Get presence data
        const presenceDays = await window.api.presence.get(parseInt(workerId, 10), year, month);
        selectedDays = new Set(presenceDays);

        // Store current selection
        currentPresenceWorkerId = parseInt(workerId, 10);
        currentPresenceYear = year;
        currentPresenceMonth = month;

        // Sync inputs to loaded data
        syncDaysToInputs();

        // Calculate and show stats
        updatePresenceStats(worker.salaire_journalier);

        // Show calendar and stats
        document.getElementById('worker-info').style.display = 'block';
        document.getElementById('calendar-container').style.display = 'block';
        document.getElementById('presence-stats').style.display = 'grid';

    } catch (error) {
        console.error('Error loading presence:', error);
        window.appUtils.showNotification('Erreur lors du chargement des présences', 'error');
    } finally {
        window.appUtils.hideLoading();
    }
}

// Display worker info
function displayWorkerInfo(worker) {
    document.getElementById('worker-info-name').textContent = worker.nom_prenom;
    document.getElementById('worker-info-type').textContent = worker.type === 'OS' ? 'Ouvrier Spécialisé (OS)' : 'Ouvrier Non Spécialisé (ONS)';
    document.getElementById('worker-info-salary').textContent = worker.salaire_journalier.toFixed(2);
}

function clampDaysCount(value, max) {
    return Math.max(0, Math.min(max, parseInt(value, 10) || 0));
}

function getMonthHalves(year, month) {
    const daysInMonth = new Date(year, month, 0).getDate();
    const firstHalfMax = 15;
    const secondHalfMax = Math.max(0, daysInMonth - 15);
    return { daysInMonth, firstHalfMax, secondHalfMax };
}

function syncDaysFromInputs() {
    const firstInput = document.getElementById('presence-days-input-first');
    const secondInput = document.getElementById('presence-days-input-second');
    if (!firstInput || !secondInput) return;

    const { year, month } = getSelectedPeriod();
    const { firstHalfMax, secondHalfMax } = getMonthHalves(year, month);
    const firstCount = clampDaysCount(firstInput.value, firstHalfMax);
    const secondCount = clampDaysCount(secondInput.value, secondHalfMax);

    const days = [];
    for (let day = 1; day <= firstCount; day++) days.push(day);
    for (let day = 16; day < 16 + secondCount; day++) days.push(day);
    selectedDays = new Set(days);

    const salaryEl = document.getElementById('worker-info-salary');
    const dailySalary = salaryEl ? parseFloat((salaryEl.textContent || '0').replace(',', '.')) : 0;
    updatePresenceStats(dailySalary);
}

function syncDaysToInputs() {
    const firstInput = document.getElementById('presence-days-input-first');
    const secondInput = document.getElementById('presence-days-input-second');
    if (!firstInput || !secondInput) return;

    const firstCount = Array.from(selectedDays).filter(day => day >= 1 && day <= 15).length;
    const secondCount = Array.from(selectedDays).filter(day => day >= 16).length;
    firstInput.value = String(firstCount);
    secondInput.value = String(secondCount);
}

function normalizeDaysInputs() {
    const firstInput = document.getElementById('presence-days-input-first');
    const secondInput = document.getElementById('presence-days-input-second');
    if (!firstInput || !secondInput) return;

    const { year, month } = getSelectedPeriod();
    const { firstHalfMax, secondHalfMax } = getMonthHalves(year, month);

    const firstCount = clampDaysCount(firstInput.value, firstHalfMax);
    const secondCount = clampDaysCount(secondInput.value, secondHalfMax);

    firstInput.value = String(firstCount);
    secondInput.value = String(secondCount);
    syncDaysFromInputs();
}

// Update presence statistics
function updatePresenceStats(dailySalary) {
    const daysWorked = selectedDays.size;
    const totalSalary = daysWorked * dailySalary;

    document.getElementById('days-worked').textContent = daysWorked;
    document.getElementById('total-salary').textContent = totalSalary.toFixed(2) + ' DH';
    const countEl = document.getElementById('presence-days-count');
    if (countEl) countEl.textContent = daysWorked;
}

// Save presence
async function savePresence() {
    if (!currentPresenceWorkerId) {
        window.appUtils.showNotification('Veuillez d\'abord charger les données d\'un ouvrier', 'error');
        return;
    }

    try {
        window.appUtils.showLoading();

        const daysArray = Array.from(selectedDays).sort((a, b) => a - b);
        const monthlyAttachmentEntries = collectMonthlyAttachmentEntries();
        const currentWorkerAttachment = monthlyAttachmentEntries.find(
            entry => entry.workerId === currentPresenceWorkerId
        )?.attachmentNumber ?? null;

        await window.api.presence.save(
            currentPresenceWorkerId,
            currentPresenceYear,
            currentPresenceMonth,
            daysArray,
            currentWorkerAttachment,
            monthlyAttachmentEntries
        );

        await loadWorkersForPresence();
        document.getElementById('presence-worker').value = String(currentPresenceWorkerId);

        window.appUtils.showNotification('Présences enregistrées avec succès');

    } catch (error) {
        console.error('Error saving presence:', error);
        window.appUtils.showNotification('Erreur lors de l\'enregistrement des présences', 'error');
    } finally {
        window.appUtils.hideLoading();
    }
}
