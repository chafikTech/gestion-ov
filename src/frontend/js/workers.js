// Workers page management

let currentWorkerId = null;
let isEditMode = false;

// Load workers on page load
// Since scripts are loaded at the end of body, DOM is already ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
        loadWorkers();
        setupEventListeners();
    });
} else {
    // DOM is already loaded, execute immediately
    loadWorkers();
    setupEventListeners();
}

// Setup event listeners
function setupEventListeners() {
    // Add worker button
    document.getElementById('add-worker-btn').addEventListener('click', () => {
        openModal(false);
    });

    // Search
    document.getElementById('search-btn').addEventListener('click', searchWorkers);
    document.getElementById('worker-search').addEventListener('keyup', (e) => {
        if (e.key === 'Enter') searchWorkers();
    });
    document.getElementById('clear-search-btn').addEventListener('click', () => {
        document.getElementById('worker-search').value = '';
        loadWorkers();
    });

    // Modal
    document.getElementById('close-modal').addEventListener('click', closeModal);
    document.getElementById('cancel-btn').addEventListener('click', closeModal);
    document.getElementById('worker-modal').addEventListener('click', (e) => {
        if (e.target.id === 'worker-modal') closeModal();
    });

    // Form submit
    document.getElementById('worker-form').addEventListener('submit', handleFormSubmit);
}

// Load all workers
async function loadWorkers() {
    try {
        window.appUtils.showLoading();
        const workers = await window.api.workers.getAll();
        displayWorkers(workers);
        updateStatistics(workers);
    } catch (error) {
        console.error('Error loading workers:', error);
        window.appUtils.showNotification('Erreur lors du chargement des ouvriers', 'error');
    } finally {
        window.appUtils.hideLoading();
    }
}

// Search workers
async function searchWorkers() {
    const query = document.getElementById('worker-search').value.trim();
    if (!query) {
        loadWorkers();
        return;
    }

    try {
        window.appUtils.showLoading();
        const workers = await window.api.workers.search(query);
        displayWorkers(workers);
        updateStatistics(workers);
    } catch (error) {
        console.error('Error searching workers:', error);
        window.appUtils.showNotification('Erreur lors de la recherche', 'error');
    } finally {
        window.appUtils.hideLoading();
    }
}

// Display workers in table
function displayWorkers(workers) {
    const tbody = document.getElementById('workers-tbody');
    tbody.innerHTML = '';

    if (workers.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="7" style="text-align: center; padding: 40px; color: var(--text-secondary);">
                    Aucun ouvrier trouvé
                </td>
            </tr>
        `;
        return;
    }

    workers.forEach(worker => {
        const age = calculateAge(worker.date_naissance);
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${escapeHtml(worker.nom_prenom)}</td>
            <td>${escapeHtml(worker.cin)}</td>
            <td>${formatDate(worker.date_naissance)}</td>
            <td>${age} ans</td>
            <td><span class="badge badge-${worker.type.toLowerCase()}">${worker.type}</span></td>
            <td>${worker.salaire_journalier.toFixed(2)} DH</td>
            <td class="actions-cell">
                <button class="btn btn-small btn-secondary" onclick="editWorker(${worker.id})">✏️ Modifier</button>
                <button class="btn btn-small btn-danger" onclick="deleteWorker(${worker.id}, '${escapeHtml(worker.nom_prenom)}')">🗑️ Supprimer</button>
            </td>
        `;
        tbody.appendChild(row);
    });
}

// Update statistics
function updateStatistics(workers) {
    const total = workers.length;
    const os = workers.filter(w => w.type === 'OS').length;
    const ons = workers.filter(w => w.type === 'ONS').length;

    document.getElementById('total-workers').textContent = total;
    document.getElementById('os-workers').textContent = os;
    document.getElementById('ons-workers').textContent = ons;

    const dashTotal = document.getElementById('dash-total-workers');
    if (dashTotal) {
        dashTotal.textContent = total;
    }
}

// Open modal for add/edit
function openModal(editMode = false, worker = null) {
    isEditMode = editMode;
    const modal = document.getElementById('worker-modal');
    const title = document.getElementById('modal-title');
    const form = document.getElementById('worker-form');

    if (editMode && worker) {
        title.textContent = 'Modifier un Ouvrier';
        currentWorkerId = worker.id;
        document.getElementById('worker-nom-prenom').value = worker.nom_prenom;
        document.getElementById('worker-cin').value = worker.cin;
        document.getElementById('worker-cin-validite').value = worker.cin_validite || '';
        document.getElementById('worker-date-naissance').value = worker.date_naissance;
        document.getElementById('worker-type').value = worker.type;
    } else {
        title.textContent = 'Ajouter un Ouvrier';
        currentWorkerId = null;
        form.reset();
    }

    modal.classList.add('show');
}

// Close modal
function closeModal() {
    const modal = document.getElementById('worker-modal');
    modal.classList.remove('show');
    document.getElementById('worker-form').reset();
    currentWorkerId = null;
}

// Handle form submit
async function handleFormSubmit(e) {
    e.preventDefault();

    const workerData = {
        nom_prenom: document.getElementById('worker-nom-prenom').value.trim(),
        cin: document.getElementById('worker-cin').value.trim(),
        cin_validite: document.getElementById('worker-cin-validite').value || null,
        date_naissance: document.getElementById('worker-date-naissance').value,
        type: document.getElementById('worker-type').value
    };

    try {
        window.appUtils.showLoading();

        if (isEditMode) {
            await window.api.workers.update(currentWorkerId, workerData);
            window.appUtils.showNotification('Ouvrier modifié avec succès');
        } else {
            await window.api.workers.create(workerData);
            window.appUtils.showNotification('Ouvrier ajouté avec succès');
        }

        closeModal();
        loadWorkers();
        
        // Refresh workers in other pages
        if (typeof loadWorkersForPresence === 'function') {
            loadWorkersForPresence();
        }
        if (typeof loadWorkersForDocuments === 'function') {
            loadWorkersForDocuments();
        }
    } catch (error) {
        console.error('Error saving worker:', error);
        window.appUtils.showNotification(error.message || 'Erreur lors de l\'enregistrement', 'error');
    } finally {
        window.appUtils.hideLoading();
    }
}

// Edit worker
async function editWorker(id) {
    try {
        window.appUtils.showLoading();
        const worker = await window.api.workers.getById(id);
        openModal(true, worker);
    } catch (error) {
        console.error('Error loading worker:', error);
        window.appUtils.showNotification('Erreur lors du chargement de l\'ouvrier', 'error');
    } finally {
        window.appUtils.hideLoading();
    }
}

// Delete worker
async function deleteWorker(id, name) {
    if (!confirm(`Êtes-vous sûr de vouloir supprimer l'ouvrier "${name}" ?\n\nCette action supprimera également toutes les données de présence associées.`)) {
        return;
    }

    try {
        window.appUtils.showLoading();
        await window.api.workers.delete(id);
        window.appUtils.showNotification('Ouvrier supprimé avec succès');
        loadWorkers();
        
        // Refresh workers in other pages
        if (typeof loadWorkersForPresence === 'function') {
            loadWorkersForPresence();
        }
        if (typeof loadWorkersForDocuments === 'function') {
            loadWorkersForDocuments();
        }
    } catch (error) {
        console.error('Error deleting worker:', error);
        window.appUtils.showNotification('Erreur lors de la suppression', 'error');
    } finally {
        window.appUtils.hideLoading();
    }
}

// Utility functions
function calculateAge(dateNaissance) {
    const birthDate = new Date(dateNaissance);
    const today = new Date();
    let age = today.getFullYear() - birthDate.getFullYear();
    const monthDiff = today.getMonth() - birthDate.getMonth();
    
    if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birthDate.getDate())) {
        age--;
    }
    
    return age;
}

function formatDate(dateString) {
    const date = new Date(dateString);
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
}

function escapeHtml(text) {
    const map = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
    };
    return text.replace(/[&<>"']/g, m => map[m]);
}

// Add CSS for badges
const badgeStyle = document.createElement('style');
badgeStyle.textContent = `
    .badge {
        padding: 4px 12px;
        border-radius: 12px;
        font-size: 12px;
        font-weight: 600;
        display: inline-block;
        background: var(--primary-soft);
        color: var(--primary);
        border: 1px solid var(--primary-border);
    }
    .badge-os {
        background: var(--success-soft);
        color: var(--success);
        border: 1px solid var(--success-border);
    }
    .badge-ons {
        background: var(--warning-soft);
        color: var(--warning);
        border: 1px solid var(--warning-border);
    }
`;
document.head.appendChild(badgeStyle);

// Make functions globally available
window.editWorker = editWorker;
window.deleteWorker = deleteWorker;
