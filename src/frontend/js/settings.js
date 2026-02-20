// Settings page: database backup/restore

function setInputValue(id, value) {
    const input = document.getElementById(id);
    if (!input) return;
    input.value = value == null ? '' : String(value);
}

function getInputValue(id) {
    const input = document.getElementById(id);
    return input ? String(input.value || '').trim() : '';
}

async function loadAdministrativeSettings() {
    if (!window.api?.settings?.get) {
        return;
    }
    try {
        const settings = await window.api.settings.get();
        setInputValue('settings-rcar', settings?.rcarAdhesionNumber);
        setInputValue('settings-chap', settings?.chap);
        setInputValue('settings-art', settings?.art);
        setInputValue('settings-prog', settings?.prog);
        setInputValue('settings-proj', settings?.proj);
        setInputValue('settings-ligne', settings?.ligne);
        setInputValue('settings-rcar-age-limit', settings?.rcarAgeLimit);
        setInputValue('settings-decision-number', settings?.decisionNumber);
        setInputValue('settings-decision-date', settings?.decisionDate);
    } catch (error) {
        console.error('Load settings error:', error);
        window.appUtils?.showNotification?.('Erreur lors du chargement des paramètres', 'error');
    }
}

function setupAdministrativeSettingsSaveButton() {
    const saveBtn = document.getElementById('settings-save');
    if (!saveBtn || !window.api?.settings?.save) {
        return;
    }

    saveBtn.addEventListener('click', async () => {
        try {
            const ageLimitRaw = getInputValue('settings-rcar-age-limit');
            const payload = {
                rcarAdhesionNumber: getInputValue('settings-rcar') || null,
                chap: getInputValue('settings-chap') || null,
                art: getInputValue('settings-art') || null,
                prog: getInputValue('settings-prog') || null,
                proj: getInputValue('settings-proj') || null,
                ligne: getInputValue('settings-ligne') || null,
                rcarAgeLimit: ageLimitRaw ? Number.parseInt(ageLimitRaw, 10) : null,
                decisionNumber: getInputValue('settings-decision-number') || null,
                decisionDate: getInputValue('settings-decision-date') || null
            };

            window.appUtils.showLoading();
            await window.api.settings.save(payload);
            window.appUtils.showNotification('Paramètres enregistrés avec succès');
            await loadAdministrativeSettings();
        } catch (error) {
            console.error('Save settings error:', error);
            window.appUtils.showNotification(error.message || 'Erreur lors de l’enregistrement', 'error');
        } finally {
            window.appUtils.hideLoading();
        }
    });
}

function setupDatabaseBackupButtons() {
    const exportBtn = document.getElementById('db-export-btn');
    const importBtn = document.getElementById('db-import-btn');

    if (exportBtn) {
        exportBtn.addEventListener('click', async () => {
            try {
                const filePath = await window.api.dialog.saveFile({
                    title: 'Exporter la base de données',
                    defaultPath: 'gestion_ouvriers.db',
                    filters: [{ name: 'SQLite Database', extensions: ['db'] }]
                });

                if (!filePath) {
                    window.appUtils.showNotification('Export annulé', 'error');
                    return;
                }

                window.appUtils.showLoading();
                await window.api.database.export(filePath);
                window.appUtils.showNotification('Base exportée avec succès');
            } catch (error) {
                console.error('Export DB error:', error);
                window.appUtils.showNotification('Erreur lors de l’export', 'error');
            } finally {
                window.appUtils.hideLoading();
            }
        });
    }

    if (importBtn) {
        importBtn.addEventListener('click', async () => {
            if (!confirm('Importer une base remplacera la base actuelle. Continuer ?')) {
                return;
            }

            try {
                const filePath = await window.api.dialog.openFile({
                    title: 'Importer la base de données',
                    filters: [{ name: 'SQLite Database', extensions: ['db'] }]
                });

                if (!filePath) {
                    window.appUtils.showNotification('Import annulé', 'error');
                    return;
                }

                window.appUtils.showLoading();
                await window.api.database.import(filePath);
                window.appUtils.showNotification('Base importée. Redémarrage...');
            } catch (error) {
                console.error('Import DB error:', error);
                window.appUtils.showNotification('Erreur lors de l’import', 'error');
                window.appUtils.hideLoading();
            }
        });
    }
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
        loadAdministrativeSettings();
        setupAdministrativeSettingsSaveButton();
        setupDatabaseBackupButtons();
    });
} else {
    loadAdministrativeSettings();
    setupAdministrativeSettingsSaveButton();
    setupDatabaseBackupButtons();
}
