// Document generation page

function getResultFileList(result) {
    if (!result) return '';
    if (Array.isArray(result.results)) {
        return result.results.map(r => {
            const name = r.docxFileName || r.fileName || r.pdfFileName || '';
            return `  • ${name}`.trimEnd();
        }).join('\n');
    }
    if (Array.isArray(result.files)) {
        return result.files.map(f => `  • ${f}`).join('\n');
    }
    if (result.docxFileName) {
        return `  • ${result.docxFileName}`;
    }
    if (result.fileName) {
        return `  • ${result.fileName}`;
    }
    return '';
}

function getSelectedDocumentsFormat() {
    return 'docx';
}

function formatBatchErrorLine(err) {
    if (!err || typeof err !== 'object') {
        return 'Erreur inconnue';
    }

    const message = String(err.error || err.message || '').trim() || 'Erreur inconnue';
    const worker = String(err.worker || '').trim();
    return worker ? `${worker}: ${message}` : message;
}

// Initialize documents page
// Since scripts are loaded at the end of body, DOM is already ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
        setupDocumentEventListeners();
        loadWorkersForDocuments();
    });
} else {
    // DOM is already loaded, execute immediately
    setupDocumentEventListeners();
    loadWorkersForDocuments();
}

// Setup event listeners
function setupDocumentEventListeners() {
    // Period type radio buttons
    document.querySelectorAll('input[name="doc-period"]').forEach(radio => {
        radio.addEventListener('change', (e) => {
            if (e.target.value === 'monthly') {
                document.getElementById('monthly-docs-section').style.display = 'block';
                document.getElementById('quarterly-docs-section').style.display = 'none';
                loadMonthlyWorkers(); // Load workers when switching to monthly
            } else {
                document.getElementById('monthly-docs-section').style.display = 'none';
                document.getElementById('quarterly-docs-section').style.display = 'block';
                loadQuarterlyWorkers(); // Load workers when switching to quarterly
            }
        });
    });
    
    // Auto-load workers when month/year changes
    document.getElementById('monthly-month').addEventListener('change', loadMonthlyWorkers);
    document.getElementById('monthly-year').addEventListener('change', loadMonthlyWorkers);
    
    // Auto-load workers when quarter/year changes
    document.getElementById('quarterly-quarter').addEventListener('change', loadQuarterlyWorkers);
    document.getElementById('quarterly-year').addEventListener('change', loadQuarterlyWorkers);

    // Monthly document buttons
    document.querySelectorAll('#monthly-docs-section .doc-btn:not(.btn-generate-all)').forEach(btn => {
        btn.addEventListener('click', () => generateMonthlyDocument(btn.dataset.doc));
    });

    const generateAllMonthlyBtn = document.getElementById('generate-all-monthly');
    if (generateAllMonthlyBtn) {
        generateAllMonthlyBtn.addEventListener('click', generateAllMonthlyDocuments);
    }

    // Quarterly document buttons
    document.querySelectorAll('#quarterly-docs-section .doc-btn:not(.btn-generate-all)').forEach(btn => {
        btn.addEventListener('click', () => generateQuarterlyDocument(btn.dataset.doc));
    });

    const generateAllQuarterlyBtn = document.getElementById('generate-all-quarterly');
    if (generateAllQuarterlyBtn) {
        generateAllQuarterlyBtn.addEventListener('click', generateAllQuarterlyDocuments);
    }
    
    // Attach event listeners to individual document buttons
    document.querySelectorAll('.doc-btn-individual').forEach(btn => {
        btn.addEventListener('click', function(event) {
            const docType = btn.dataset.doc;
            console.log('🔔 Document button clicked:', docType);
            
            if (docType === 'all') {
                generateAllMonthlyDocumentsNew();
            } else if (docType === 'depense-regie-salaire-combined') {
                generateCombinedMonthly();
            } else if (docType === 'recu-combined') {
                generateRecuCombined();
            } else {
                generateSingleMonthlyDocument(docType);
            }
        });
    });
    
    const quarterlyCombinedBtn = document.getElementById('generate-combined-quarterly');
    console.log('Looking for quarterly button...', quarterlyCombinedBtn);
    if (quarterlyCombinedBtn) {
        console.log('✓ Quarterly button found, attaching event listener');
        quarterlyCombinedBtn.addEventListener('click', function(event) {
            console.log('🔔 QUARTERLY BUTTON CLICKED!', event);
            generateCombinedQuarterly();
        }, false);
    } else {
        console.error('✗ Quarterly button NOT found in DOM!');
    }
}

// Load workers for document generation (REMOVED - no longer needed)
async function loadWorkersForDocuments() {
    // This function is kept for compatibility but does nothing
    // Workers are now auto-loaded based on presence data
}

// Load workers with presence for selected month
async function loadMonthlyWorkers() {
    const month = parseInt(document.getElementById('monthly-month').value);
    const year = parseInt(document.getElementById('monthly-year').value);
    
    try {
        const workers = await window.api.documents.getAllWorkers(year, month);
        
        const infoDiv = document.getElementById('monthly-workers-info');
        const listDiv = document.getElementById('monthly-workers-list');
        
        if (workers.length === 0) {
            infoDiv.style.display = 'block';
            listDiv.innerHTML = '<span style="color: var(--danger);">❌ Aucun ouvrier avec présence pour cette période</span>';
        } else {
            infoDiv.style.display = 'block';
            const totalDays = workers.reduce((sum, w) => sum + (w.daysWorked || 0), 0);
            const totalSalary = workers.reduce((sum, w) => sum + (w.totalSalary || 0), 0);
            
            listDiv.innerHTML = `
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <div>
                        <strong>${workers.length}</strong> ouvrier(s) • 
                        <strong>${totalDays}</strong> jours travaillés • 
                        Total: <strong style="color: var(--success);">${totalSalary.toFixed(2)} DH</strong>
                    </div>
                </div>
                <div style="margin-top: 10px; max-height: 150px; overflow-y: auto;">
                    ${workers.map(w => `
                        <div style="padding: 5px; border-bottom: 1px solid var(--border);">
                            👤 ${w.nom_prenom} (${w.cin}) - ${w.type} - ${w.daysWorked || 0} jours - ${(w.totalSalary || 0).toFixed(2)} DH
                        </div>
                    `).join('')}
                </div>
            `;
        }
    } catch (error) {
        console.error('Error loading monthly workers:', error);
        const infoDiv = document.getElementById('monthly-workers-info');
        const listDiv = document.getElementById('monthly-workers-list');
        infoDiv.style.display = 'block';
        listDiv.innerHTML = '<span style="color: var(--danger);">❌ Erreur lors du chargement des ouvriers</span>';
    }
}

// Load workers with presence for selected quarter
async function loadQuarterlyWorkers() {
    const quarter = parseInt(document.getElementById('quarterly-quarter').value);
    const year = parseInt(document.getElementById('quarterly-year').value);
    
    try {
        const workersWithPresence = await window.api.documents.getAllQuarterlyWorkers(year, quarter);
        
        const infoDiv = document.getElementById('quarterly-workers-info');
        const listDiv = document.getElementById('quarterly-workers-list');
        
        if (workersWithPresence.length === 0) {
            infoDiv.style.display = 'block';
            listDiv.innerHTML = '<span style="color: var(--danger);">❌ Aucun ouvrier avec présence pour ce trimestre</span>';
        } else {
            infoDiv.style.display = 'block';
            const totalDays = workersWithPresence.reduce((sum, w) => sum + (w.totalDaysWorked || w.daysWorked || 0), 0);
            const totalSalary = workersWithPresence.reduce((sum, w) => sum + (w.totalSalary || 0), 0);
            
            listDiv.innerHTML = `
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <div>
                        <strong>${workersWithPresence.length}</strong> ouvrier(s) • 
                        <strong>${totalDays}</strong> jours travaillés • 
                        Total: <strong style="color: var(--success);">${totalSalary.toFixed(2)} DH</strong>
                    </div>
                </div>
                <div style="margin-top: 10px; max-height: 150px; overflow-y: auto;">
                    ${workersWithPresence.map(w => `
                        <div style="padding: 5px; border-bottom: 1px solid var(--border);">
                            👤 ${w.workerName || w.nom_prenom} (${w.cin}) - ${w.workerType || w.type} - ${w.totalDaysWorked || w.daysWorked || 0} jours - ${(w.totalSalary || 0).toFixed(2)} DH
                        </div>
                    `).join('')}
                </div>
            `;
        }
    } catch (error) {
        console.error('Error loading quarterly workers:', error);
        const infoDiv = document.getElementById('quarterly-workers-info');
        const listDiv = document.getElementById('quarterly-workers-list');
        infoDiv.style.display = 'block';
        listDiv.innerHTML = '<span style="color: var(--danger);">❌ Erreur lors du chargement des ouvriers</span>';
    }
}

// Generate monthly document (DEPRECATED - use combined generation instead)
async function generateMonthlyDocument(documentType) {
    window.appUtils.showNotification('Cette fonction n\'est plus disponible. Utilisez le bouton de génération combinée.', 'error');
    return;

    try {
        window.appUtils.showLoading();
        
        const result = await window.api.documents.generateMonthly(
            documentType,
            parseInt(workerId),
            year,
            month
        );

        showGenerationStatus(
            `Document "${documentType}" généré avec succès!\n\nFichier: ${result.fileName}`,
            'success'
        );

    } catch (error) {
        console.error('Error generating document:', error);
        showGenerationStatus(
            `Erreur lors de la génération du document: ${error.message}`,
            'error'
        );
    } finally {
        window.appUtils.hideLoading();
    }
}

// Generate all monthly documents (DEPRECATED - use combined generation instead)
async function generateAllMonthlyDocuments() {
    window.appUtils.showNotification('Cette fonction n\'est plus disponible. Utilisez le bouton de génération combinée.', 'error');
    return;

    const documentTypes = [
        'depense-regie-salaire',
        'recu',
        'demande-autorisation',
        'certificat-paiement',
        'ordre-paiement',
        'depense-regie-recapitulatif',
        'reference-values'
    ];

    try {
        window.appUtils.showLoading();
        
        let successCount = 0;
        let failCount = 0;
        const results = [];

        for (const docType of documentTypes) {
            try {
                const result = await window.api.documents.generateMonthly(
                    docType,
                    parseInt(workerId),
                    year,
                    month
                );
                successCount++;
                results.push(`✓ ${result.fileName}`);
            } catch (error) {
                failCount++;
                results.push(`✗ ${docType}: ${error.message}`);
            }
        }

        showGenerationStatus(
            `Génération terminée!\n\nRéussis: ${successCount}\nÉchecs: ${failCount}\n\n${results.join('\n')}`,
            failCount === 0 ? 'success' : 'error'
        );

    } catch (error) {
        console.error('Error generating documents:', error);
        showGenerationStatus(`Erreur générale: ${error.message}`, 'error');
    } finally {
        window.appUtils.hideLoading();
    }
}

// Generate quarterly document (DEPRECATED - use combined generation instead)
async function generateQuarterlyDocument(documentType) {
    window.appUtils.showNotification('Cette fonction n\'est plus disponible. Utilisez le bouton de génération combinée.', 'error');
    return;

    try {
        window.appUtils.showLoading();
        
        const result = await window.api.documents.generateQuarterly(
            documentType,
            parseInt(workerId),
            year,
            quarter
        );

        showGenerationStatus(
            `Document "${documentType}" généré avec succès!\n\nFichier: ${result.fileName}`,
            'success'
        );

    } catch (error) {
        console.error('Error generating document:', error);
        showGenerationStatus(
            `Erreur lors de la génération du document: ${error.message}`,
            'error'
        );
    } finally {
        window.appUtils.hideLoading();
    }
}

// Generate all quarterly documents (DEPRECATED - use combined generation instead)
async function generateAllQuarterlyDocuments() {
    window.appUtils.showNotification('Cette fonction n\'est plus disponible. Utilisez le bouton de génération combinée.', 'error');
    return;

    const documentTypes = [
        'rcar-salariale',
        'rcar-patronale'
    ];

    try {
        window.appUtils.showLoading();
        
        let successCount = 0;
        let failCount = 0;
        const results = [];

        for (const docType of documentTypes) {
            try {
                const result = await window.api.documents.generateQuarterly(
                    docType,
                    parseInt(workerId),
                    year,
                    quarter
                );
                successCount++;
                results.push(`✓ ${result.fileName}`);
            } catch (error) {
                failCount++;
                results.push(`✗ ${docType}: ${error.message}`);
            }
        }

        showGenerationStatus(
            `Génération terminée!\n\nRéussis: ${successCount}\nÉchecs: ${failCount}\n\n${results.join('\n')}`,
            failCount === 0 ? 'success' : 'error'
        );

    } catch (error) {
        console.error('Error generating documents:', error);
        showGenerationStatus(`Erreur générale: ${error.message}`, 'error');
    } finally {
        window.appUtils.hideLoading();
    }
}

// Generate batch monthly documents for all workers
async function generateBatchMonthly() {
    const month = parseInt(document.getElementById('monthly-month').value);
    const year = parseInt(document.getElementById('monthly-year').value);

    // Ask user to select output directory
    const outputDir = await window.api.dialog.selectDirectory();
    if (!outputDir) {
        window.appUtils.showNotification('Aucun dossier sélectionné', 'error');
        return;
    }

    const documentTypes = [
        { id: 'depense-regie-salaire', name: 'Dépense en Régie - Salaire' },
        // 'recu' removed - now uses combined version (recu-combined) which generates ONE receipt for all workers
        { id: 'demande-autorisation', name: 'Demande d\'autorisation' },
        { id: 'certificat-paiement', name: 'Certificat de paiement' },
        { id: 'ordre-paiement', name: 'Ordre de paiement' },
        { id: 'depense-regie-recapitulatif', name: 'Dépense en Régie (récapitulatif)' },
        { id: 'reference-values', name: 'Valeurs de référence' }
    ];

    try {
        window.appUtils.showLoading();
        
        let totalSuccess = 0;
        let totalFailed = 0;
        const allResults = [];

        for (const docType of documentTypes) {
            try {
                const result = await window.api.documents.generateMonthlyBatch(
                    docType.id,
                    year,
                    month,
                    outputDir
                );
                
                totalSuccess += result.success;
                totalFailed += result.failed;
                allResults.push(`\n📄 ${docType.name}:\n  ✓ Réussis: ${result.success}\n  ✗ Échecs: ${result.failed}`);
                
                if (result.errors.length > 0) {
                    result.errors.forEach(err => {
                        allResults.push(`    • ${formatBatchErrorLine(err)}`);
                    });
                }
            } catch (error) {
                allResults.push(`\n📄 ${docType.name}: ✗ Erreur - ${error.message}`);
                totalFailed++;
            }
        }

        showGenerationStatus(
            `🎉 GÉNÉRATION BATCH TERMINÉE!\n\nDossier: ${outputDir}\n\nRésumé Global:\n✓ Documents réussis: ${totalSuccess}\n✗ Documents échoués: ${totalFailed}\n\nDétails:${allResults.join('\n')}`,
            totalFailed === 0 ? 'success' : 'error'
        );

    } catch (error) {
        console.error('Error generating batch documents:', error);
        showGenerationStatus(`Erreur générale: ${error.message}`, 'error');
    } finally {
        window.appUtils.hideLoading();
    }
}

// Generate batch quarterly documents for all workers
async function generateBatchQuarterly() {
    const quarter = parseInt(document.getElementById('quarterly-quarter').value);
    const year = parseInt(document.getElementById('quarterly-year').value);

    // Ask user to select output directory
    const outputDir = await window.api.dialog.selectDirectory();
    if (!outputDir) {
        window.appUtils.showNotification('Aucun dossier sélectionné', 'error');
        return;
    }

    const documentTypes = [
        { id: 'rcar-salariale', name: 'RCAR - Cotisation Salariale' },
        { id: 'rcar-patronale', name: 'RCAR - Cotisation Patronale' }
    ];

    try {
        window.appUtils.showLoading();
        
        let totalSuccess = 0;
        let totalFailed = 0;
        const allResults = [];

        for (const docType of documentTypes) {
            try {
                const result = await window.api.documents.generateQuarterlyBatch(
                    docType.id,
                    year,
                    quarter,
                    outputDir
                );
                
                totalSuccess += result.success;
                totalFailed += result.failed;
                allResults.push(`\n📄 ${docType.name}:\n  ✓ Réussis: ${result.success}\n  ✗ Échecs: ${result.failed}`);
                
                if (result.errors.length > 0) {
                    result.errors.forEach(err => {
                        allResults.push(`    • ${formatBatchErrorLine(err)}`);
                    });
                }
            } catch (error) {
                allResults.push(`\n📄 ${docType.name}: ✗ Erreur - ${error.message}`);
                totalFailed++;
            }
        }

        showGenerationStatus(
            `🎉 GÉNÉRATION BATCH TERMINÉE!\n\nDossier: ${outputDir}\n\nRésumé Global:\n✓ Documents réussis: ${totalSuccess}\n✗ Documents échoués: ${totalFailed}\n\nDétails:${allResults.join('\n')}`,
            totalFailed === 0 ? 'success' : 'error'
        );

    } catch (error) {
        console.error('Error generating batch documents:', error);
        showGenerationStatus(`Erreur générale: ${error.message}`, 'error');
    } finally {
        window.appUtils.hideLoading();
    }
}

// Generate a SINGLE monthly document type for all workers
async function generateSingleMonthlyDocument(documentType) {
    console.log('📋 Generate Single Monthly Document:', documentType);
    
    const month = parseInt(document.getElementById('monthly-month').value);
    const year = parseInt(document.getElementById('monthly-year').value);

    try {
        // Ask user to select output directory
        const outputDir = await window.api.dialog.selectDirectory();
        
        if (!outputDir) {
            window.appUtils.showNotification('Aucun dossier sélectionné', 'error');
            return;
        }

        window.appUtils.showLoading();
        
        const format = getSelectedDocumentsFormat();
        const result = await window.api.documents.generateMonthlyBatch(
            documentType,
            year,
            month,
            outputDir,
            { format }
        );
        
        window.appUtils.hideLoading();
        const filesList = getResultFileList(result);

        showGenerationStatus(
            `🎉 DOCUMENT GÉNÉRÉ!\n\nType: ${documentType}\nDossier: ${outputDir}\n\nFichiers:\n${filesList}\n\n✓ Réussis: ${result.success}\n✗ Échecs: ${result.failed}\n\nTotal ouvriers: ${result.total}`,
            result.failed === 0 ? 'success' : 'error'
        );

    } catch (error) {
        console.error('❌ Error generating document:', error);
        window.appUtils.hideLoading();
        showGenerationStatus(`Erreur: ${error.message}`, 'error');
    }
}

// Generate ALL monthly documents for all workers
async function generateAllMonthlyDocumentsNew() {
    console.log('📋 Generate ALL Monthly Documents');
    
    const month = parseInt(document.getElementById('monthly-month').value);
    const year = parseInt(document.getElementById('monthly-year').value);

    const documentTypes = [
        { id: 'depense-regie-salaire', name: 'Dépense en Régie - Salaire', mode: 'batch' },
        { id: 'recu-combined', name: 'Reçu (Total)', mode: 'combined' },
        { id: 'demande-autorisation', name: 'Demande d\'Autorisation', mode: 'batch' },
        { id: 'certificat-paiement', name: 'Certificat de Paiement', mode: 'batch' },
        { id: 'ordre-paiement', name: 'Ordre de Paiement', mode: 'batch' },
        { id: 'mandat-paiement', name: 'Mandat de Paiement', mode: 'batch' },
        { id: 'bordereau', name: 'Bordereau', mode: 'batch' }
    ];

    try {
        // Ask user to select output directory
        const outputDir = await window.api.dialog.selectDirectory();
        
        if (!outputDir) {
            window.appUtils.showNotification('Aucun dossier sélectionné', 'error');
            return;
        }

        window.appUtils.showLoading();
        
        let totalSuccess = 0;
        let totalFailed = 0;
        const allResults = [];

        for (const docType of documentTypes) {
            try {
                if (docType.mode === 'combined') {
                    const format = getSelectedDocumentsFormat();
                    const result = await window.api.documents.generateCombinedMonthly(
                        docType.id,
                        year,
                        month,
                        outputDir,
                        { format }
                    );
                    totalSuccess += 1;
                    const filesList = getResultFileList(result);
                    allResults.push(`\n📄 ${docType.name}:\n  ✓ Généré:\n${filesList}`);
                } else {
                    const format = getSelectedDocumentsFormat();
                    const result = await window.api.documents.generateMonthlyBatch(
                        docType.id,
                        year,
                        month,
                        outputDir,
                        { format }
                    );
                    
                    totalSuccess += result.success;
                    totalFailed += result.failed;
                    const filesList = getResultFileList(result);
                    allResults.push(`\n📄 ${docType.name}:\n  ✓ Réussis: ${result.success}\n  ✗ Échecs: ${result.failed}\n  ✓ Fichiers:\n${filesList}`);
                    
                    if (result.errors && result.errors.length > 0) {
                        result.errors.forEach(err => {
                            allResults.push(`    • ${formatBatchErrorLine(err)}`);
                        });
                    }
                }
            } catch (error) {
                allResults.push(`\n📄 ${docType.name}: ✗ Erreur - ${error.message}`);
                totalFailed++;
            }
        }

        window.appUtils.hideLoading();

        showGenerationStatus(
            `🎉 GÉNÉRATION COMPLÈTE!\n\nDossier: ${outputDir}\n\nRésumé Global:\n✓ Documents réussis: ${totalSuccess}\n✗ Documents échoués: ${totalFailed}\n\nDétails:${allResults.join('\n')}`,
            totalFailed === 0 ? 'success' : 'error'
        );

    } catch (error) {
        console.error('❌ Error generating all documents:', error);
        window.appUtils.hideLoading();
        showGenerationStatus(`Erreur générale: ${error.message}`, 'error');
    }
}

// Generate COMBINED monthly document (ONE document for ALL workers)
async function generateCombinedMonthly() {
    console.log('📋 Generate Combined Monthly clicked!');
    
    const month = parseInt(document.getElementById('monthly-month').value);
    const year = parseInt(document.getElementById('monthly-year').value);
    
    console.log(`Selected period: ${month}/${year}`);

    try {
        // Ask user to select output directory
        console.log('Opening directory selector...');
        const outputDir = await window.api.dialog.selectDirectory();
        console.log('Selected directory:', outputDir);
        
        if (!outputDir) {
            console.log('No directory selected, aborting');
            window.appUtils.showNotification('Aucun dossier sélectionné', 'error');
            return;
        }

        window.appUtils.showLoading();
        console.log('Generating combined Dépense en Régie (Salaire)...');
        
        // Generate only Dépense en Régie (Salaire) - Combined
        // Note: Reçu now has its own dedicated button
        const format = getSelectedDocumentsFormat();
        const resultSalaire = await window.api.documents.generateCombinedMonthly(
            'depense-regie-salaire-combined',
            year,
            month,
            outputDir,
            { format }
        );
        
        console.log('Document generated successfully:', resultSalaire);
        const filesList = getResultFileList(resultSalaire);

        showGenerationStatus(
            `🎉 DOCUMENT GÉNÉRÉ!\n\nDossier: ${outputDir}\n\n` +
            `📄 Dépense en Régie (Salaire):\n${filesList}\n  Ouvriers: ${resultSalaire.workersCount}\n  Total: ${(resultSalaire.totalAmount || 0).toFixed(2)} DH\n\n` +
            `UN SEUL document contenant TOUS les ouvriers!`,
            'success'
        );

    } catch (error) {
        console.error('❌ Error generating combined document:', error);
        showGenerationStatus(`Erreur: ${error.message}`, 'error');
    } finally {
        window.appUtils.hideLoading();
    }
}

// Generate ONLY the combined Reçu (ONE receipt for ALL workers)
async function generateRecuCombined() {
    console.log('📋 Generate Reçu Combined clicked!');
    
    const month = parseInt(document.getElementById('monthly-month').value);
    const year = parseInt(document.getElementById('monthly-year').value);
    
    console.log(`Selected period: ${month}/${year}`);

    try {
        // Ask user to select output directory
        console.log('Opening directory selector...');
        const outputDir = await window.api.dialog.selectDirectory();
        console.log('Selected directory:', outputDir);
        
        if (!outputDir) {
            console.log('No directory selected, aborting');
            window.appUtils.showNotification('Aucun dossier sélectionné', 'error');
            return;
        }

        window.appUtils.showLoading();
        console.log('Generating Reçu combined...');
        
        const format = getSelectedDocumentsFormat();
        const resultRecu = await window.api.documents.generateCombinedMonthly(
            'recu-combined',
            year,
            month,
            outputDir,
            { format }
        );
        
        console.log('Reçu generated successfully:', resultRecu);
        const filesList = getResultFileList(resultRecu);

        showGenerationStatus(
            `🎉 REÇU GÉNÉRÉ!\n\nDossier: ${outputDir}\n\n` +
            `📄 Reçu N° ${month}/${year}:\n` +
            `${filesList}\n` +
            `  Nombre d'ouvriers: ${resultRecu.totalWorkers}\n` +
            `  Total Net: ${resultRecu.totalNetSalary} DH\n\n` +
            `UN SEUL reçu contenant le TOTAL de tous les ouvriers!`,
            'success'
        );

    } catch (error) {
        console.error('❌ Error generating Reçu:', error);
        showGenerationStatus(`Erreur: ${error.message}`, 'error');
    } finally {
        window.appUtils.hideLoading();
    }
}

// Generate COMBINED quarterly document (ONE document for ALL workers)
async function generateCombinedQuarterly() {
    const quarter = parseInt(document.getElementById('quarterly-quarter').value);
    const year = parseInt(document.getElementById('quarterly-year').value);

    // Ask user to select output directory
    const outputDir = await window.api.dialog.selectDirectory();
    if (!outputDir) {
        window.appUtils.showNotification('Aucun dossier sélectionné', 'error');
        return;
    }

    try {
        window.appUtils.showLoading();
        
        const format = getSelectedDocumentsFormat();
        const result = await window.api.documents.generateCombinedQuarterly(
            'rcar-combined',
            year,
            quarter,
            outputDir,
            { format }
        );

        const filesList = (result.files || []).map(file => `  • ${file}`).join('\n');
        showGenerationStatus(
            `🎉 DOCUMENTS RCAR GÉNÉRÉS!\n\nDossier: ${outputDir}\n\nFichiers:\n${filesList}\n\nDEUX documents contenant TOUS les ouvriers!`,
            'success'
        );

    } catch (error) {
        console.error('Error generating combined document:', error);
        showGenerationStatus(`Erreur: ${error.message}`, 'error');
    } finally {
        window.appUtils.hideLoading();
    }
}

// Show generation status
function showGenerationStatus(message, type) {
    const statusDiv = document.getElementById('generation-status');
    statusDiv.textContent = message;
    statusDiv.className = `status-message ${type}`;
    statusDiv.style.display = 'block';
    statusDiv.style.whiteSpace = 'pre-wrap';

    if (type === 'success') {
        const current = Number(localStorage.getItem('docs-generated') || 0);
        const next = current + 1;
        localStorage.setItem('docs-generated', String(next));
        const dashDocs = document.getElementById('dash-docs-count');
        if (dashDocs) dashDocs.textContent = next;

        const activity = document.getElementById('recent-activity');
        if (activity) {
            activity.innerHTML = `<div class=\"activity-item\">${new Date().toLocaleString('fr-FR')} - Génération réussie</div>`;
        }
    }

    // Auto hide after 10 seconds
    setTimeout(() => {
        statusDiv.style.display = 'none';
    }, 10000);
}
