/**
 * Application Configuration
 * Customize your document templates here (DOCX-only)
 */

const config = {
    // Organization Information
    organization: {
        name: 'Nom de votre établissement',
        address: 'Adresse complète de l\'établissement',
        city: 'Ville',
        phone: 'Téléphone: +212 XXX-XXXXXX',
        email: 'email@etablissement.ma',
        // Logo path (relative to project root or absolute path)
        // Example: 'assets/logo.png' or '/path/to/logo.png'
        logoPath: null
    },

    // Responsible persons (for signatures)
    responsibles: {
        regisseur: {
            name: 'Nom du Régisseur',
            title: 'Le Régisseur'
        },
        director: {
            name: 'Nom du Directeur',
            title: 'Le Directeur'
        },
        comptable: {
            name: 'Nom du Chef Comptable',
            title: 'Le Chef Comptable'
        },
        tresorier: {
            name: 'Nom du Trésorier',
            title: 'Le Trésorier'
        }
    },

    // Document specific settings
    documents: {
        // Add watermark to documents (optional)
        watermark: {
            enabled: false,
            text: 'CONFIDENTIEL',
            opacity: 0.1
        },

        // Add custom footer text
        footerText: 'Document officiel généré automatiquement',

        // Date format
        dateFormat: 'DD/MM/YYYY', // Options: 'DD/MM/YYYY', 'YYYY-MM-DD', etc.

        // CERTIFICAT DE PAIEMENT decision reference (managed in code/config, not user input)
        // Use "{year}" placeholder for automatic yearly update.
        certificatPaiementDecision: {
            number: '2/{year}',
            date: '1/2/{year}'
        }
    },

    // RCAR settings (required for quarterly forms)
    rcar: {
        adhesionNumber: '35160001' // Numéro d’adhésion RCAR (digits only)
    }
};

/**
 * Get configuration
 */
function getConfig() {
    return config;
}

module.exports = {
    ...config,
    getConfig
};
