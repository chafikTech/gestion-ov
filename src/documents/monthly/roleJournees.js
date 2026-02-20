const { REFERENCE_VALUES, RCAR_AGE_LIMIT, RCAR_RATE } = require('../../backend/constants');
const fs = require('node:fs');
const path = require('node:path');
const { app } = require('electron');
const pythonBridge = require('../../backend/pythonBridge');

const PYTHON_ROLE_SCRIPT = path.join('src', 'python', 'generate_role.py');

/**
 * Generate "ROLE DES JOURNEES D'OUVRIERS EMPLOYES" - Combined document for ALL workers
 * This matches the exact template structure
 */
class RoleJourneesGenerator {
  constructor() {
    this.outputDir = '';
  }

  getDefaultOutputDir() {
    return path.join(app.getPath('documents'), 'Gestion_Ouvriers', 'Documents');
  }

  normalizeDateInput(value) {
    if (!value) {
      return '';
    }
    if (value instanceof Date) {
      return this.formatShortDate(value);
    }
    return String(value).trim();
  }

  formatShortDate(date) {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
  }

  parseDate(value) {
    if (!value) {
      return new Date(NaN);
    }
    if (value instanceof Date) {
      return value;
    }
    const str = String(value).trim();
    if (str.includes('/')) {
      const parts = str.split('/').map(p => p.trim());
      if (parts.length === 3) {
        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10) - 1;
        const year = parseInt(parts[2], 10);
        return new Date(year, month, day);
      }
    }
    return new Date(str);
  }

  calculateAge(dateNaissance, periodEndDate) {
    if (!dateNaissance || !(periodEndDate instanceof Date) || Number.isNaN(periodEndDate.getTime())) {
      return null;
    }
    const birthDate = this.parseDate(dateNaissance);
    if (Number.isNaN(birthDate.getTime())) {
      return null;
    }
    let age = periodEndDate.getFullYear() - birthDate.getFullYear();
    const monthDiff = periodEndDate.getMonth() - birthDate.getMonth();
    if (monthDiff < 0 || (monthDiff === 0 && periodEndDate.getDate() < birthDate.getDate())) {
      age -= 1;
    }
    return age;
  }

  round2(value) {
    return Number(Number(value || 0).toFixed(2));
  }

  calculateIgrDeduction(dateNaissance, periodEndDate, grossSalary, ageLimit = RCAR_AGE_LIMIT) {
    const effectiveAgeLimit = Number.isFinite(Number(ageLimit)) ? Number(ageLimit) : RCAR_AGE_LIMIT;
    const age = this.calculateAge(dateNaissance, periodEndDate);
    if (age === null || age <= effectiveAgeLimit) {
      return this.round2(grossSalary * RCAR_RATE);
    }
    return 0;
  }

  async generate(data) {
    const { report, year, month, startDate, endDate, options = {} } = data;
    const workers = report?.rows || [];
    
    // Validate we have workers
    if (!workers || workers.length === 0) {
      throw new Error('Aucun ouvrier avec présence pour cette période');
    }

    // Filter to ensure ONLY real workers (exclude administrative staff)
    const realWorkers = workers.filter(w => 
      w.workerId && // Must have valid ID
      w.nom_prenom && // Must have name
      w.type && (w.type === 'OS' || w.type === 'ONS') && // Must be OS or ONS
      w.salaire_journalier > 0 // Must have valid salary
    );

    if (realWorkers.length === 0) {
      throw new Error('Aucun ouvrier valide trouvé (seuls les ouvriers OS/ONS sont inclus)');
    }

    const periodStart = this.normalizeDateInput(
      options.dateFrom || startDate || this.formatShortDate(new Date(year, month - 1, 1))
    );
    const periodEnd = this.normalizeDateInput(
      options.dateTo || endDate || this.formatShortDate(new Date(year, month, 0))
    );
    const monthIndex = month - 1;
    const monthEndDay = new Date(year, month, 0).getDate();

    const buildSection = (startDay, endDay) => {
      const sectionStartDate = new Date(year, monthIndex, startDay);
      const sectionEndDate = new Date(year, monthIndex, endDay);
      const startDateStr = this.formatShortDate(sectionStartDate);
      const endDateStr = this.formatShortDate(sectionEndDate);
      const documentDate = options.documentDate || endDateStr;
      const payDate = options.payDate || documentDate;

      let totalDays = 0;
      let totalGross = 0;
      let totalDeduction = 0;
      let totalNet = 0;

      const computedRows = realWorkers.map(worker => {
        const presenceDays = Array.isArray(worker.presenceDays) ? worker.presenceDays : [];
        const daysWorked = presenceDays.filter(day => day >= startDay && day <= endDay).length;
        if (daysWorked === 0) {
          return null;
        }
        const dailyRate = Number(worker.salaire_journalier || 0);
        const grossSalary = this.round2(daysWorked * dailyRate);
        const deduction = this.calculateIgrDeduction(worker.date_naissance, sectionEndDate, grossSalary, options.rcarAgeLimit);
        const netSalary = this.round2(grossSalary - deduction);
        const cinValidity = worker.cin_validite ? this.formatShortDate(new Date(worker.cin_validite)) : '';

        totalDays += daysWorked;
        totalGross += grossSalary;
        totalDeduction += deduction;
        totalNet += netSalary;

        return {
          ...worker,
          daysWorked,
          grossSalary,
          deduction,
          netSalary,
          cin_validite: cinValidity
        };
      }).filter(Boolean);

      return {
        startDate: startDateStr,
        endDate: endDateStr,
        documentDate,
        payDate,
        workers: computedRows,
        totalDays: this.round2(totalDays),
        totalGross: this.round2(totalGross),
        totalDeduction: this.round2(totalDeduction),
        totalNet: this.round2(totalNet)
      };
    };

    const sections = [
      buildSection(1, 15),
      buildSection(16, monthEndDay)
    ];

    const totals = sections.reduce(
      (acc, section) => {
        acc.totalDays += section.totalDays;
        acc.totalGross += section.totalGross;
        acc.totalDeduction += section.totalDeduction;
        acc.totalNet += section.totalNet;
        return acc;
      },
      { totalDays: 0, totalGross: 0, totalDeduction: 0, totalNet: 0 }
    );
    totals.totalGross = this.round2(totals.totalGross);
    totals.totalDeduction = this.round2(totals.totalDeduction);
    totals.totalNet = this.round2(totals.totalNet);

    // Regisseur name is fixed and immutable
    const regisseurName = 'MAJDA TAKNOUTI';

    const referenceValues = {
      chapitre: (options.chap || REFERENCE_VALUES.chapitre),
      article: (options.art || REFERENCE_VALUES.article),
      programme: (options.prog || REFERENCE_VALUES.programme),
      projet: (options.proj || REFERENCE_VALUES.projet),
      ligne: (options.ligne || REFERENCE_VALUES.ligne)
    };

    const outputDir = (this.outputDir || '').toString().trim() || this.getDefaultOutputDir();
    if (outputDir && !fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    const safeStart = periodStart.replace(/\//g, '-');
    const safeEnd = periodEnd.replace(/\//g, '-');

    const payload = {
      report,
      year,
      month,
      periodStart,
      periodEnd,
      safeStart,
      safeEnd,
      outputDir: outputDir || '',
      format: 'docx',
      decimalComma: options.decimalComma !== undefined ? !!options.decimalComma : true,
      regisseurName,
      referenceValues,
      splitIndex: 8,
      continuationRows: 6,
      sections
    };

    const pythonResult = await pythonBridge.runPythonJson(PYTHON_ROLE_SCRIPT, payload);
    const result = {
      success: true,
      docxFileName: pythonResult.docxFileName,
      docxFilePath: pythonResult.docxFilePath,
      files: [pythonResult.docxFileName]
    };

    return {
      ...result,
      documentType: 'role-journees-combined',
      workersCount: realWorkers.length,
      totalGross: totals.totalGross,
      totalDeduction: totals.totalDeduction,
      totalNet: totals.totalNet,
      totalAmount: totals.totalNet
    };
  }
}

module.exports = new RoleJourneesGenerator();
