/**
 * Business constants for the application
 */

// Worker types and their daily salaries
const WORKER_TYPES = {
  OS: {
    code: 'OS',
    name: 'Ouvrier Spécialisé',
    dailySalary: 100.40
  },
  ONS: {
    code: 'ONS',
    name: 'Ouvrier Non Spécialisé',
    dailySalary: 93.00
  }
};

// Fixed reference values for documents
const REFERENCE_VALUES = {
  chapitre: '10',
  article: '20',
  programme: '20',
  projet: '10',
  ligne: '14'
};

// RCAR age limit for quarterly documents
const RCAR_AGE_LIMIT = 64;

// RCAR deduction rate (6% of gross salary)
const RCAR_RATE = 0.06;

// Quarter definitions (months)
const QUARTERS = {
  1: [1, 2, 3],      // Q1: January, February, March
  2: [4, 5, 6],      // Q2: April, May, June
  3: [7, 8, 9],      // Q3: July, August, September
  4: [10, 11, 12]    // Q4: October, November, December
};

// Month names in French
const MONTH_NAMES = {
  1: 'Janvier',
  2: 'Février',
  3: 'Mars',
  4: 'Avril',
  5: 'Mai',
  6: 'Juin',
  7: 'Juillet',
  8: 'Août',
  9: 'Septembre',
  10: 'Octobre',
  11: 'Novembre',
  12: 'Décembre'
};

// Quarter names in French
const QUARTER_NAMES = {
  1: 'Premier Trimestre',
  2: 'Deuxième Trimestre',
  3: 'Troisième Trimestre',
  4: 'Quatrième Trimestre'
};

/**
 * Get daily salary for a worker type
 * @param {string} type - Worker type (OS or ONS)
 * @returns {number} Daily salary
 */
function getDailySalary(type) {
  return WORKER_TYPES[type]?.dailySalary || 0;
}

/**
 * Get worker type name
 * @param {string} type - Worker type code
 * @returns {string} Worker type name
 */
function getWorkerTypeName(type) {
  return WORKER_TYPES[type]?.name || type;
}

/**
 * Get quarter from month
 * @param {number} month - Month number (1-12)
 * @returns {number} Quarter number (1-4)
 */
function getQuarterFromMonth(month) {
  for (const [quarter, months] of Object.entries(QUARTERS)) {
    if (months.includes(month)) {
      return parseInt(quarter);
    }
  }
  return 1;
}

/**
 * Get months in a quarter
 * @param {number} quarter - Quarter number (1-4)
 * @returns {number[]} Array of month numbers
 */
function getMonthsInQuarter(quarter) {
  return QUARTERS[quarter] || [];
}

/**
 * Get month name in French
 * @param {number} month - Month number (1-12)
 * @returns {string} Month name
 */
function getMonthName(month) {
  return MONTH_NAMES[month] || '';
}

/**
 * Get quarter name in French
 * @param {number} quarter - Quarter number (1-4)
 * @returns {string} Quarter name
 */
function getQuarterName(quarter) {
  return QUARTER_NAMES[quarter] || '';
}

/**
 * Calculate age from date of birth
 * @param {string} dateNaissance - Date of birth (YYYY-MM-DD)
 * @returns {number} Age in years
 */
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

/**
 * Calculate age at a given reference date
 * @param {string} dateNaissance - Date of birth (YYYY-MM-DD)
 * @param {Date} referenceDate - Reference date
 * @returns {number|null} Age in years or null if invalid
 */
function calculateAgeAt(dateNaissance, referenceDate) {
  if (!dateNaissance || !(referenceDate instanceof Date) || isNaN(referenceDate.getTime())) {
    return null;
  }
  const birthDate = new Date(dateNaissance);
  if (isNaN(birthDate.getTime())) {
    return null;
  }
  let age = referenceDate.getFullYear() - birthDate.getFullYear();
  const monthDiff = referenceDate.getMonth() - birthDate.getMonth();
  if (monthDiff < 0 || (monthDiff === 0 && referenceDate.getDate() < birthDate.getDate())) {
    age--;
  }
  return age;
}

/**
 * Check if worker is eligible for RCAR (age <= 64)
 * @param {string} dateNaissance - Date of birth (YYYY-MM-DD)
 * @returns {boolean} True if eligible
 */
function isEligibleForRCAR(dateNaissance) {
  const age = calculateAge(dateNaissance);
  return age <= RCAR_AGE_LIMIT;
}

/**
 * Get number of days in a month
 * @param {number} year - Year
 * @param {number} month - Month (1-12)
 * @returns {number} Number of days
 */
function getDaysInMonth(year, month) {
  return new Date(year, month, 0).getDate();
}

/**
 * Get half-month ranges for a given month
 * @param {number} year - Year
 * @param {number} month - Month (1-12)
 * @returns {Object} Half ranges and last day
 */
function getMonthHalfRanges(year, month) {
  const lastDay = getDaysInMonth(year, month);
  return {
    first: { startDay: 1, endDay: 15 },
    second: { startDay: 16, endDay: lastDay },
    combined: { startDay: 1, endDay: lastDay },
    lastDay
  };
}

/**
 * Filter presence days for a given range
 * @param {number[]} presenceDays - Days worked in month
 * @param {number} startDay - Start day
 * @param {number} endDay - End day
 * @returns {number[]} Filtered days
 */
function filterPresenceDays(presenceDays, startDay, endDay) {
  const days = Array.isArray(presenceDays) ? presenceDays : [];
  return days.filter(day => day >= startDay && day <= endDay);
}

/**
 * Calculate presence stats for a range
 * @param {number[]} presenceDays - Days worked in month
 * @param {number} dailySalary - Daily salary
 * @param {number} startDay - Start day
 * @param {number} endDay - End day
 * @returns {{daysWorked:number,totalSalary:number,presenceDays:number[]}}
 */
function calculatePresenceStatsForRange(presenceDays, dailySalary, startDay, endDay) {
  const filteredDays = filterPresenceDays(presenceDays, startDay, endDay);
  const daysWorked = filteredDays.length;
  const totalSalary = daysWorked * (dailySalary || 0);
  return {
    daysWorked,
    totalSalary,
    presenceDays: filteredDays
  };
}

/**
 * Format currency (Moroccan Dirham)
 * @param {number} amount - Amount to format
 * @returns {string} Formatted currency string
 */
function formatCurrency(amount) {
  return `${amount.toFixed(2)} DH`;
}

/**
 * Format date in French format (DD/MM/YYYY)
 * @param {Date|string} date - Date to format
 * @returns {string} Formatted date string
 */
function formatDate(date) {
  const d = new Date(date);
  const day = String(d.getDate()).padStart(2, '0');
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const year = d.getFullYear();
  return `${day}/${month}/${year}`;
}

/**
 * Calculate RCAR deduction (6% of gross salary)
 * @param {number} grossSalary - Gross salary amount
 * @returns {number} RCAR deduction amount
 */
function calculateRCARDeduction(grossSalary) {
  return grossSalary * RCAR_RATE;
}

/**
 * Calculate net salary (gross - RCAR deduction)
 * @param {number} grossSalary - Gross salary amount
 * @returns {number} Net salary amount
 */
function calculateNetSalary(grossSalary) {
  return grossSalary - calculateRCARDeduction(grossSalary);
}

/**
 * Calculate IGR deduction (6% of gross salary) for workers up to an age limit
 * @param {string} dateNaissance - Date of birth (YYYY-MM-DD)
 * @param {Date} referenceDate - Reference date (period end)
 * @param {number} grossSalary - Gross salary amount
 * @param {number} [ageLimit] - Max age (inclusive) to apply the 6% deduction
 * @returns {number} IGR deduction amount
 */
function calculateIgrDeduction(dateNaissance, referenceDate, grossSalary, ageLimit = RCAR_AGE_LIMIT) {
  const effectiveAgeLimit = Number.isFinite(Number(ageLimit)) ? Number(ageLimit) : RCAR_AGE_LIMIT;
  const age = calculateAgeAt(dateNaissance, referenceDate);
  if (age === null || age <= effectiveAgeLimit) {
    return grossSalary * RCAR_RATE;
  }
  return 0;
}

/**
 * Calculate net salary with IGR deduction
 * @param {string} dateNaissance - Date of birth (YYYY-MM-DD)
 * @param {Date} referenceDate - Reference date (period end)
 * @param {number} grossSalary - Gross salary amount
 * @param {number} [ageLimit] - Max age (inclusive) to apply the 6% deduction
 * @returns {number} Net salary amount
 */
function calculateNetSalaryWithIgr(dateNaissance, referenceDate, grossSalary, ageLimit = RCAR_AGE_LIMIT) {
  return grossSalary - calculateIgrDeduction(dateNaissance, referenceDate, grossSalary, ageLimit);
}

module.exports = {
  WORKER_TYPES,
  REFERENCE_VALUES,
  RCAR_AGE_LIMIT,
  RCAR_RATE,
  QUARTERS,
  MONTH_NAMES,
  QUARTER_NAMES,
  getDailySalary,
  getWorkerTypeName,
  getQuarterFromMonth,
  getMonthsInQuarter,
  getMonthName,
  getQuarterName,
  calculateAge,
  calculateAgeAt,
  isEligibleForRCAR,
  getDaysInMonth,
  getMonthHalfRanges,
  filterPresenceDays,
  calculatePresenceStatsForRange,
  formatCurrency,
  formatDate,
  calculateRCARDeduction,
  calculateNetSalary,
  calculateIgrDeduction,
  calculateNetSalaryWithIgr
};
