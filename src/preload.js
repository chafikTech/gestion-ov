const { contextBridge, ipcRenderer } = require('electron');

// Expose protected methods that allow the renderer process to use
// the ipcRenderer without exposing the entire object
contextBridge.exposeInMainWorld('api', {
  // Workers API
  workers: {
    getAll: () => ipcRenderer.invoke('workers:getAll'),
    getById: (id) => ipcRenderer.invoke('workers:getById', id),
    search: (query) => ipcRenderer.invoke('workers:search', query),
    create: (workerData) => ipcRenderer.invoke('workers:create', workerData),
    update: (id, workerData) => ipcRenderer.invoke('workers:update', id, workerData),
    delete: (id) => ipcRenderer.invoke('workers:delete', id)
  },
  
  // Presence API
  presence: {
    get: (workerId, year, month) => ipcRenderer.invoke('presence:get', workerId, year, month),
    save: (workerId, year, month, days, attachmentNumber, monthlyAttachmentEntries) =>
      ipcRenderer.invoke('presence:save', workerId, year, month, days, attachmentNumber, monthlyAttachmentEntries),
    calculate: (workerId, year, month) => ipcRenderer.invoke('presence:calculate', workerId, year, month),
    getQuarterly: (workerId, year, quarter) => ipcRenderer.invoke('presence:getQuarterly', workerId, year, quarter),
    getAttachmentNumber: (workerId, year, month) => ipcRenderer.invoke('presence:getAttachmentNumber', workerId, year, month),
    saveAttachmentNumber: (workerId, year, month, attachmentNumber) =>
      ipcRenderer.invoke('presence:saveAttachmentNumber', workerId, year, month, attachmentNumber),
    getMonthlyAttachmentNumbers: (year, month) => ipcRenderer.invoke('presence:getMonthlyAttachmentNumbers', year, month),
    saveMonthlyAttachmentNumbers: (year, month, entries) =>
      ipcRenderer.invoke('presence:saveMonthlyAttachmentNumbers', year, month, entries),
    getWorkersForMonth: (year, month) => ipcRenderer.invoke('presence:getWorkersForMonth', year, month)
  },
  
  // Documents API
  documents: {
    generateMonthly: (documentType, workerId, year, month) => 
      ipcRenderer.invoke('documents:generateMonthly', documentType, workerId, year, month),
    generateQuarterly: (documentType, workerId, year, quarter) => 
      ipcRenderer.invoke('documents:generateQuarterly', documentType, workerId, year, quarter),
    getAllWorkers: (year, month) => ipcRenderer.invoke('documents:getAllWorkers', year, month),
    getAllQuarterlyWorkers: (year, quarter) => ipcRenderer.invoke('documents:getAllQuarterlyWorkers', year, quarter),
    generateMonthlyBatch: (documentType, year, month, outputDir, options) => 
      ipcRenderer.invoke('documents:generateMonthlyBatch', documentType, year, month, outputDir, options),
    generateQuarterlyBatch: (documentType, year, quarter, outputDir, options) => 
      ipcRenderer.invoke('documents:generateQuarterlyBatch', documentType, year, quarter, outputDir, options),
    generateCombinedMonthly: (documentType, year, month, outputDir, options) => 
      ipcRenderer.invoke('documents:generateCombinedMonthly', documentType, year, month, outputDir, options),
    generateCombinedQuarterly: (documentType, year, quarter, outputDir, options) => 
      ipcRenderer.invoke('documents:generateCombinedQuarterly', documentType, year, quarter, outputDir, options)
  },
  
  // File Dialog API
  dialog: {
    selectDirectory: () => ipcRenderer.invoke('dialog:selectDirectory'),
    saveFile: (options) => ipcRenderer.invoke('dialog:saveFile', options),
    openFile: (options) => ipcRenderer.invoke('dialog:openFile', options)
  },

  // Application Settings API
  settings: {
    get: () => ipcRenderer.invoke('settings:get'),
    save: (settings) => ipcRenderer.invoke('settings:save', settings)
  },

  database: {
    export: (filePath) => ipcRenderer.invoke('database:export', filePath),
    import: (filePath) => ipcRenderer.invoke('database:import', filePath)
  }
});
