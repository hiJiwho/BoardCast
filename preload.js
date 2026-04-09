const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
    getTimetable: () => ipcRenderer.invoke('get-timetable'),
    getSettings: () => ipcRenderer.invoke('get-settings'),
    saveSettings: (data) => ipcRenderer.invoke('save-settings', data),
    getSubjects: () => ipcRenderer.invoke('get-subjects'),
    selectImageFolder: () => ipcRenderer.invoke('select-image-folder'),
    copyImages: (folder) => ipcRenderer.invoke('copy-images', folder),
    getTbimagesList: () => ipcRenderer.invoke('get-tbimages-list'),
    saveSubjectOverride: (index, newSubject) => ipcRenderer.invoke('save-subject-override', index, newSubject),
    resetSubjectOverrides: () => ipcRenderer.invoke('reset-subject-overrides'),

    importTBImages: async () => {
        const folder = await ipcRenderer.invoke('select-image-folder');
        if (!folder) return null;
        const copyRes = await ipcRenderer.invoke('copy-images', folder);
        if (!copyRes.success) return copyRes;
        const files = await ipcRenderer.invoke('get-tbimages-list');
        return { success: true, count: copyRes.count, files };
    },

    setUIMode: (mode, data) => ipcRenderer.send('set-ui-mode', mode, data),
    minimizeAndFloat: (data) => ipcRenderer.send('minimize-and-float', data),
    appQuit: () => ipcRenderer.send('app-quit'),
    setWindowLayout: (layout) => ipcRenderer.send('set-window-layout', layout),
    controlExplorer: (action) => ipcRenderer.send('control-explorer', action),
    openDevTools: () => ipcRenderer.send('open-devtools'),

    // Widget specific
    onWidgetData: (callback) => ipcRenderer.on('update-widget-data', (event, data) => callback(data)),
    onUSBEvent: (callback) => ipcRenderer.on('usb-event', callback),
    getUSBFiles: (drivePath) => ipcRenderer.invoke('list-usb-files', drivePath),
    openUSBFile: (filePath) => ipcRenderer.invoke('open-usb-file', filePath),
    openUSBExplorer: (drivePath) => ipcRenderer.send('open-usb-explorer', drivePath),
    restoreFullscreen: () => ipcRenderer.send('restore-to-fullscreen')
});
