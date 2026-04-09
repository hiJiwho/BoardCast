const { app, BrowserWindow, ipcMain, shell, dialog, screen } = require('electron')
const path = require('node:path')
const fs = require('fs')
const { exec, spawn } = require('child_process')
const Timetable = require('comcigan-parser')
const QRCode = require('qrcode')

// Fix black screen in Sandbox/VM
app.disableHardwareAcceleration();

function isAprilFools() {
  const now = new Date();
  return (now.getMonth() === 3 && now.getDate() === 1);
}

// Single Instance Lock
const gotTheLock = app.requestSingleInstanceLock();
if (!gotTheLock) {
  app.quit();
} else {
  app.on('second-instance', () => {
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore();
      mainWindow.focus();
    }
  });

  ipcMain.on('open-devtools', () => {
    if (mainWindow) {
      mainWindow.webContents.openDevTools({ mode: 'detach' });
    }
  });

  ipcMain.on('app-quit', () => {
    app.quit();
  });
}

// Global Error Handling to prevent silent crashes
process.on('uncaughtException', (error) => {
  console.error('[System] Uncaught Exception:', error);
});

process.on('unhandledRejection', (reason, promise) => {
  console.error('[System] Unhandled Rejection at:', promise, 'reason:', reason);
});

let knownDrives = new Set();

let mainWindow = null;
let usbExplorerWindow = null;
let currentUSBDrive = null;
let originalWallpaper = ''; // To store user's original wallpaper
let shouldOpenDevTools = false; // Flag to open DevTools for the current session

const createWindow = () => {
  const userDataPath = app.getPath('userData');
  const infoPath = path.join(userDataPath, 'Info.json');
  let info = {};
  let hasSettings = false;
  try {
    if (fs.existsSync(infoPath)) {
      info = JSON.parse(fs.readFileSync(infoPath, 'utf8'));
      hasSettings = !!(info && info.schoolName);
    }
  } catch (e) { /* ignore */ }

  // Intelligent Startup Mode Detection
  let initialLayout = 'break'; // Default to fullscreen
  let targetBounds = null;

  if (hasSettings) {
    const now = new Date();
    const curMins = now.getHours() * 60 + now.getMinutes();
    const startTimeStr = info.StartTime || "09:00";
    const [stH, stM] = startTimeStr.split(':').map(Number);
    const startMins = stH * 60 + stM;
    const classDur = parseInt(info.ClassDuration) || 45;
    const breakMins = parseInt(info.Breaktime) || 10;
    const lunchDur = parseInt(info.LunchDuration) || 50;

    let inClass = false;
    let marker = startMins;
    for (let i = 0; i < 8; i++) {
      const start = marker;
      const end = start + classDur;

      // We are in class if time is between start and end of a period
      if (curMins >= start && curMins < end) {
        inClass = true;
        break;
      }

      // Move marker to next period (end + break + possible lunch)
      marker = end + breakMins;
      if (i === 3) { // After 4th period, add lunch time
        marker += (lunchDur > breakMins ? lunchDur - breakMins : 0);
      }
    }

    // Special cases where it should NOT be class mode (조회 is before startMins)
    if (curMins < startMins) inClass = false;
    if (curMins >= marker) inClass = false;

    if (inClass) initialLayout = 'class';
    else initialLayout = 'break';
  }

  const primaryDisplay = screen.getPrimaryDisplay();
  const { width: screenWidth, height: screenHeight } = primaryDisplay.workAreaSize;
  const desktopWidth = Math.floor(screenWidth * 0.8);

  mainWindow = new BrowserWindow({
    width: initialLayout === 'class' ? desktopWidth : screenWidth,
    height: screenHeight,
    x: initialLayout === 'class' ? screenWidth - desktopWidth : 0,
    y: 0,
    frame: false,
    autoHideMenuBar: true,
    fullscreen: initialLayout === 'break' && hasSettings,
    icon: path.join(__dirname, 'assets', 'icon.ico'), // Set official icon
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true
    }
  })

  if (hasSettings && initialLayout === 'break') {
    if (!isAprilFools()) {
      console.log("[System] Startup in break mode (fullscreen), killing explorer.exe");
      exec('taskkill /F /IM explorer.exe');
    } else {
      console.log("[System] April Fools! Skipping explorer kill on startup.");
    }
  } else if (hasSettings && initialLayout === 'class') {
    console.log("[System] Startup in class mode (desktop), keeping explorer.exe");
  }

  // Completely hide the native menu bar
  mainWindow.setMenuBarVisibility(false);
  mainWindow.autoHideMenuBar = true;

  // Open DevTools if requested
  if (shouldOpenDevTools) {
    mainWindow.webContents.openDevTools({ mode: 'detach' });
  } else {
    try {
      const userDataPath = app.getPath('userData');
      const infoPath = path.join(userDataPath, 'Info.json');
      if (fs.existsSync(infoPath)) {
        const info = JSON.parse(fs.readFileSync(infoPath, 'utf8'));
        if (info.Dev === true) {
          mainWindow.webContents.openDevTools({ mode: 'detach' });
        }
      }
    } catch (e) { console.error("DevTools Auto-Open Error:", e); }
  }

  mainWindow.loadFile('index.html')
}

function ensureDir(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
}

function readJson(filePath, defaultObj) {
  try {
    if (fs.existsSync(filePath)) {
      const data = fs.readFileSync(filePath, 'utf8');
      return JSON.parse(data);
    }
  } catch (err) {
    console.error(`Error reading ${filePath}:`, err);
  }
  return defaultObj !== undefined ? defaultObj : {};
}

function writeJson(filePath, data) {
  try {
    ensureDir(path.dirname(filePath));
    fs.writeFileSync(filePath, JSON.stringify(data, null, 2), 'utf8');
    return true;
  } catch (err) {
    console.error(`Error writing ${filePath}:`, err);
    return false;
  }
}

function resetDebugOption(infoPath) {
  try {
    const info = readJson(infoPath);
    if (info && info.Dev === true) {
      info.Dev = false;
      writeJson(infoPath, info);
      return true;
    }
  } catch (e) {
    console.error("[System] Failed to reset Debug option:", e);
  }
  return false;
}

app.whenReady().then(() => {
  const userDataPath = app.getPath('userData');
  const infoPath = path.join(userDataPath, 'Info.json');
  const tbImagePath = path.join(userDataPath, 'TBimage.json');
  const tbImageDir = path.join(userDataPath, 'TBimage');
  const tempPath = path.join(userDataPath, 'Temp.json');
  ensureDir(tbImageDir);

  // Initialize with empty files if they don't exist
  if (!fs.existsSync(infoPath)) writeJson(infoPath, { region: "", schoolName: "", grade: 1, class: 1, Breaktime: 10, ClassDuration: 45, AutoStart: true, homeTeacher: "", Dev: false });
  if (!fs.existsSync(tbImagePath)) writeJson(tbImagePath, []);
  if (!fs.existsSync(tempPath)) writeJson(tempPath, { overrides: {} });

  // Apply AutoStart settings on launch
  const currentInfo = readJson(infoPath);
  if (currentInfo) {
    // Reset Debug mode on restart as requested
    if (currentInfo.Dev === true) {
      console.log("[System] Debug mode detected. Resetting for next launch...");
      shouldOpenDevTools = true;
      resetDebugOption(infoPath);
    }

    if (currentInfo.AutoStart !== undefined) {
      try {
        app.setLoginItemSettings({
          openAtLogin: currentInfo.AutoStart,
          path: app.getPath('exe')
        });
        console.log("[System] Startup registration synced:", currentInfo.AutoStart);
      } catch (e) {
        console.warn("[System] Failed to sync startup registration:", e.message);
      }
    }
  }

  // Remove legacy Name.json if it exists
  const legacyNamePath = path.join(userDataPath, 'Name.json');
  if (fs.existsSync(legacyNamePath)) {
    try { fs.unlinkSync(legacyNamePath); } catch (e) { /* ignore */ }
  }

  // Helper: find entry in TBimage array by id
  function findTbEntry(tbImageArr, id) {
    return tbImageArr.find(e => e.id === id);
  }

  // Detect Original Wallpaper on start
  exec(`powershell -command "Get-ItemProperty -Path 'HKCU:\\Control Panel\\Desktop' -Name Wallpaper | Select-Object -ExpandProperty Wallpaper"`, (err, stdout) => {
    if (!err && stdout.trim()) {
      originalWallpaper = stdout.trim();
      console.log("[System] Original wallpaper detected:", originalWallpaper);
    }
  });

  function setWallpaper(imagePath) {
    if (!imagePath || !fs.existsSync(imagePath)) return;
    const psCommand = `Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public class Wallpaper { [DllImport("user32.dll", CharSet = CharSet.Auto)] public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni); }'; [Wallpaper]::SystemParametersInfo(20, 0, "${imagePath.replace(/\\/g, '\\\\')}", 3)`;
    exec(`powershell -command "${psCommand}"`);
  }

  ipcMain.handle('get-timetable', async (event, overrideDate) => {
    console.log("[IPC] get-timetable invoked", overrideDate ? `with override ${overrideDate}` : "");
    try {
      const info = readJson(infoPath);
      let tbImageArr = readJson(tbImagePath, []);
      // Ensure it's an array (migration from old format)
      if (!Array.isArray(tbImageArr)) {
        tbImageArr = [];
        writeJson(tbImagePath, tbImageArr);
      }

      console.log("[IPC] get-timetable: info loaded", info);
      if (!info || !info.schoolName) {
        console.warn("[IPC] get-timetable: No valid schoolName found.");
        return { success: false, needsMapping: true, message: "초기 설정이 필요합니다." };
      }

      console.log("[IPC] get-timetable: initializing Timetable()");
      const timetable = new Timetable();

      const timeoutPromise = (ms, promise) => {
        return new Promise((resolve, reject) => {
          const timeoutId = setTimeout(() => {
            reject(new Error("API Timeout (" + ms + "ms) - check network"));
          }, ms);
          if (!promise || typeof promise.then !== 'function') {
            clearTimeout(timeoutId);
            reject(new Error("Invalid API promise"));
            return;
          }
          promise.then(res => { clearTimeout(timeoutId); resolve(res); })
            .catch(err => { clearTimeout(timeoutId); reject(err); });
        });
      };

      await timeoutPromise(5000, timetable.init());

      console.log("[IPC] get-timetable: timetable initialized");
      const schoolList = await timeoutPromise(5000, timetable.search(info.schoolName));
      console.log(`[Search] Found ${schoolList.length} schools for "${info.schoolName}"`);
      let targetSchool = null;
      if (schoolList.length === 1) {
        targetSchool = schoolList[0];
      } else if (schoolList.length > 1) {
        targetSchool = schoolList.find((school) => {
          return school.region.includes(info.region) && school.name.includes(info.schoolName);
        });
        if (!targetSchool) {
          targetSchool = schoolList.find(s => s.name.includes(info.schoolName));
        }
        if (!targetSchool) {
          targetSchool = schoolList[0];
        }
      }

      if (!targetSchool) {
        throw new Error(`'${info.schoolName}' 학교를 찾을 수 없습니다. 정확한 이름을 입력해주세요.`);
      }

      timetable.setSchool(targetSchool.code);

      console.log("[IPC] get-timetable: timetable.setSchool done");

      console.log("[IPC] get-timetable: timetable.getTimetable()");
      const result = await timeoutPromise(5000, timetable.getTimetable());

      console.log("[IPC] get-timetable: timetable.getClassTime()");
      const classTimes = await timeoutPromise(3000, timetable.getClassTime());
      console.log("[IPC] get-timetable: API data fetched successfully");

      const date = overrideDate ? new Date(overrideDate) : new Date();
      let dayOfWeek = date.getDay();
      if (dayOfWeek === 0 || dayOfWeek === 6) {
        dayOfWeek = 5;
      }

      const weekdayIndex = dayOfWeek - 1;

      const weekData = result[info.grade][info.class] || [];
      const tempData = readJson(tempPath, { overrides: {} });

      let missingMappingDetails = { images: [], teachers: [] };
      let tbArrChanged = false;
      const weeklySubjectsSet = new Set();

      // Scan the entire week to discover subjects and teachers
      for (let d = 0; d < 5; d++) {
        const dayTimetable = weekData[d] || [];

        dayTimetable.forEach(item => {
          if (!item.subject && !item.teacher) return;

          const shortName = item.teacher || "";
          const subject = item.subject || "";

          if (subject) {
            const baseSubj = subject.replace(/[\dABab]+$/, '').trim() || subject;
            weeklySubjectsSet.add(baseSubj);
          }

          // Skip mapping for subjects that don't need books/teachers
          if (subject === "자율" || subject === "동아리") return;

          const pureShortName = shortName.replace(/\*/g, '').trim();
          const mapKey = `${subject}${pureShortName}`;

          // Check if entry exists in TBimage array
          let existing = findTbEntry(tbImageArr, mapKey);
          if (!existing) {
            // New entry: Attempt Auto-Mapping by subject name
            let autoImage = "none";
            const possibleExts = ['.png', '.jpg', '.jpeg'];
            const baseSubj = subject.replace(/[\dABab]+$/, '').trim() || subject;

            for (const ext of possibleExts) {
              // 1. Try exact match
              let testPath = path.join(tbImageDir, `${subject}${ext}`);
              if (fs.existsSync(testPath)) {
                autoImage = `${subject}${ext}`;
                break;
              }
              // 2. Try base subject match (e.g., "국어" for "국어2")
              testPath = path.join(tbImageDir, `${baseSubj}${ext}`);
              if (fs.existsSync(testPath)) {
                autoImage = `${baseSubj}${ext}`;
                break;
              }
            }

            existing = {
              id: mapKey,
              subject: subject,
              teacher: "",
              image: autoImage
            };
            tbImageArr.push(existing);
            tbArrChanged = true;
          }

          if (!existing.teacher && !missingMappingDetails.teachers.includes(mapKey)) {
            missingMappingDetails.teachers.push(mapKey);
          }
          if ((!existing.image || existing.image === "none") && !missingMappingDetails.images.includes(subject)) {
            missingMappingDetails.images.push(subject);
          }
        });
      }

      const hasUnmappedEntry = missingMappingDetails.images.length > 0 || missingMappingDetails.teachers.length > 0;

      if (hasUnmappedEntry) {
        if (tbArrChanged) writeJson(tbImagePath, tbImageArr);
        return {
          success: false,
          needsMapping: true,
          missingDetails: missingMappingDetails,
          message: "교과서 이미지 및 교사 이름 매핑이 필요합니다."
        };
      }

      if (tbArrChanged) writeJson(tbImagePath, tbImageArr);

      let todayTimetable = weekData[weekdayIndex] || [];
      todayTimetable = todayTimetable.filter(item => item.subject || item.teacher);

      const mappedTimetable = todayTimetable.map((item, index) => {
        let subject = item.subject || "";
        let isOverridden = false;

        // Apply override if exists in Temp.json
        if (tempData.overrides && tempData.overrides[index]) {
          subject = tempData.overrides[index];
          isOverridden = true;
        }

        const shortName = item.teacher || "";

        // Subjects without textbook mapping
        if (subject === "자율" || subject === "동아리") {
          return {
            ...item,
            mappedTeacher: subject,
            subjectImage: null,
            hideImage: true,
            isOverridden: isOverridden
          };
        }

        const pureShortName = shortName.replace(/\*/g, '').trim();
        const mapKey = `${subject}${pureShortName}`;

        // Find teacher name from TBimage array
        const entry = findTbEntry(tbImageArr, mapKey);
        let mappedName = shortName;
        if (entry && entry.teacher) {
          mappedName = entry.teacher;
        }

        // Determine subject image — with A/B/1/2 fallback
        let subjectImage = null;
        let hideImage = false;

        if (subject) {
          // Find image: first try exact subject match, then base subject
          let imageValue = null;

          // Try exact subject match in entries
          if (entry && entry.image && entry.image.trim() && entry.image.trim().toLowerCase() !== 'none') {
            imageValue = entry.image.trim();
          }

          // Fallback: try base subject (strip trailing digits/A/B)
          if (!imageValue) {
            const baseMatch = subject.match(/^(.+?)\s*[\dABab]*$/);
            if (baseMatch && baseMatch[1] !== subject) {
              const baseEntry = tbImageArr.find(e => e.subject === baseMatch[1] && e.image && e.image.trim() && e.image.trim().toLowerCase() !== 'none');
              if (baseEntry) {
                imageValue = baseEntry.image.trim();
              }
            }
          }

          if (imageValue) {
            if (imageValue.toLowerCase() === 'none') {
              hideImage = true;
            } else {
              const imagePath = path.join(tbImageDir, imageValue);
              if (fs.existsSync(imagePath)) {
                subjectImage = `file:///${imagePath.replace(/\\/g, '/')}`;
              }
            }
          }
        }

        return {
          ...item,
          subject: subject,
          isOverridden: isOverridden,
          mappedTeacher: mappedName,
          subjectImage: subjectImage,
          hideImage: hideImage
        };
      });

      // Generate QR Code if Share-notion is configured
      let shareQrCode = null;
      if (info["Share-notion"]) {
        try {
          shareQrCode = await QRCode.toDataURL(info["Share-notion"], { width: 200, margin: 1 });
        } catch (e) {
          console.error('Failed to generate QR:', e);
        }
      }

      console.log("[IPC] get-timetable: finished mapping, returning to renderer");
      return JSON.parse(JSON.stringify({
        success: true,
        info: info,
        classTimes: classTimes,
        timetable: mappedTimetable,
        weekday: weekdayIndex,
        shareQrCode: shareQrCode,
        weeklySubjects: [...weeklySubjectsSet].sort()
      }));
    } catch (error) {
      console.error("[IPC] get-timetable Error caught:", error);
      return JSON.parse(JSON.stringify({ success: false, error: error ? (error.message || String(error)) : "Unknown IPC Error" }));
    }
  });

  // Settings IPC Handlers
  ipcMain.handle('get-settings', async () => {
    const tbImageArr = readJson(tbImagePath, []);
    return JSON.parse(JSON.stringify({
      info: readJson(infoPath),
      tbImageArr: Array.isArray(tbImageArr) ? tbImageArr : []
    }));
  });

  ipcMain.handle('save-settings', async (event, { info, tbImageArr }) => {
    try {
      console.log("[IPC] save-settings invoked", info ? "with info" : "no info", tbImageArr ? "with tbImageArr" : "no tbImageArr");

      if (info) {
        writeJson(infoPath, info);
        // Handle AutoStart Toggle separately to not break the whole save if it fails
        if (info.AutoStart !== undefined) {
          try {
            app.setLoginItemSettings({
              openAtLogin: info.AutoStart,
              path: app.getPath('exe')
            });
          } catch (autoStartErr) {
            console.warn("[System] Failed to set login item settings:", autoStartErr.message);
          }
        }
      }

      if (tbImageArr) {
        writeJson(tbImagePath, tbImageArr);
      }

      return { success: true };
    } catch (e) {
      console.error("[IPC] save-settings Error:", e);
      return { success: false, error: e.message || String(e) };
    }
  });

  ipcMain.handle('save-subject-override', async (event, index, newSubject) => {
    try {
      const tempData = readJson(tempPath, { overrides: {} });
      if (newSubject === null) {
        delete tempData.overrides[index];
      } else {
        tempData.overrides[index] = newSubject;
      }
      writeJson(tempPath, tempData);
      return JSON.parse(JSON.stringify({ success: true }));
    } catch (e) {
      return JSON.parse(JSON.stringify({ success: false, error: e.message || String(e) }));
    }
  });

  ipcMain.handle('reset-subject-overrides', async () => {
    try {
      writeJson(tempPath, { overrides: {} });
      return JSON.parse(JSON.stringify({ success: true }));
    } catch (e) {
      return JSON.parse(JSON.stringify({ success: false, error: e.message || String(e) }));
    }
  });

  // Get unique subjects from Comcigan API for this class (excludes duplicate suffixes)
  ipcMain.handle('get-subjects', async () => {
    try {
      const info = readJson(infoPath);
      if (!info.schoolName) return { success: false, error: "학교 설정이 필요합니다." };

      const timetable = new Timetable();
      await timetable.init();

      const schoolList = await timetable.search(info.schoolName);
      let targetSchool = null;
      if (schoolList.length === 1) {
        targetSchool = schoolList[0];
      } else if (schoolList.length > 1) {
        targetSchool = schoolList.find(s => s.region.includes(info.region) && s.name.includes(info.schoolName))
          || schoolList.find(s => s.name.includes(info.schoolName))
          || schoolList[0];
      }

      if (!targetSchool) throw new Error("학교를 찾을 수 없습니다.");

      timetable.setSchool(targetSchool.code);
      const result = await timetable.getTimetable();
      const weekData = result[info.grade][info.class] || [];

      // Collect all unique subjects and subject+teacher combos
      const subjectSet = new Set();
      const entries = []; // { id, subject, teacherShort }

      for (let d = 0; d < 5; d++) {
        const dayTimetable = weekData[d] || [];
        dayTimetable.forEach(item => {
          if (!item.subject) return;
          const subject = item.subject;
          const shortName = (item.teacher || "").replace(/\*/g, '').trim();

          // Skip "자율" and "동아리"
          if (subject === "자율" || subject === "동아리") return;

          // Strip trailing duplicate markers (1,2,A,B,a,b) for display
          const baseSubject = subject.replace(/[\dABab]+$/, '').trim() || subject;
          subjectSet.add(baseSubject);

          const mapKey = `${subject}${shortName}`;
          if (!entries.find(e => e.id === mapKey)) {
            entries.push({
              id: mapKey,
              subject: subject,
              baseSubject: baseSubject,
              teacherShort: shortName
            });
          }
        });
      }

      return {
        success: true,
        subjects: [...subjectSet].sort(),
        entries: entries
      };
    } catch (e) {
      return { success: false, error: e.message };
    }
  });

  ipcMain.handle('select-image-folder', async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
      properties: ['openDirectory']
    });
    if (result.canceled) return null;
    return result.filePaths[0];
  });

  ipcMain.handle('copy-images', async (event, sourceFolder) => {
    try {
      const files = fs.readdirSync(sourceFolder);
      let copiedCount = 0;
      files.forEach(file => {
        if (file.toLowerCase().endsWith('.png') || file.toLowerCase().endsWith('.jpg') || file.toLowerCase().endsWith('.jpeg')) {
          const srcPath = path.join(sourceFolder, file);
          const destPath = path.join(tbImageDir, file);
          fs.copyFileSync(srcPath, destPath);
          copiedCount++;
        }
      });
      return { success: true, count: copiedCount };
    } catch (e) {
      return { success: false, error: e.message };
    }
  });

  ipcMain.handle('get-tbimages-list', async () => {
    try {
      const files = fs.readdirSync(tbImageDir);
      return files.filter(f => f.toLowerCase().match(/\.(png|jpe?g)$/));
    } catch (e) {
      return [];
    }
  });

  // --- USB File Explorer IPC ---
  ipcMain.handle('list-usb-files', async (event, drivePath) => {
    try {
      const files = fs.readdirSync(drivePath, { withFileTypes: true });
      return files.map(f => ({
        name: f.name,
        isDirectory: f.isDirectory(),
        path: path.join(drivePath, f.name)
      })).filter(f => !f.name.startsWith('$') && f.name !== 'System Volume Information'); // Filter system files
    } catch (e) {
      console.error("[USB] list-usb-files error:", e);
      return [];
    }
  });

  ipcMain.handle('open-usb-file', async (event, filePath) => {
    try {
      if (fs.statSync(filePath).isDirectory()) {
         shell.openPath(filePath);
      } else {
         shell.openPath(filePath);
      }
      
      // Close USB Explorer window after opening a file or folder
      if (usbExplorerWindow) {
        usbExplorerWindow.close();
      }
      
      return { success: true };
    } catch (e) {
      return { success: false, error: e.message };
    }
  });

  createWindow()

  ipcMain.on('open-usb-explorer', (event, drivePath) => {
    createUSBExplorerWindow(drivePath);
  });

  function createUSBExplorerWindow(drivePath) {
    if (usbExplorerWindow) {
      usbExplorerWindow.focus();
      // Update path if different? For now, just focus.
      return;
    }

    const primaryDisplay = screen.getPrimaryDisplay();
    const { width: screenWidth, height: screenHeight } = primaryDisplay.workAreaSize;

    usbExplorerWindow = new BrowserWindow({
      width: 600,
      height: 700,
      x: screenWidth - 650,
      y: 50,
      frame: false,
      transparent: true,
      alwaysOnTop: true,
      skipTaskbar: true,
      webPreferences: {
        preload: path.join(__dirname, 'preload.js'),
        nodeIntegration: false,
        contextIsolation: true
      }
    });

    // Pass the drivePath via query param or just rely on a new IPC
    usbExplorerWindow.loadFile('usb-explorer.html', { query: { drive: drivePath } });

    usbExplorerWindow.on('closed', () => {
      usbExplorerWindow = null;
    });
  }

  ipcMain.on('restore-to-fullscreen', () => {
    if (mainWindow) {
      mainWindow.show();
      mainWindow.restore();
      mainWindow.setFullScreen(true);
    }
  });

  ipcMain.on('app-quit', () => {
    // Restore original wallpaper and explorer before quitting
    if (originalWallpaper) setWallpaper(originalWallpaper);
    exec('start explorer.exe');
    app.quit();
  });

  // --- Phase 3: New Layout & System Control ---

  // 1. Explorer & Wallpaper Control
  ipcMain.on('control-explorer', (event, action) => {
    if (action === 'kill') {
      if (isAprilFools()) {
        console.log("[System] April Fools! Skipping explorer kill.");
        return;
      }
      console.log("[System] Killing explorer.exe");
      exec('taskkill /F /IM explorer.exe');
    } else if (action === 'start') {
      console.log("[System] Starting explorer.exe and restoring wallpaper");
      exec('start explorer.exe');
      if (originalWallpaper) setWallpaper(originalWallpaper);
    }
  });

  // 2. Dynamic Layout Resize
  ipcMain.on('set-window-layout', (event, layout) => {
    if (!mainWindow) return;

    if (layout === 'break') {
      // Fullscreen for Break Time - Acts like a wallpaper
      mainWindow.setFullScreen(true);
      mainWindow.setAlwaysOnTop(false);
      mainWindow.setIgnoreMouseEvents(false); // Enable interaction
    } else if (layout === 'class') {
      // 3/4 Width for Class Time (Right Side)
      const primaryDisplay = screen.getPrimaryDisplay();
      const { width: screenWidth, height: screenHeight } = primaryDisplay.workAreaSize;
      const targetWidth = Math.floor(screenWidth * (0.8)); // 4/5 width

      mainWindow.setFullScreen(false);
      mainWindow.setBounds({
        x: screenWidth - targetWidth,
        y: 0,
        width: targetWidth,
        height: screenHeight
      });
      // Allow interaction in windowed mode
      mainWindow.setAlwaysOnTop(false);
      mainWindow.setIgnoreMouseEvents(false);
    }
  });

  // 3. Forced Wallpaper Generation (Hidden Window Approach)
  async function generateAndSetWallpaper() {
    const { width, height } = screen.getPrimaryDisplay().bounds;
    let wallWin = new BrowserWindow({
      width: width,
      height: height,
      show: false,
      frame: false,
      webPreferences: { offscreen: true }
    });

    // Simple HTML for wallpaper
    const bgHTML = `
      <body style="margin:0; padding:0; overflow:hidden; background:#121212; color:white; font-family:'Malgun Gothic', sans-serif;">
        <div style="display:flex; width:100vw; height:100vh;">
          <div style="width:25%; background:#1a1a1a;"></div>
          <div style="width:75%; display:flex; flex-direction:column; align-items:center; justify-content:center; text-align:center;">
             <h1 style="font-size:50px; margin-bottom:20px;">BoardCast (시간표 프로그램)이 켜지는 중입니다.</h1>
             <p style="font-size:30px; color:#aaa;">안켜진다면 직접 실행해주세요.</p>
          </div>
        </div>
      </body>
    `;

    await wallWin.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(bgHTML)}`);
    const image = await wallWin.webContents.capturePage();
    const wallpaperPath = path.join(app.getPath('userData'), 'forced_wallpaper.png');
    fs.writeFileSync(wallpaperPath, image.toPNG());
    wallWin.close();

    setWallpaper(wallpaperPath);
  }

  ipcMain.on('apply-forced-wallpaper', () => {
    generateAndSetWallpaper().catch(err => console.error("Wallpaper error:", err));
  });

  ipcMain.on('restore-original-wallpaper', () => {
    if (originalWallpaper) setWallpaper(originalWallpaper);
  });

  // Set wallpaper on startup
  generateAndSetWallpaper().catch(err => console.error("Wallpaper error:", err));

  // --- USB System ---
  function startUSBWatcher() {
    console.log("[USB] Watcher started");
    setInterval(() => {
      const cmd = `powershell -command "Get-WmiObject Win32_Volume | Where-Object { $_.DriveType -eq 2 } | Select-Object -ExpandProperty DriveLetter"`;
      exec(cmd, (err, stdout) => {
        if (err) return;
        const currentDrives = stdout.split(/\r?\n/).map(d => d.trim()).filter(Boolean).map(d => d.toUpperCase());
        
        // Detection: New Drives
        currentDrives.forEach(driveUpper => {
          if (!knownDrives.has(driveUpper)) {
            console.log(`[USB] New drive detected: ${driveUpper}`);
            handleNewUSB(driveUpper);
            knownDrives.add(driveUpper);
          }
        });

        // Detection: Removed Drives
        knownDrives.forEach(drive => {
          if (!currentDrives.includes(drive)) {
            console.log(`[USB] Drive removed: ${drive}`);
            knownDrives.delete(drive);
            
            // Notify Renderer
            if (mainWindow) {
              mainWindow.webContents.send('usb-event', { type: 'removed', drive: drive });
            }
            
            // Close Explorer Window if it was looking at this drive
            if (usbExplorerWindow && currentUSBDrive === drive) {
               usbExplorerWindow.close();
               currentUSBDrive = null;
            }
          }
        });
      });
    }, 5000);
  }

  async function getCurrentClassInfo() {
    try {
      const info = readJson(infoPath);
      if (!info || !info.schoolName) return null;

      const now = new Date();
      if (now.getDay() === 0 || now.getDay() === 6) return null;

      const timetable = new Timetable();
      
      const timeoutPromise = (ms, promise) => {
        return new Promise((resolve, reject) => {
          const timeoutId = setTimeout(() => {
            reject(new Error("API Timeout (" + ms + "ms)"));
          }, ms);
          if (!promise || typeof promise.then !== 'function') {
            clearTimeout(timeoutId);
            reject(new Error("Invalid API promise"));
            return;
          }
          promise.then(res => { clearTimeout(timeoutId); resolve(res); })
            .catch(err => { clearTimeout(timeoutId); reject(err); });
        });
      };

      await timeoutPromise(5000, timetable.init());
      const schoolList = await timeoutPromise(5000, timetable.search(info.schoolName));
      if (!schoolList || schoolList.length === 0) return null;
      
      let targetSchool = schoolList[0];
      if (schoolList.length > 1) {
        targetSchool = schoolList.find(s => s.region.includes(info.region)) || schoolList[0];
      }
      
      timetable.setSchool(targetSchool.code);
      const result = await timeoutPromise(5000, timetable.getTimetable());
      const weekData = result[info.grade][info.class] || [];
      const todayData = weekData[now.getDay() - 1] || [];

      const startTime = info.StartTime || "09:00";
      const [stH, stM] = startTime.split(':').map(Number);
      const startMins = stH * 60 + stM;
      const breakMins = info.Breaktime || 10;
      const classDur = info.ClassDuration || 45;
      const lunchDur = info.LunchDuration || 50;

      const curMins = now.getHours() * 60 + now.getMinutes();

      let marker = startMins;
      for (let i = 0; i < todayData.length; i++) {
        if (i === 4) marker += (lunchDur > breakMins ? lunchDur - breakMins : 0);
        const start = marker;
        const end = start + classDur;
        if (curMins >= start && curMins < end) {
          return { subject: todayData[i].subject, grade: info.grade };
        }
        marker = end + breakMins;
      }
    } catch (e) { console.error("[BoardBook] Error calculating class:", e); }
    return null;
  }

  async function launchBoardBook(drivePath) {
    const exePath = path.join(drivePath, "BoardBook-App", "BoardBook.exe");
    if (!fs.existsSync(exePath)) {
      console.warn(`[BoardBook] Executable not found at ${exePath}`);
      return;
    }

    const classInfo = await getCurrentClassInfo();
    const subject = classInfo ? classInfo.subject : "기본";
    const grade = classInfo ? `${classInfo.grade}학년` : "1학년";

    console.log(`[BoardBook] Launching: ${exePath} --subject="${subject}" --grade="${grade}"`);
    const child = spawn(exePath, [`--subject=${subject}`, `--grade=${grade}`], {
      detached: true,
      stdio: 'ignore'
    });
    child.unref();
  }

  function safeJsonRead(filePath) {
    try {
      if (!fs.existsSync(filePath)) return null;
      let content = fs.readFileSync(filePath, 'utf8');
      // Strip UTF-8 BOM if present
      if (content.charCodeAt(0) === 0xFEFF) content = content.slice(1);
      return JSON.parse(content);
    } catch (e) {
      console.error(`[USB] Error parsing JSON at ${filePath}:`, e);
      return null;
    }
  }

  async function handleNewUSB(driveLetter) {
    const drivePath = driveLetter + "\\";
    const usbJsonPath = path.join(drivePath, "Board-USB.json");
    const boardBookJsonPath = path.join(drivePath, "BoardBook.json");
    const boardBookExeSubPath = path.join(drivePath, "BoardBook-App", "BoardBook.exe");
    const updaterExePath = path.join(drivePath, "BoardCastUpdater.exe");

    let usbType = 'unrecognized'; // unrecognized, boardbook, updater

    // 1. Detection Logic
    const config = safeJsonRead(usbJsonPath);
    if (config) {
      const isBoardBook = config.BoardBook === true || config.BoardBook === "true" || config.boardbook === true || config.boardbook === "true";
      const isUpdate = config.Update === true || config.Update === "true" || config.update === true || config.update === "true";

      if (isBoardBook) {
        usbType = 'boardbook';
      } else if (isUpdate) {
        usbType = 'updater';
      }
    } else {
      // Direct file check if no Board-USB.json or parsing failed
      if (fs.existsSync(boardBookJsonPath)) usbType = 'boardbook';
      else if (fs.existsSync(boardBookExeSubPath)) usbType = 'boardbook';
      else if (fs.existsSync(updaterExePath)) usbType = 'updater';
    }

    console.log(`[USB] Detected ${driveLetter} as ${usbType} (config: ${!!config})`);
    currentUSBDrive = driveLetter;

    if (usbType === 'boardbook') {
      if (mainWindow) mainWindow.webContents.send('usb-event', { type: 'boardbook-detected', drive: driveLetter });
      launchBoardBook(drivePath);
    } else if (usbType === 'updater') {
      // Existing update system
      let updateSource = drivePath; 
      if (config && config.UpdatePath) {
        updateSource = path.join(drivePath, config.UpdatePath);
      }
      
      if (mainWindow) mainWindow.webContents.send('usb-event', { type: 'updater-detected', drive: driveLetter });
      if (fs.existsSync(updateSource)) {
        setTimeout(() => { performUSBUpdate(updateSource); }, 3000);
      }
    } else {
      // Unrecognized USB
      if (mainWindow) mainWindow.webContents.send('usb-event', { type: 'inserted', unrecognized: true, drive: driveLetter });
      createUSBExplorerWindow(driveLetter);
    }
  }

  function performUSBUpdate(sourcePath) {
    if (mainWindow) mainWindow.webContents.send('usb-event', { type: 'update-start' });
    const appPath = path.dirname(app.getPath('exe'));
    const userDataPath = app.getPath('userData');
    const batPath = path.join(userDataPath, 'Update.bat');
    
    // Normalize paths: ensure backslashes and no trailing backslashes
    const src = sourcePath.replace(/[\\\/]+$/, '').replace(/\//g, '\\');
    const dst = appPath.replace(/[\\\/]+$/, '').replace(/\//g, '\\');

    // Remove erroneously created C:\Program if it exists
    try {
      if (fs.existsSync('C:\\Program') && dst.toLowerCase() !== 'c:\\program') {
        const stats = fs.statSync('C:\\Program');
        if (stats.isDirectory()) fs.rmdirSync('C:\\Program', { recursive: true });
        else fs.unlinkSync('C:\\Program');
      }
    } catch (e) { console.warn("[Cleanup] Failed to remove C:\\Program", e.message); }

    // ASCII-only script to avoid encoding issues with CMD
    const batContent = `@echo off
echo ========================================
echo [BoardCast Update System]
echo ========================================
echo Source: "${src}"
echo Target: "${dst}"
echo.

:: Wait for app to close
timeout /t 3 /nobreak > nul

echo [Update] Killing processes...
taskkill /F /IM BoardCast.exe /T 2>nul
taskkill /F /IM explorer.exe /T 2>nul

echo [Update] Copying files (Robocopy)...
:: Robocopy exit codes: 0-7 are success.
robocopy "${src}" "${dst}" /E /Z /ZB /R:5 /W:5 /MT:8

if %ERRORLEVEL% GEQ 8 (
    echo.
    echo [Update] ERROR: Robocopy failed with code %ERRORLEVEL%
    pause
    start explorer.exe
    exit /b 1
)

echo [Update] Restarting...
start explorer.exe
cd /d "${dst}"
start "" "BoardCast.exe"

echo [Update] Success!
timeout /t 3 /nobreak > nul
del "%~f0" & exit
`;

    // Writing as ASCII ensures no BOM and maximum compatibility with CMD
    fs.writeFileSync(batPath, batContent, 'ascii');

    // Refined PowerShell command: simple and clean
    const psCmd = `Start-Process cmd -ArgumentList '/c ""${batPath}""' -Verb RunAs`;
    
    console.log("[Update] Executing:", psCmd);
    exec(`powershell -command "${psCmd}"`, (err) => {
      if (err) console.error("[Update] PowerShell failed", err);
      setTimeout(() => { app.quit(); }, 1000);
    });
  }

  startUSBWatcher();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow()
    }
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})
