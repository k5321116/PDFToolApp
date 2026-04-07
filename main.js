const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const { PythonShell } = require('python-shell');
const path = require('path');

function createWindow() {
  const win = new BrowserWindow({
    width: 1000, height: 800,
    webPreferences: { 
      nodeIntegration: true, 
      contextIsolation: false, 
      sandbox: false 
    }
  });
  win.loadFile('index.html');
}

// --- フォルダ選択用 ---
ipcMain.handle('select-directory', async () => {
  const r = await dialog.showOpenDialog({ properties: ['openDirectory'] });
  return r.filePaths[0];
});

// --- 保存先指定用 ---
ipcMain.handle('select-save-path', async () => {
  const r = await dialog.showSaveDialog({ 
    title: '結合後のPDFを保存',
    defaultPath: 'merged_result.pdf', 
    filters: [{ name: 'PDF', extensions: ['pdf'] }] 
  });
  return r.filePath;
});

// --- ファイルとフォルダ両方選択用 ---
ipcMain.handle('select-files-and-dirs', async () => {
  const r = await dialog.showOpenDialog({
    properties: ['openFile', 'openDirectory', 'multiSelections'],
    filters: [{ name: 'PDF Files', extensions: ['pdf'] }]
  });
  return r.filePaths;
});

// --- Word変換の実行 ---
ipcMain.on('run-python', (event, data) => {
  const { exec, spawn } = require('child_process');
  exec('taskkill /f /im WINWORD.EXE', () => {
    const isPackaged = app.isPackaged;
    const scriptPath = isPackaged
      ? path.join(process.resourcesPath, 'python/Word_to_PDF.exe')
      : path.join(__dirname, 'python/Word_to_PDF.py');

    let child;
    if (isPackaged) {
      // exe を直接起動
      child = spawn(scriptPath, [data.sourceDir, data.destDir], { encoding: 'utf8' });
    } else {
      child = spawn('python', ['-u', scriptPath, data.sourceDir, data.destDir], { encoding: 'utf8' });
    }

    let msg = "";
    child.stdout.on('data', (chunk) => {
      const lines = chunk.toString('utf8').split('\n');
      lines.forEach(m => {
        m = m.trim();
        if (!m) return;
        if (m.startsWith('PROGRESS:')) {
          event.reply('python-progress', m.split(':')[1]);
        } else {
          msg = m;
        }
      });
    });

    child.on('close', () => {
      event.reply('python-progress', 100);
      event.reply('python-result', msg || "変換完了");
    });
  });
});
// --- PDF結合の実行 ---
ipcMain.on('run-merge', (event, data) => {
  const { spawn } = require('child_process');
  const isPackaged = app.isPackaged;
  const scriptPath = isPackaged
    ? path.join(process.resourcesPath, 'python/pdf_merger.exe')
    : path.join(__dirname, 'python/pdf_merger.py');

  const pathsString = data.filePaths.join('|');
  const excludeFilesString = data.excludeFiles.join('|');
  const args = [
    pathsString,
    data.savePath,
    data.addPageNum ? "True" : "False",
    data.excludePages || "",
    excludeFilesString || ""
  ];

  let child;
  if (isPackaged) {
    child = spawn(scriptPath, args, { encoding: 'utf8' });
  } else {
    child = spawn('python', ['-u', scriptPath, ...args], { encoding: 'utf8' });
  }

  let lastMsg = "";
  child.stdout.on('data', (chunk) => {
    const lines = chunk.toString('utf8').split('\n');
    lines.forEach(m => {
      m = m.trim();
      if (!m) return;
      if (m.startsWith('PROGRESS:')) {
        event.reply('python-progress', m.split(':')[1]);
      } else {
        lastMsg = m;
      }
    });
  });

  child.on('close', () => {
    event.reply('python-progress', 100);
    event.reply('python-result', lastMsg || "結合完了");
  });
});

app.whenReady().then(createWindow);