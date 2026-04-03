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
  // 実行前に裏で残っているWordを掃除する
  const { exec } = require('child_process');
  exec('taskkill /f /im WINWORD.EXE', () => {
    
    const isPackaged = app.isPackaged;
    const scriptPath = isPackaged 
        ? path.join(process.resourcesPath, 'python/Word_to_PDF.exe') // 配布時
        : path.join(__dirname, 'python/Word_to_PDF.py');            // 開発時

    let options = {
        // パッケージ済み（exe使用）の場合は pythonPath を空にするなどの調整が必要な場合があります
        mode: 'text',
        args: [data.sourceDir, data.destDir]
    };
    if (isPackaged) {
        options.pythonPath = scriptPath; // exe自体をインタープリタとして指定
        // あるいは直接外部プロセスとして起動する設定
    }
    // 呼び出し部分
    let py = new PythonShell(scriptPath, options);
    let msg = "";

    py.on('message', m => {
      if (m.startsWith('PROGRESS:')) {
        event.reply('python-progress', m.split(':')[1]);
      } else {
        msg = m;
      }
    });

    py.end(err => {
      event.reply('python-progress', 100);
      event.reply('python-result', (err && !msg.includes('✅') && !msg.includes('完了')) ? "変換エラー" : msg);
    });
  });
});
// --- PDF結合の実行 ---
ipcMain.on('run-merge', (event, data) => {
  const isPackaged = app.isPackaged;
  
  // 実行するパスの切り分け
  const scriptPath = isPackaged 
      ? path.join(process.resourcesPath, 'python/pdf_merger.exe') 
      : path.join(__dirname, 'python/pdf_merger.py');

  const pathsString = data.filePaths.join('|'); 
  const excludeFilesString = data.excludeFiles.join('|'); 
  
  let options = {
    mode: 'text',
    pythonOptions: isPackaged ? [] : ['-u'],
    args: [
        pathsString, 
        data.savePath, 
        data.addPageNum ? "True" : "False",
        data.excludePages || "",
        excludeFilesString || ""
    ]
  };

  // パッケージ化時は pythonPath を exe 自身に向ける（Python本体不要にするため）
  if (isPackaged) {
    options.pythonPath = scriptPath;
  }
  let pyShell = new PythonShell(scriptPath, options);
  let lastMsg = "";

  pyShell.on('message', (m) => {
    if (m.startsWith('PROGRESS:')) {
      event.reply('python-progress', m.split(':')[1]);
    } else {
      lastMsg = m;
    }
  });

  pyShell.end((err) => {
    event.reply('python-progress', 100);
    event.reply('python-result', (err && !lastMsg.includes('✅')) ? "結合エラー" : lastMsg);
  });
});

app.whenReady().then(createWindow);