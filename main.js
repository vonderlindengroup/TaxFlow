const { app, BrowserWindow, Menu, shell, ipcMain, dialog } = require('electron');
const path  = require('path');
const fs    = require('fs');
const os    = require('os');
const { pathToFileURL } = require('url');

let mainWindow;
const isDev = process.argv.includes('--dev');

const gotLock = app.requestSingleInstanceLock();
if (!gotLock) { app.quit(); }
else {
  app.on('second-instance', () => {
    if (mainWindow) { if (mainWindow.isMinimized()) mainWindow.restore(); mainWindow.focus(); }
  });
}

// ── pdfjs ─────────────────────────────────────────────────────────────────
let _pdfjs = null;
async function getPdfjs() {
  if (_pdfjs) return _pdfjs;
  const pdfjsBase   = path.join(__dirname, 'node_modules', 'pdfjs-dist', 'legacy', 'build');
  const pdfjsMain   = path.join(pdfjsBase, 'pdf.mjs');
  const pdfjsWorker = path.join(pdfjsBase, 'pdf.worker.mjs');
  if (!fs.existsSync(pdfjsMain)) { console.error('pdfjs-dist not found'); return null; }
  try {
    if (typeof globalThis.DOMMatrix === 'undefined') {
      globalThis.DOMMatrix = class { constructor(){this.a=1;this.b=0;this.c=0;this.d=1;this.e=0;this.f=0;} };
    }
    if (typeof globalThis.ImageData === 'undefined') {
      globalThis.ImageData = class { constructor(w,h){this.width=w;this.height=h;this.data=new Uint8ClampedArray(w*h*4);} };
    }
    if (typeof globalThis.Path2D === 'undefined') { globalThis.Path2D = class {}; }
    _pdfjs = await import(pathToFileURL(pdfjsMain).href);
    _pdfjs.GlobalWorkerOptions.workerSrc = pathToFileURL(pdfjsWorker).href;
    console.log('pdfjs-dist loaded, version:', _pdfjs.version);
    return _pdfjs;
  } catch(e) { console.error('pdfjs load error:', e.message); return null; }
}

async function extractPdfText(buffer) {
  const pdfjs = await getPdfjs();
  if (!pdfjs) return null;
  try {
    const data = new Uint8Array(buffer);
    const pdf  = await pdfjs.getDocument({ data, useSystemFonts:true, verbosity:0, disableFontFace:true, isEvalSupported:false, useWorkerFetch:false }).promise;
    let fullText = '';
    for (let p = 1; p <= pdf.numPages; p++) {
      const page = await pdf.getPage(p);
      const ct   = await page.getTextContent();
      const lineMap = {};
      ct.items.forEach(item => {
        const y = Math.round(item.transform[5]);
        if (!lineMap[y]) lineMap[y] = [];
        lineMap[y].push({ x: Math.round(item.transform[4]), t: item.str });
      });
      Object.keys(lineMap).sort((a,b)=>b-a).forEach(y => {
        fullText += lineMap[y].sort((a,b)=>a.x-b.x).map(i=>i.t).join(' ') + '\n';
      });
    }
    return fullText;
  } catch(e) { console.error('PDF parse error:', e.message); return null; }
}

// ── Window ────────────────────────────────────────────────────────────────
function createWindow() {
  mainWindow = new BrowserWindow({
    width:1280, height:820, minWidth:960, minHeight:640, title:'TaxFlow',
    icon: path.join(__dirname,'assets', process.platform==='win32'?'icon.ico':process.platform==='darwin'?'icon.icns':'icon.png'),
    backgroundColor:'#f8fafc', show:false,
    webPreferences:{ preload:path.join(__dirname,'preload.js'), contextIsolation:true, nodeIntegration:false, sandbox:false, webSecurity:false },
    titleBarStyle: process.platform==='darwin'?'hiddenInset':'default',
    trafficLightPosition:{x:16,y:16},
  });
  mainWindow.loadFile(path.join(__dirname,'src','index.html'));
  mainWindow.once('ready-to-show', ()=>{ mainWindow.show(); });
  mainWindow.webContents.setWindowOpenHandler(({url})=>{ shell.openExternal(url); return {action:'deny'}; });
  mainWindow.on('closed', ()=>{ mainWindow=null; });
}

// ── IPC ────────────────────────────────────────────────────────────────────
ipcMain.handle('ping', async () => {
  const pdfjsExists = fs.existsSync(path.join(__dirname,'node_modules','pdfjs-dist','legacy','build','pdf.mjs'));
  const xlsxExists  = fs.existsSync(path.join(__dirname,'node_modules','xlsx'));
  return { ok:true, appVersion:'1.0.3', platform:process.platform, nodeVersion:process.version, pdfjsExists, xlsxExists };
});

ipcMain.handle('extract-pdf-text', async (event, {fileName, buffer}) => {
  try {
    const buf = Buffer.isBuffer(buffer)?buffer:buffer instanceof Uint8Array?Buffer.from(buffer):Buffer.from(Object.values(buffer));
    const text = await extractPdfText(buf);
    if (!text) return {success:false, error:'Could not extract text'};
    return {success:true, text};
  } catch(e) { return {success:false, error:e.message}; }
});

ipcMain.handle('parse-xlsx', async (event, {fileName, buffer}) => {
  try {
    const XLSX = require(path.join(__dirname,'node_modules','xlsx'));
    const buf  = Buffer.isBuffer(buffer)?buffer:buffer instanceof Uint8Array?Buffer.from(buffer):Buffer.from(Object.values(buffer));
    const wb   = XLSX.read(buf, {type:'buffer'});
    const csv  = XLSX.utils.sheet_to_csv(wb.Sheets[wb.SheetNames[0]]);
    return {success:true, csv};
  } catch(e) { return {success:false, error:e.message}; }
});

ipcMain.handle('save-file', async (event, {defaultName, buffer, mimeType}) => {
  const ext = defaultName.split('.').pop();
  const {canceled, filePath} = await dialog.showSaveDialog(mainWindow, {
    title:'Save File',
    defaultPath:path.join(app.getPath('documents'), defaultName),
    filters: ext==='xlsx'?[{name:'Excel Workbook',extensions:['xlsx']}]:[{name:'CSV File',extensions:['csv']}],
  });
  if (canceled||!filePath) return {success:false};
  try { fs.writeFileSync(filePath, Buffer.from(buffer)); return {success:true, filePath}; }
  catch(e) { return {success:false, error:e.message}; }
});

ipcMain.handle('read-file', async (event, filePath) => {
  try {
    const buf   = fs.readFileSync(filePath);
    const ext   = filePath.toLowerCase().split('.').pop();
    const fname = path.basename(filePath);
    if (ext==='pdf') {
      const text = await extractPdfText(buf);
      return {success:true, isPdf:true, text:text||'', name:fname};
    }
    if (ext==='xlsx'||ext==='xls') {
      const XLSX = require(path.join(__dirname,'node_modules','xlsx'));
      const wb   = XLSX.read(buf, {type:'buffer'});
      const csv  = XLSX.utils.sheet_to_csv(wb.Sheets[wb.SheetNames[0]]);
      return {success:true, isXlsx:true, csv, name:fname};
    }
    return {success:true, text:buf.toString('utf8'), name:fname};
  } catch(e) { return {success:false, error:e.message}; }
});

// ── Menu ───────────────────────────────────────────────────────────────────
function buildMenu() {
  const isMac = process.platform==='darwin';
  const template = [
    ...(isMac?[{label:app.name,submenu:[{role:'about',label:'About TaxFlow'},{type:'separator'},{label:'Preferences…',accelerator:'Cmd+,',click:()=>mainWindow.webContents.executeJavaScript("showTab('settings')")},{type:'separator'},{role:'services'},{type:'separator'},{role:'hide'},{role:'hideOthers'},{role:'unhide'},{type:'separator'},{role:'quit',label:'Quit TaxFlow'}]}]:[]),
    {label:'File',submenu:[
      {label:'Import Bank Statement…',accelerator:'CmdOrCtrl+O',click:async()=>{
        const {canceled,filePaths}=await dialog.showOpenDialog(mainWindow,{title:'Import Bank Statement',filters:[{name:'Bank Statements',extensions:['pdf','csv','xlsx','xls']}],properties:['openFile','multiSelections']});
        if(!canceled&&filePaths.length) mainWindow.webContents.send('import-files',filePaths);
      }},
      {type:'separator'},
      {label:'Export to Excel…',accelerator:'CmdOrCtrl+E',click:()=>mainWindow.webContents.executeJavaScript('exportAll()')},
      {label:'Export to CSV…',accelerator:'CmdOrCtrl+Shift+E',click:()=>mainWindow.webContents.executeJavaScript('exportCsv()')},
      {type:'separator'},
      isMac?{role:'close'}:{role:'quit',label:'Exit TaxFlow'},
    ]},
    {label:'Edit',submenu:[{role:'undo'},{role:'redo'},{type:'separator'},{role:'cut'},{role:'copy'},{role:'paste'},{role:'selectAll'}]},
    {label:'View',submenu:[
      {label:'Statements',accelerator:'CmdOrCtrl+1',click:()=>mainWindow.webContents.executeJavaScript("showTab('statements')")},
      {label:'Summary',accelerator:'CmdOrCtrl+2',click:()=>mainWindow.webContents.executeJavaScript("showTab('summary')")},
      {type:'separator'},{role:'reload'},{role:'forceReload'},{type:'separator'},
      {role:'resetZoom'},{role:'zoomIn'},{role:'zoomOut'},{type:'separator'},{role:'togglefullscreen'},
      ...(isDev?[{type:'separator'},{role:'toggleDevTools'}]:[]),
    ]},
    {label:'Window',submenu:[{role:'minimize'},{role:'zoom'},...(isMac?[{type:'separator'},{role:'front'}]:[{role:'close'}])]},
    {role:'help',submenu:[
      {label:'TaxFlow Website',click:()=>shell.openExternal('https://gettaxflow.com')},
      {label:'Contact Support',click:()=>shell.openExternal('mailto:support@taxflow.app')},
      {type:'separator'},{label:`Version ${app.getVersion()}`,enabled:false},
    ]},
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

app.whenReady().then(()=>{ buildMenu(); createWindow(); getPdfjs(); app.on('activate',()=>{ if(BrowserWindow.getAllWindows().length===0) createWindow(); }); });
app.on('window-all-closed',()=>{ if(process.platform!=='darwin') app.quit(); });
app.on('web-contents-created',(event,contents)=>{ contents.on('will-navigate',(event,url)=>{ if(!url.startsWith('file://')) { event.preventDefault(); shell.openExternal(url); } }); });
