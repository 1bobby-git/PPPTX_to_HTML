'use strict';

const {app, BrowserWindow, dialog, ipcMain, Menu, protocol, session} = require('electron');
const fs = require('node:fs/promises');
const path = require('node:path');
const {DesktopController, describe} = require('./controller.cjs');
const root = path.resolve(__dirname, '..');
const appPackage = require(path.join(root, 'package.json'));
const APP_URL = 'pptx-app://desktop/index.html';

protocol.registerSchemesAsPrivileged([{scheme: 'pptx-app', privileges: {standard: true, secure: true}}]);
app.setName('PPTX to HTML');
app.setAppUserModelId('kr.ppptx.desktop');

let window;
let controller;
let dialogActive = false;
const busy = () => dialogActive || controller?.busy;
const error = code => { const e = new Error(code); e.code = code; return e; };

function trusted(event) {
  return window && !window.isDestroyed() && event.sender === window.webContents &&
    event.senderFrame === window.webContents.mainFrame && event.senderFrame.url.split('#')[0] === APP_URL;
}
function register(name, callback) {
  ipcMain.handle('pptx-desktop:' + name, async (event, ...args) => {
    if (!trusted(event)) return {ok: false, code: 'INVALID_CALLER', error: '허용되지 않은 요청입니다.'};
    try { return await callback(...args); }
    catch (failure) { return describe(failure); }
  });
}
async function installProtocol(browserSession) {
  const assets = new Map([
    ['/index.html', ['index.html', 'text/html; charset=utf-8']],
    ['/renderer.js', ['renderer.js', 'text/javascript; charset=utf-8']],
    ['/styles.css', ['styles.css', 'text/css; charset=utf-8']]
  ]);
  browserSession.protocol.handle('pptx-app', async request => {
    const url = new URL(request.url);
    const headers = {'X-Content-Type-Options': 'nosniff', 'Cache-Control': 'no-store'};
    if (request.method !== 'GET' || url.username || url.password || url.port) return new Response('Not allowed', {status: 405, headers});
    if (url.hostname === 'desktop' && assets.has(url.pathname) && !url.search) {
      const [name, type] = assets.get(url.pathname);
      try { return new Response(await fs.readFile(path.join(__dirname, name)), {headers: {...headers, 'Content-Type': type}}); }
      catch { return new Response('Application file unavailable', {status: 500, headers}); }
    }
    if (url.hostname === 'preview') {
      const html = controller.previewHtml(request.url);
      if (html) return new Response(html, {headers: {...headers, 'Content-Type': 'text/html; charset=utf-8'}});
    }
    return new Response('Not found', {status: 404, headers});
  });
}
async function start() {
  Menu.setApplicationMenu(null);
  const browserSession = session.fromPartition('pptx-desktop');
  browserSession.setPermissionRequestHandler((_contents, _permission, callback) => callback(false));
  browserSession.setPermissionCheckHandler(() => false);
  browserSession.on('will-download', event => event.preventDefault());
  browserSession.webRequest.onBeforeRequest({urls: ['http://*/*', 'https://*/*', 'ws://*/*', 'wss://*/*', 'ftp://*/*']},
    (_details, callback) => callback({cancel: true}));
  controller = new DesktopController(root, message => {
    if (window && !window.isDestroyed()) window.webContents.send('pptx-desktop:progress', message);
  });
  await installProtocol(browserSession);
  window = new BrowserWindow({
    title: 'PPTX to HTML', width: 1180, height: 840, minWidth: 780, minHeight: 640,
    backgroundColor: '#f5f7fb', show: false,
    webPreferences: {session: browserSession, preload: path.join(__dirname, 'preload.cjs'),
      nodeIntegration: false, nodeIntegrationInWorker: false, nodeIntegrationInSubFrames: false,
      contextIsolation: true, sandbox: true, webviewTag: false, webSecurity: true,
      navigateOnDragDrop: false, spellcheck: false}
  });
  window.webContents.setWindowOpenHandler(() => ({action: 'deny'}));
  window.webContents.on('will-navigate', event => event.preventDefault());
  window.webContents.on('will-attach-webview', event => event.preventDefault());
  window.webContents.on('will-frame-navigate', event => {
    const url = event.url;
    if (url === APP_URL || url === 'about:blank') return;
    try { if (controller.previewHtml(url)) return; } catch {}
    event.preventDefault();
  });
  window.on('closed', () => { controller.dispose(); window = null; app.quit(); });
  register('info', async () => {
    const engine = JSON.parse(await fs.readFile(path.join(root, 'node_modules/@briank-dev/pptx-to-html/package.json'), 'utf8'));
    return {version: appPackage.version, engineVersion: engine.version, platform: process.platform};
  });
  register('choose', async () => {
    if (busy()) throw error('BUSY');
    dialogActive = true;
    try {
      const selected = await dialog.showOpenDialog(window, {title: '변환할 PPTX 파일 선택',
        properties: ['openFile'], filters: [{name: 'PowerPoint 문서', extensions: ['pptx']}, {name: '모든 파일', extensions: ['*']}]});
      if (selected.canceled || !selected.filePaths.length) return {ok: false, cancelled: true};
      return await controller.selectFile(selected.filePaths[0]);
    } finally { dialogActive = false; }
  });
  register('convert', async options => { if (dialogActive) throw error('BUSY'); return controller.convert(options); });
  register('cancel', () => controller.cancel());
  register('preview', () => controller.preview());
  register('save', async () => {
    if (busy()) throw error('BUSY');
    const defaultPath = controller.suggestedOutput();
    dialogActive = true;
    try {
      const saved = await dialog.showSaveDialog(window, {title: 'HTML 저장', defaultPath,
        filters: [{name: 'HTML 문서', extensions: ['html']}], properties: ['showOverwriteConfirmation']});
      if (saved.canceled || !saved.filePath) return {ok: false, cancelled: true};
      const output = /\.html?$/i.test(saved.filePath) ? saved.filePath : saved.filePath + '.html';
      // The native overwrite prompt covers only the path returned by the dialog.
      // If the extension changes it, confirm the actual destination separately.
      if (output !== saved.filePath) {
        let exists = false;
        try { await fs.stat(output); exists = true; }
        catch (failure) { if (failure.code !== 'ENOENT') throw failure; }
        if (exists) {
          const choice = await dialog.showMessageBox(window, {type: 'question',
            title: '기존 HTML 파일 바꾸기', message: path.basename(output) + ' 파일이 이미 있습니다.',
            detail: 'HTML 확장자를 붙인 최종 저장 위치의 파일을 바꾸시겠습니까?',
            buttons: ['취소', '덮어쓰기'], defaultId: 0, cancelId: 0, noLink: true});
          if (choice.response !== 1) return {ok: false, cancelled: true};
        }
      }
      return await controller.saveHtml(output);
    } finally { dialogActive = false; }
  });
  window.once('ready-to-show', () => window.show());
  await window.loadURL(APP_URL);
}

if (!app.requestSingleInstanceLock()) app.quit();
else {
  app.on('second-instance', () => { if (window) { if (window.isMinimized()) window.restore(); window.show(); window.focus(); } });
  app.on('before-quit', () => controller?.dispose());
  app.on('window-all-closed', () => app.quit());
  app.whenReady().then(start).catch(() => {
    dialog.showErrorBox('PPTX to HTML', '프로그램을 시작하지 못했습니다. 압축을 모두 푼 폴더에서 실행 파일을 열어 주세요.');
    app.exit(1);
  });
}
