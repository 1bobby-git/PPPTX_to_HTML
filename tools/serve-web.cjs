#!/usr/bin/env node
'use strict';

// The browser sends documents only to this service. The service launches the
// existing native CLI directly; it never launches PowerPoint, a shell, or a browser.
const http = require('node:http');
const fs = require('node:fs');
const fsp = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');
const crypto = require('node:crypto');
const {spawn} = require('node:child_process');
const {pipeline} = require('node:stream/promises');
const busboy = require('busboy');

const ROOT = path.resolve(__dirname, '..');
const VERSION = require('../package.json').version;
const MAX_INPUT = 64 * 1024 * 1024;
const MAX_OUTPUT = 256 * 1024 * 1024;
const MAX_OPTIONS = 4096;
const MAX_HEADER = 11500;
const ERROR_MESSAGES = Object.freeze({
  INVALID_REQUEST: '요청 형식을 확인하지 못했습니다. 페이지를 새로고침한 뒤 다시 시도해 주세요.',
  INVALID_OPTIONS: '변환 설정이 올바르지 않습니다. 제목·슬라이드 범위·암호를 확인해 주세요.',
  INVALID_INPUT: '비어 있지 않은 .pptx 파일을 선택해 주세요.',
  INPUT_LIMIT: '파일 크기는 최대 64 MiB까지 지원합니다.',
  FORBIDDEN: '이 페이지에서 시작한 요청만 처리할 수 있습니다. 변환 페이지를 다시 열어 주세요.',
  BUSY: '다른 변환이 진행 중입니다. 잠시 뒤 다시 시도해 주세요.',
  UPLOAD_TIMEOUT: '파일 전송 시간이 1분을 초과했습니다. 연결 상태와 파일 크기를 확인해 주세요.',
  TIMEOUT: '변환 시간이 제한을 초과해 중지했습니다. 슬라이드 범위를 줄여 다시 시도해 주세요.',
  CANCELLED: '변환을 취소했습니다.',
  SHUTDOWN: '변환 서비스를 종료하는 중입니다. 잠시 뒤 다시 접속해 주세요.',
  PASSWORD_REQUIRED: '암호로 보호된 Office 문서입니다. 문서를 열 때 사용하는 암호를 입력해 주세요.',
  INCORRECT_PASSWORD: '문서 열기 암호가 맞지 않습니다. 암호를 확인한 뒤 다시 시도해 주세요.',
  DRM_PROTECTED: '조직의 DRM·IRM으로 보호된 문서입니다. 허용된 일반 PPTX 사본이나 조직의 공식 연동이 필요합니다.',
  LEGACY_PPT: '예전 PowerPoint PPT 형식입니다. .pptx 형식으로 저장된 사본이 필요합니다.',
  UNKNOWN_CONTAINER: '파일 내부 형식을 인식하지 못했습니다. 다른 형식이거나 보안 프로그램으로 보호된 문서일 수 있습니다.',
  INVALID_CONTAINER: '문서 구조 또는 암호화 무결성 정보를 확인하지 못했습니다. 원본 파일을 확인해 주세요.',
  INVALID_PACKAGE: 'PPTX 내부 데이터를 읽지 못했습니다. 원본 파일을 확인해 주세요.',
  UNSUPPORTED_PACKAGE: '지원하는 PPTX 구성 요소를 찾지 못했습니다. 원본 문서의 형식을 확인해 주세요.',
  UNSUPPORTED_ENCRYPTION: '이 문서의 암호화 방식은 지원하지 않습니다. 사용이 허용된 일반 PPTX 사본이 필요합니다.',
  INVALID_ARGUMENT: '슬라이드 범위 또는 변환 설정을 확인해 주세요. 원본 슬라이드 번호를 입력해야 합니다.',
  EMPTY_SELECTION: '선택한 범위에 표시할 슬라이드가 없습니다. 범위와 숨긴 슬라이드 설정을 확인해 주세요.',
  SLIDE_LIMIT: '1개 이상, 1,000개 이하의 슬라이드가 있는 PPTX를 선택해 주세요.',
  INVALID_DIMENSIONS: '지원하지 않는 슬라이드 크기입니다.',
  OUTPUT_LIMIT: '변환 결과가 256 MiB 한도를 넘었습니다. 슬라이드 범위를 줄여 다시 시도해 주세요.',
  SLIDE_COUNT_MISMATCH: '일부 슬라이드를 확인하지 못해 변환을 중지했습니다. 다른 범위로 다시 시도해 주세요.',
  CONVERSION_FAILED: '파일을 변환하지 못했습니다. 문서 형식과 선택한 범위를 확인해 주세요.',
  SERVER_ERROR: '변환 서비스에서 오류가 발생했습니다. 관리자에게 서비스 상태 확인을 요청해 주세요.'
});

function failure(code, status = 400) {
  const error = new Error(ERROR_MESSAGES[code] || ERROR_MESSAGES.SERVER_ERROR);
  error.code = ERROR_MESSAGES[code] ? code : 'SERVER_ERROR';
  error.status = status;
  return error;
}
function loopback(host) { return ['127.0.0.1', 'localhost', '::1', '[::1]'].includes(String(host).toLowerCase()); }
function positiveInteger(value, fallback, max) {
  if (value === undefined || value === '') return fallback;
  const number = Number(value);
  if (!Number.isInteger(number) || number < 1 || number > max) throw new Error('Invalid numeric service configuration.');
  return number;
}
function parseOrigin(value) {
  if (!value) return null;
  let parsed;
  try { parsed = new URL(value); } catch { throw new Error('PPTX_WEB_ORIGIN must be an HTTP(S) origin.'); }
  if (!['http:', 'https:'].includes(parsed.protocol) || parsed.username || parsed.password ||
      parsed.pathname !== '/' || parsed.search || parsed.hash) {
    throw new Error('PPTX_WEB_ORIGIN must contain only scheme, hostname, and optional port.');
  }
  return parsed;
}
function validateOptions(value) {
  let data;
  try { data = JSON.parse(value || '{}'); } catch { throw failure('INVALID_OPTIONS'); }
  if (!data || Array.isArray(data) || typeof data !== 'object' ||
      Object.keys(data).some(key => !['title', 'slideRange', 'includeHidden', 'password'].includes(key))) {
    throw failure('INVALID_OPTIONS');
  }
  const options = {title: data.title === undefined ? 'presentation' : data.title,
    slideRange: data.slideRange === undefined ? '' : data.slideRange,
    includeHidden: data.includeHidden === undefined ? false : data.includeHidden,
    password: data.password === undefined ? '' : data.password};
  data.password = '';
  if (typeof options.title !== 'string' || options.title.length > 300 || options.title.includes('\0') ||
      typeof options.slideRange !== 'string' || options.slideRange.length > 2048 ||
      typeof options.includeHidden !== 'boolean' || typeof options.password !== 'string' ||
      options.password.length > 255 || Buffer.byteLength(options.password, 'utf8') > 1022) throw failure('INVALID_OPTIONS');
  if (options.slideRange.trim()) {
    for (const item of options.slideRange.split(',')) {
      const match = /^\s*(\d+)(?:\s*-\s*(\d+))?\s*$/.exec(item);
      const first = match && Number(match[1]);
      const last = match && Number(match[2] || match[1]);
      if (!match || !Number.isSafeInteger(first) || !Number.isSafeInteger(last) || first < 1 || last < first) {
        options.password = '';
        throw failure('INVALID_OPTIONS');
      }
    }
  }
  return options;
}

function receiveUpload(req, inputPath, signal, config) {
  return new Promise((resolve, reject) => {
    let parser;
    try {
      parser = busboy({headers: req.headers, defParamCharset: 'utf8', limits: {
        // busboy emits limit at equality, so one sentinel byte/part allows the
        // documented inclusive limits while the handlers reject the excess.
        fileSize: config.maxInputBytes + 1, files: 1, fields: 1, parts: 3,
        fieldSize: MAX_OPTIONS + 1, fieldNameSize: 30, headerPairs: 50
      }});
    } catch { reject(failure('INVALID_REQUEST')); return; }
    let settled = false;
    let bytes = 0;
    let fileBytes = 0;
    let fileSeen = false;
    let optionsSeen = false;
    let options = null;
    const streams = new Set();
    const writes = [];
    const timer = setTimeout(() => stop(failure('UPLOAD_TIMEOUT', 408)), config.uploadTimeoutMs);
    const onAbort = () => stop(signal.reason instanceof Error ? signal.reason : failure('CANCELLED', 499));
    const onRequestError = () => stop(failure('INVALID_REQUEST'));
    const onData = chunk => {
      bytes += chunk.length;
      if (bytes > config.maxInputBytes + MAX_OPTIONS + 64 * 1024) stop(failure('INPUT_LIMIT', 413));
    };
    function detach() {
      clearTimeout(timer);
      signal.removeEventListener('abort', onAbort);
      req.off('error', onRequestError);
      req.off('data', onData);
    }
    function stop(error) {
      if (settled) return;
      settled = true;
      detach();
      if (options) options.password = '';
      req.unpipe(parser);
      // A limit event runs inside busboy's parser. Defer destruction until that
      // callback finishes so its own stream bookkeeping is not invalidated.
      queueMicrotask(() => {
        for (const stream of streams) stream.destroy();
        parser.destroy();
        // Let the HTTP response close the connection; destroying req here would
        // discard the useful JSON error before it reaches the browser.
        req.resume();
        Promise.all(writes).then(() => reject(error));
      });
    }
    parser.on('file', (name, file, info) => {
      streams.add(file);
      if (settled) { file.resume(); return; }
      if (name !== 'file' || fileSeen || !/\.pptx$/i.test(info.filename || '')) {
        file.resume();
        stop(failure('INVALID_INPUT'));
        return;
      }
      fileSeen = true;
      file.on('data', chunk => {
        fileBytes += chunk.length;
        if (fileBytes > config.maxInputBytes) stop(failure('INPUT_LIMIT', 413));
      });
      file.once('limit', () => stop(failure('INPUT_LIMIT', 413)));
      const output = fs.createWriteStream(inputPath, {flags: 'wx', mode: 0o600});
      streams.add(output);
      // These promises always settle successfully; stop() can await open file
      // handles before the request's finally block removes its directory.
      writes.push(pipeline(file, output).then(() => null, () => { stop(failure('SERVER_ERROR', 500)); return null; }));
    });
    parser.on('field', (name, value, info) => {
      if (settled) return;
      if (name !== 'options' || optionsSeen || info.nameTruncated || info.valueTruncated || Buffer.byteLength(value, 'utf8') > MAX_OPTIONS) {
        stop(failure('INVALID_OPTIONS'));
        return;
      }
      optionsSeen = true;
      try { options = validateOptions(value); } catch (error) { stop(error); }
    });
    for (const event of ['partsLimit', 'filesLimit', 'fieldsLimit']) parser.on(event, () => stop(failure('INVALID_REQUEST')));
    parser.on('error', () => stop(failure('INVALID_REQUEST')));
    parser.on('close', async () => {
      if (settled) return;
      await Promise.all(writes);
      if (settled) return;
      if (!fileSeen || fileBytes < 1) { stop(failure('INVALID_INPUT')); return; }
      settled = true;
      detach();
      resolve(options || validateOptions('{}'));
    });
    req.on('data', onData);
    req.once('error', onRequestError);
    signal.addEventListener('abort', onAbort, {once: true});
    if (signal.aborted) { onAbort(); return; }
    req.pipe(parser);
  });
}

function runCli(input, output, options, signal, config) {
  return new Promise((resolve, reject) => {
    if (signal.aborted) { options.password = ''; reject(signal.reason || failure('CANCELLED', 499)); return; }
    const args = [config.cliPath, input, '-o', output, '--json', '--password-stdin',
      '--timeout', String(Math.max(1, Math.ceil(config.conversionTimeoutMs / 1000)))];
    if (options.title) args.push('--title', options.title);
    if (options.slideRange.trim()) args.push('--slides', options.slideRange.trim());
    if (options.includeHidden) args.push('--include-hidden');
    let child;
    try { child = spawn(process.execPath, args, {cwd: ROOT, shell: false, windowsHide: true, stdio: ['pipe', 'pipe', 'pipe']}); }
    catch { options.password = ''; reject(failure('SERVER_ERROR', 500)); return; }
    let ended = false;
    let reason = null;
    let killTimer;
    let stdoutSize = 0;
    let stderrSize = 0;
    const stdout = [];
    const timer = setTimeout(() => stop(failure('TIMEOUT', 504)), config.conversionTimeoutMs);
    const onAbort = () => stop(signal.reason instanceof Error ? signal.reason : failure('CANCELLED', 499));
    function stop(error) {
      if (ended || reason) return;
      reason = error;
      child.kill('SIGTERM');
      killTimer = setTimeout(() => { if (!ended) child.kill('SIGKILL'); }, config.shutdownGraceMs);
    }
    child.once('error', () => stop(failure('SERVER_ERROR', 500)));
    child.stdin.on('error', error => { if (error.code !== 'EPIPE') stop(failure('SERVER_ERROR', 500)); });
    child.stdout.on('data', chunk => {
      stdoutSize += chunk.length;
      if (stdoutSize > 512 * 1024) stop(failure('SERVER_ERROR', 500));
      else if (!reason) stdout.push(chunk);
    });
    child.stderr.on('data', chunk => {
      stderrSize += chunk.length;
      if (stderrSize > 512 * 1024) stop(failure('SERVER_ERROR', 500));
    });
    child.once('close', code => {
      ended = true;
      clearTimeout(timer);
      clearTimeout(killTimer);
      signal.removeEventListener('abort', onAbort);
      if (reason) { reject(reason); return; }
      let result;
      try { result = JSON.parse(Buffer.concat(stdout).toString('utf8').trim()); }
      catch { reject(failure('SERVER_ERROR', 500)); return; }
      if (code !== 0 || !result || result.ok !== true) {
        if (result && ['UNSUPPORTED_NODE', 'DEPENDENCY_MISSING', 'WORKER_EXIT', 'EXPORTER_ERROR',
          'CHECK_TIMEOUT', 'EACCES', 'ENOENT', 'ENOSPC', 'EIO', 'EMFILE', 'ENFILE'].includes(result.code)) {
          reject(failure('SERVER_ERROR', 500));
          return;
        }
        const known = result && Object.hasOwn(ERROR_MESSAGES, result.code) ? result.code : 'CONVERSION_FAILED';
        reject(failure(known, known === 'TIMEOUT' ? 504 : 422));
        return;
      }
      resolve(result);
    });
    signal.addEventListener('abort', onAbort, {once: true});
    if (signal.aborted) onAbort();
    child.stdin.end(options.password + '\n', 'utf8');
    options.password = '';
  });
}

function metadataHeader(result) {
  if (!Number.isInteger(result.slideCount) || result.slideCount < 1 || result.slideCount > 1000 ||
      ![result.width, result.height].every(value => Number.isFinite(value) && value > 0 && value <= 30000)) {
    throw failure('SERVER_ERROR', 500);
  }
  const sourceSlides = Array.isArray(result.sourceSlides) && result.sourceSlides.length === result.slideCount &&
    result.sourceSlides.every(value => Number.isInteger(value) && value > 0 && value <= 1000) ? result.sourceSlides : null;
  const inputWarnings = Array.isArray(result.warnings) ? result.warnings : [];
  const warnings = inputWarnings.slice(0, 20).map(item => ({
    code: String(item && item.code || 'NOTICE').slice(0, 64),
    message: String(item && item.message || '변환 알림').slice(0, 300)
  }));
  const metadata = {slideCount: result.slideCount, width: result.width, height: result.height, warnings, sourceSlides};
  const truncated = {code: 'MORE_NOTICES', message: '변환 알림이 많아 일부를 생략했습니다. 결과를 확인해 주세요.'};
  let omitted = inputWarnings.length > warnings.length;
  const encode = () => Buffer.from(JSON.stringify(metadata), 'utf8').toString('base64url');
  if (omitted) warnings.push(truncated);
  let encoded = encode();
  while (encoded.length > MAX_HEADER && warnings.length) {
    if (warnings[warnings.length - 1] === truncated) warnings.pop();
    warnings.pop();
    omitted = true;
    warnings.push(truncated);
    encoded = encode();
    if (warnings.length === 1) break;
  }
  if (encoded.length > MAX_HEADER) throw failure('SERVER_ERROR', 500);
  return encoded;
}

function createWebServer(options = {}) {
  const config = {
    host: options.host || '127.0.0.1',
    port: options.port === 0 ? 0 : positiveInteger(options.port, 3080, 65535),
    origin: parseOrigin(options.origin),
    tempRoot: path.resolve(options.tempRoot || os.tmpdir()),
    cliPath: path.resolve(options.cliPath || path.join(__dirname, 'convert-pptx.cjs')),
    maxInputBytes: positiveInteger(options.maxInputBytes, MAX_INPUT, MAX_INPUT),
    uploadTimeoutMs: positiveInteger(options.uploadTimeoutMs, 60000, 60000),
    conversionTimeoutMs: positiveInteger(options.conversionTimeoutMs, 120000, 120000),
    responseTimeoutMs: positiveInteger(options.responseTimeoutMs, 60000, 60000),
    shutdownGraceMs: positiveInteger(options.shutdownGraceMs, 1000, 10000)
  };
  if (!loopback(config.host) && !config.origin) throw new Error('Set PPTX_WEB_ORIGIN before binding the service to a non-loopback interface.');
  const token = crypto.randomBytes(32).toString('hex');
  const mode = loopback(config.host) && (!config.origin || loopback(config.origin.hostname)) ? 'local' : 'server';
  const activeJobs = new Set();
  let closing = false;
  let closePromise;
  const viewerHash = crypto.createHash('sha256').update(fs.readFileSync(path.join(ROOT, 'src/viewer.js'))).digest('base64');
  const csp = "default-src 'none'; script-src 'self' 'sha256-" + viewerHash + "'; style-src 'self' 'unsafe-inline'; " +
    "img-src data: blob:; font-src data:; media-src data: blob:; connect-src 'self'; frame-src blob:; " +
    "worker-src 'none'; object-src 'none'; base-uri 'none'; form-action 'none'; frame-ancestors 'none'";
  const assets = new Map([
    ['/assets/styles.css', ['src/styles.css', 'text/css;charset=utf-8']],
    ['/assets/server-client.js', ['src/server-client.js', 'text/javascript;charset=utf-8']],
    ['/assets/app.js', ['src/app.js', 'text/javascript;charset=utf-8']]
  ]);
  function headers(res) {
    res.setHeader('Cache-Control', 'no-store');
    res.setHeader('X-Content-Type-Options', 'nosniff');
    res.setHeader('Referrer-Policy', 'no-referrer');
    res.setHeader('X-Frame-Options', 'DENY');
    res.setHeader('Content-Security-Policy', csp);
    res.setHeader('Permissions-Policy', 'camera=(), microphone=(), geolocation=()');
  }
  function json(res, status, value, close = false) {
    if (res.destroyed || res.writableEnded) return;
    if (close) res.setHeader('Connection', 'close');
    res.writeHead(status, {'Content-Type': 'application/json;charset=utf-8'});
    res.end(JSON.stringify(value));
  }
  function sendError(res, error, close = false) {
    const code = error && Object.hasOwn(ERROR_MESSAGES, error.code) ? error.code : 'SERVER_ERROR';
    json(res, error && error.status || 500, {ok: false, code, error: ERROR_MESSAGES[code]}, close);
  }
  function allowedOrigin(req) {
    const host = String(req.headers.host || '').toLowerCase();
    if (config.origin) return host === config.origin.host.toLowerCase() ? config.origin.origin : null;
    const address = server.address();
    const port = address && address.port;
    const suffix = port === 80 ? '' : ':' + port;
    const hosts = ['127.0.0.1' + suffix, 'localhost' + suffix, '[::1]' + suffix];
    return hosts.includes(host) ? 'http://' + host : null;
  }
  async function convertRequest(req, res) {
    if (closing) { sendError(res, failure('SHUTDOWN', 503), true); return; }
    if (activeJobs.size) { res.setHeader('Retry-After', '2'); sendError(res, failure('BUSY', 429), true); return; }
    const contentLength = req.headers['content-length'];
    if (contentLength && (!/^\d+$/.test(contentLength) || Number(contentLength) > config.maxInputBytes + MAX_OPTIONS + 64 * 1024)) {
      sendError(res, failure('INPUT_LIMIT', 413), true);
      return;
    }
    const controller = new AbortController();
    let resolveDone;
    const job = {controller, done: new Promise(resolve => { resolveDone = resolve; })};
    activeJobs.add(job);
    const abort = () => controller.abort(failure('CANCELLED', 499));
    const closed = () => { if (!res.writableFinished) abort(); };
    req.once('aborted', abort);
    res.once('close', closed);
    let directory;
    let conversionOptions;
    let responseDeadline;
    try {
      directory = await fsp.mkdtemp(path.join(config.tempRoot, 'pptx-web-'));
      await fsp.chmod(directory, 0o700);
      const input = path.join(directory, 'input.pptx');
      const output = path.join(directory, 'output.html');
      conversionOptions = await receiveUpload(req, input, controller.signal, config);
      const result = await runCli(input, output, conversionOptions, controller.signal, config);
      const stat = await fsp.lstat(output);
      if (!stat.isFile() || stat.size < 1 || stat.size > MAX_OUTPUT) throw failure('OUTPUT_LIMIT', 422);
      if (controller.signal.aborted) throw controller.signal.reason;
      const metadata = metadataHeader(result);
      responseDeadline = setTimeout(() => controller.abort(failure('TIMEOUT', 504)), config.responseTimeoutMs);
      res.writeHead(200, {'Content-Type': 'text/html;charset=utf-8', 'Content-Length': stat.size,
        'Content-Disposition': 'attachment; filename="presentation.html"', 'X-Pptx-Metadata': metadata});
      await pipeline(fs.createReadStream(output), res, {signal: controller.signal});
    } catch (error) {
      if (!res.headersSent) sendError(res, error, true);
      else if (!res.writableFinished) res.destroy();
    } finally {
      clearTimeout(responseDeadline);
      if (conversionOptions) conversionOptions.password = '';
      req.off('aborted', abort);
      res.off('close', closed);
      if (directory) {
        await fsp.rm(directory, {recursive: true, force: true, maxRetries: 3, retryDelay: 100}).catch(() => {
          console.error('Temporary conversion files could not be removed. Check service storage permissions.');
        });
      }
      activeJobs.delete(job);
      resolveDone();
    }
  }
  const server = http.createServer(async (req, res) => {
    headers(res);
    try {
      const origin = allowedOrigin(req);
      if (!origin) { sendError(res, failure('FORBIDDEN', 403), true); return; }
      if (req.url === '/api/convert' && req.method === 'POST') {
        if (req.headers.origin !== origin || req.headers['x-pptx-token'] !== token ||
            req.headers['sec-fetch-site'] === 'cross-site') {
          sendError(res, failure('FORBIDDEN', 403), true);
          return;
        }
        await convertRequest(req, res);
        return;
      }
      if (req.method !== 'GET' && req.method !== 'HEAD') {
        res.setHeader('Allow', 'GET, HEAD');
        json(res, 405, {ok: false, code: 'METHOD_NOT_ALLOWED', error: '지원하지 않는 요청 방식입니다.'}, true);
        return;
      }
      if (req.url === '/api/status') {
        if (req.method === 'HEAD') { res.writeHead(200); res.end(); return; }
        json(res, 200, {ok: true, engine: 'native', version: VERSION, mode, csrfToken: token, busy: activeJobs.size > 0,
          maxInputBytes: config.maxInputBytes, conversionTimeoutMs: config.conversionTimeoutMs,
          uploadTimeoutMs: config.uploadTimeoutMs, responseTimeoutMs: config.responseTimeoutMs});
        return;
      }
      let body;
      let type;
      if (req.url === '/' || req.url === '/index.html') {
        const shell = await fsp.readFile(path.join(ROOT, 'src/web-shell.html'), 'utf8');
        if (shell.split('<!-- WEB_CSRF -->').length !== 2) throw failure('SERVER_ERROR', 500);
        body = Buffer.from(shell.replace('<!-- WEB_CSRF -->', '<meta name="pptx-csrf" content="' + token + '">'));
        type = 'text/html;charset=utf-8';
      } else if (assets.has(req.url)) {
        const [file, contentType] = assets.get(req.url);
        body = await fsp.readFile(path.join(ROOT, file));
        type = contentType;
      } else if (req.url === '/favicon.ico') {
        res.writeHead(204); res.end(); return;
      } else {
        json(res, 404, {ok: false, code: 'NOT_FOUND', error: '페이지를 찾을 수 없습니다.'});
        return;
      }
      res.writeHead(200, {'Content-Type': type, 'Content-Length': body.length});
      res.end(req.method === 'HEAD' ? undefined : body);
    } catch (error) { if (!res.headersSent) sendError(res, error, true); else res.destroy(); }
  });
  server.requestTimeout = config.uploadTimeoutMs + 5000;
  server.headersTimeout = Math.min(15000, server.requestTimeout);
  server.keepAliveTimeout = 5000;
  function close() {
    if (closePromise) return closePromise;
    closePromise = (async () => {
      closing = true;
      const stopped = new Promise(resolve => server.close(() => resolve()));
      for (const job of activeJobs) job.controller.abort(failure('SHUTDOWN', 503));
      server.closeIdleConnections();
      await Promise.all([...activeJobs].map(job => job.done));
      server.closeAllConnections();
      await stopped;
    })();
    return closePromise;
  }
  return {server, close, token, getState: () => ({busy: activeJobs.size > 0, activeJobs: activeJobs.size}), config};
}

async function main() {
  if (process.argv.slice(2).some(arg => arg === '--help' || arg === '-h')) {
    process.stdout.write('Usage: node tools/serve-web.cjs\n\nEnvironment: PPTX_WEB_HOST (127.0.0.1), PPTX_WEB_PORT or PORT (3080),\nPPTX_WEB_ORIGIN (required for non-loopback binding). End users open the web URL.\n');
    return;
  }
  if (process.argv.length > 2) throw new Error('Use environment variables to configure the web service; see --help.');
  const instance = createWebServer({host: process.env.PPTX_WEB_HOST || '127.0.0.1',
    port: process.env.PPTX_WEB_PORT || process.env.PORT || 3080, origin: process.env.PPTX_WEB_ORIGIN});
  let stopping = false;
  const stop = () => { if (!stopping) { stopping = true; instance.close().catch(() => { process.exitCode = 1; }); } };
  process.once('SIGINT', stop);
  process.once('SIGTERM', stop);
  await new Promise((resolve, reject) => {
    instance.server.once('error', reject);
    instance.server.listen(instance.config.port, instance.config.host, resolve);
  });
  const boundHost = instance.config.host.includes(':') ? '[' + instance.config.host + ']' : instance.config.host;
  const displayOrigin = instance.config.origin ? instance.config.origin.origin : 'http://' + boundHost + ':' + instance.server.address().port;
  process.stdout.write('PPTX web converter v' + VERSION + ' ready: ' + displayOrigin + '\n');
}

module.exports = {createWebServer};
if (require.main === module) main().catch(error => {
  process.stderr.write('Could not start PPTX web service: ' + error.message + '\n');
  process.exitCode = 1;
});
