'use strict';

const fsp = require('node:fs/promises');
const path = require('node:path');
const crypto = require('node:crypto');
const {Worker} = require('node:worker_threads');

const MAX_INPUT = 64 * 1024 * 1024;
const MAX_OUTPUT = 256 * 1024 * 1024;
const MESSAGES = Object.freeze({
  BUSY: '현재 작업을 마친 뒤 다시 시도해 주세요.',
  NO_INPUT: '먼저 PPTX 파일을 선택해 주세요.',
  NO_RESULT: '먼저 파일을 HTML로 변환해 주세요.',
  INPUT_LIMIT: '비어 있지 않은 64 MiB 이하의 파일을 선택해 주세요.',
  INVALID_INPUT: '선택한 파일을 읽을 수 없습니다. 일반 파일인지 확인해 주세요.',
  INVALID_ARGUMENT: '변환 설정을 확인해 주세요. 슬라이드 범위는 1,3-5 형식으로 입력합니다.',
  PASSWORD_REQUIRED: '이 문서는 열기 암호가 필요합니다. 문서 암호를 입력하고 다시 변환해 주세요.',
  INCORRECT_PASSWORD: '문서 암호가 맞지 않습니다. 확인 후 다시 입력해 주세요.',
  UNKNOWN_CONTAINER: '파일 내부 형식을 읽지 못했습니다. 다른 형식이거나 회사 보안 프로그램이 보호된 데이터를 전달했을 수 있습니다. 이 메시지만으로 DRM 여부를 확정할 수는 없습니다.',
  DRM_PROTECTED: '문서에서 DRM·IRM 보호 정보를 확인했습니다. 이 프로그램의 문서 읽기와 HTML 저장이 회사 정책에서 허용돼야 합니다.',
  LEGACY_PPT: '이 파일은 이전 PowerPoint PPT 형식입니다. PPTX 형식의 문서가 필요합니다.',
  INVALID_CONTAINER: '문서 내부 구조나 무결성 정보를 읽지 못했습니다. 원본 파일 상태를 확인해 주세요.',
  INVALID_PACKAGE: 'PPTX 압축 구조를 끝까지 읽지 못했습니다. 파일이 불완전하거나 지원되지 않는 구조일 수 있습니다.',
  UNSUPPORTED_PACKAGE: '파일에서 PPTX 변환에 필요한 문서 구성을 찾지 못했습니다.',
  UNSUPPORTED_ENCRYPTION: '현재 지원하지 않는 암호화 형식입니다.',
  CANCELLED: '변환을 취소했습니다.',
  TIMEOUT: '변환 시간이 120초를 초과했습니다. 슬라이드 범위를 줄여 다시 시도해 주세요.',
  OUTPUT_LIMIT: '변환 결과가 256 MiB를 초과합니다. 슬라이드 범위를 줄여 주세요.',
  PATH_COLLISION: '원본 문서와 같은 파일에는 저장할 수 없습니다. 다른 이름이나 위치를 선택해 주세요.',
  EACCES: '파일에 접근할 수 없습니다. 파일 읽기 또는 저장 권한을 확인해 주세요.',
  EPERM: '파일 작업이 허용되지 않았습니다. 저장 위치나 회사 보안 정책을 확인해 주세요.',
  ENOENT: '파일이나 폴더를 찾지 못했습니다. 다시 선택해 주세요.',
  ENOSPC: '저장 공간이 부족합니다. 다른 저장 위치를 선택해 주세요.',
  WORKER_EXIT: '변환 작업이 종료됐습니다. 파일이나 슬라이드 범위를 확인하고 다시 시도해 주세요.',
  ERR_WORKER_OUT_OF_MEMORY: '변환에 필요한 메모리가 부족합니다. 슬라이드 범위를 줄여 주세요.'
});
function failure(code, message) { const error = new Error(message || MESSAGES[code] || '변환을 완료하지 못했습니다.'); error.code = code; return error; }
function describe(error) {
  const code = typeof error?.code === 'string' ? error.code.slice(0, 80) : 'CONVERSION_FAILED';
  return {ok: false, code, error: MESSAGES[code] || '작업을 완료하지 못했습니다. 파일과 변환 설정을 확인하고 다시 시도해 주세요.',
    ...(code === 'CANCELLED' ? {cancelled: true} : {})};
}
function progressText(message) {
  if (/^Checking/.test(message)) return '문서 형식을 확인하고 있습니다…';
  if (/^Loading/.test(message)) return '변환 엔진을 준비하고 있습니다…';
  if (/^Sanitizing/.test(message)) return 'HTML을 정리하고 있습니다…';
  const counts = /(?:slide|슬라이드)\s*(\d+)\s*(?:of|\/)\s*(\d+)/i.exec(message);
  return counts ? `슬라이드 변환 중 (${counts[1]} / ${counts[2]})…` : '슬라이드를 HTML로 변환하고 있습니다…';
}
function validateOptions(input) {
  if (!input || typeof input !== 'object' || Array.isArray(input)) throw failure('INVALID_ARGUMENT');
  const {slides = '', includeHidden = false, title = '', password = ''} = input;
  if (typeof slides !== 'string' || slides.length > 2048 ||
      (slides && !/^\s*\d+(?:\s*-\s*\d+)?(?:\s*,\s*\d+(?:\s*-\s*\d+)?)*\s*$/.test(slides)) ||
      typeof includeHidden !== 'boolean' || typeof title !== 'string' || title.length > 300 ||
      typeof password !== 'string' || password.length > 255) throw failure('INVALID_ARGUMENT');
  return {slides, includeHidden, title: title.trim(), password};
}
function comparable(file) { return process.platform === 'win32' ? file.toLowerCase() : file; }

class DesktopController {
  constructor(root, onProgress = () => {}) {
    this.root = root;
    this.onProgress = onProgress;
    this.selected = null;
    this.result = null;
    this.job = null;
    this.disposed = false;
  }
  get busy() { return Boolean(this.job); }
  async selectFile(file) {
    if (this.busy || this.disposed) throw failure('BUSY');
    const input = path.resolve(file);
    const stat = await fsp.stat(input);
    if (!stat.isFile()) throw failure('INVALID_INPUT');
    if (!stat.size || stat.size > MAX_INPUT) throw failure('INPUT_LIMIT');
    const canonical = await fsp.realpath(input);
    if (this.busy || this.disposed) throw failure('BUSY');
    this.selected = {path: input, canonical, name: path.basename(input), size: stat.size};
    this.result = null;
    return {ok: true, file: {name: this.selected.name, size: stat.size}};
  }
  async readSelected(job) {
    const handle = await fsp.open(this.selected.path, 'r');
    const chunks = [];
    let stream;
    try {
      const stat = await handle.stat();
      if (!stat.isFile()) throw failure('INVALID_INPUT');
      if (!stat.size || stat.size > MAX_INPUT) throw failure('INPUT_LIMIT');
      // Keep the identity of the file actually opened. Saving a completed result
      // must still work if that source is later moved or its drive disconnects.
      job.sourceIdentity = {path: this.selected.path, canonical: this.selected.canonical,
        dev: stat.dev, ino: stat.ino};
      if (job.cancelCode) throw failure(job.cancelCode);
      stream = handle.createReadStream({autoClose: false, highWaterMark: 64 * 1024, signal: job.abort.signal});
      let length = 0;
      for await (const chunk of stream) {
        if (job.cancelCode) throw failure(job.cancelCode);
        length += chunk.length;
        if (length > MAX_INPUT) throw failure('INPUT_LIMIT');
        chunks.push(chunk);
      }
      if (!length) throw failure('INPUT_LIMIT');
      return Buffer.concat(chunks, length);
    } finally {
      for (const chunk of chunks) chunk.fill(0);
      if (stream) stream.destroy();
      await handle.close().catch(() => {});
    }
  }
  async convert(input) {
    if (this.busy || this.disposed) throw failure('BUSY');
    if (!this.selected) throw failure('NO_INPUT');
    const options = validateOptions(input);
    // Do not keep the IPC payload's password reference for the duration of conversion.
    input.password = '';
    const job = {abort: new AbortController(), worker: null, reject: null, cancelCode: null};
    this.job = job;
    this.result = null;
    const timer = setTimeout(() => this.stopJob(job, 'TIMEOUT'), 120000);
    try {
      this.onProgress({message: '선택한 파일을 읽고 있습니다…'});
      const bytes = await this.readSelected(job);
      if (job.cancelCode) { bytes.fill(0); throw failure(job.cancelCode); }
      const data = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
      bytes.fill(0);
      const converted = await new Promise((resolve, reject) => {
        let settled = false;
        const finish = (error, result) => {
          if (settled) return;
          settled = true;
          error ? reject(error) : resolve(result);
        };
        job.reject = error => finish(error);
        const worker = new Worker(path.join(this.root, 'src/native-cli-worker.cjs'), {
          workerData: {bytes: data, password: options.password, slides: options.slides,
            includeHidden: options.includeHidden, title: options.title || path.parse(this.selected.name).name},
          transferList: [data], resourceLimits: {maxOldGenerationSizeMb: 768}
        });
        job.worker = worker;
        options.password = '';
        worker.on('message', message => {
          if (settled || this.job !== job || job.cancelCode) return;
          if (message.type === 'progress') this.onProgress({message: progressText(String(message.message || ''))});
          else if (message.type === 'result') finish(null, message);
          else if (message.type === 'error') finish(failure(message.code || 'CONVERSION_FAILED'));
        });
        worker.once('error', error => finish(error));
        worker.once('exit', () => { if (!settled) finish(failure(job.cancelCode || 'WORKER_EXIT')); });
      });
      if (job.cancelCode || this.disposed) throw failure(job.cancelCode || 'CANCELLED');
      if (typeof converted.html !== 'string' || Buffer.byteLength(converted.html, 'utf8') > MAX_OUTPUT) throw failure('OUTPUT_LIMIT');
      if (!Number.isSafeInteger(converted.slideCount) || converted.slideCount < 1) throw failure('CONVERSION_FAILED');
      this.result = {...converted, token: crypto.randomUUID(), sourceIdentity: job.sourceIdentity};
      return {ok: true, slideCount: converted.slideCount, width: converted.width, height: converted.height,
        warnings: converted.warnings || [], sourceSlides: converted.sourceSlides || []};
    } catch (error) {
      throw job.cancelCode ? failure(job.cancelCode) : error;
    } finally {
      options.password = '';
      clearTimeout(timer);
      if (job.worker) await job.worker.terminate().catch(() => {});
      if (this.job === job) this.job = null;
    }
  }
  stopJob(job, code) {
    if (!job || this.job !== job || job.cancelCode) return;
    job.cancelCode = code;
    job.abort.abort();
    job.reject?.(failure(code));
    if (job.worker) job.worker.terminate().catch(() => {});
  }
  cancel() { this.stopJob(this.job, 'CANCELLED'); return {ok: true}; }
  preview() {
    if (this.busy) throw failure('BUSY');
    if (!this.result) throw failure('NO_RESULT');
    return {ok: true, url: `pptx-app://preview/${this.result.token}/index.html`};
  }
  previewHtml(url) {
    if (!this.result || this.busy) return null;
    const parsed = new URL(url);
    return parsed.protocol === 'pptx-app:' && parsed.hostname === 'preview' && !parsed.port && !parsed.username && !parsed.password &&
      parsed.pathname === `/${this.result.token}/index.html` && !parsed.search
      ? this.result.html : null;
  }
  suggestedOutput() {
    if (!this.selected || !this.result) throw failure('NO_RESULT');
    return path.join(path.dirname(this.selected.path), path.parse(this.selected.name).name + '.html');
  }
  async saveHtml(output) {
    if (this.busy || this.disposed) throw failure('BUSY');
    if (!this.result || !this.selected) throw failure('NO_RESULT');
    const source = this.result.sourceIdentity;
    output = path.resolve(output);
    const parentReal = await fsp.realpath(path.dirname(output));
    const targetCanonical = comparable(path.join(parentReal, path.basename(output)));
    if (comparable(source.canonical) === targetCanonical || comparable(source.path) === comparable(output)) throw failure('PATH_COLLISION');
    let targetStat;
    try { targetStat = await fsp.stat(output); } catch (error) { if (error.code !== 'ENOENT') throw error; }
    if (targetStat) {
      if (source.dev === targetStat.dev && source.ino === targetStat.ino) throw failure('PATH_COLLISION');
      if (!targetStat.isFile()) throw failure('INVALID_INPUT');
    }
    const temp = path.join(parentReal, `.pptx-html-${crypto.randomUUID()}.tmp`);
    try {
      await fsp.writeFile(temp, this.result.html, {encoding: 'utf8', flag: 'wx', mode: 0o600});
      await fsp.rename(temp, output);
    } finally { await fsp.unlink(temp).catch(() => {}); }
    return {ok: true, name: path.basename(output), path: output};
  }
  dispose() { this.disposed = true; this.cancel(); this.result = null; }
}

module.exports = {DesktopController, describe, MAX_INPUT, MAX_OUTPUT};
