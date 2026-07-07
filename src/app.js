const WAYBILL_PATTERN = /JDM\d{12}/i;

const elements = {
  dataCount: document.querySelector('#data-count'),
  dataSource: document.querySelector('#data-source'),
  excelInput: document.querySelector('#excel-input'),
  video: document.querySelector('#camera-preview'),
  placeholder: document.querySelector('#camera-placeholder'),
  scanButton: document.querySelector('#scan-button'),
  stopButton: document.querySelector('#stop-button'),
  scanStatus: document.querySelector('#scan-status'),
  manualInput: document.querySelector('#manual-input'),
  manualButton: document.querySelector('#manual-button'),
  resultCard: document.querySelector('#result-card'),
  resultLabel: document.querySelector('#result-label'),
  waybillNumber: document.querySelector('#waybill-number'),
  segmentCode: document.querySelector('#segment-code'),
  resultNote: document.querySelector('#result-note')
};

let waybillMap = new Map();
let mediaStream = null;
let codeReader = null;
let scanTimer = null;
let lastScanValue = '';
let audioContext = null;

document.addEventListener('DOMContentLoaded', () => {
  bindEvents();
  loadEmbeddedData();
});

function bindEvents() {
  elements.scanButton.addEventListener('click', startScanner);
  elements.stopButton.addEventListener('click', stopScanner);
  elements.manualButton.addEventListener('click', () => lookup(elements.manualInput.value, '手动查询'));
  elements.manualInput.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      lookup(elements.manualInput.value, '手动查询');
    }
  });
  elements.excelInput.addEventListener('change', handleExcelUpload);
}

function loadEmbeddedData() {
  const records = Array.isArray(window.WAYBILL_SEGMENTS) ? window.WAYBILL_SEGMENTS : [];
  const nextMap = new Map();

  records.forEach((record) => {
    const waybill = normalizeWaybill(record.waybill);
    const segment = String(record.segment || '').trim();
    if (waybill && segment) {
      nextMap.set(waybill, segment);
    }
  });

  if (!nextMap.size) {
    elements.dataCount.textContent = '待导入';
    elements.scanStatus.textContent = '内置数据未载入，请点击“重新导入 Excel”。';
    return;
  }

  waybillMap = nextMap;
  elements.dataCount.textContent = `${waybillMap.size} 条`;
  elements.dataSource.textContent = '已载入：内置运单段码对照表';
  elements.scanStatus.textContent = '数据已就绪，可以开始扫码。';
}

function waitForXlsx() {
  return new Promise((resolve, reject) => {
    const startedAt = Date.now();
    const timer = window.setInterval(() => {
      if (window.XLSX) {
        window.clearInterval(timer);
        resolve();
      } else if (Date.now() - startedAt > 8000) {
        window.clearInterval(timer);
        reject(new Error('Excel 解析库加载失败，请检查网络后刷新。'));
      }
    }, 80);
  });
}

async function handleExcelUpload(event) {
  const [file] = event.target.files;
  if (!file) {
    return;
  }

  try {
    await waitForXlsx();
    const buffer = await file.arrayBuffer();
    loadWorkbook(buffer, file.name);
  } catch (error) {
    elements.scanStatus.textContent = error.message || 'Excel 导入失败。';
  }
}

function loadWorkbook(buffer, sourceName) {
  const workbook = window.XLSX.read(buffer, { type: 'array' });
  const firstSheetName = workbook.SheetNames[0];
  const rows = window.XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], { defval: '' });
  const nextMap = new Map();

  rows.forEach((row) => {
    const waybill = normalizeWaybill(row['运单号'] || row['单号'] || row['waybill'] || row['Waybill']);
    const segment = String(row['正确段码'] || row['段码'] || row['segment'] || row['Segment'] || '').trim();
    if (waybill && segment) {
      nextMap.set(waybill, segment);
    }
  });

  if (nextMap.size === 0) {
    throw new Error('Excel 中没有识别到“运单号”和“正确段码”数据。');
  }

  waybillMap = nextMap;
  elements.dataCount.textContent = `${waybillMap.size} 条`;
  elements.dataSource.textContent = `已载入：${sourceName}`;
  elements.scanStatus.textContent = '数据已就绪，可以开始扫码。';
}

async function startScanner() {
  if (!waybillMap.size) {
    showResult('', '', 'warn', '请先导入运单段码对照表。');
    return;
  }

  stopScanner();
  lastScanValue = '';

  if (window.ZXing) {
    await startZxingScanner();
    return;
  }

  elements.scanStatus.textContent = '扫码组件未加载成功，请刷新页面；也可以先手动输入单号。';
}

async function startZxingScanner() {
  try {
    if (!window.isSecureContext) {
      throw new Error('InsecureContext');
    }
    await warmUpAudio();
    codeReader = new window.ZXing.BrowserMultiFormatReader(getDecodeHints(), 120);
    elements.placeholder.hidden = true;
    elements.scanButton.disabled = true;
    elements.stopButton.disabled = false;
    elements.scanStatus.textContent = '正在请求摄像头权限，请允许浏览器使用摄像头。';

    mediaStream = await navigator.mediaDevices.getUserMedia({
      video: {
        facingMode: { ideal: 'environment' },
        width: { ideal: 1280 },
        height: { ideal: 720 }
      },
      audio: false
    });
    elements.video.srcObject = mediaStream;
    await elements.video.play();
    codeReader.decodeFromVideoElementContinuously(elements.video, handleDecodeResult);

    elements.scanStatus.textContent = '摄像头已打开，请把条码横向放进取景框。';
  } catch (error) {
    stopScanner();
    elements.scanStatus.textContent = getCameraErrorMessage(error);
  }
}

function getDecodeHints() {
  const hints = new Map();
  const formats = window.ZXing.BarcodeFormat;
  hints.set(window.ZXing.DecodeHintType.POSSIBLE_FORMATS, [
    formats.CODE_128,
    formats.CODE_39,
    formats.ITF,
    formats.CODABAR,
    formats.QR_CODE
  ]);
  hints.set(window.ZXing.DecodeHintType.TRY_HARDER, true);
  return hints;
}

function handleDecodeResult(result, error) {
  if (error && error.name && error.name !== 'NotFoundException') {
    elements.scanStatus.textContent = '正在扫描，请让条码保持清晰并横向放入取景框。';
  }
  if (!result?.text || result.text === lastScanValue) {
    return;
  }

  const waybill = normalizeWaybill(result.text);
  if (!waybill) {
    return;
  }

  lastScanValue = result.text;
  playBeep();
  window.navigator.vibrate?.(70);
  lookup(waybill, '扫码识别');
}

function stopScanner() {
  if (scanTimer) {
    window.clearInterval(scanTimer);
    scanTimer = null;
  }
  codeReader?.reset();
  codeReader = null;
  mediaStream?.getTracks().forEach((track) => track.stop());
  mediaStream = null;
  elements.video.srcObject = null;
  elements.video.hidden = false;
  elements.placeholder.hidden = false;
  elements.scanButton.disabled = false;
  elements.stopButton.disabled = true;
}

async function warmUpAudio() {
  if (!audioContext) {
    audioContext = new (window.AudioContext || window.webkitAudioContext)();
  }
  if (audioContext.state === 'suspended') {
    await audioContext.resume();
  }
}

function playBeep() {
  if (!audioContext) {
    return;
  }
  const oscillator = audioContext.createOscillator();
  const gain = audioContext.createGain();
  oscillator.type = 'square';
  oscillator.frequency.setValueAtTime(1320, audioContext.currentTime);
  gain.gain.setValueAtTime(0.0001, audioContext.currentTime);
  gain.gain.exponentialRampToValueAtTime(0.22, audioContext.currentTime + 0.01);
  gain.gain.exponentialRampToValueAtTime(0.0001, audioContext.currentTime + 0.12);
  oscillator.connect(gain);
  gain.connect(audioContext.destination);
  oscillator.start();
  oscillator.stop(audioContext.currentTime + 0.13);
}

function getCameraErrorMessage(error) {
  const message = String(error?.message || error || '');
  if (error?.name === 'NotAllowedError' || message.includes('NotAllowedError') || message.includes('Permission')) {
    return '摄像头权限被拒绝，请在浏览器地址栏或系统设置中允许相机权限。';
  }
  if (error?.name === 'NotFoundError' || message.includes('NotFoundError')) {
    return '没有找到可用摄像头，请确认浏览器有相机权限。';
  }
  if (!window.isSecureContext) {
    return '当前不是安全 HTTPS 页面，浏览器会禁止摄像头。请使用 GitHub Pages 的 https 链接访问。';
  }
  return '摄像头启动失败，请刷新页面后重试，或先手动输入单号。';
}

function lookup(rawText, sourceLabel) {
  const waybill = normalizeWaybill(rawText);
  elements.manualInput.value = waybill || rawText.trim();

  if (!waybill) {
    showResult('--', '--', 'warn', '没有识别到 JDM 加 12 位数字的运单号，请调整角度或手动输入。');
    return;
  }

  const segment = waybillMap.get(waybill);
  if (segment) {
    showResult(waybill, segment, 'hit', `${sourceLabel}成功，已匹配到正确段码。`);
    return;
  }

  showResult(waybill, '未命中', 'missing', '这张单不在当前问题清单中，请复核面单或重新导入最新对照表。');
}

function normalizeWaybill(value) {
  const text = String(value || '').replace(/\s+/g, '').toUpperCase();
  const match = text.match(WAYBILL_PATTERN);
  return match ? match[0] : '';
}

function showResult(waybill, segment, state, note) {
  elements.resultCard.dataset.state = state;
  elements.resultLabel.textContent = state === 'hit' ? '正确段码' : '查询结果';
  elements.waybillNumber.textContent = waybill || '--';
  elements.segmentCode.textContent = segment || '--';
  elements.resultNote.textContent = note;
  elements.scanStatus.textContent = note;
}