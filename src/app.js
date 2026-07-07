const WAYBILL_PATTERN = /JDM[0-9A-Z]{6,}/i;

const elements = {
  dataCount: document.querySelector('#data-count'),
  dataSource: document.querySelector('#data-source'),
  excelInput: document.querySelector('#excel-input'),
  video: document.querySelector('#camera-preview'),
  html5Reader: document.querySelector('#html5-reader'),
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
let detector = null;
let html5Scanner = null;
let scanTimer = null;
let lastScanValue = '';

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

  if (window.Html5Qrcode) {
    await startHtml5Scanner();
    return;
  }

  if ('BarcodeDetector' in window) {
    await startNativeScanner();
    return;
  }

  elements.scanStatus.textContent = '扫码组件未加载成功，请刷新页面；也可以先手动输入单号。';
}

async function startNativeScanner() {
  try {
    detector = new window.BarcodeDetector({
      formats: ['qr_code', 'code_128', 'code_39', 'codabar', 'ean_13', 'itf']
    });
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
    elements.placeholder.hidden = true;
    elements.scanButton.disabled = true;
    elements.stopButton.disabled = false;
    elements.scanStatus.textContent = '正在扫码，请让条码或二维码进入取景框。';
    scanTimer = window.setInterval(scanFrame, 280);
  } catch (error) {
    stopScanner();
    elements.scanStatus.textContent = error.name === 'NotAllowedError'
      ? '摄像头权限被拒绝，请允许浏览器使用摄像头。'
      : '摄像头启动失败，请改用手动输入。';
  }
}

async function startHtml5Scanner() {
  try {
    html5Scanner = new window.Html5Qrcode(elements.html5Reader.id, {
      formatsToSupport: getHtml5Formats()
    });
    elements.placeholder.hidden = true;
    elements.video.hidden = true;
    elements.scanButton.disabled = true;
    elements.stopButton.disabled = false;
    elements.scanStatus.textContent = '正在请求摄像头权限，请允许浏览器使用摄像头。';

    await html5Scanner.start(
      { facingMode: 'environment' },
      { fps: 10, qrbox: getQrbox, aspectRatio: 1.35 },
      (decodedText) => {
        if (decodedText && decodedText !== lastScanValue) {
          lastScanValue = decodedText;
          lookup(decodedText, '扫码识别');
          window.navigator.vibrate?.(60);
        }
      }
    );

    elements.scanStatus.textContent = '摄像头已打开，请对准面单条码或二维码。';
  } catch (error) {
    stopScanner();
    elements.scanStatus.textContent = getCameraErrorMessage(error);
  }
}

function getHtml5Formats() {
  if (!window.Html5QrcodeSupportedFormats) {
    return undefined;
  }

  const formats = window.Html5QrcodeSupportedFormats;
  return [
    formats.QR_CODE,
    formats.CODE_128,
    formats.CODE_39,
    formats.CODABAR,
    formats.EAN_13,
    formats.ITF
  ].filter(Boolean);
}

function getQrbox(viewfinderWidth, viewfinderHeight) {
  const width = Math.floor(viewfinderWidth * 0.82);
  const height = Math.floor(Math.min(viewfinderHeight * 0.52, 260));
  return { width, height };
}

async function scanFrame() {
  if (!detector || !elements.video.videoWidth) {
    return;
  }

  try {
    const codes = await detector.detect(elements.video);
    if (!codes.length) {
      return;
    }
    const rawValue = codes[0].rawValue || '';
    if (rawValue && rawValue !== lastScanValue) {
      lastScanValue = rawValue;
      lookup(rawValue, '扫码识别');
      window.navigator.vibrate?.(60);
    }
  } catch (error) {
    elements.scanStatus.textContent = '扫码识别中断，请重新开始扫码。';
    stopScanner();
  }
}

function stopScanner() {
  if (scanTimer) {
    window.clearInterval(scanTimer);
    scanTimer = null;
  }
  if (html5Scanner) {
    const scanner = html5Scanner;
    html5Scanner = null;
    scanner.stop()
      .then(() => scanner.clear())
      .catch(() => scanner.clear());
  }
  mediaStream?.getTracks().forEach((track) => track.stop());
  mediaStream = null;
  detector = null;
  elements.video.srcObject = null;
  elements.video.hidden = false;
  elements.placeholder.hidden = false;
  elements.scanButton.disabled = false;
  elements.stopButton.disabled = true;
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
    showResult('--', '--', 'warn', '没有识别到 JDM 开头的运单号，请调整角度或手动输入。');
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