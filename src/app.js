const WAYBILL_PATTERN = /JDM[0-9A-Z]{6,}/i;

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
let detector = null;
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

  if (!('BarcodeDetector' in window)) {
    elements.scanStatus.textContent = '当前浏览器不支持直接扫码，请使用 Chrome/Edge 手机版，或先手动输入单号。';
    return;
  }

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
  mediaStream?.getTracks().forEach((track) => track.stop());
  mediaStream = null;
  detector = null;
  elements.video.srcObject = null;
  elements.placeholder.hidden = false;
  elements.scanButton.disabled = false;
  elements.stopButton.disabled = true;
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