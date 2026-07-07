const WAYBILL_PATTERN = /JDM\d{12}/i;

const elements = {
  dataCount: document.querySelector('#data-count'),
  dataSource: document.querySelector('#data-source'),
  excelInput: document.querySelector('#excel-input'),
  cameraCard: document.querySelector('.camera-card'),
  video: document.querySelector('#camera-preview'),
  canvas: document.querySelector('#photo-canvas'),
  placeholder: document.querySelector('#camera-placeholder'),
  photoButton: document.querySelector('#photo-button'),
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
let audioContext = null;
let photoCaptured = false;

document.addEventListener('DOMContentLoaded', () => {
  bindEvents();
  loadEmbeddedData();
  startCameraPreview();
});

function bindEvents() {
  elements.photoButton.addEventListener('click', handlePhotoButtonClick);
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
  elements.scanStatus.textContent = '数据已就绪，正在准备摄像头。';
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

async function startCameraPreview() {
  if (!waybillMap.size) {
    return;
  }

  try {
    if (!window.isSecureContext) {
      throw new Error('InsecureContext');
    }
    elements.scanStatus.textContent = '正在请求摄像头权限，请允许浏览器使用摄像头。';
    mediaStream = await navigator.mediaDevices.getUserMedia({
      video: {
        facingMode: { ideal: 'environment' },
        width: { ideal: 1920 },
        height: { ideal: 1080 },
        focusMode: { ideal: 'continuous' }
      },
      audio: false
    });
    elements.video.srcObject = mediaStream;
    await elements.video.play();
    codeReader = new window.ZXing.BrowserMultiFormatReader(getDecodeHints(), 120);
    elements.placeholder.hidden = true;
    elements.cameraCard.classList.add('preview-ready');
    elements.cameraCard.classList.remove('photo-captured');
    elements.photoButton.disabled = false;
    elements.scanStatus.textContent = '摄像头已打开，请把面单条码或二维码放进取景框后拍照。';
  } catch (error) {
    stopCameraPreview();
    elements.scanStatus.textContent = getCameraErrorMessage(error);
  }
}

async function handlePhotoButtonClick() {
  if (photoCaptured) {
    resetPhotoMode();
    await startCameraPreview();
    return;
  }

  await captureAndDecodePhoto();
}

async function captureAndDecodePhoto() {
  if (!elements.video.videoWidth || !elements.video.videoHeight) {
    elements.scanStatus.textContent = '摄像头尚未准备好，请稍后再拍照。';
    return;
  }

  try {
    await warmUpAudio();
    const canvas = elements.canvas;
    canvas.width = elements.video.videoWidth;
    canvas.height = elements.video.videoHeight;
    const context = canvas.getContext('2d', { willReadFrequently: true });
    context.drawImage(elements.video, 0, 0, canvas.width, canvas.height);
    canvas.hidden = false;
    elements.video.hidden = true;
    elements.cameraCard.classList.add('photo-captured');
    elements.cameraCard.classList.remove('preview-ready');
    photoCaptured = true;
    elements.photoButton.textContent = '重拍';
    elements.scanStatus.textContent = '正在识别照片中的条码或二维码。';
    stopCameraPreview(false);

    const decodedText = await decodePhoto(canvas);
    const waybill = resolveWaybill(decodedText);
    if (!waybill) {
      elements.scanStatus.textContent = `已识别：${shortenRawText(decodedText)}，但未解析出可匹配单号。`;
      return;
    }

    playBeep();
    window.navigator.vibrate?.(70);
    lookup(waybill, '拍照识别');
  } catch (error) {
    elements.scanStatus.textContent = '照片未识别到条码或二维码，请点击“重拍”后靠近面单重新拍照。';
  }
}

async function decodePhoto(canvas) {
  const candidates = createDecodeCanvases(canvas);
  for (const candidate of candidates) {
    const zxingText = await decodeWithZxing(candidate).catch(() => '');
    if (resolveWaybill(zxingText)) {
      return zxingText;
    }
    const quaggaText = await decodeWithQuagga(candidate).catch(() => '');
    if (resolveWaybill(quaggaText)) {
      return quaggaText;
    }
  }
  throw new Error('No barcode found');
}

async function decodeWithZxing(canvas) {
  if (!window.ZXing) {
    throw new Error('ZXing is unavailable');
  }
  if (!codeReader) {
    codeReader = new window.ZXing.BrowserMultiFormatReader(getDecodeHints(), 120);
  }
  const result = await codeReader.decodeFromCanvas(canvas);
  return result?.text || '';
}

function decodeWithQuagga(canvas) {
  if (!window.Quagga) {
    return Promise.reject(new Error('Quagga is unavailable'));
  }

  return new Promise((resolve, reject) => {
    window.Quagga.decodeSingle({
      decoder: {
        readers: ['code_128_reader', 'code_39_reader', 'i2of5_reader', 'codabar_reader'],
        multiple: false
      },
      locate: true,
      locator: {
        patchSize: 'medium',
        halfSample: false
      },
      numOfWorkers: 0,
      src: canvas.toDataURL('image/png')
    }, (result) => {
      const code = result?.codeResult?.code || '';
      if (code) {
        resolve(code);
      } else {
        reject(new Error('Quagga found no barcode'));
      }
    });
  });
}

function createDecodeCanvases(sourceCanvas) {
  const full = cloneCanvas(sourceCanvas);
  const crop = cropCanvas(sourceCanvas, 0.06, 0.12, 0.88, 0.58);
  return [
    crop,
    enhanceCanvas(crop),
    full,
    enhanceCanvas(full)
  ];
}

function cloneCanvas(sourceCanvas) {
  const canvas = document.createElement('canvas');
  canvas.width = sourceCanvas.width;
  canvas.height = sourceCanvas.height;
  canvas.getContext('2d').drawImage(sourceCanvas, 0, 0);
  return canvas;
}

function cropCanvas(sourceCanvas, leftRatio, topRatio, widthRatio, heightRatio) {
  const sourceX = Math.floor(sourceCanvas.width * leftRatio);
  const sourceY = Math.floor(sourceCanvas.height * topRatio);
  const sourceWidth = Math.floor(sourceCanvas.width * widthRatio);
  const sourceHeight = Math.floor(sourceCanvas.height * heightRatio);
  const scale = Math.max(1, Math.min(2.2, 1800 / sourceWidth));
  const canvas = document.createElement('canvas');
  canvas.width = Math.floor(sourceWidth * scale);
  canvas.height = Math.floor(sourceHeight * scale);
  canvas.getContext('2d').drawImage(
    sourceCanvas,
    sourceX,
    sourceY,
    sourceWidth,
    sourceHeight,
    0,
    0,
    canvas.width,
    canvas.height
  );
  return canvas;
}

function enhanceCanvas(sourceCanvas) {
  const canvas = cloneCanvas(sourceCanvas);
  const context = canvas.getContext('2d', { willReadFrequently: true });
  const imageData = context.getImageData(0, 0, canvas.width, canvas.height);
  const data = imageData.data;
  for (let index = 0; index < data.length; index += 4) {
    const gray = data[index] * 0.299 + data[index + 1] * 0.587 + data[index + 2] * 0.114;
    const value = gray > 150 ? 255 : 0;
    data[index] = value;
    data[index + 1] = value;
    data[index + 2] = value;
  }
  context.putImageData(imageData, 0, 0);
  return canvas;
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

function stopCameraPreview(resetReader = true) {
  if (resetReader) {
    codeReader?.reset();
    codeReader = null;
  }
  mediaStream?.getTracks().forEach((track) => track.stop());
  mediaStream = null;
  elements.video.srcObject = null;
  elements.cameraCard.classList.remove('preview-ready');
}

function resetPhotoMode() {
  stopCameraPreview();
  photoCaptured = false;
  elements.canvas.hidden = true;
  elements.video.hidden = false;
  elements.placeholder.hidden = false;
  elements.cameraCard.classList.remove('photo-captured', 'preview-ready');
  elements.photoButton.textContent = '拍照';
  elements.photoButton.disabled = true;
  elements.scanStatus.textContent = '正在重新打开摄像头。';
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
  const lookupResult = sourceLabel === '手动查询'
    ? resolveManualWaybill(rawText)
    : { status: 'hit', waybill: resolveWaybill(rawText) };
  const waybill = lookupResult.waybill;
  elements.manualInput.value = waybill || rawText.trim();

  if (lookupResult.status === 'multiple') {
    showResult('--', '--', 'warn', '查到多条数据，请完善单号');
    return;
  }

  if (!waybill) {
    showResult('--', '--', 'warn', '没有匹配到单号，请检查输入或重新拍照。');
    return;
  }

  const segment = waybillMap.get(waybill);
  if (segment) {
    showResult(waybill, segment, 'hit', `${sourceLabel}成功，已匹配到正确段码。`);
    return;
  }

  showResult(waybill, '未命中', 'missing', '这张单不在当前问题清单中，请复核面单或重新导入最新对照表。');
}

function resolveManualWaybill(value) {
  const exactWaybill = resolveWaybill(value);
  if (exactWaybill && waybillMap.has(exactWaybill)) {
    return { status: 'hit', waybill: exactWaybill };
  }

  const query = String(value || '').replace(/\s+/g, '').toUpperCase();
  const digits = query.replace(/\D/g, '');
  const token = digits || query;
  if (!token) {
    return { status: 'none', waybill: '' };
  }

  const matches = Array.from(waybillMap.keys()).filter((waybill) => {
    if (digits) {
      return waybill.replace(/^JDM/, '').endsWith(digits);
    }
    return waybill.includes(token);
  });

  if (matches.length === 1) {
    return { status: 'hit', waybill: matches[0] };
  }
  if (matches.length > 1) {
    return { status: 'multiple', waybill: '' };
  }
  return { status: 'none', waybill: '' };
}

function normalizeWaybill(value) {
  const text = String(value || '').replace(/\s+/g, '').toUpperCase();
  const directMatch = text.match(WAYBILL_PATTERN);
  if (directMatch) {
    return directMatch[0];
  }

  const spacedPrefixMatch = text.match(/JDM\D*(\d{12})/);
  if (spacedPrefixMatch) {
    return `JDM${spacedPrefixMatch[1]}`;
  }

  const digits = text.replace(/\D/g, '');
  return digits.length === 12 ? `JDM${digits}` : '';
}

function resolveWaybill(value) {
  const normalized = normalizeWaybill(value);
  if (normalized) {
    return normalized;
  }

  const digits = String(value || '').replace(/\D/g, '');
  for (let index = 0; index <= digits.length - 12; index += 1) {
    const candidate = `JDM${digits.slice(index, index + 12)}`;
    if (waybillMap.has(candidate)) {
      return candidate;
    }
  }

  return '';
}

function shortenRawText(value) {
  const text = String(value || '').trim();
  if (text.length <= 28) {
    return text || '--';
  }
  return `${text.slice(0, 12)}...${text.slice(-10)}`;
}

function showResult(waybill, segment, state, note) {
  elements.resultCard.dataset.state = state;
  elements.resultLabel.textContent = state === 'hit' ? '正确段码' : '查询结果';
  elements.waybillNumber.textContent = waybill || '--';
  elements.segmentCode.textContent = segment || '--';
  elements.resultNote.textContent = note;
  elements.scanStatus.textContent = note;
}