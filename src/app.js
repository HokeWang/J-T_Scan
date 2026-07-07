const WAYBILL_PATTERN = /JDM\d{12}/i;

const translations = {
  zh: {
    title: '运单段码快查',
    cameraPlaceholder: '正在打开摄像头，请允许相机权限',
    photoButton: '拍照',
    retakeButton: '重拍',
    loadingData: '正在加载运单段码对照表。',
    waitingPhoto: '等待拍照',
    resultHint: '拍照或手输后，这里会显示正确段码。',
    manualLabel: '手动输入或粘贴单号',
    queryButton: '查询',
    embeddedReady: '数据已就绪，正在准备摄像头。',
    noData: '内置数据未载入，请点击“重新导入 Excel”。',
    cameraPermission: '正在请求摄像头权限，请允许浏览器使用摄像头。',
    cameraReady: '摄像头已打开，请把面单条码或二维码放进取景框后拍照。',
    cameraNotReady: '摄像头尚未准备好，请稍后再拍照。',
    decodingPhoto: '正在识别照片中的条码或二维码。',
    retakeOpening: '正在重新打开摄像头。',
    notFoundRetake: '未查到，请重拍',
    multipleMatches: '查到多条数据，请完善单号',
    noManualMatch: '没有匹配到单号，请检查输入或重新拍照。',
    hitLabel: '正确段码',
    resultLabel: '查询结果',
    photoSuccess: '拍照识别成功，已匹配到正确段码。',
    manualSuccess: '手动查询成功，已匹配到正确段码。',
    notInList: '这张单不在当前问题清单中，请复核面单或重新导入最新对照表。',
    cameraDenied: '摄像头权限被拒绝，请在浏览器地址栏或系统设置中允许相机权限。',
    cameraMissing: '没有找到可用摄像头，请确认浏览器有相机权限。',
    insecureContext: '当前不是安全 HTTPS 页面，浏览器会禁止摄像头。请使用 GitHub Pages 的 https 链接访问。',
    cameraFailed: '摄像头启动失败，请刷新页面后重试，或先手动输入单号。'
  },
  es: {
    title: 'Consulta rápida de código de tramo',
    cameraPlaceholder: 'Abriendo la cámara. Permita el acceso a la cámara.',
    photoButton: 'Tomar foto',
    retakeButton: 'Repetir foto',
    loadingData: 'Cargando la tabla de guías y códigos de tramo.',
    waitingPhoto: 'Esperando foto',
    resultHint: 'Después de tomar una foto o ingresar la guía, aquí se mostrará el código correcto.',
    manualLabel: 'Ingresar o pegar número de guía',
    queryButton: 'Buscar',
    embeddedReady: 'Datos listos. Preparando la cámara.',
    noData: 'No se cargaron datos integrados. Importe el Excel de nuevo.',
    cameraPermission: 'Solicitando permiso de cámara. Permita el acceso en el navegador.',
    cameraReady: 'Cámara abierta. Coloque el código de barras o QR dentro del marco y tome la foto.',
    cameraNotReady: 'La cámara aún no está lista. Intente de nuevo en unos segundos.',
    decodingPhoto: 'Reconociendo el código de barras o QR en la foto.',
    retakeOpening: 'Abriendo la cámara de nuevo.',
    notFoundRetake: 'No encontrado. Repita la foto.',
    multipleMatches: 'Se encontraron varios registros. Complete el número de guía.',
    noManualMatch: 'No se encontró la guía. Revise el dato ingresado o repita la foto.',
    hitLabel: 'Código correcto',
    resultLabel: 'Resultado',
    photoSuccess: 'Foto reconocida. Se encontró el código correcto.',
    manualSuccess: 'Consulta manual exitosa. Se encontró el código correcto.',
    notInList: 'Esta guía no está en la lista actual. Revise la etiqueta o importe la tabla más reciente.',
    cameraDenied: 'Permiso de cámara denegado. Permita la cámara en el navegador o en la configuración del sistema.',
    cameraMissing: 'No se encontró una cámara disponible. Confirme el permiso de cámara.',
    insecureContext: 'La página no usa HTTPS seguro. El navegador bloqueará la cámara. Use el enlace HTTPS de GitHub Pages.',
    cameraFailed: 'No se pudo iniciar la cámara. Actualice la página o ingrese la guía manualmente.'
  }
};

const elements = {
  languageButtons: document.querySelectorAll('.language-button'),
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
  resultValues: document.querySelector('#result-values'),
  waybillNumber: document.querySelector('#waybill-number'),
  segmentCode: document.querySelector('#segment-code'),
  resultNote: document.querySelector('#result-note')
};

let waybillMap = new Map();
let mediaStream = null;
let codeReader = null;
let audioContext = null;
let photoCaptured = false;
let currentLanguage = 'zh';

document.addEventListener('DOMContentLoaded', () => {
  bindEvents();
  setLanguage('zh');
  loadEmbeddedData();
  startCameraPreview();
});

function bindEvents() {
  elements.languageButtons.forEach((button) => {
    button.addEventListener('click', () => setLanguage(button.dataset.language));
  });
  elements.photoButton.addEventListener('click', handlePhotoButtonClick);
  elements.manualButton.addEventListener('click', () => lookup(elements.manualInput.value, '手动查询'));
  elements.manualInput.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      lookup(elements.manualInput.value, '手动查询');
    }
  });
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
    elements.scanStatus.textContent = t('noData');
    return;
  }

  waybillMap = nextMap;
  elements.scanStatus.textContent = t('embeddedReady');
}

async function startCameraPreview() {
  if (!waybillMap.size) {
    return;
  }

  try {
    if (!window.isSecureContext) {
      throw new Error('InsecureContext');
    }
    elements.scanStatus.textContent = t('cameraPermission');
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
    elements.scanStatus.textContent = t('cameraReady');
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
    elements.scanStatus.textContent = t('cameraNotReady');
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
    elements.photoButton.textContent = t('retakeButton');
    elements.scanStatus.textContent = t('decodingPhoto');
    stopCameraPreview(false);

    const decodedText = await decodePhoto(canvas);
    const waybill = resolveWaybill(decodedText);
    if (!waybill) {
      showResult('--', '--', 'missing', t('notFoundRetake'));
      return;
    }

    playBeep();
    window.navigator.vibrate?.(70);
    lookup(waybill, '拍照识别');
  } catch (error) {
    showResult('--', '--', 'missing', t('notFoundRetake'));
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
  elements.photoButton.textContent = t('photoButton');
  elements.photoButton.disabled = true;
  clearResult();
  elements.scanStatus.textContent = t('retakeOpening');
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
    return t('cameraDenied');
  }
  if (error?.name === 'NotFoundError' || message.includes('NotFoundError')) {
    return t('cameraMissing');
  }
  if (!window.isSecureContext) {
    return t('insecureContext');
  }
  return t('cameraFailed');
}

function lookup(rawText, sourceLabel) {
  const lookupResult = sourceLabel === '手动查询'
    ? resolveManualWaybill(rawText)
    : { status: 'hit', waybill: resolveWaybill(rawText) };
  const waybill = lookupResult.waybill;
  elements.manualInput.value = waybill || rawText.trim();

  if (lookupResult.status === 'multiple') {
    showMultipleResults(lookupResult.matches, t('multipleMatches'));
    return;
  }

  if (!waybill) {
    showResult('--', '--', 'warn', t('noManualMatch'));
    return;
  }

  const segment = waybillMap.get(waybill);
  if (segment) {
    showResult(waybill, segment, 'hit', sourceLabel === '手动查询' ? t('manualSuccess') : t('photoSuccess'));
    return;
  }

  showResult(waybill, currentLanguage === 'es' ? 'Sin coincidencia' : '未命中', 'missing', t('notInList'));
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
    return { status: 'multiple', waybill: '', matches };
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
  elements.resultLabel.textContent = state === 'hit' ? t('hitLabel') : t('resultLabel');
  elements.resultValues.replaceChildren(createWaybillLine(waybill || '--'), createSegmentLine(segment || '--'));
  elements.waybillNumber = elements.resultValues.querySelector('.waybill-number');
  elements.segmentCode = elements.resultValues.querySelector('.segment-code');
  elements.resultNote.textContent = note;
  elements.scanStatus.textContent = note;
}

function showMultipleResults(matches, note) {
  elements.resultCard.dataset.state = 'hit';
  elements.resultLabel.textContent = t('hitLabel');
  elements.resultValues.replaceChildren(...matches.flatMap((waybill) => [
    createWaybillLine(waybill),
    createSegmentLine(waybillMap.get(waybill) || '--')
  ]));
  elements.waybillNumber = elements.resultValues.querySelector('.waybill-number');
  elements.segmentCode = elements.resultValues.querySelector('.segment-code');
  elements.resultNote.textContent = note;
  elements.scanStatus.textContent = note;
}

function createWaybillLine(value) {
  const line = document.createElement('div');
  line.className = 'waybill-number';
  line.textContent = value;
  return line;
}

function createSegmentLine(value) {
  const line = document.createElement('div');
  line.className = 'segment-code';
  line.textContent = value;
  return line;
}

function clearResult(updateStatus = true) {
  elements.resultCard.dataset.state = 'idle';
  elements.resultLabel.textContent = t('waitingPhoto');
  elements.resultValues.replaceChildren(createWaybillLine('--'), createSegmentLine('--'));
  elements.waybillNumber = elements.resultValues.querySelector('.waybill-number');
  elements.segmentCode = elements.resultValues.querySelector('.segment-code');
  elements.resultNote.textContent = t('resultHint');
  if (updateStatus) {
    elements.scanStatus.textContent = t('retakeOpening');
  }
}

function t(key) {
  return translations[currentLanguage][key] || translations.zh[key] || key;
}

function setLanguage(language) {
  currentLanguage = translations[language] ? language : 'zh';
  document.documentElement.lang = currentLanguage === 'es' ? 'es' : 'zh-CN';
  elements.languageButtons.forEach((button) => {
    button.classList.toggle('active', button.dataset.language === currentLanguage);
  });
  document.querySelectorAll('[data-i18n]').forEach((node) => {
    node.textContent = t(node.dataset.i18n);
  });
  elements.manualInput.placeholder = currentLanguage === 'es' ? 'Ej. JDM000001507469' : '例如 JDM000001507469';
  refreshResultLanguage();
}

function refreshResultLanguage() {
  const state = elements.resultCard.dataset.state;
  if (state === 'hit') {
    elements.resultLabel.textContent = t('hitLabel');
    return;
  }
  if (state === 'idle') {
    clearResult(false);
    return;
  }
  elements.resultLabel.textContent = t('resultLabel');
}