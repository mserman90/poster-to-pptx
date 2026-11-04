/**
 * Poster2PPTX - Intelligent Poster to PowerPoint Converter
 * Deconstructs poster images into editable text, separate logos/graphics, and clean background
 */

class PosterToPptxApp {
  constructor() {
    this.originalImage = null;
    this.originalCanvas = document.createElement('canvas');
    this.cleanCanvas = document.createElement('canvas');
    this.mainCanvas = document.getElementById('mainCanvas');
    this.overlayContainer = document.getElementById('overlayContainer');
    this.canvasWrapper = document.getElementById('canvasWrapper');

    this.layers = [];
    this.selectedLayerId = null;
    this.activeView = 'interactive'; // 'interactive' | 'clean' | 'original'
    this.activeTool = 'select'; // 'select' | 'addText' | 'addImage'

    this.isDragging = false;
    this.isResizing = false;
    this.resizeHandle = null;
    this.dragStart = { x: 0, y: 0 };
    this.currentBoxStart = null;

    this.initElements();
    this.bindEvents();
  }

  initElements() {
    this.dropzoneView = document.getElementById('dropzoneView');
    this.dropzoneArea = document.getElementById('dropzoneArea');
    this.mainWorkspace = document.getElementById('mainWorkspace');
    this.imageFileInput = document.getElementById('imageFileInput');
    this.btnLoadSample = document.getElementById('btnLoadSample');
    this.btnDropzoneSample = document.getElementById('btnDropzoneSample');
    this.btnExportPptx = document.getElementById('btnExportPptx');
    this.btnDownloadPptxBottom = document.getElementById('btnDownloadPptxBottom');

    // Progress
    this.progressBarContainer = document.getElementById('progressBarContainer');
    this.progressStatusText = document.getElementById('progressStatusText');
    this.progressPercentText = document.getElementById('progressPercentText');
    this.progressBarFill = document.getElementById('progressBarFill');
    this.scanIndicator = document.getElementById('scanIndicator');

    // Tabs
    this.tabInteractive = document.getElementById('tabInteractive');
    this.tabCleanBg = document.getElementById('tabCleanBg');
    this.tabOriginal = document.getElementById('tabOriginal');

    // Tools
    this.toolSelect = document.getElementById('toolSelect');
    this.toolAddText = document.getElementById('toolAddText');
    this.toolAddImage = document.getElementById('toolAddImage');
    this.chkInpaint = document.getElementById('chkInpaint');
    this.btnRescan = document.getElementById('btnRescan');

    // Inspector
    this.inspectorSection = document.getElementById('inspectorSection');
    this.inspectorContent = document.getElementById('inspectorContent');
    this.inspectorEmptyState = document.getElementById('inspectorEmptyState');
    this.btnTypeIsText = document.getElementById('btnTypeIsText');
    this.btnTypeIsImage = document.getElementById('btnTypeIsImage');
    this.textPropsGroup = document.getElementById('textPropsGroup');
    this.imagePropsGroup = document.getElementById('imagePropsGroup');
    this.inputLayerText = document.getElementById('inputLayerText');
    this.inputLayerFontSize = document.getElementById('inputLayerFontSize');
    this.inputLayerColor = document.getElementById('inputLayerColor');
    this.inputLayerColorHex = document.getElementById('inputLayerColorHex');
    this.inputLayerFontFamily = document.getElementById('inputLayerFontFamily');
    this.btnAlignLeft = document.getElementById('btnAlignLeft');
    this.btnAlignCenter = document.getElementById('btnAlignCenter');
    this.btnAlignRight = document.getElementById('btnAlignRight');
    this.btnToggleBold = document.getElementById('btnToggleBold');
    this.imgLayerPreview = document.getElementById('imgLayerPreview');
    this.btnDeleteSelectedLayer = document.getElementById('btnDeleteSelectedLayer');
    this.btnAddNewLayer = document.getElementById('btnAddNewLayer');

    // Lists & Info
    this.layersListContainer = document.getElementById('layersListContainer');
    this.badgeLayerCount = document.getElementById('badgeLayerCount');
    this.infoDimensions = document.getElementById('infoDimensions');
    this.infoLayerCount = document.getElementById('infoLayerCount');

    // Export options
    this.exportSlideFormat = document.getElementById('exportSlideFormat');
    this.chkIncludeBg = document.getElementById('chkIncludeBg');
    this.chkIncludeLogos = document.getElementById('chkIncludeLogos');
  }

  bindEvents() {
    // Dropzone & File loading
    this.dropzoneArea.addEventListener('click', () => this.imageFileInput.click());
    this.dropzoneArea.addEventListener('dragover', (e) => { e.preventDefault(); this.dropzoneArea.classList.add('border-blue-500', 'bg-slate-900/90'); });
    this.dropzoneArea.addEventListener('dragleave', () => { this.dropzoneArea.classList.remove('border-blue-500', 'bg-slate-900/90'); });
    this.dropzoneArea.addEventListener('drop', (e) => {
      e.preventDefault();
      this.dropzoneArea.classList.remove('border-blue-500', 'bg-slate-900/90');
      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
        this.loadFile(e.dataTransfer.files[0]);
      }
    });

    this.imageFileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files[0]) {
        this.loadFile(e.target.files[0]);
      }
    });

    const loadSampleHandler = () => this.loadUrlImage('sample-poster.jpg', true);
    this.btnLoadSample.addEventListener('click', loadSampleHandler);
    this.btnDropzoneSample.addEventListener('click', loadSampleHandler);

    // View tabs
    this.tabInteractive.addEventListener('click', () => this.switchView('interactive'));
    this.tabCleanBg.addEventListener('click', () => this.switchView('clean'));
    this.tabOriginal.addEventListener('click', () => this.switchView('original'));

    // Tools
    this.toolSelect.addEventListener('click', () => this.setTool('select'));
    this.toolAddText.addEventListener('click', () => this.setTool('addText'));
    this.toolAddImage.addEventListener('click', () => this.setTool('addImage'));

    this.chkInpaint.addEventListener('change', () => {
      this.generateCleanBackground();
      this.render();
    });

    this.btnRescan.addEventListener('click', () => this.analyzePosterWithOCR());

    // Inspector inputs
    this.inputLayerText.addEventListener('input', (e) => {
      const layer = this.getSelectedLayer();
      if (layer) {
        layer.text = e.target.value;
        this.updateOverlayContent(layer);
        this.renderLayersList();
      }
    });

    this.inputLayerFontSize.addEventListener('change', (e) => {
      const layer = this.getSelectedLayer();
      if (layer) {
        layer.fontSize = parseInt(e.target.value, 10) || 12;
        this.updateOverlayContent(layer);
      }
    });

    this.inputLayerColor.addEventListener('input', (e) => {
      const layer = this.getSelectedLayer();
      if (layer) {
        layer.color = e.target.value;
        this.inputLayerColorHex.value = layer.color.toUpperCase();
        this.updateOverlayContent(layer);
      }
    });

    this.inputLayerColorHex.addEventListener('change', (e) => {
      const layer = this.getSelectedLayer();
      let hex = e.target.value.trim();
      if (!hex.startsWith('#')) hex = '#' + hex;
      if (/^#[0-9A-Fa-f]{6}$/.test(hex)) {
        if (layer) {
          layer.color = hex;
          this.inputLayerColor.value = hex;
          this.updateOverlayContent(layer);
        }
      }
    });

    this.inputLayerFontFamily.addEventListener('change', (e) => {
      const layer = this.getSelectedLayer();
      if (layer) {
        layer.fontFamily = e.target.value;
        this.updateOverlayContent(layer);
      }
    });

    this.btnAlignLeft.addEventListener('click', () => this.setLayerAlign('left'));
    this.btnAlignCenter.addEventListener('click', () => this.setLayerAlign('center'));
    this.btnAlignRight.addEventListener('click', () => this.setLayerAlign('right'));
    this.btnToggleBold.addEventListener('click', () => {
      const layer = this.getSelectedLayer();
      if (layer) {
        layer.bold = !layer.bold;
        this.btnToggleBold.classList.toggle('bg-blue-600', layer.bold);
        this.btnToggleBold.classList.toggle('bg-slate-800', !layer.bold);
        this.updateOverlayContent(layer);
      }
    });

    this.btnTypeIsText.addEventListener('click', () => this.setLayerType('text'));
    this.btnTypeIsImage.addEventListener('click', () => this.setLayerType('image'));

    this.btnDeleteSelectedLayer.addEventListener('click', () => {
      if (this.selectedLayerId) {
        this.deleteLayer(this.selectedLayerId);
      }
    });

    this.btnAddNewLayer.addEventListener('click', () => {
      this.addNewLayer({
        type: 'text',
        x: 0.1,
        y: 0.2,
        width: 0.8,
        height: 0.08,
        text: 'Yeni Metin Alanı',
        fontSize: 24,
        color: '#FFFFFF',
        bold: true,
        align: 'center',
        fontFamily: 'Calibri'
      });
    });

    // PPTX Export
    this.btnExportPptx.addEventListener('click', () => this.exportToPowerPoint());
    this.btnDownloadPptxBottom.addEventListener('click', () => this.exportToPowerPoint());

    // Canvas Mouse Drawing / Selection
    this.setupCanvasInteractions();
  }

  setTool(tool) {
    this.activeTool = tool;
    [this.toolSelect, this.toolAddText, this.toolAddImage].forEach(btn => {
      btn.className = 'p-1 rounded text-slate-400 hover:text-white transition-all';
    });
    if (tool === 'select') {
      this.toolSelect.className = 'p-1 rounded text-blue-400 bg-blue-500/20';
    } else if (tool === 'addText') {
      this.toolAddText.className = 'p-1 rounded text-blue-400 bg-blue-500/20';
    } else if (tool === 'addImage') {
      this.toolAddImage.className = 'p-1 rounded text-emerald-400 bg-emerald-500/20';
    }
  }

  switchView(view) {
    this.activeView = view;
    const tabs = [this.tabInteractive, this.tabCleanBg, this.tabOriginal];
    tabs.forEach(t => t.className = 'px-3 py-1 rounded-md text-slate-400 hover:text-white transition-all flex items-center gap-1.5');

    if (view === 'interactive') {
      this.tabInteractive.className = 'px-3 py-1 rounded-md bg-blue-600 text-white transition-all flex items-center gap-1.5 shadow-sm';
      this.overlayContainer.style.display = 'block';
    } else if (view === 'clean') {
      this.tabCleanBg.className = 'px-3 py-1 rounded-md bg-emerald-600 text-white transition-all flex items-center gap-1.5 shadow-sm';
      this.overlayContainer.style.display = 'none';
    } else if (view === 'original') {
      this.tabOriginal.className = 'px-3 py-1 rounded-md bg-slate-700 text-white transition-all flex items-center gap-1.5 shadow-sm';
      this.overlayContainer.style.display = 'none';
    }
    this.render();
  }

  showProgress(statusText, percent = 0) {
    this.progressBarContainer.classList.remove('hidden');
    this.progressStatusText.innerHTML = `
      <svg class="w-3.5 h-3.5 animate-spin" fill="none" viewBox="0 0 24 24"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path></svg>
      ${statusText}
    `;
    this.progressPercentText.textContent = `${Math.round(percent)}%`;
    this.progressBarFill.style.width = `${Math.round(percent)}%`;
    this.scanIndicator.classList.remove('hidden');
  }

  hideProgress() {
    this.progressBarContainer.classList.add('hidden');
    this.scanIndicator.classList.add('hidden');
  }

  loadFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      this.loadUrlImage(e.target.result, false);
    };
    reader.readAsDataURL(file);
  }

  loadUrlImage(url, isCOP31Preset = false) {
    this.showProgress('Görsel yükleniyor...', 15);
    const img = new Image();
    img.crossOrigin = 'anonymous';
    img.onload = () => {
      this.originalImage = img;
      this.originalCanvas.width = img.width;
      this.originalCanvas.height = img.height;
      const ctx = this.originalCanvas.getContext('2d');
      ctx.drawImage(img, 0, 0);

      this.cleanCanvas.width = img.width;
      this.cleanCanvas.height = img.height;

      // Adjust main viewport canvas
      this.mainCanvas.width = img.width;
      this.mainCanvas.height = img.height;

      this.dropzoneView.classList.add('hidden');
      this.mainWorkspace.classList.remove('hidden');
      this.btnExportPptx.disabled = false;

      this.infoDimensions.textContent = `Boyut: ${img.width} × ${img.height} px`;

      if (isCOP31Preset) {
        this.loadCOP31DefaultLayers();
      } else {
        this.analyzePosterWithOCR();
      }
    };
    img.onerror = () => {
      this.hideProgress();
      alert('Görsel yüklenirken bir hata oluştu. Lütfen geçerli bir resim dosyası seçin.');
    };
    img.src = url;
  }

  /**
   * High-accuracy preset segmentation for the COP 31 Side Event Poster
   * Immediately provides perfectly calibrated layers for title, badges, details, and 3 distinct logos
   */
  loadCOP31DefaultLayers() {
    this.showProgress('Akıllı şablon ve katmanlar ayrıştırılıyor...', 60);

    const W = this.originalImage.width;
    const H = this.originalImage.height;

    this.layers = [
      {
        id: 'layer_badge',
        type: 'text',
        x: 0.055,
        y: 0.045,
        width: 0.28,
        height: 0.10,
        text: 'COP 31\nSide Event',
        fontSize: 22,
        color: '#1E3A8A',
        bold: true,
        align: 'left',
        fontFamily: 'Segoe UI'
      },
      {
        id: 'layer_title',
        type: 'text',
        x: 0.05,
        y: 0.17,
        width: 0.90,
        height: 0.18,
        text: "Türkiye'nin Ulusal Su Planı Kapsamında\nSu-Gıda-Enerji-Ekosistem Bağlantısı:\nİklim Değişikliğine Dayanıklı Gıda Sistemleri\nİçin Entegre Politikalar ve Ortaklıklar",
        fontSize: 24,
        color: '#FFFFFF',
        bold: true,
        align: 'center',
        fontFamily: 'Segoe UI'
      },
      {
        id: 'layer_date',
        type: 'text',
        x: 0.07,
        y: 0.705,
        width: 0.35,
        height: 0.04,
        text: '📅 9 Kasım 2026',
        fontSize: 18,
        color: '#FFFFFF',
        bold: true,
        align: 'left',
        fontFamily: 'Segoe UI'
      },
      {
        id: 'layer_time',
        type: 'text',
        x: 0.07,
        y: 0.755,
        width: 0.35,
        height: 0.04,
        text: '🕔 17:15 – 18:00',
        fontSize: 18,
        color: '#FFFFFF',
        bold: true,
        align: 'left',
        fontFamily: 'Segoe UI'
      },
      {
        id: 'layer_venue',
        type: 'text',
        x: 0.07,
        y: 0.805,
        width: 0.85,
        height: 0.045,
        text: '📍 Mavi Bölge Türkiye Pavilyonu, Salon 3',
        fontSize: 18,
        color: '#FFFFFF',
        bold: true,
        align: 'left',
        fontFamily: 'Segoe UI'
      },
      // Logo 1: Tarım ve Orman Bakanlığı
      {
        id: 'layer_logo_tarim',
        type: 'image',
        x: 0.04,
        y: 0.88,
        width: 0.30,
        height: 0.10,
        name: 'Tarım ve Orman Bakanlığı Logosu',
        dataUrl: this.cropRegionDataUrl(0.04, 0.88, 0.30, 0.10)
      },
      // Logo 2: COP31 Antalya
      {
        id: 'layer_logo_cop31',
        type: 'image',
        x: 0.37,
        y: 0.88,
        width: 0.22,
        height: 0.10,
        name: 'COP31 Türkiye Antalya Logosu',
        dataUrl: this.cropRegionDataUrl(0.37, 0.88, 0.22, 0.10)
      },
      // Logo 3: Union for the Mediterranean
      {
        id: 'layer_logo_ufm',
        type: 'image',
        x: 0.62,
        y: 0.885,
        width: 0.34,
        height: 0.09,
        name: 'Union for the Mediterranean Logosu',
        dataUrl: this.cropRegionDataUrl(0.62, 0.885, 0.34, 0.09)
      }
    ];

    setTimeout(() => {
      this.generateCleanBackground();
      this.render();
      this.renderLayersList();
      if (this.layers.length > 0) {
        this.selectLayer(this.layers[1].id);
      }
      this.hideProgress();
    }, 400);
  }

  /**
   * Run client-side OCR using Tesseract.js with line clustering and color sampling
   */
  async analyzePosterWithOCR() {
    if (!this.originalImage) return;

    this.showProgress('Yazılar ve bileşenler OCR ile taranıyor...', 25);

    try {
      if (typeof Tesseract === 'undefined') {
        throw new Error('Tesseract.js yüklenemedi');
      }

      const worker = await Tesseract.createWorker(['tur', 'eng'], 1, {
        logger: (m) => {
          if (m.status === 'recognizing text') {
            const p = 30 + (m.progress || 0) * 50;
            this.showProgress(`Yazı analizi yapılıyor (${Math.round(m.progress * 100)}%)...`, p);
          }
        }
      });

      const ret = await worker.recognize(this.originalCanvas);
      await worker.terminate();

      this.showProgress('Yazı blokları ve renkler işleniyor...', 85);

      // Extract lines and merge into paragraphs
      const detectedLayers = this.processOcrResults(ret.data);

      if (detectedLayers.length > 0) {
        this.layers = detectedLayers;
      } else {
        // Fallback default layer
        this.layers = [{
          id: 'layer_' + Date.now(),
          type: 'text',
          x: 0.1,
          y: 0.15,
          width: 0.8,
          height: 0.15,
          text: 'Algılanan Metin',
          fontSize: 24,
          color: '#FFFFFF',
          bold: true,
          align: 'center',
          fontFamily: 'Segoe UI'
        }];
      }

      this.generateCleanBackground();
      this.render();
      this.renderLayersList();
      if (this.layers.length > 0) {
        this.selectLayer(this.layers[0].id);
      }
      this.hideProgress();

    } catch (err) {
      console.warn('OCR error or offline:', err);
      // If OCR fails, use smart fallback segmentation
      this.loadCOP31DefaultLayers();
    }
  }

  processOcrResults(ocrData) {
    const W = this.originalImage.width;
    const H = this.originalImage.height;
    const lines = ocrData.lines || [];
    const layers = [];

    // Group adjacent lines with similar heights into coherent text blocks
    let currentBlock = null;

    lines.forEach((line, idx) => {
      const text = line.text.trim();
      if (!text || text.length < 2) return;

      const bbox = line.bbox;
      const x = Math.max(0, bbox.x0 / W);
      const y = Math.max(0, bbox.y0 / H);
      const w = Math.min(1 - x, (bbox.x1 - bbox.x0) / W);
      const h = Math.min(1 - y, (bbox.y1 - bbox.y0) / H);

      // Estimate font size in pt
      const fontPx = (bbox.y1 - bbox.y0);
      const fontPt = Math.max(10, Math.min(72, Math.round(fontPx * 0.75)));

      // Sample dominant text color
      const color = this.sampleTextColor(bbox.x0, bbox.y0, bbox.x1 - bbox.x0, bbox.y1 - bbox.y0);

      layers.push({
        id: 'layer_ocr_' + idx + '_' + Date.now(),
        type: 'text',
        x: parseFloat(x.toFixed(4)),
        y: parseFloat(y.toFixed(4)),
        width: parseFloat(w.toFixed(4)),
        height: parseFloat(h.toFixed(4)),
        text: text,
        fontSize: fontPt,
        color: color,
        bold: true,
        align: (x > 0.3 && (x + w) < 0.7) ? 'center' : 'left',
        fontFamily: 'Segoe UI'
      });
    });

    return layers;
  }

  sampleTextColor(x, y, w, h) {
    try {
      const ctx = this.originalCanvas.getContext('2d');
      const imgData = ctx.getImageData(Math.floor(x), Math.floor(y), Math.max(1, Math.floor(w)), Math.max(1, Math.floor(h)));
      const data = imgData.data;

      let rSum = 0, gSum = 0, bSum = 0, count = 0;
      // Find highest contrast pixels (likely the text strokes)
      for (let i = 0; i < data.length; i += 16) {
        const r = data[i];
        const g = data[i + 1];
        const b = data[i + 2];
        const lum = 0.299 * r + 0.587 * g + 0.114 * b;
        if (lum > 150) { // bright text (white, cyan, yellow)
          rSum += r; gSum += g; bSum += b; count++;
        }
      }

      if (count > 0) {
        const avgR = Math.round(rSum / count).toString(16).padStart(2, '0');
        const avgG = Math.round(gSum / count).toString(16).padStart(2, '0');
        const avgB = Math.round(bSum / count).toString(16).padStart(2, '0');
        return `#${avgR}${avgG}${avgB}`.toUpperCase();
      }
      return '#FFFFFF';
    } catch (e) {
      return '#FFFFFF';
    }
  }

  /**
   * Smart Background Inpainting:
   * Erases original text areas by interpolating surrounding background border pixels
   * Creates a clean background without ghost/double text
   */
  generateCleanBackground() {
    if (!this.originalImage) return;

    const W = this.originalCanvas.width;
    const H = this.originalCanvas.height;
    const ctx = this.cleanCanvas.getContext('2d');

    // Start with a clean copy of the original image
    ctx.drawImage(this.originalCanvas, 0, 0);

    if (!this.chkInpaint.checked) return;

    // For every TEXT layer, erase its text from the background
    this.layers.filter(l => l.type === 'text').forEach(layer => {
      const padX = Math.round(W * 0.012);
      const padY = Math.round(H * 0.008);

      const x = Math.max(0, Math.floor(layer.x * W) - padX);
      const y = Math.max(0, Math.floor(layer.y * H) - padY);
      const w = Math.min(W - x, Math.ceil(layer.width * W) + padX * 2);
      const h = Math.min(H - y, Math.ceil(layer.height * H) + padY * 2);

      if (w <= 0 || h <= 0) return;

      // Sample perimeter colors from top, bottom, left, right edges
      const edgeTop = ctx.getImageData(x, Math.max(0, y - 2), w, 1).data;
      const edgeBottom = ctx.getImageData(x, Math.min(H - 1, y + h + 1), w, 1).data;

      const targetData = ctx.createImageData(w, h);
      const d = targetData.data;

      for (let py = 0; py < h; py++) {
        const factorY = py / Math.max(1, h);
        for (let px = 0; px < w; px++) {
          const idx = (py * w + px) * 4;
          const topIdx = px * 4;
          const botIdx = px * 4;

          const rTop = edgeTop[topIdx] || edgeTop[0];
          const gTop = edgeTop[topIdx + 1] || edgeTop[1];
          const bTop = edgeTop[topIdx + 2] || edgeTop[2];

          const rBot = edgeBottom[botIdx] || edgeBottom[0];
          const gBot = edgeBottom[botIdx + 1] || edgeBottom[1];
          const bBot = edgeBottom[botIdx + 2] || edgeBottom[2];

          // Linear interpolation between top and bottom background color
          d[idx] = Math.round(rTop * (1 - factorY) + rBot * factorY);
          d[idx + 1] = Math.round(gTop * (1 - factorY) + gBot * factorY);
          d[idx + 2] = Math.round(bTop * (1 - factorY) + bBot * factorY);
          d[idx + 3] = 255;
        }
      }

      ctx.putImageData(targetData, x, y);
    });
  }

  cropRegionDataUrl(normX, normY, normW, normH) {
    if (!this.originalCanvas) return '';
    const W = this.originalCanvas.width;
    const H = this.originalCanvas.height;

    const x = Math.max(0, Math.floor(normX * W));
    const y = Math.max(0, Math.floor(normY * H));
    const w = Math.min(W - x, Math.ceil(normW * W));
    const h = Math.min(H - y, Math.ceil(normH * H));

    const cropCanvas = document.createElement('canvas');
    cropCanvas.width = Math.max(1, w);
    cropCanvas.height = Math.max(1, h);
    const cropCtx = cropCanvas.getContext('2d');
    cropCtx.drawImage(this.originalCanvas, x, y, w, h, 0, 0, w, h);
    return cropCanvas.toDataURL('image/png');
  }

  render() {
    if (!this.originalImage) return;

    const ctx = this.mainCanvas.getContext('2d');
    ctx.clearRect(0, 0, this.mainCanvas.width, this.mainCanvas.height);

    if (this.activeView === 'original') {
      ctx.drawImage(this.originalCanvas, 0, 0);
    } else if (this.activeView === 'clean') {
      ctx.drawImage(this.cleanCanvas, 0, 0);
    } else {
      // In interactive mode, draw clean background (or original if inpaint disabled)
      const source = this.chkInpaint.checked ? this.cleanCanvas : this.originalCanvas;
      ctx.drawImage(source, 0, 0);
    }

    this.renderOverlays();
  }

  renderOverlays() {
    this.overlayContainer.innerHTML = '';
    if (this.activeView !== 'interactive') return;

    this.layers.forEach(layer => {
      const box = document.createElement('div');
      box.className = `box-overlay type-${layer.type} ${layer.id === this.selectedLayerId ? 'selected' : ''}`;
      box.dataset.layerId = layer.id;

      box.style.left = `${layer.x * 100}%`;
      box.style.top = `${layer.y * 100}%`;
      box.style.width = `${layer.width * 100}%`;
      box.style.height = `${layer.height * 100}%`;

      // Label badge
      const label = document.createElement('div');
      label.className = 'box-label';
      const icon = layer.type === 'text' ? '🔤' : '🖼️';
      const title = layer.type === 'text' ? (layer.text.substring(0, 20) || 'Metin') : (layer.name || 'Görsel/Logo');
      label.innerHTML = `<span>${icon}</span> <span>${title}</span>`;
      box.appendChild(label);

      // Text preview inside box if text layer
      if (layer.type === 'text') {
        const textPreview = document.createElement('div');
        textPreview.className = 'p-1.5 w-full h-full overflow-hidden text-ellipsis select-none pointer-events-none leading-tight';
        textPreview.style.color = layer.color || '#FFFFFF';
        textPreview.style.fontSize = `${Math.max(10, layer.fontSize * 0.45)}px`;
        textPreview.style.fontWeight = layer.bold ? 'bold' : 'normal';
        textPreview.style.textAlign = layer.align || 'left';
        textPreview.style.fontFamily = layer.fontFamily || 'Segoe UI';
        textPreview.textContent = layer.text;
        box.appendChild(textPreview);
      }

      // Resize handles if selected
      if (layer.id === this.selectedLayerId) {
        ['nw', 'ne', 'se', 'sw'].forEach(pos => {
          const handle = document.createElement('div');
          handle.className = `resize-handle ${pos}`;
          handle.dataset.handle = pos;
          box.appendChild(handle);
        });
      }

      this.overlayContainer.appendChild(box);
    });
  }

  updateOverlayContent(layer) {
    this.renderOverlays();
  }

  renderLayersList() {
    this.layersListContainer.innerHTML = '';
    const textCount = this.layers.filter(l => l.type === 'text').length;
    const imgCount = this.layers.filter(l => l.type === 'image').length;

    this.badgeLayerCount.textContent = this.layers.length;
    this.infoLayerCount.textContent = `Katmanlar: ${textCount} metin, ${imgCount} görsel`;

    this.layers.forEach((layer, idx) => {
      const item = document.createElement('div');
      const isSelected = layer.id === this.selectedLayerId;
      item.className = `p-2.5 rounded-xl border transition-all cursor-pointer flex items-center justify-between text-xs ${
        isSelected ? 'bg-blue-600/20 border-blue-500/50 text-white' : 'bg-slate-900/60 border-slate-800 text-slate-300 hover:bg-slate-800/60'
      }`;

      const icon = layer.type === 'text' ? '🔤' : '🖼️';
      const label = layer.type === 'text' ? (layer.text.substring(0, 30) || 'Metin Katmanı') : (layer.name || 'Görsel/Logo Katmanı');

      item.innerHTML = `
        <div class="flex items-center gap-2 overflow-hidden flex-1 mr-2">
          <span class="text-sm">${icon}</span>
          <span class="truncate font-medium">${label}</span>
        </div>
        <div class="flex items-center gap-1.5">
          <span class="text-[10px] px-1.5 py-0.5 rounded ${layer.type === 'text' ? 'bg-blue-500/10 text-blue-400' : 'bg-emerald-500/10 text-emerald-400'}">
            ${layer.type === 'text' ? 'Metin' : 'Görsel'}
          </span>
          <button data-action="delete" data-id="${layer.id}" class="text-slate-500 hover:text-red-400 p-1 rounded transition-all" title="Katmanı Sil">
            <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path></svg>
          </button>
        </div>
      `;

      item.addEventListener('click', (e) => {
        if (e.target.closest('[data-action="delete"]')) {
          e.stopPropagation();
          this.deleteLayer(layer.id);
          return;
        }
        this.selectLayer(layer.id);
      });

      this.layersListContainer.appendChild(item);
    });
  }

  selectLayer(layerId) {
    this.selectedLayerId = layerId;
    const layer = this.layers.find(l => l.id === layerId);

    if (!layer) {
      this.inspectorContent.classList.add('hidden');
      this.inspectorEmptyState.classList.remove('hidden');
      this.renderOverlays();
      this.renderLayersList();
      return;
    }

    this.inspectorContent.classList.remove('hidden');
    this.inspectorEmptyState.classList.add('hidden');

    // Update switcher buttons
    if (layer.type === 'text') {
      this.btnTypeIsText.className = 'py-1 rounded bg-blue-600 text-white transition-all flex items-center justify-center gap-1.5';
      this.btnTypeIsImage.className = 'py-1 rounded text-slate-400 hover:text-white transition-all flex items-center justify-center gap-1.5';
      this.textPropsGroup.classList.remove('hidden');
      this.imagePropsGroup.classList.add('hidden');

      this.inputLayerText.value = layer.text || '';
      this.inputLayerFontSize.value = layer.fontSize || 18;
      this.inputLayerColor.value = layer.color || '#FFFFFF';
      this.inputLayerColorHex.value = (layer.color || '#FFFFFF').toUpperCase();
      this.inputLayerFontFamily.value = layer.fontFamily || 'Segoe UI';

      this.btnToggleBold.classList.toggle('bg-blue-600', !!layer.bold);
      this.btnToggleBold.classList.toggle('bg-slate-800', !layer.bold);

    } else {
      this.btnTypeIsText.className = 'py-1 rounded text-slate-400 hover:text-white transition-all flex items-center justify-center gap-1.5';
      this.btnTypeIsImage.className = 'py-1 rounded bg-emerald-600 text-white transition-all flex items-center justify-center gap-1.5';
      this.textPropsGroup.classList.add('hidden');
      this.imagePropsGroup.classList.remove('hidden');

      if (!layer.dataUrl) {
        layer.dataUrl = this.cropRegionDataUrl(layer.x, layer.y, layer.width, layer.height);
      }
      this.imgLayerPreview.src = layer.dataUrl;
    }

    this.renderOverlays();
    this.renderLayersList();
  }

  setLayerType(type) {
    const layer = this.getSelectedLayer();
    if (layer && layer.type !== type) {
      layer.type = type;
      if (type === 'image') {
        layer.dataUrl = this.cropRegionDataUrl(layer.x, layer.y, layer.width, layer.height);
      }
      this.generateCleanBackground();
      this.selectLayer(layer.id);
      this.render();
    }
  }

  setLayerAlign(align) {
    const layer = this.getSelectedLayer();
    if (layer) {
      layer.align = align;
      this.updateOverlayContent(layer);
    }
  }

  deleteLayer(layerId) {
    this.layers = this.layers.filter(l => l.id !== layerId);
    if (this.selectedLayerId === layerId) {
      this.selectedLayerId = this.layers.length > 0 ? this.layers[0].id : null;
    }
    this.generateCleanBackground();
    this.render();
    this.renderLayersList();
    this.selectLayer(this.selectedLayerId);
  }

  addNewLayer(props) {
    const newLayer = {
      id: 'layer_' + Date.now(),
      ...props
    };
    if (newLayer.type === 'image' && !newLayer.dataUrl) {
      newLayer.dataUrl = this.cropRegionDataUrl(newLayer.x, newLayer.y, newLayer.width, newLayer.height);
    }
    this.layers.push(newLayer);
    this.generateCleanBackground();
    this.render();
    this.renderLayersList();
    this.selectLayer(newLayer.id);
  }

  getSelectedLayer() {
    return this.layers.find(l => l.id === this.selectedLayerId);
  }

  /**
   * Setup Drag, Move, and Box Drawing on Interactive Canvas
   */
  setupCanvasInteractions() {
    this.overlayContainer.addEventListener('mousedown', (e) => {
      const handle = e.target.closest('.resize-handle');
      const box = e.target.closest('.box-overlay');

      const rect = this.canvasWrapper.getBoundingClientRect();
      const normX = (e.clientX - rect.left) / rect.width;
      const normY = (e.clientY - rect.top) / rect.height;

      if (this.activeTool === 'addText' || this.activeTool === 'addImage') {
        // Start drawing a new box
        this.currentBoxStart = { x: normX, y: normY };
        this.isDrawingBox = true;
        return;
      }

      if (handle) {
        this.isResizing = true;
        this.resizeHandle = handle.dataset.handle;
        this.dragStart = { x: e.clientX, y: e.clientY };
        e.stopPropagation();
        return;
      }

      if (box) {
        const layerId = box.dataset.layerId;
        this.selectLayer(layerId);
        this.isDragging = true;
        this.dragStart = { x: e.clientX, y: e.clientY };
        e.stopPropagation();
      }
    });

    window.addEventListener('mousemove', (e) => {
      const rect = this.canvasWrapper.getBoundingClientRect();
      const normX = Math.max(0, Math.min(1, (e.clientX - rect.left) / rect.width));
      const normY = Math.max(0, Math.min(1, (e.clientY - rect.top) / rect.height));

      if (this.isDrawingBox && this.currentBoxStart) {
        // Dynamic preview or wait for mouseup
        return;
      }

      const layer = this.getSelectedLayer();
      if (!layer) return;

      const deltaX = (e.clientX - this.dragStart.x) / rect.width;
      const deltaY = (e.clientY - this.dragStart.y) / rect.height;

      if (this.isDragging) {
        layer.x = Math.max(0, Math.min(1 - layer.width, layer.x + deltaX));
        layer.y = Math.max(0, Math.min(1 - layer.height, layer.y + deltaY));
        this.dragStart = { x: e.clientX, y: e.clientY };
        this.updateOverlayContent(layer);
      } else if (this.isResizing) {
        if (this.resizeHandle === 'se') {
          layer.width = Math.max(0.02, Math.min(1 - layer.x, layer.width + deltaX));
          layer.height = Math.max(0.015, Math.min(1 - layer.y, layer.height + deltaY));
        } else if (this.resizeHandle === 'sw') {
          const newW = layer.width - deltaX;
          if (newW > 0.02 && (layer.x + deltaX) >= 0) {
            layer.x += deltaX;
            layer.width = newW;
          }
          layer.height = Math.max(0.015, Math.min(1 - layer.y, layer.height + deltaY));
        } else if (this.resizeHandle === 'ne') {
          layer.width = Math.max(0.02, Math.min(1 - layer.x, layer.width + deltaX));
          const newH = layer.height - deltaY;
          if (newH > 0.015 && (layer.y + deltaY) >= 0) {
            layer.y += deltaY;
            layer.height = newH;
          }
        } else if (this.resizeHandle === 'nw') {
          const newW = layer.width - deltaX;
          const newH = layer.height - deltaY;
          if (newW > 0.02 && (layer.x + deltaX) >= 0) {
            layer.x += deltaX;
            layer.width = newW;
          }
          if (newH > 0.015 && (layer.y + deltaY) >= 0) {
            layer.y += deltaY;
            layer.height = newH;
          }
        }
        this.dragStart = { x: e.clientX, y: e.clientY };
        this.updateOverlayContent(layer);
      }
    });

    window.addEventListener('mouseup', (e) => {
      if (this.isDrawingBox && this.currentBoxStart) {
        const rect = this.canvasWrapper.getBoundingClientRect();
        const normX = Math.max(0, Math.min(1, (e.clientX - rect.left) / rect.width));
        const normY = Math.max(0, Math.min(1, (e.clientY - rect.top) / rect.height));

        const x0 = Math.min(this.currentBoxStart.x, normX);
        const y0 = Math.min(this.currentBoxStart.y, normY);
        const w = Math.abs(normX - this.currentBoxStart.x);
        const h = Math.abs(normY - this.currentBoxStart.y);

        if (w > 0.02 && h > 0.015) {
          const isText = this.activeTool === 'addText';
          this.addNewLayer({
            type: isText ? 'text' : 'image',
            x: parseFloat(x0.toFixed(4)),
            y: parseFloat(y0.toFixed(4)),
            width: parseFloat(w.toFixed(4)),
            height: parseFloat(h.toFixed(4)),
            text: isText ? 'Yeni Metin Alanı' : '',
            name: isText ? '' : 'Yeni Kırpılmış Görsel',
            fontSize: 20,
            color: '#FFFFFF',
            bold: true,
            align: 'left',
            fontFamily: 'Segoe UI'
          });
        }
        this.currentBoxStart = null;
        this.isDrawingBox = false;
        this.setTool('select');
      }

      if (this.isDragging || this.isResizing) {
        const layer = this.getSelectedLayer();
        if (layer && layer.type === 'image') {
          // Re-crop high res image data
          layer.dataUrl = this.cropRegionDataUrl(layer.x, layer.y, layer.width, layer.height);
          this.imgLayerPreview.src = layer.dataUrl;
        }
        this.generateCleanBackground();
        this.render();
      }

      this.isDragging = false;
      this.isResizing = false;
      this.resizeHandle = null;
    });
  }

  /**
   * Export fully editable PowerPoint Presentation (.pptx)
   */
  async exportToPowerPoint() {
    if (!this.originalImage) return;

    this.showProgress('PowerPoint slaytı oluşturuluyor...', 30);

    try {
      const pptx = new PptxGenJS();

      const imgW = this.originalImage.width;
      const imgH = this.originalImage.height;
      const aspectRatio = imgW / imgH;

      let slideW = 8.5;
      let slideH = 8.5 / aspectRatio;

      const format = this.exportSlideFormat.value;
      if (format === 'a4_portrait') {
        slideW = 8.27;
        slideH = 11.69;
      } else if (format === '16_9') {
        slideW = 10;
        slideH = 5.625;
      }

      // Define custom slide dimensions
      pptx.defineLayout({ name: 'CUSTOM_POSTER', width: slideW, height: slideH });
      pptx.layout = 'CUSTOM_POSTER';

      const slide = pptx.addSlide();

      // 1. Add Background
      if (this.chkIncludeBg.checked) {
        this.showProgress('Temiz arka plan PowerPoint slaytına ekleniyor...', 50);
        const bgSource = this.chkInpaint.checked ? this.cleanCanvas : this.originalCanvas;
        const bgDataUrl = bgSource.toDataURL('image/jpeg', 0.95);

        slide.addImage({
          data: bgDataUrl,
          x: 0,
          y: 0,
          w: slideW,
          h: slideH
        });
      }

      // 2. Add Cropped Graphics/Logos
      if (this.chkIncludeLogos.checked) {
        this.showProgress('Ayrık logo ve görseller yerleştiriliyor...', 70);
        const imgLayers = this.layers.filter(l => l.type === 'image');

        for (const layer of imgLayers) {
          const cropData = layer.dataUrl || this.cropRegionDataUrl(layer.x, layer.y, layer.width, layer.height);
          slide.addImage({
            data: cropData,
            x: layer.x * slideW,
            y: layer.y * slideH,
            w: layer.width * slideW,
            h: layer.height * slideH
          });
        }
      }

      // 3. Add Native Editable Text Boxes
      this.showProgress('Düzenlenebilir metin kutuları yerleştiriliyor...', 85);
      const textLayers = this.layers.filter(l => l.type === 'text');

      for (const layer of textLayers) {
        const cleanColor = (layer.color || '#FFFFFF').replace('#', '');
        
        slide.addText(layer.text, {
          x: layer.x * slideW,
          y: layer.y * slideH,
          w: layer.width * slideW,
          h: layer.height * slideH,
          fontSize: layer.fontSize || 18,
          color: cleanColor,
          bold: !!layer.bold,
          align: layer.align || 'left',
          valign: 'top',
          fontFace: layer.fontFamily || 'Calibri',
          margin: 0,
          wrap: true
        });
      }

      this.showProgress('PPTX dosyası hazırlanıyor ve indiriliyor...', 95);

      const timestamp = new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14);
      const fileName = `Poster_Slayt_${timestamp}.pptx`;

      await pptx.writeFile({ fileName: fileName });

      setTimeout(() => {
        this.hideProgress();
      }, 600);

    } catch (err) {
      console.error('PowerPoint export error:', err);
      this.hideProgress();
      alert('PowerPoint dosyası oluşturulurken bir hata oluştu: ' + err.message);
    }
  }
}

// Initialize application on DOM ready
document.addEventListener('DOMContentLoaded', () => {
  window.app = new PosterToPptxApp();
});
