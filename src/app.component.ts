import { Component, ChangeDetectionStrategy, signal, computed, ElementRef, afterNextRender, effect, viewChild, viewChildren } from '@angular/core';

// Declare global variables from CDN scripts
declare var pdfjsLib: any;
declare var PDFLib: any;

interface TextSettings {
  font: string;
  size: number;
  color: string;
  bold: boolean;
}

interface TextAnnotation {
  type: 'text';
  x: number;
  y: number;
  text: string;
  size: number;
  color: string;
  font: string;
  bold: boolean;
  width: number;
  height: number;
}

interface DrawAnnotation {
  type: 'draw';
  path: { x: number; y: number }[];
  color: string;
  lineWidth: number;
}

interface ImageAnnotation {
  type: 'image';
  x: number;
  y: number;
  width: number;
  height: number;
  imageData: string; // base64
  originalAspectRatio: number;
}

type Annotation = TextAnnotation | DrawAnnotation | ImageAnnotation;

type PageNumberPosition = 'top-left' | 'top-center' | 'top-right' | 'middle-left' | 'middle-center' | 'middle-right' | 'bottom-left' | 'bottom-center' | 'bottom-right';

interface PageNumberSettings {
  visible: boolean;
  position: PageNumberPosition;
  margin: 'small' | 'medium' | 'large';
  startPage: number;
}

interface WatermarkSettings {
  visible: boolean;
  text: string;
  font: string; // e.g., 'Helvetica'
  size: number;
  color: string; // hex #RRGGBB
  opacity: number; // 0 to 1
}


type DragMode = 'move' | 'resize' | 'none';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: {
    '(window:keydown.Delete)': 'onDeleteKeyPressed()',
    '(window:keydown.Backspace)': 'onDeleteKeyPressed()'
  }
})
export class AppComponent {
  // PDF state
  pdfDoc = signal<any>(null);
  pdfLibDoc = signal<any>(null);
  currentPageNum = signal(1);
  totalPages = signal(0);
  isPdfLoaded = computed(() => !!this.pdfDoc());
  originalFileName = signal<string>('');
  scale = signal(1.5);

  // UI state
  activeTab = signal<'edit' | 'organize' | 'security'>('edit');
  activeTool = signal<'select' | 'text' | 'draw'>('select');
  isSaving = signal(false);
  isPageNumberPanelOpen = signal(false);
  isPreviewPanelCollapsed = signal(false);
  
  // Page Number Settings
  pageNumberSettings = signal<PageNumberSettings>({ visible: false, position: 'bottom-center', margin: 'medium', startPage: 1 });
  tempPageNumberSettings = signal<PageNumberSettings>(this.pageNumberSettings());
  readonly pageNumberPositions: {id: PageNumberPosition, iconPath: string}[] = [
    { id: 'top-left', iconPath: 'M3,3H9V5H5V9H3V3Z' }, { id: 'top-center', iconPath: 'M10.5,3H13.5V5H10.5V3Z' }, { id: 'top-right', iconPath: 'M15,3H21V9H19V5H15V3Z' },
    { id: 'middle-left', iconPath: 'M3,10.5H5V13.5H3V10.5Z' }, { id: 'middle-center', iconPath: 'M10.5,10.5H13.5V13.5H10.5V10.5Z' }, { id: 'middle-right', iconPath: 'M19,10.5H21V13.5H19V10.5Z' },
    { id: 'bottom-left', iconPath: 'M3,15H5V19H9V21H3V15Z' }, { id: 'bottom-center', iconPath: 'M10.5,19H13.5V21H10.5V19Z' }, { id: 'bottom-right', iconPath: 'M15,19V15H21V21H15V19H19V17H17V19Z' },
  ];

  // Watermark Settings
  isWatermarkPanelOpen = signal(false);
  watermarkSettings = signal<WatermarkSettings>({
    visible: false,
    text: 'Sample',
    font: 'Helvetica-Bold',
    size: 100,
    color: '#ff0000',
    opacity: 0.5,
  });
  readonly availableFonts = [
      { id: 'Helvetica', name: 'Arial' },
      { id: 'Helvetica-Bold', name: 'Arial Bold' },
      { id: 'Times-Roman', name: 'Times New Roman' },
      { id: 'Courier', name: 'Courier' },
  ];
  
  // Signature State
  isSignatureModalOpen = signal(false);
  signaturePadPaths = signal<{ path: {x: number, y: number}[], color: string }[]>([]);
  private signaturePadCurrentPath: {x: number, y: number}[] = [];
  private isSigning = false;
  signatureColor = signal('black');
  private signatureUndoStack = signal<{ path: {x: number, y: number}[], color: string }[][]>([]);
  private signatureRedoStack = signal<{ path: {x: number, y: number}[], color: string }[][]>([]);
  canUndoSignature = computed(() => this.signatureUndoStack().length > 0);
  canRedoSignature = computed(() => this.signatureRedoStack().length > 0);
  
  // Password Protection State
  isPasswordModalOpen = signal(false);
  password = signal('');
  confirmPassword = signal('');
  passwordError = signal<string | null>(null);


  // Page order and drag/drop
  pageOrder = signal<number[]>([]); 
  draggedPageIndex = signal<number>(-1);
  dragOverIndex = signal<number>(-1);
  private dragScrollInterval: any = null;

  // PDF insertion
  insertionIndex = signal<number>(0);
  isDraggingOverInsert = signal<number>(-1);

  // Editing state
  annotations = signal<Annotation[][]>([]);
  isDrawing = false;
  currentDrawingPath: { x: number; y: number }[] = [];
  private imageCache = new Map<string, HTMLImageElement>();
  
  // Text Editing State
  isEditingText = signal(false);
  textEditingState = signal<{x: number, y: number, width: number, text: string} | null>(null);
  textSettings = signal<TextSettings>({ font: 'Helvetica', size: 16, color: '#000000', bold: false });


  // Interactive annotation state
  selectedAnnotationIndex = signal<number>(-1);
  selectedAnnotation = computed(() => {
    const pageIndex = this.pageOrder().indexOf(this.currentPageNum());
    const annIndex = this.selectedAnnotationIndex();
    if (pageIndex > -1 && annIndex > -1) {
      return this.annotations()[pageIndex][annIndex];
    }
    return null;
  });
  isTextToolbarVisible = computed(() => this.activeTool() === 'text' || this.selectedAnnotation()?.type === 'text');
  hoveredAnnotationInfo = signal<{index: number; part: 'body' | 'resize'} | null>(null);
  dragMode: DragMode = 'none';
  dragStartPos = { x: 0, y: 0 };
  dragStartAnn = { x: 0, y: 0, width: 0, height: 0 };


  // History
  undoStack = signal<Annotation[][][]>([]);
  redoStack = signal<Annotation[][][]>([]);
  canUndo = computed(() => this.undoStack().length > 0);
  canRedo = computed(() => this.redoStack().length > 0);

  // Canvas elements
  pdfCanvasRef = viewChild<ElementRef<HTMLCanvasElement>>('pdfCanvas');
  drawCanvasRef = viewChild<ElementRef<HTMLCanvasElement>>('drawCanvas');
  signatureCanvasRef = viewChild<ElementRef<HTMLCanvasElement>>('signatureCanvas');
  textInputRef = viewChild<ElementRef<HTMLTextAreaElement>>('textInput');
  thumbnailCanvases = viewChildren<ElementRef<HTMLCanvasElement>>('thumbnailCanvas');
  insertPdfInputRef = viewChild.required<ElementRef<HTMLInputElement>>('insertPdfInput');
  imageInputRef = viewChild.required<ElementRef<HTMLInputElement>>('imageInput');
  viewerContainerRef = viewChild.required<ElementRef<HTMLDivElement>>('viewerContainer');
  previewPanel = viewChild.required<ElementRef<HTMLElement>>('previewPanel');
  
  private pdfRenderTask: any = null;
  private thumbnailRenderTasks = new Map<HTMLCanvasElement, any>();

  constructor() {
    afterNextRender(() => {
      pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.3.136/pdf.worker.mjs`;
    });

    effect(() => {
      const pdf = this.pdfDoc();
      const pageNum = this.currentPageNum();
      this.scale(); // dependency
      if (pdf && pageNum > 0 && pageNum <= pdf.numPages) {
        this.renderPage(pageNum);
      }
    });

    effect(() => {
      const allCanvases = this.thumbnailCanvases();
      const pdf = this.pdfDoc();
      
      if (pdf && allCanvases.length > 0) {
        const canvasMap = new Map<string, HTMLCanvasElement>(allCanvases.map(c => [c.nativeElement.id, c.nativeElement]));
        for (const pageNum of this.pageOrder()) {
          const canvasEl = canvasMap.get('thumb-canvas-' + pageNum);
          if (canvasEl) {
            this.renderThumbnail(pdf, pageNum, canvasEl);
          }
        }
      }
    });

    effect(() => {
      const textInput = this.textInputRef();
      if (textInput) {
        textInput.nativeElement.focus();
      }
    });
  }
  
  canvasCursor = computed(() => {
    if (this.dragMode === 'move') return 'grabbing';
    if (this.dragMode === 'resize') return 'nwse-resize';
    
    const hoverInfo = this.hoveredAnnotationInfo();
    if (hoverInfo) {
      if (hoverInfo.part === 'body') return 'grab';
      if (hoverInfo.part === 'resize') return 'nwse-resize';
    }

    if (this.activeTab() !== 'edit' || this.isPageNumberPanelOpen()) return 'default';
    
    switch (this.activeTool()) {
      case 'text': return 'text';
      case 'draw': return 'crosshair';
      default: return 'default';
    }
  });

  async onFileSelected(event: Event): Promise<void> {
    const input = event.target as HTMLInputElement;
    if (!input.files || input.files.length === 0) return;
    const file = input.files[0];
    this.originalFileName.set(file.name.replace(/\.pdf$/i, ''));
    const buffer = await file.arrayBuffer();
    await this.loadPdf(buffer, true);
  }

  async loadPdf(pdfBytes: ArrayBuffer, isNewFile: boolean, pageToLoad?: number): Promise<void> {
    this.isSaving.set(true);
    if (this.pdfRenderTask) this.pdfRenderTask.cancel();
    this.pdfDoc.set(null);
    this.pdfLibDoc.set(null);

    try {
      const { PDFDocument } = PDFLib;
      const loadingTask = pdfjsLib.getDocument({data: new Uint8Array(pdfBytes)});
      const [pdf, pdfLibDoc] = await Promise.all([loadingTask.promise, PDFDocument.load(pdfBytes)]);
      
      this.totalPages.set(pdf.numPages);
      
      if (isNewFile) {
        this.annotations.set(Array(pdf.numPages).fill([]));
        this.pageOrder.set(Array.from({ length: pdf.numPages }, (_, i) => i + 1));
        this.currentPageNum.set(pageToLoad ?? 1);
        this.pageNumberSettings.set({ visible: false, position: 'bottom-center', margin: 'medium', startPage: 1 });
        this.isPageNumberPanelOpen.set(false);
        this.undoStack.set([]);
        this.redoStack.set([]);
        this.selectedAnnotationIndex.set(-1);
      } else {
         this.currentPageNum.set(pageToLoad ?? this.currentPageNum());
      }

      this.pdfDoc.set(pdf);
      this.pdfLibDoc.set(pdfLibDoc);
      
    } catch (error) {
      console.error('Error loading PDF:', error);
      alert('PDFファイルの読み込みに失敗しました。');
      this.totalPages.set(0);
    } finally {
      this.isSaving.set(false);
    }
  }

  async renderPage(num: number): Promise<void> {
    const pdf = this.pdfDoc(), pdfCanvasRef = this.pdfCanvasRef(), drawCanvasRef = this.drawCanvasRef();
    if (!pdf || !pdfCanvasRef || !drawCanvasRef) return;
    if (this.pdfRenderTask) this.pdfRenderTask.cancel();

    try {
      const page = await pdf.getPage(num);
      const pixelRatio = window.devicePixelRatio || 1;
      const viewport = page.getViewport({ scale: this.scale() * pixelRatio });

      const pdfCanvas = pdfCanvasRef.nativeElement;
      const drawCanvas = drawCanvasRef.nativeElement;
      const container = pdfCanvas.parentElement;

      const canvasWidth = viewport.width;
      const canvasHeight = viewport.height;
      
      const cssWidth = canvasWidth / pixelRatio;
      const cssHeight = canvasHeight / pixelRatio;

      if (container) {
          container.style.width = `${cssWidth}px`;
          container.style.height = `${cssHeight}px`;
      }

      pdfCanvas.width = canvasWidth;
      pdfCanvas.height = canvasHeight;
      pdfCanvas.style.width = `${cssWidth}px`;
      pdfCanvas.style.height = `${cssHeight}px`;

      drawCanvas.width = canvasWidth;
      drawCanvas.height = canvasHeight;
      drawCanvas.style.width = `${cssWidth}px`;
      drawCanvas.style.height = `${cssHeight}px`;
      
      this.pdfRenderTask = page.render({ 
        canvasContext: pdfCanvas.getContext('2d')!, 
        viewport 
      });
      await this.pdfRenderTask.promise;
      this.drawAnnotations();
    } catch (error) {
        if ((error as any).name !== 'RenderingCancelledException') console.error('Error rendering page:', error);
    } finally {
        this.pdfRenderTask = null;
    }
  }

  renderThumbnail = async (pdf: any, pageNum: number, canvas: HTMLCanvasElement): Promise<void> => {
    this.thumbnailRenderTasks.get(canvas)?.cancel();

    let renderTask: any;
    try {
      const page = await pdf.getPage(pageNum);
      const pixelRatio = window.devicePixelRatio || 1;
      const viewport = page.getViewport({ scale: 0.25 * pixelRatio });
      canvas.height = viewport.height;
      canvas.width = viewport.width;
      const context = canvas.getContext('2d');
      if (!context) return;
      context.clearRect(0, 0, canvas.width, canvas.height);
      
      renderTask = page.render({ canvasContext: context, viewport });
      this.thumbnailRenderTasks.set(canvas, renderTask);
      
      await renderTask.promise;
    } catch (e: any) {
      if (e.name !== 'RenderingCancelledException') {
        console.error(`Error rendering thumbnail for page ${pageNum}`, e);
      }
    } finally {
        if (this.thumbnailRenderTasks.get(canvas) === renderTask) {
            this.thumbnailRenderTasks.delete(canvas);
        }
    }
  }
  
  drawAnnotations(): void {
    const drawCanvas = this.drawCanvasRef()?.nativeElement; if (!drawCanvas) return;
    const context = drawCanvas.getContext('2d')!;
    context.clearRect(0, 0, drawCanvas.width, drawCanvas.height);
    const pageIndex = this.pageOrder().indexOf(this.currentPageNum()); if (pageIndex === -1) return;
    const pageAnnotations = this.annotations()[pageIndex] || [];
    const pixelRatio = window.devicePixelRatio || 1;
    const currentScale = this.scale() * pixelRatio;

    for (const [i, ann] of pageAnnotations.entries()) {
      if (this.isEditingText() && this.selectedAnnotationIndex() === i) {
        continue; // Skip drawing the annotation being edited
      }
      if (ann.type === 'image') {
        let img = this.imageCache.get(ann.imageData);
        if (!img) {
          img = new Image();
          img.src = ann.imageData;
          img.onload = () => this.drawAnnotations();
          this.imageCache.set(ann.imageData, img);
        }
        if (img.complete) {
          context.drawImage(img, ann.x * currentScale, ann.y * currentScale, ann.width * currentScale, ann.height * currentScale);
          if (this.selectedAnnotationIndex() === i) {
            context.strokeStyle = 'blue';
            context.lineWidth = 2 * pixelRatio;
            context.strokeRect(ann.x * currentScale, ann.y * currentScale, ann.width * currentScale, ann.height * currentScale);
            const handleSize = 10 * pixelRatio;
            context.fillStyle = 'white';
            context.fillRect(ann.x * currentScale + ann.width * currentScale - handleSize / 2, ann.y * currentScale + ann.height * currentScale - handleSize / 2, handleSize, handleSize);
            context.strokeRect(ann.x * currentScale + ann.width * currentScale - handleSize / 2, ann.y * currentScale + ann.height * currentScale - handleSize / 2, handleSize, handleSize);
          }
        }
      } else if (ann.type === 'text') {
        context.font = `${ann.bold ? 'bold ' : ''}${ann.size * currentScale}px ${this.getCanvasFontFamily(ann.font)}`;
        context.fillStyle = ann.color;
        context.textBaseline = 'top';

        const lines = ann.text.split('\n');
        const lineHeight = ann.size * 1.2 * currentScale;
        lines.forEach((line, index) => {
            context.fillText(line, ann.x * currentScale, ann.y * currentScale + (index * lineHeight));
        });

        if (this.selectedAnnotationIndex() === i) {
            context.strokeStyle = 'rgba(0, 0, 255, 0.7)';
            context.lineWidth = 1 * pixelRatio;
            context.strokeRect(ann.x * currentScale, ann.y * currentScale, ann.width * currentScale, ann.height * currentScale);
        }
        context.textBaseline = 'alphabetic'; // Reset
      } else if (ann.type === 'draw' && ann.path.length > 1) {
        context.beginPath();
        context.moveTo(ann.path[0].x * currentScale, ann.path[0].y * currentScale);
        ann.path.slice(1).forEach(p => context.lineTo(p.x * currentScale, p.y * currentScale));
        context.strokeStyle = ann.color;
        context.lineWidth = ann.lineWidth * currentScale;
        context.lineCap = 'round'; context.lineJoin = 'round';
        context.stroke();
      }
    }
    
    const settings = this.isPageNumberPanelOpen() ? this.tempPageNumberSettings() : this.pageNumberSettings();
    if (settings.visible) {
        const displayPageIndex = this.pageOrder().indexOf(this.currentPageNum()); if (displayPageIndex === -1) return;
        const pageNumText = `${displayPageIndex + settings.startPage}`;
        const fontSize = 12 * this.scale() * pixelRatio;
        context.font = `${fontSize}px sans-serif`; context.fillStyle = 'black';
        const marginValues = { small: 18, medium: 36, large: 54 };
        const margin = marginValues[settings.margin] * this.scale() * pixelRatio;
        let x = 0, y = 0;
        if (settings.position.includes('left')) { context.textAlign = 'left'; x = margin; } 
        else if (settings.position.includes('center')) { context.textAlign = 'center'; x = drawCanvas.width / 2; } 
        else if (settings.position.includes('right')) { context.textAlign = 'right'; x = drawCanvas.width - margin; }
        if (settings.position.includes('top')) { context.textBaseline = 'top'; y = margin; } 
        else if (settings.position.includes('middle')) { context.textBaseline = 'middle'; y = drawCanvas.height / 2; } 
        else if (settings.position.includes('bottom')) { context.textBaseline = 'bottom'; y = drawCanvas.height - margin; }
        context.fillText(pageNumText, x, y);
        context.textAlign = 'start'; context.textBaseline = 'alphabetic';
    }

    this.drawWatermark();

    if (this.isDrawing && this.currentDrawingPath.length > 1) {
        context.beginPath();
        const pathInPixels = this.currentDrawingPath.map(p => ({ x: p.x * currentScale, y: p.y * currentScale }));
        pathInPixels.forEach((p, i) => i === 0 ? context.moveTo(p.x, p.y) : context.lineTo(p.x, p.y));
        context.strokeStyle = 'red'; context.lineWidth = 2 * pixelRatio; context.stroke();
    }
  }

  private drawWatermark(): void {
    const settings = this.watermarkSettings();
    if (!settings.visible || !settings.text) return;

    const drawCanvas = this.drawCanvasRef()?.nativeElement;
    if (!drawCanvas) return;
    const context = drawCanvas.getContext('2d')!;
    const pixelRatio = window.devicePixelRatio || 1;

    context.save();

    const centerX = drawCanvas.width / 2;
    const centerY = drawCanvas.height / 2;
    
    context.translate(centerX, centerY);
    context.rotate(-Math.PI / 4); // -45 degrees, counter-clockwise
    context.translate(-centerX, -centerY);

    const fontFamily = this.getCanvasFontFamily(settings.font);
    const fontWeight = settings.font.includes('Bold') ? 'bold' : 'normal';
    
    context.globalAlpha = settings.opacity;
    context.font = `${fontWeight} ${settings.size * (this.scale() / 1.5) * pixelRatio}px ${fontFamily}`;
    context.fillStyle = settings.color;
    context.textAlign = 'center';
    context.textBaseline = 'middle';

    context.fillText(settings.text, centerX, centerY);

    context.restore();
  }

  getCanvasFontFamily(fontId: string): string {
    if (fontId.includes('Helvetica')) return 'Arial, Helvetica, sans-serif';
    if (fontId.includes('Times')) return '"Times New Roman", Times, serif';
    if (fontId.includes('Courier')) return '"Courier New", Courier, monospace';
    return 'Arial, Helvetica, sans-serif';
  }

  togglePreviewPanel = () => this.isPreviewPanelCollapsed.update(v => !v);

  togglePageNumberPanel(): void {
    this.isPageNumberPanelOpen.update(v => !v);
    if (this.isPageNumberPanelOpen()) this.tempPageNumberSettings.set(this.pageNumberSettings());
    else this.drawAnnotations();
  }
  updateTempSetting = (key: keyof PageNumberSettings, value: any) => { this.tempPageNumberSettings.update(s => ({ ...s, [key]: value })); this.drawAnnotations(); };
  applyPageNumbers(): void { this.pageNumberSettings.set({ ...this.tempPageNumberSettings(), visible: true }); this.isPageNumberPanelOpen.set(false); this.drawAnnotations(); }
  removePageNumbers(): void { this.pageNumberSettings.update(s => ({ ...s, visible: false })); this.isPageNumberPanelOpen.set(false); this.drawAnnotations(); }

  toggleWatermarkPanel(): void {
    this.isWatermarkPanelOpen.update(v => !v);
  }

  updateWatermarkSetting(key: keyof WatermarkSettings, value: any): void {
    this.watermarkSettings.update(s => ({ ...s, [key]: value }));
    this.drawAnnotations();
  }

  zoomIn = () => this.scale.update(s => Math.min(5, s + 0.25));
  zoomOut = () => this.scale.update(s => Math.max(0.25, s - 0.25));

  async fitToPage() {
    const pdf = this.pdfDoc(), viewerContainer = this.viewerContainerRef()?.nativeElement; if (!pdf || !viewerContainer) return;
    try {
        const page = await pdf.getPage(this.currentPageNum());
        const viewport = page.getViewport({ scale: 1 });
        const padding = 32;
        const scale = Math.min((viewerContainer.clientWidth - padding) / viewport.width, (viewerContainer.clientHeight - padding) / viewport.height);
        this.scale.set(scale);
    } catch (error) { console.error("Failed to fit page:", error); }
  }

  onPageInput(event: Event) {
    const pageNum = parseInt((event.target as HTMLInputElement).value, 10);
    if (!isNaN(pageNum) && pageNum >= 1 && pageNum <= this.totalPages()) this.goToPage(pageNum);
    else (event.target as HTMLInputElement).value = (this.pageOrder().indexOf(this.currentPageNum()) + 1).toString();
  }
  goToPage = (num: number) => { this.currentPageNum.set(this.pageOrder()[num - 1]); this.selectedAnnotationIndex.set(-1); };
  goToPrevPage = () => { if (this.pageOrder().indexOf(this.currentPageNum()) > 0) this.goToPage(this.pageOrder().indexOf(this.currentPageNum())); };
  goToNextPage = () => { if (this.pageOrder().indexOf(this.currentPageNum()) < this.pageOrder().length - 1) this.goToPage(this.pageOrder().indexOf(this.currentPageNum()) + 2); };
  
  selectTab = (tab: 'edit' | 'organize' | 'security') => this.activeTab.set(tab);
  selectTool = (tool: 'select' | 'text' | 'draw') => { this.activeTool.set(tool); this.selectedAnnotationIndex.set(-1); };

  addImage(): void { if (!this.isPdfLoaded()) return; this.imageInputRef().nativeElement.value = ''; this.imageInputRef().nativeElement.click(); }
  onImageSelected(event: Event): void {
    const file = (event.target as HTMLInputElement).files?.[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = e => {
      const imageData = e.target?.result as string;
      const img = new Image();
      img.onload = () => {
        const MAX_WIDTH_POINTS = 200, MAX_HEIGHT_POINTS = 200;
        let { width, height } = img;
        const scale = this.scale();
        let pointWidth = width / scale;
        let pointHeight = height / scale;

        if (pointWidth > pointHeight) { 
            if (pointWidth > MAX_WIDTH_POINTS) { 
                pointHeight *= MAX_WIDTH_POINTS / pointWidth; 
                pointWidth = MAX_WIDTH_POINTS; 
            } 
        } else { 
            if (pointHeight > MAX_HEIGHT_POINTS) { 
                pointWidth *= MAX_HEIGHT_POINTS / pointHeight; 
                pointHeight = MAX_HEIGHT_POINTS; 
            } 
        }
        
        const drawCanvas = this.drawCanvasRef()?.nativeElement;
        if (!drawCanvas) return;
        
        this.addAnnotation({
          type: 'image', 
          x: (drawCanvas.clientWidth / scale - pointWidth) / 2, 
          y: (drawCanvas.clientHeight / scale - pointHeight) / 2,
          width: pointWidth, 
          height: pointHeight, 
          imageData, 
          originalAspectRatio: img.height / img.width
        });

        const pageIndex = this.pageOrder().indexOf(this.currentPageNum());
        this.selectedAnnotationIndex.set(this.annotations()[pageIndex].length - 1);
      };
      img.src = imageData;
    };
    reader.readAsDataURL(file);
  }

  private getMousePos = (event: MouseEvent, canvasEl: HTMLCanvasElement | undefined): { x: number; y: number } | null => {
    if (!canvasEl) return null;
    const rect = canvasEl.getBoundingClientRect();
    return rect ? { x: event.clientX - rect.left, y: event.clientY - rect.top } : null;
  };

  onMouseDown(event: MouseEvent) {
    if (this.isEditingText() || this.activeTab() !== 'edit' || this.isPageNumberPanelOpen()) return;
    
    const pos = this.getMousePos(event, this.pdfCanvasRef()?.nativeElement);
    if (!pos) return;

    const posUnscaled = { x: pos.x / this.scale(), y: pos.y / this.scale() };
    const pageIndex = this.pageOrder().indexOf(this.currentPageNum());
    if (pageIndex === -1) return;
    const pageAnns = this.annotations()[pageIndex];

    let newSelection: number = -1;
    for (let i = pageAnns.length - 1; i >= 0; i--) {
        const ann = pageAnns[i];
        if (ann.type === 'text' || ann.type === 'image') {
            const annRect = { x: ann.x, y: ann.y, width: ann.width, height: ann.height };
            
            // Check for resize handle on images first, as it's more specific
            if (ann.type === 'image') {
                const handleSize = 10 / this.scale();
                const handleRect = { x: ann.x + ann.width - handleSize/2, y: ann.y + ann.height - handleSize/2, width: handleSize, height: handleSize };
                if (posUnscaled.x >= handleRect.x && posUnscaled.x <= handleRect.x + handleRect.width && posUnscaled.y >= handleRect.y && posUnscaled.y <= handleRect.y + handleRect.height) {
                    this.dragMode = 'resize'; 
                    newSelection = i; 
                    break;
                }
            }

            // Check for selection/move on the body of the annotation
            if (posUnscaled.x >= annRect.x && posUnscaled.x <= annRect.x + annRect.width && posUnscaled.y >= annRect.y && posUnscaled.y <= annRect.y + annRect.height) {
                this.dragMode = 'move';
                newSelection = i; 
                break;
            }
        }
    }
    
    if (newSelection > -1) {
        this.activeTool.set('select');
        this.selectedAnnotationIndex.set(newSelection);
        const ann = pageAnns[newSelection];
        this.dragStartPos = posUnscaled;
        if (ann.type === 'image' || ann.type === 'text') {
            this.dragStartAnn = { x: ann.x, y: ann.y, width: ann.width, height: ann.height };
        }
        if (ann.type === 'text') {
            this.textSettings.set({ font: ann.font, size: ann.size, color: ann.color, bold: ann.bold });
        }
    } else {
        this.selectedAnnotationIndex.set(-1);
        if (this.activeTool() === 'draw') {
            this.isDrawing = true; this.currentDrawingPath = [posUnscaled];
        } else if (this.activeTool() === 'text') {
            this.isEditingText.set(true);
            this.textEditingState.set({ x: pos.x, y: pos.y, width: 150, text: '' });
        }
    }
  }

  onMouseMove(event: MouseEvent) {
    if (this.dragMode === 'none' && !this.isDrawing) return;
    const pos = this.getMousePos(event, this.pdfCanvasRef()?.nativeElement);
    if (!pos) return;
    
    if (this.dragMode !== 'none') {
        const posUnscaled = { x: pos.x / this.scale(), y: pos.y / this.scale() };
        const annIndex = this.selectedAnnotationIndex();
        const pageIndex = this.pageOrder().indexOf(this.currentPageNum());
        if (annIndex === -1 || pageIndex === -1) return;
        const dx = posUnscaled.x - this.dragStartPos.x, dy = posUnscaled.y - this.dragStartPos.y;

        this.annotations.update(anns => {
            const newAnns = [...anns];
            const pageAnns = [...newAnns[pageIndex]];
            const ann = { ...pageAnns[annIndex] };

            if (this.dragMode === 'move' && (ann.type === 'image' || ann.type === 'text')) {
                ann.x = this.dragStartAnn.x + dx;
                ann.y = this.dragStartAnn.y + dy;
            } else if (this.dragMode === 'resize' && ann.type === 'image') {
                const newWidth = this.dragStartAnn.width + dx;
                if (newWidth > 10) { ann.width = newWidth; ann.height = newWidth * ann.originalAspectRatio; }
            }
            pageAnns[annIndex] = ann; newAnns[pageIndex] = pageAnns; return newAnns;
        });
        this.drawAnnotations();
    } else if (this.isDrawing) {
        const posUnscaled = { x: pos.x / this.scale(), y: pos.y / this.scale() };
        this.currentDrawingPath.push(posUnscaled);
        this.drawAnnotations();
    }
  }

  onMouseUp() {
    this.dragMode = 'none';
    if (this.isDrawing) {
      this.isDrawing = false;
      if(this.currentDrawingPath.length > 1) {
        const lineWidthInPoints = 2 / this.scale();
        this.addAnnotation({ type: 'draw', path: this.currentDrawingPath, color: 'red', lineWidth: lineWidthInPoints });
      }
      this.currentDrawingPath = [];
    }
  }

  onMouseLeave = () => this.onMouseUp();

  onDblClick(event: MouseEvent) {
    if (this.activeTab() !== 'edit' || this.isPageNumberPanelOpen()) return;

    const pos = this.getMousePos(event, this.pdfCanvasRef()?.nativeElement);
    if (!pos) return;

    const posUnscaled = { x: pos.x / this.scale(), y: pos.y / this.scale() };
    const pageIndex = this.pageOrder().indexOf(this.currentPageNum());
    if (pageIndex === -1) return;
    
    const pageAnns = this.annotations()[pageIndex];
    let editIndex = -1;

    for (let i = pageAnns.length - 1; i >= 0; i--) {
        const ann = pageAnns[i];
        if (ann.type === 'text') {
            if (posUnscaled.x >= ann.x && posUnscaled.x <= ann.x + ann.width &&
                posUnscaled.y >= ann.y && posUnscaled.y <= ann.y + ann.height) {
                editIndex = i;
                break;
            }
        }
    }

    if (editIndex > -1) {
        const annToEdit = pageAnns[editIndex] as TextAnnotation;
        this.selectedAnnotationIndex.set(editIndex);
        this.textSettings.set({ font: annToEdit.font, size: annToEdit.size, color: annToEdit.color, bold: annToEdit.bold });
        
        this.isEditingText.set(true);
        this.textEditingState.set({
            x: annToEdit.x * this.scale(),
            y: annToEdit.y * this.scale(),
            width: annToEdit.width * this.scale() + 20, // Add some padding
            text: annToEdit.text
        });
        
        this.drawAnnotations(); // Redraw to hide the annotation being edited
    }
  }

  addAnnotation(annotation: Annotation) {
    this.undoStack.update(stack => [...stack, JSON.parse(JSON.stringify(this.annotations()))]);
    this.redoStack.set([]);
    this.annotations.update(allAnns => {
      const pageIndex = this.pageOrder().indexOf(this.currentPageNum());
      if (pageIndex > -1) allAnns[pageIndex] = [...allAnns[pageIndex], annotation];
      return [...allAnns];
    });
    this.drawAnnotations();
  }
  
  onDeleteKeyPressed(): void {
    const annIndex = this.selectedAnnotationIndex();
    if (annIndex > -1) {
        this.undoStack.update(stack => [...stack, JSON.parse(JSON.stringify(this.annotations()))]);
        this.redoStack.set([]);
        this.annotations.update(anns => {
            const pageIndex = this.pageOrder().indexOf(this.currentPageNum());
            if(pageIndex === -1) return anns;
            const newAnns = [...anns];
            const pageAnns = [...newAnns[pageIndex]];
            pageAnns.splice(annIndex, 1);
            newAnns[pageIndex] = pageAnns;
            return newAnns;
        });
        this.selectedAnnotationIndex.set(-1);
        this.drawAnnotations();
    } else if (this.activeTab() === 'organize' && this.isPdfLoaded() && this.totalPages() > 1) {
       this.deletePage(this.currentPageNum());
    }
  }

  undo() {
    if (!this.canUndo()) return;
    const currentStack = this.undoStack(); const lastState = currentStack[currentStack.length - 1];
    this.redoStack.update(stack => [...stack, JSON.parse(JSON.stringify(this.annotations()))]);
    this.annotations.set(lastState); this.undoStack.update(stack => stack.slice(0, -1));
    this.selectedAnnotationIndex.set(-1); this.drawAnnotations();
  }
  redo() {
    if (!this.canRedo()) return;
    const currentStack = this.redoStack(); const nextState = currentStack[currentStack.length - 1];
    this.undoStack.update(stack => [...stack, JSON.parse(JSON.stringify(this.annotations()))]);
    this.annotations.set(nextState); this.redoStack.update(stack => stack.slice(0, -1));
    this.selectedAnnotationIndex.set(-1); this.drawAnnotations();
  }

  onDragStart = (index: number) => this.draggedPageIndex.set(index);
  onDragEnd = () => { 
    this.draggedPageIndex.set(-1); 
    this.dragOverIndex.set(-1);
    this.clearDragScroll();
  }
  onDragOver = (event: DragEvent, index: number) => { event.preventDefault(); this.dragOverIndex.set(index); };
  onDragLeave = () => this.dragOverIndex.set(-1);
  onDrop = (dropIndex: number) => { const dragIndex = this.draggedPageIndex(); if (dragIndex > -1) this.reorderPages(dragIndex, dropIndex); this.onDragEnd(); }
  reorderPages(dragIndex: number, dropIndex: number) {
      if (dropIndex === dragIndex || dropIndex === dragIndex + 1) return;
      const adjustedDropIndex = dragIndex < dropIndex ? dropIndex - 1 : dropIndex;
      this.pageOrder.update(currentOrder => { const newOrder = [...currentOrder], [movedPage] = newOrder.splice(dragIndex, 1); newOrder.splice(adjustedDropIndex, 0, movedPage); return newOrder; });
      this.annotations.update(currentAnns => { const newAnns = [...currentAnns], [movedAnn] = newAnns.splice(dragIndex, 1); newAnns.splice(adjustedDropIndex, 0, movedAnn); return newAnns; });
  }

  handleDragScroll(event: DragEvent): void {
    const panel = this.previewPanel().nativeElement;
    const rect = panel.getBoundingClientRect();
    
    const y = event.clientY - rect.top;
    const scrollZone = 50;
    const scrollSpeed = 10;
    const intervalTime = 25;
  
    if (y < scrollZone) {
      if (!this.dragScrollInterval) { // Only set if not already scrolling
        this.dragScrollInterval = setInterval(() => {
          panel.scrollTop -= scrollSpeed;
        }, intervalTime);
      }
    } else if (y > rect.height - scrollZone) {
      if (!this.dragScrollInterval) { // Only set if not already scrolling
        this.dragScrollInterval = setInterval(() => {
          panel.scrollTop += scrollSpeed;
        }, intervalTime);
      }
    } else {
      this.clearDragScroll(); // In the safe zone, clear interval
    }
  }
  
  clearDragScroll(): void {
    if (this.dragScrollInterval) {
      clearInterval(this.dragScrollInterval);
      this.dragScrollInterval = null;
    }
  }

  promptInsertPdf(index: number): void { const input = this.insertPdfInputRef()?.nativeElement; if (!input) return; this.insertionIndex.set(index); input.value = ''; input.click(); }
  async onInsertFileSelected(event: Event): Promise<void> {
      const input = event.target as HTMLInputElement; if (!input.files?.length) return;
      const newPdfBuffer = await input.files[0].arrayBuffer();
      await this.processAndInsertPdf(newPdfBuffer, this.insertionIndex());
  }
  private async processAndInsertPdf(pdfBuffer: ArrayBuffer, insertionIndex: number): Promise<void> {
    await this.confirmAndProceedWithOrganize(async () => {
        this.isSaving.set(true);
        try {
            const { PDFDocument } = PDFLib; const existingPdfDoc = this.pdfLibDoc(); if (!existingPdfDoc) return;
            const newPdfDoc = await PDFDocument.load(pdfBuffer);
            const copiedPages = await existingPdfDoc.copyPages(newPdfDoc, newPdfDoc.getPageIndices());
            copiedPages.forEach((page, i) => existingPdfDoc.insertPage(insertionIndex + i, page));
            const pdfBytes = await existingPdfDoc.save();
            await this.loadPdf(pdfBytes, true, insertionIndex + 1);
        } catch (error) { console.error('Error inserting PDF:', error); alert('PDFの挿入に失敗しました。'); } 
        finally { this.isSaving.set(false); }
    });
  }
  onDragOverInsert(event: DragEvent, index: number): void { event.preventDefault(); if (event.dataTransfer) event.dataTransfer.dropEffect = 'copy'; this.isDraggingOverInsert.set(index); }
  onDragLeaveInsert = () => this.isDraggingOverInsert.set(-1);
  async onDropInsert(event: DragEvent, index: number): Promise<void> {
      event.preventDefault(); this.isDraggingOverInsert.set(-1);
      const file = event.dataTransfer?.files[0];
      if (file?.type === 'application/pdf') await this.processAndInsertPdf(await file.arrayBuffer(), index);
      else if (file) alert('PDFファイルのみドロップできます。');
  }

  private confirmAndProceedWithOrganize = async (action: () => Promise<void>): Promise<void> => { if (this.annotations().some(p => p.length > 0) && !confirm('この操作を実行すると、追加した注釈は失われます。よろしいですか？')) return; await action(); }
  
  async rotatePage(degrees: 90 | -90, pageNumToRotate?: number) {
    const pageToRotate = pageNumToRotate ?? this.currentPageNum();
    await this.confirmAndProceedWithOrganize(async () => {
      const pdfLibDoc = this.pdfLibDoc();
      if (!pdfLibDoc) return;
      this.isSaving.set(true);
      try {
        const { degrees: degreesFn } = PDFLib;
        const page = pdfLibDoc.getPage(pageToRotate - 1);
        page.setRotation(degreesFn(page.getRotation().angle + degrees));
        await this.loadPdf(await pdfLibDoc.save(), true, this.currentPageNum());
      } catch (e) {
        console.error('Error rotating page:', e);
        alert('ページの回転に失敗しました。');
      } finally {
        this.isSaving.set(false);
      }
    });
  }

  async deletePage(pageNumToDelete: number) {
    if (this.totalPages() <= 1) { alert('最後のページは削除できません。'); return; }
    await this.confirmAndProceedWithOrganize(async () => {
        const pdfLibDoc = this.pdfLibDoc(); if (!pdfLibDoc) return; this.isSaving.set(true);
        try {
            pdfLibDoc.removePage(pageNumToDelete - 1);
            const newPageNum = Math.min(pageNumToDelete, pdfLibDoc.getPageCount());
            await this.loadPdf(await pdfLibDoc.save(), true, newPageNum);
        } catch (e) { console.error('Error deleting page:', e); alert('ページの削除に失敗しました。'); } finally { this.isSaving.set(false); }
    });
  }
  async duplicatePage(pageNumToDuplicate: number) {
    await this.confirmAndProceedWithOrganize(async () => {
      const pdfLibDoc = this.pdfLibDoc();
      if (!pdfLibDoc) return;
  
      this.isSaving.set(true);
      try {
        const pageIndexToDuplicate = pageNumToDuplicate - 1;
        const [copiedPage] = await pdfLibDoc.copyPages(pdfLibDoc, [pageIndexToDuplicate]);
        pdfLibDoc.insertPage(pageIndexToDuplicate + 1, copiedPage);
  
        const pdfBytes = await pdfLibDoc.save();
        const newCurrentPage = pageIndexToDuplicate + 2; 
  
        await this.loadPdf(pdfBytes, true, newCurrentPage);
  
      } catch (e) {
        console.error('Error duplicating page:', e);
        alert('ページの複製に失敗しました。');
      } finally {
        this.isSaving.set(false);
      }
    });
  }
  async extractPage() {
    const pdfLibDoc = this.pdfLibDoc(); if (!pdfLibDoc) return; this.isSaving.set(true);
    try {
        const { PDFDocument } = PDFLib; const newPdfDoc = await PDFDocument.create();
        const [copiedPage] = await newPdfDoc.copyPages(pdfLibDoc, [this.currentPageNum() - 1]);
        newPdfDoc.addPage(copiedPage);
        const pdfBytes = await newPdfDoc.save();
        const blob = new Blob([pdfBytes], { type: 'application/pdf' });
        const link = document.createElement('a'); link.href = URL.createObjectURL(blob);
        link.download = `${this.originalFileName()}(ページ ${this.currentPageNum()} 抽出).pdf`;
        link.click(); 
        URL.revokeObjectURL(link.href);
    } catch (e) { console.error('Error extracting page:', e); alert('ページの抽出に失敗しました。'); } finally { this.isSaving.set(false); }
  }

  // --- Text Methods ---
  onTextInput(event: Event) {
    const textarea = event.target as HTMLTextAreaElement;
    this.textEditingState.update(state => state ? {...state, text: textarea.value} : null);
    textarea.style.height = 'auto'; // Reset height
    textarea.style.height = `${textarea.scrollHeight}px`; // Set to scroll height
  }
  
  finalizeTextAnnotation() {
    if (!this.isEditingText()) return;
    const state = this.textEditingState();
    if (!state || !state.text.trim()) {
      this.cancelTextAnnotation();
      return;
    }

    const settings = this.textSettings();
    const { width, height } = this.measureText(state.text, settings);
    const trimmedText = state.text.trim();
    
    const editIndex = this.selectedAnnotationIndex();
    const pageIndex = this.pageOrder().indexOf(this.currentPageNum());
    const annToEdit = (editIndex > -1 && pageIndex > -1) ? this.annotations()[pageIndex][editIndex] : null;
    
    const pixelRatio = window.devicePixelRatio || 1;
    const totalScale = this.scale() * pixelRatio;

    if (annToEdit && annToEdit.type === 'text') {
      this.undoStack.update(stack => [...stack, JSON.parse(JSON.stringify(this.annotations()))]);
      this.redoStack.set([]);
      this.annotations.update(anns => {
        const newAnns = [...anns];
        const pageAnns = [...newAnns[pageIndex]];
        pageAnns[editIndex] = {
          ...annToEdit,
          text: trimmedText,
          ...settings,
          width: width / totalScale,
          height: height / totalScale,
        };
        newAnns[pageIndex] = pageAnns;
        return newAnns;
      });
    } else {
      this.addAnnotation({
        type: 'text', text: trimmedText,
        x: state.x / this.scale(), y: state.y / this.scale(),
        ...settings,
        width: width / totalScale, height: height / totalScale
      });
    }

    this.isEditingText.set(false);
    this.textEditingState.set(null);
    this.drawAnnotations();
  }
  
  cancelTextAnnotation() {
    this.isEditingText.set(false);
    this.textEditingState.set(null);
    this.drawAnnotations();
  }

  handleTextEnter(event: KeyboardEvent) {
    if (event.key === 'Enter' && !event.shiftKey) {
      event.preventDefault();
      this.finalizeTextAnnotation();
    }
  }

  measureText(text: string, settings: TextSettings): { width: number, height: number } {
    const drawCanvas = this.drawCanvasRef()?.nativeElement;
    if (!drawCanvas) return { width: 100, height: settings.size };
    const context = drawCanvas.getContext('2d')!;
    const lines = text.split('\n');
    const pixelRatio = window.devicePixelRatio || 1;
    const currentScale = this.scale() * pixelRatio;
    const font = `${settings.bold ? 'bold ' : ''}${settings.size * currentScale}px ${this.getCanvasFontFamily(settings.font)}`;
    context.font = font;
    const lineHeight = settings.size * 1.2 * currentScale;
    const height = lines.length * lineHeight;
    const width = Math.max(...lines.map(line => context.measureText(line).width));
    return { width, height };
  }

  updateTextSetting(key: keyof TextSettings, value: any) {
    this.textSettings.update(s => ({ ...s, [key]: value }));
    
    const selectedIndex = this.selectedAnnotationIndex();
    const pageIndex = this.pageOrder().indexOf(this.currentPageNum());

    if (selectedIndex !== -1 && pageIndex !== -1) {
      this.annotations.update(anns => {
        const pageAnns = anns[pageIndex];
        const ann = pageAnns?.[selectedIndex];
        if (ann?.type === 'text') {
            const newAnns = [...anns];
            const newPageAnns = [...pageAnns];
            const updatedAnn = { ...ann, ...this.textSettings() };

            const { width, height } = this.measureText(updatedAnn.text, updatedAnn);
            const pixelRatio = window.devicePixelRatio || 1;
            const totalScale = this.scale() * pixelRatio;
            updatedAnn.width = width / totalScale;
            updatedAnn.height = height / totalScale;
            
            newPageAnns[selectedIndex] = updatedAnn;
            newAnns[pageIndex] = newPageAnns;
            return newAnns;
        }
        return anns;
      });
      this.drawAnnotations();
    }
  }

  // --- Signature Methods ---
  openSignatureModal() {
    this.isSignatureModalOpen.set(true);
    this.signaturePadPaths.set([]);
    this.signatureUndoStack.set([]);
    this.signatureRedoStack.set([]);
    this.signatureColor.set('black');

    afterNextRender(() => {
      const sigCanvas = this.signatureCanvasRef()?.nativeElement;
      // It's possible the modal was closed before this render hook runs
      if (sigCanvas && this.isSignatureModalOpen()) {
        // Use the canvas's own client dimensions to set its resolution.
        // This ensures the drawing buffer size matches the CSS display size,
        // correcting coordinate mismatches caused by borders or padding on the parent.
        sigCanvas.width = sigCanvas.clientWidth;
        sigCanvas.height = sigCanvas.clientHeight;
        this.drawSignaturePad();
      }
    });
  }

  closeSignatureModal() {
    this.isSignatureModalOpen.set(false);
  }

  drawSignaturePad() {
    const canvas = this.signatureCanvasRef()?.nativeElement;
    if (!canvas) return;
    const ctx = canvas.getContext('2d')!;
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    for (const stroke of this.signaturePadPaths()) {
        if (stroke.path.length === 0) continue;
        ctx.beginPath();
        ctx.moveTo(stroke.path[0].x, stroke.path[0].y);
        stroke.path.slice(1).forEach(p => ctx.lineTo(p.x, p.y));
        ctx.strokeStyle = stroke.color;
        ctx.lineWidth = 2;
        ctx.lineCap = 'round';
        ctx.lineJoin = 'round';
        ctx.stroke();
    }
  }

  onSignatureMouseDown(event: MouseEvent) {
    const pos = this.getSignatureMousePos(event);
    if (!pos) return;

    this.isSigning = true;
    this.signaturePadCurrentPath = [pos];
  }

  onSignatureMouseMove(event: MouseEvent) {
    if (!this.isSigning) return;
    
    const pos = this.getSignatureMousePos(event);
    if (!pos) return;

    this.signaturePadCurrentPath.push(pos);

    const canvas = this.signatureCanvasRef()?.nativeElement;
    if (!canvas) return;
    const ctx = canvas.getContext('2d')!;
    
    // Clear and redraw for live preview
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    // Redraw completed paths
    for (const stroke of this.signaturePadPaths()) {
        if (stroke.path.length === 0) continue;
        ctx.beginPath();
        ctx.moveTo(stroke.path[0].x, stroke.path[0].y);
        stroke.path.slice(1).forEach(p => ctx.lineTo(p.x, p.y));
        ctx.strokeStyle = stroke.color;
        ctx.lineWidth = 2;
        ctx.lineCap = 'round';
        ctx.lineJoin = 'round';
        ctx.stroke();
    }

    // Draw current in-progress path
    if (this.signaturePadCurrentPath.length > 1) {
      ctx.beginPath();
      ctx.moveTo(this.signaturePadCurrentPath[0].x, this.signaturePadCurrentPath[0].y);
      this.signaturePadCurrentPath.slice(1).forEach(p => ctx.lineTo(p.x, p.y));
      ctx.strokeStyle = this.signatureColor();
      ctx.lineWidth = 2;
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      ctx.stroke();
    }
  }

  onSignatureMouseUp() {
    if (!this.isSigning) return;
    this.isSigning = false;
    if (this.signaturePadCurrentPath.length > 1) {
      this.signatureUndoStack.update(stack => [...stack, this.signaturePadPaths()]);
      this.signatureRedoStack.set([]);
      this.signaturePadPaths.update(paths => [...paths, { path: this.signaturePadCurrentPath, color: this.signatureColor() }]);
    }
    this.signaturePadCurrentPath = [];
    this.drawSignaturePad(); // Redraw to consolidate the new path
  }

  onSignatureMouseLeave() {
    if(this.isSigning) this.onSignatureMouseUp();
  }
  
  undoSignature() {
    if (!this.canUndoSignature()) return;
    const currentStack = this.signatureUndoStack();
    const lastState = currentStack[currentStack.length - 1];
    this.signatureRedoStack.update(stack => [...stack, this.signaturePadPaths()]);
    this.signaturePadPaths.set(lastState);
    this.signatureUndoStack.update(stack => stack.slice(0, -1));
    this.drawSignaturePad();
  }
  
  redoSignature() {
    if (!this.canRedoSignature()) return;
    const currentStack = this.signatureRedoStack();
    const nextState = currentStack[currentStack.length - 1];
    this.signatureUndoStack.update(stack => [...stack, this.signaturePadPaths()]);
    this.signaturePadPaths.set(nextState);
    this.signatureRedoStack.update(stack => stack.slice(0, -1));
    this.drawSignaturePad();
  }

  clearSignature() {
    if (this.signaturePadPaths().length > 0) {
      this.signatureUndoStack.update(stack => [...stack, this.signaturePadPaths()]);
    }
    this.signatureRedoStack.set([]);
    this.signaturePadPaths.set([]);
    this.drawSignaturePad();
  }
  
  createSignature() {
    const canvas = this.signatureCanvasRef()?.nativeElement;
    if (!canvas || this.signaturePadPaths().length === 0) return;

    const bounds = this.getSignatureBounds(canvas);
    if (!bounds) return;

    const trimmedCanvas = document.createElement('canvas');
    const trimmedCtx = trimmedCanvas.getContext('2d')!;
    const padding = 10;
    trimmedCanvas.width = bounds.width + padding * 2;
    trimmedCanvas.height = bounds.height + padding * 2;
    trimmedCtx.drawImage(canvas, bounds.x, bounds.y, bounds.width, bounds.height, padding, padding, bounds.width, bounds.height);

    const imageData = trimmedCanvas.toDataURL('image/png');
    const img = new Image();
    img.onload = () => {
        const drawCanvas = this.drawCanvasRef()?.nativeElement;
        if (!drawCanvas) return;
        const scale = this.scale();
        const aspectRatio = img.height / img.width;
        let pointWidth = 150 / scale;
        let pointHeight = pointWidth * aspectRatio;

        this.addAnnotation({
          type: 'image',
          x: (drawCanvas.clientWidth / scale - pointWidth) / 2,
          y: (drawCanvas.clientHeight / scale - pointHeight) / 2,
          width: pointWidth,
          height: pointHeight,
          imageData,
          originalAspectRatio: aspectRatio
        });
        const pageIndex = this.pageOrder().indexOf(this.currentPageNum());
        this.selectedAnnotationIndex.set(this.annotations()[pageIndex].length - 1);
        this.closeSignatureModal();
    };
    img.src = imageData;
  }

  private getSignatureMousePos(event: MouseEvent): { x: number; y: number } | null {
    const canvas = this.signatureCanvasRef()?.nativeElement;
    if (!canvas) return null;
    const rect = canvas.getBoundingClientRect();

    // Scale the mouse coordinates from the CSS display size to the canvas's internal resolution.
    // This corrects any discrepancies caused by borders, padding, or browser sub-pixel rendering.
    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;

    return {
      x: (event.clientX - rect.left) * scaleX,
      y: (event.clientY - rect.top) * scaleY,
    };
  }
  
  private getSignatureBounds(canvas: HTMLCanvasElement) {
    const ctx = canvas.getContext('2d')!;
    const w = canvas.width, h = canvas.height;
    const imageData = ctx.getImageData(0, 0, w, h).data;
    let minX = w, minY = h, maxX = 0, maxY = 0;
    let foundPixel = false;
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        const alpha = imageData[(y * w + x) * 4 + 3];
        if (alpha > 0) {
          minX = Math.min(minX, x); minY = Math.min(minY, y);
          maxX = Math.max(maxX, x); maxY = Math.max(maxY, y);
          foundPixel = true;
        }
      }
    }
    if (!foundPixel) return null;
    return { x: minX, y: minY, width: maxX - minX, height: maxY - minY };
  }

  private hexToRgb(hex: string): { r: number; g: number; b: number } {
    const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    return result ? {
        r: parseInt(result[1], 16) / 255,
        g: parseInt(result[2], 16) / 255,
        b: parseInt(result[3], 16) / 255,
    } : { r: 0, g: 0, b: 0 };
  }
  
  // --- Password Protection ---
  openPasswordModal() {
    this.password.set('');
    this.confirmPassword.set('');
    this.passwordError.set(null);
    this.isPasswordModalOpen.set(true);
  }

  closePasswordModal() {
    this.isPasswordModalOpen.set(false);
  }

  applyPasswordAndSave() {
    const pw = this.password();
    const confirmPw = this.confirmPassword();

    if (!pw) {
      this.passwordError.set('パスワードを入力してください。');
      return;
    }
    if (pw !== confirmPw) {
      this.passwordError.set('パスワードが一致しません。');
      return;
    }

    this.passwordError.set(null);
    this.savePdf(pw);
    this.closePasswordModal();
  }

  async savePdf(password?: string) {
    const pdfLibDoc = this.pdfLibDoc(); if (!pdfLibDoc) return; this.isSaving.set(true);
    try {
      const { PDFDocument, rgb, StandardFonts, degrees, pushGraphicsState, popGraphicsState, translate, rotateDegrees } = PDFLib;
      const newPdfDoc = await PDFDocument.create();
      
      const fontCache = new Map<string, any>();
      const getFont = async (fontId: string) => {
          if (fontCache.has(fontId)) return fontCache.get(fontId);
          const font = await newPdfDoc.embedFont(fontId as any);
          fontCache.set(fontId, font);
          return font;
      }
      
      const pageIndices = this.pageOrder().map(p => p - 1);
      const copiedPages = await newPdfDoc.copyPages(pdfLibDoc, pageIndices);
      copiedPages.forEach(p => newPdfDoc.addPage(p));
      const pages = newPdfDoc.getPages();
      
      const settings = this.pageNumberSettings();
      if (settings.visible) {
          const helveticaFont = await getFont(StandardFonts.Helvetica);
          const margin = { small: 18, medium: 36, large: 54 }[settings.margin], fontSize = 12;
          for (let i = 0; i < pages.length; i++) {
              const page = pages[i], { width, height } = page.getSize();
              const pageNumText = `${i + settings.startPage}`;
              const textWidth = helveticaFont.widthOfTextAtSize(pageNumText, fontSize);
              let x = 0, y = 0;
              if (settings.position.includes('left')) x = margin;
              else if (settings.position.includes('center')) x = (width - textWidth) / 2;
              else if (settings.position.includes('right')) x = width - textWidth - margin;
              if (settings.position.includes('top')) y = height - fontSize - margin;
              else if (settings.position.includes('middle')) y = (height - fontSize) / 2;
              else if (settings.position.includes('bottom')) y = margin;
              page.drawText(pageNumText, { x, y, font: helveticaFont, size: fontSize, color: rgb(0, 0, 0) });
          }
      }

      const watermark = this.watermarkSettings();
      if (watermark.visible && watermark.text) {
          const watermarkFont = await getFont(watermark.font as any || StandardFonts.Helvetica);
          const rgbColor = this.hexToRgb(watermark.color);
          for (const page of pages) {
              const { width, height } = page.getSize();
              const text = watermark.text;
              const fontSize = watermark.size;

              const textWidth = watermarkFont.widthOfTextAtSize(text, fontSize);
              const ascent = watermarkFont.heightAtSize(fontSize);
              const totalHeight = watermarkFont.heightAtSize(fontSize, { descender: true });
              const descent = totalHeight - ascent;
              const yOffset = -(ascent - descent) / 2;

              page.pushOperators(
                pushGraphicsState(),
                translate(width / 2, height / 2),
                rotateDegrees(45)
              );

              page.drawText(text, {
                x: -textWidth / 2,
                y: yOffset,
                font: watermarkFont,
                size: fontSize,
                color: rgb(rgbColor.r, rgbColor.g, rgbColor.b),
                opacity: watermark.opacity,
              });

              page.pushOperators(popGraphicsState());
          }
      }

      for (let i = 0; i < pages.length; i++) {
        const pageAnns = this.annotations()[i]; if (!pageAnns?.length) continue;
        const page = pages[i];
        const { height: pdfPageHeight } = page.getSize();
        
        for (const ann of pageAnns) {
          if (ann.type === 'image') {
            const imageBytes = ann.imageData.startsWith('data:image/png') ? 
              await newPdfDoc.embedPng(ann.imageData) : await newPdfDoc.embedJpg(ann.imageData);
            page.drawImage(imageBytes, { x: ann.x, y: pdfPageHeight - ann.y - ann.height, width: ann.width, height: ann.height });
          } else if (ann.type === 'text') {
            const fontId = ann.bold ? (ann.font.includes('Bold') ? ann.font : `${ann.font}-Bold`) : (ann.font.replace('-Bold', ''));
            const pdfFont = await getFont(fontId as any);
            const rgbColor = this.hexToRgb(ann.color);
            page.drawText(ann.text, { x: ann.x, y: pdfPageHeight - ann.y - ann.size, font: pdfFont, size: ann.size, color: rgb(rgbColor.r, rgbColor.g, rgbColor.b), lineHeight: ann.size * 1.2 });
          } else if (ann.type === 'draw') {
            const svgPath = `M ${ann.path[0].x} ${pdfPageHeight - ann.path[0].y} ` + ann.path.slice(1).map(p => `L ${p.x} ${pdfPageHeight - p.y}`).join(' ');
            page.drawSvgPath(svgPath, { borderColor: rgb(1, 0, 0), borderWidth: ann.lineWidth });
          }
        }
      }

      if (password) {
        if (typeof newPdfDoc.encrypt === 'function') {
            await newPdfDoc.encrypt({
              userPassword: password,
              ownerPassword: password,
              permissions: {
                printing: 'highResolution',
                modifying: true,
                copying: true,
                annotating: true,
                fillingForms: true,
                contentAccessibility: true,
                documentAssembly: true,
              }
            });
        } else {
            alert('この環境では暗号化機能（パスワード保護）がサポートされていません。HTTPS環境で実行しているか確認してください。');
        }
      }
      
      const pdfBytes = await newPdfDoc.save();
      const blob = new Blob([pdfBytes], { type: 'application/pdf' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      const fileNameSuffix = password ? '(編集・保護済み)' : '(編集済み)';
      link.download = `${this.originalFileName()}${fileNameSuffix}.pdf`;
      link.click();
      URL.revokeObjectURL(link.href);
    } catch (e) { console.error('Failed to save PDF:', e); alert('PDFの保存中にエラーが発生しました。'); } 
    finally { this.isSaving.set(false); }
  }
}