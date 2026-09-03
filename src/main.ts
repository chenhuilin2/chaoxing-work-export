import { AppController } from './application/app-controller';
import { ChapterExtractionService } from './application/chapter-extraction-service';
import { ExtractionService } from './application/extraction-service';
import { ExportService } from './exporters/export-service';
import { CompositeExtractor } from './extractors/composite-extractor';
import { FrameAgent, FrameBridge } from './infrastructure/frame-bridge';
import { PanelView } from './ui/panel-view';

function inIframe(): boolean {
  try {
    return window.self !== window.top;
  } catch {
    return true;
  }
}

function bootTopWindow(): void {
  const frameBridge = new FrameBridge();
  const extractionService = new ExtractionService(new CompositeExtractor(), frameBridge);
  const chapterService = new ChapterExtractionService(extractionService);
  const controller = new AppController(
    new PanelView(),
    extractionService,
    chapterService,
    new ExportService(),
    frameBridge,
  );
  controller.start();
}

function boot(): void {
  if (inIframe()) {
    new FrameAgent().start();
    return;
  }
  if (document.body) bootTopWindow();
  else document.addEventListener('DOMContentLoaded', bootTopWindow, { once: true });
}

boot();
