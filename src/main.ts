import { AppController } from './application/app-controller';
import { ChapterExtractionService } from './application/chapter-extraction-service';
import { ChapterLocator } from './application/chapter-locator';
import { ExtractionService } from './application/extraction-service';
import { TaskTabLocator } from './application/task-tab-locator';
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
  // 章节切换与任务卡切换共用同一个定位器实例
  const taskTabs = new TaskTabLocator();
  const chapterService = new ChapterExtractionService(extractionService, new ChapterLocator(), taskTabs);
  const controller = new AppController(
    new PanelView(),
    extractionService,
    chapterService,
    new ExportService(),
    frameBridge,
    taskTabs,
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
