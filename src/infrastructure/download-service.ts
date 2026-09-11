import type { ExportArtifact } from '../domain/export-options';

export class DownloadService {
  // 原版下载方式：创建临时 <a> 点击后立即释放 URL
  async download(artifact: ExportArtifact): Promise<void> {
    const url = URL.createObjectURL(artifact.blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = artifact.filename;
    a.click();
    URL.revokeObjectURL(url);
  }

  // 多文件下载间隔 300ms，避免多个下载被浏览器拦截
  async downloadMany(artifacts: readonly ExportArtifact[]): Promise<void> {
    for (const artifact of artifacts) {
      await this.download(artifact);
      // 浏览器下载间隔，避免多个下载被拦截
      await new Promise((resolve) => setTimeout(resolve, 300));
    }
  }
}
