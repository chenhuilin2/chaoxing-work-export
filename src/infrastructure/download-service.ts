import type { ExportArtifact } from '../domain/export-options';
import { delay } from '../utils/async';

export class DownloadService {
  async download(artifact: ExportArtifact): Promise<void> {
    const url = URL.createObjectURL(artifact.blob);
    try {
      const anchor = document.createElement('a');
      anchor.href = url;
      anchor.download = artifact.filename;
      anchor.style.display = 'none';
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
    } finally {
      window.setTimeout(() => URL.revokeObjectURL(url), 2_000);
    }
  }

  async downloadMany(artifacts: readonly ExportArtifact[]): Promise<void> {
    for (const artifact of artifacts) {
      await this.download(artifact);
      await delay(350);
    }
  }
}
