/**
 * 「当前章节」变化检测，供「切换章节自动提取」使用。
 *
 * 学习通切章走 `getTeacherAjax()`：先把 `#mainid` 清空重建，再把激活态类移到新的
 * `#coursetree .posCatalog_select` 上 —— **URL 不变、页面不刷新**，脚本也不会重新注入，
 * 所以「进入页面自动提取」那一次之后再没有任何触发点，只能在 DOM 侧观察。
 *
 * 判定用的指纹是「激活章节的标识」（容器 id 即 `cur<chapterId>`），而不是「DOM 是否变化」：
 * 章节页每几秒就有自己的 DOM 变动（进度上报、提示气泡），按 DOM 变化触发会疯狂误提取。
 *
 * 两条边界：
 * - 读不到标识时（不在章节页 / 目录树还没渲染）既不更新基线也不上报，短暂读空不会被当成切章；
 * - 首次读到标识只用来建立基线，因此「进入页面后目录树才渲染出来」不会被当成切章。
 */
export class ChapterChangeWatcher {
  private baseline: string | null = null;

  constructor(private readonly readKey: () => string | null) {}

  /** 把当前章节记为基线：用于跳过「自己触发的切换」（批量提取逐章遍历） */
  reset(): void {
    this.baseline = this.read();
  }

  /** 当前章节标识；读不到（或读取失败）返回 null */
  currentKey(): string | null {
    return this.read();
  }

  /** 同步一次；确认已经切到别的章节时返回 true（同一章节或读不到时返回 false） */
  poll(): boolean {
    const key = this.read();
    if (key === null) return false;
    if (this.baseline === null) {
      this.baseline = key;
      return false;
    }
    if (key === this.baseline) return false;
    this.baseline = key;
    return true;
  }

  private read(): string | null {
    try {
      return this.readKey();
    } catch {
      // 宿主页面结构异常不应打断轮询
      return null;
    }
  }
}
