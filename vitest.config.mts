import { defineConfig } from 'vitest/config'

/**
 * 覆盖率门禁。只设全局下限，不做 per-file 强制：`src/excel-audit.ts`（1271 行）
 * 当前 60.58%，是最大的覆盖短板；若设 per-file 阈值会逼着为了数字去补边缘分支，
 * 反而偏离价值。先防整体倒退，单模块提升另作议程。
 *
 * 基线（2026-11-21）：整体 76.03% lines；
 * docx 97.0 / excel 86.8 / excel-audit 60.6 / fonts 81.3 / pdf 87.9 / ppt 79.7。
 */
export default defineConfig({
  test: {
    coverage: {
      provider: 'v8',
      include: ['src/**/*.ts'],
      reporter: ['text-summary'],
      thresholds: { lines: 76, statements: 76 },
    },
  },
})
