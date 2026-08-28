// Shared write helpers for the office tools.

import { mkdirSync } from 'node:fs'
import { dirname } from 'node:path'

/**
 * Create the parent directory of `filePath` (and its ancestors) when missing.
 * Call before every tool-side file write so a destination inside a
 * not-yet-existing directory succeeds instead of failing with ENOENT.
 *
 * @param filePath Destination file path whose parent directory is ensured.
 */
export function ensureParentDir(filePath: string): void {
  mkdirSync(dirname(filePath), { recursive: true })
}
