// Builds a workbook in a module Web Worker and transfers its bytes back. The
// WebAssembly loader uses top-level await, so the worker announces when it has
// loaded instead of relying on messages posted before its listener exists.
import { Workbook } from '../../../wrapper.browser.mjs';

self.addEventListener('message', (event) => {
  try {
    const workbook = new Workbook();
    workbook.addWorksheet().writeRows(0, 0, event.data.rows);
    const bytes = new Uint8Array(workbook.saveToBuffer());
    // oxlint-disable-next-line unicorn/require-post-message-target-origin -- a worker's postMessage has no target origin
    self.postMessage({ type: 'workbook', bytes }, [bytes.buffer]);
  } catch (error) {
    // oxlint-disable-next-line unicorn/require-post-message-target-origin -- a worker's postMessage has no target origin
    self.postMessage({ type: 'error', error: String(error) });
  }
});
// oxlint-disable-next-line unicorn/require-post-message-target-origin -- a worker's postMessage has no target origin
self.postMessage({ type: 'ready' });
