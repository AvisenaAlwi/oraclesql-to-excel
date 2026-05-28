# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [1.1.3] - 2026-05-28

### Fixed

- **Memory still grew linearly behind a reverse proxy** — when Node.js is behind nginx or any reverse proxy, the proxy accepts data from Node.js instantly (local socket), so `write()` on the output stream always returns `true`. The event-driven drain loop introduced in `v1.1.2` never triggered. Data accumulated in Node.js's internal `outputData` buffer (visible as RSS growth) while the proxy slowly forwarded data to the slow end client.

  Added a second, RSS-based fallback: after each Oracle fetch batch the library checks `process.memoryUsage().rss`. If it exceeds the configured threshold, it polls every 200 ms and waits until RSS drops before fetching the next batch. This approach works regardless of proxy topology because it monitors the process's own memory, not stream events.

### Changed

- **`.backpressureThreshold(bytes)` is no longer a no-op** — repurposed as the RSS threshold for the polling fallback. Default changed from `16 MB` to `512 MB`. Set lower if your PM2 `--max-memory-restart` limit is below 1 GB.

---

## [1.1.2] - 2026-05-28

### Fixed

- **Archiver buffer accumulation not fully drained between batches** — `v1.1.1` waited for one `drain` event on the output stream before fetching the next Oracle batch. However, after a TCP `drain`, the ExcelJS internal archiver immediately flushes its own queued data back into the stream, re-filling the buffer. This cycle meant only one TCP buffer's worth of data (~4–8 MB) was drained per batch pause, while the archiver continued to accumulate. Over many batches this still caused out-of-memory on slow clients.

  The drain wait is now a loop: after each `drain` event the code yields one event-loop turn (via `setImmediate`) to let the archiver flush pending data, then checks for backpressure again. The loop exits only when no new data arrives in one turn — meaning the archiver is truly empty and the next Oracle fetch can safely begin.

- **`bytesSinceDrain` threshold could trigger false-positive drain waits** — the proactive byte-count check introduced in `v1.1.0` set `needsDrain = true` based on bytes written since last drain, even when `write()` was still returning `true` (no actual backpressure). With the drain loop now in place, this check is unnecessary and was removed. Backpressure detection now relies solely on `write()` returning `false`, which is the authoritative Node.js signal.

### Deprecated

- **`.backpressureThreshold(bytes)`** — now a no-op. Kept for API compatibility with `v1.1.0`. See fix above.

---

## [1.1.1] - 2026-05-26

### Fixed

- **Backpressure check was using the wrong signal** — `v1.1.0` checked `writableLength` on the output stream (`res`), but data actually accumulates in ExcelJS's internal archiver buffer which sits upstream of `res`. This caused the check to never trigger and memory still grew unboundedly on slow clients.

  The fix intercepts `write()` on the output stream directly. When `write()` returns `false` — Node.js's authoritative backpressure signal — the Oracle fetch is paused until the stream drains.

- **Client disconnect leaked Oracle connection for up to 30 seconds** — when a browser cancelled a download mid-stream, the `drain` event never fired, and the export continued fetching from Oracle until the 30-second safety timeout. The stream `close` and `error` events now abort the fetch immediately, closing the Oracle connection and result set right away.

---

## [1.1.0] - 2026-05-26

### Fixed

- **Backpressure not respected during `.pipe()` streaming** — when using `.pipe(res)` with a slow HTTP client, Oracle rows were fetched and committed to the ExcelJS stream faster than the client could receive them. This caused the Node.js stream internal buffers to grow unboundedly, leading to process memory exhaustion on large exports (1M+ rows).

  The library now checks the output stream's `writableLength` after each Oracle fetch batch. If it exceeds the threshold, the fetch is paused until the stream drains before continuing.

### Added

- **`.backpressureThreshold(bytes)`** — configurable byte threshold that controls when the Oracle fetch is paused to allow the output stream to drain. Default: `16777216` (16 MB). Only applies to `.pipe()`; has no effect on `.run()` or `.toBuffer()`.

---

## [1.0.0] - 2026-05-01

### Added

- Initial release.
- Streaming Oracle SQL → Excel (`.xlsx`) via `ExcelJS` streaming writer.
- Fluent builder API: `.connectionFactory()`, `.sheet()`, `.run()`, `.pipe()`, `.toBuffer()`.
- Per-sheet: `.sql()`, `.columns()`, `.fetchSize()`, `.maxRowsPerSheet()`, `.freezeHeader()`, `.autoFilter()`, `.headerStyle()`, `.docHeader()`, `.showTotalRows()`, `.onRowError()`.
- Multi-sheet and automatic split sheets.
- Document header rows with simple and multi-column (merge) modes.
- Row range summary via parallel `COUNT(*)` query.
- Progress callback (`.onProgress()`).
- ZIP compression toggle (`.compress()`).
- Auto-detect columns from Oracle metadata when `.columns()` is omitted.
