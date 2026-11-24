/**
 * Memory-efficient Streaming CSV Parser for huge files (up to 1 GB).
 * Reads files sequentially in 4 MB slices using File.slice() + TextDecoder({ stream: true })
 * so browser memory stays low even on massive multi-lakh-row exports.
 */

export interface ParseProgress {
  percent: number;
  loadedBytes: number;
  totalBytes: number;
  rowsParsed: number;
  stage: string;
}

export function yieldToMainThread(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

/**
 * Parses a single CSV line handling RFC-4180 quotes and escaped double quotes ("").
 */
export function parseCsvLine(line: string, delimiter = ','): string[] {
  const fields: string[] = [];
  let current = '';
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const ch = line[i];

    if (inQuotes) {
      if (ch === '"') {
        if (i + 1 < line.length && line[i + 1] === '"') {
          current += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        current += ch;
      }
    } else {
      if (ch === '"') {
        inQuotes = true;
      } else if (ch === delimiter) {
        fields.push(current.trim());
        current = '';
      } else {
        current += ch;
      }
    }
  }

  fields.push(current.trim());
  return fields;
}

/**
 * Detects delimiter from the first few non-empty lines (supports comma, semicolon, tab).
 */
function detectCsvDelimiter(sampleLines: string[]): string {
  const candidates = [',', '\t', ';', '|'];
  let bestDelim = ',';
  let bestScore = 0;

  for (const delim of candidates) {
    let score = 0;
    for (const line of sampleLines) {
      if (!line.trim()) continue;
      score += parseCsvLine(line, delim).length;
    }
    if (score > bestScore) {
      bestScore = score;
      bestDelim = delim;
    }
  }

  return bestDelim;
}

/**
 * Streams a CSV File in 4 MB chunks and delivers batches of parsed string[] rows
 * to the caller without loading the entire file into RAM.
 */
export async function streamCsvFileInBatches(
  file: File,
  onBatch: (rows: string[][], batchIndex: number) => void,
  onProgress?: (progress: ParseProgress) => void,
  chunkSizeBytes = 4 * 1024 * 1024 // 4 MB slices
): Promise<number> {
  const totalBytes = file.size;
  const decoder = new TextDecoder('utf-8');
  let offset = 0;
  let leftover = '';
  let delimiter: string | null = null;
  let totalRows = 0;
  let batchIndex = 0;

  // Strip UTF-8 BOM if present at byte 0
  let isFirstChunk = true;

  while (offset < totalBytes) {
    const end = Math.min(offset + chunkSizeBytes, totalBytes);
    const slice = file.slice(offset, end);
    const arrayBuf = await slice.arrayBuffer();
    const isLastChunk = end >= totalBytes;

    let text = leftover + decoder.decode(arrayBuf, { stream: !isLastChunk });
    if (isFirstChunk) {
      if (text.charCodeAt(0) === 0xfeff) {
        text = text.slice(1);
      }
      isFirstChunk = false;
    }

    const rawLines: string[] = [];
    let startIdx = 0;
    let inQuotes = false;

    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (ch === '"') {
        inQuotes = !inQuotes;
      } else if (!inQuotes && (ch === '\n' || ch === '\r')) {
        if (ch === '\r' && i + 1 < text.length && text[i + 1] === '\n') {
          rawLines.push(text.slice(startIdx, i));
          i++;
          startIdx = i + 1;
        } else {
          rawLines.push(text.slice(startIdx, i));
          startIdx = i + 1;
        }
      }
    }

    if (isLastChunk) {
      const tail = text.slice(startIdx);
      if (tail.trim().length > 0) {
        rawLines.push(tail);
      }
      leftover = '';
    } else {
      leftover = text.slice(startIdx);
    }

    if (rawLines.length > 0) {
      if (!delimiter) {
        delimiter = detectCsvDelimiter(rawLines.slice(0, 10));
      }

      const parsedBatch: string[][] = [];
      for (let i = 0; i < rawLines.length; i++) {
        const line = rawLines[i];
        if (!line || !line.trim()) continue;
        parsedBatch.push(parseCsvLine(line, delimiter));
      }

      if (parsedBatch.length > 0) {
        totalRows += parsedBatch.length;
        onBatch(parsedBatch, batchIndex++);
      }
    }

    offset = end;

    if (onProgress) {
      onProgress({
        percent: Math.min(95, Math.round((offset / Math.max(totalBytes, 1)) * 90)),
        loadedBytes: offset,
        totalBytes,
        rowsParsed: totalRows,
        stage: 'Streaming file chunks...',
      });
    }

    await yieldToMainThread();
  }

  return totalRows;
}
