import { closeSync, constants, openSync, readSync, statSync } from 'node:fs';
import path from 'node:path';
import logger from './logger.js';

/** Shared context for MCP `initialize.instructions` (hosts that forward it to the model). */
export type McpInstructionsContext = {
  orgMode: boolean;
  readOnly: boolean;
  multiAccount: boolean;
};

/** UTF-8 byte cap for operator-supplied instructions appended at startup. */
export const EXTRA_INSTRUCTIONS_MAX_BYTES = 32 * 1024;

function defaultExtraInstructionsWarn(message: string): void {
  logger.warn(message);
  console.error(message);
}

/**
 * Drop a trailing UTF-8 sequence that was cut off when `buf` is only a prefix
 * of a longer file. A complete prefix is returned unchanged.
 */
function trimIncompleteUtf8Prefix(buf: Buffer): Buffer {
  let i = buf.length - 1;
  let continuations = 0;
  while (i >= 0) {
    const byte = buf[i];
    if (byte === undefined || (byte & 0xc0) !== 0x80) break;
    continuations++;
    i--;
  }
  if (i < 0) return Buffer.alloc(0);
  const lead = buf[i];
  if (lead === undefined) return Buffer.alloc(0);
  const width = lead < 0x80 ? 1 : lead < 0xe0 ? 2 : lead < 0xf0 ? 3 : 4;
  if (1 + continuations < width) return buf.subarray(0, i);
  return buf;
}

/**
 * Read at most `maxBytes` from a regular file. Callers must already have
 * rejected non-files: opening a FIFO would block, and reading /dev/zero would not end.
 */
function readCappedUtf8(filePath: string, maxBytes: number): { text: string; truncated: boolean } {
  const info = statSync(filePath);
  if (!info.isFile()) {
    throw new Error('not a regular file');
  }
  if (info.size === 0) return { text: '', truncated: false };

  const toRead = Math.min(info.size, maxBytes);
  const buf = Buffer.alloc(toRead);
  const fd = openSync(filePath, constants.O_RDONLY | constants.O_NONBLOCK);
  try {
    let offset = 0;
    while (offset < toRead) {
      const n = readSync(fd, buf, offset, toRead - offset, offset);
      if (n <= 0) break;
      offset += n;
    }
    const bytes =
      info.size > maxBytes
        ? trimIncompleteUtf8Prefix(buf.subarray(0, offset))
        : buf.subarray(0, offset);
    return { text: bytes.toString('utf8'), truncated: info.size > maxBytes };
  } finally {
    closeSync(fd);
  }
}

/**
 * Read operator instructions from a UTF-8 file. Missing, unreadable, empty,
 * relative, or non-regular paths warn and return '' so startup continues with
 * the built-in text. Over-long files are truncated to {@link EXTRA_INSTRUCTIONS_MAX_BYTES}.
 */
export function loadExtraInstructions(
  filePath: string | undefined,
  options?: {
    warn?: (message: string) => void;
    maxBytes?: number;
  }
): string {
  if (typeof filePath !== 'string') return '';
  const trimmedPath = filePath.trim();
  if (!trimmedPath) return '';

  const warn = options?.warn ?? defaultExtraInstructionsWarn;
  const maxBytes = options?.maxBytes ?? EXTRA_INSTRUCTIONS_MAX_BYTES;

  if (!path.isAbsolute(trimmedPath)) {
    warn(
      `Extra MCP instructions path "${trimmedPath}" is not absolute. Pass an absolute path. Continuing with default instructions.`
    );
    return '';
  }

  let text: string;
  let truncated = false;
  try {
    const info = statSync(trimmedPath);
    if (!info.isFile()) {
      warn(
        `Extra MCP instructions path "${trimmedPath}" is not a regular file. Continuing with default instructions.`
      );
      return '';
    }
    const capped = readCappedUtf8(trimmedPath, maxBytes);
    text = capped.text;
    truncated = capped.truncated;
  } catch (error) {
    const detail = error instanceof Error ? error.message : String(error);
    warn(
      `Could not read extra MCP instructions file "${trimmedPath}": ${detail}. Continuing with default instructions.`
    );
    return '';
  }

  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
  if (truncated) {
    warn(
      `Extra MCP instructions file "${trimmedPath}" exceeds ${maxBytes} bytes; truncating to ${maxBytes} bytes.`
    );
  }
  if (text.trim() === '') {
    warn(
      `Extra MCP instructions file "${trimmedPath}" is empty. Continuing with default instructions.`
    );
    return '';
  }
  return text;
}

function appendExtraInstructions(base: string, extra: string | undefined): string {
  const trimmed = extra?.trim() ?? '';
  if (!trimmed) return base;
  return `${base}\n\n${trimmed}`;
}

function buildGeneralMcpInstructions(opts: McpInstructionsContext): string {
  const parts = [
    'Microsoft 365 MCP exposes Microsoft Graph through MCP tools. Use each tool name, description, and parameter schema as the source of truth.',
    'Microsoft Graph OData: do not combine $filter with $search on the same request. For lists, prefer modest $top (or top) and $select; avoid very large pages unless the user needs them.',
    'Mail and message $search uses KQL; the $search query parameter value must be double-quoted per Graph (see search-query-parameter in Microsoft Graph docs).',
    'When you need an organizational user or recipient address, resolve it with list-users (or another directory tool); do not invent SMTP addresses.',
    'Directory $search on collections such as /users or /groups requires ConsistencyLevel: eventual when the tool exposes that header.',
    'Teams chat and channel messages: prefer HTML contentType in the body; plain text is often mangled by Graph.',
    'Files / binary content: for large drive/SharePoint file content, prefer get-download-url to resolve a pre-authenticated URL for out-of-band download. Use download-bytes for authenticated byte reads such as mail attachments, profile photos, Teams hosted content, and meeting recordings. In stdio mode (or HTTP with --http-local-file-tools), download-bytes-to-file writes those same authenticated bytes straight to a local absolute path instead of returning base64 — the only out-of-band option for large mail attachments and meeting recordings, which get-download-url cannot handle. These tools take relative Microsoft Graph paths, not absolute URLs. For uploads, prefer create-upload-session. Pass driveItemId unencoded, such as root:/folder/my file.docx: (trailing colon; do not percent-encode). The server encodes that segment to /drives/{drive-id}/items/root%3A%2Ffolder%2Fmy%20file.docx%3A/createUploadSession. driveId comes from list-drives (your drive or a SharePoint site drive). Body { item: { "@microsoft.graph.conflictBehavior": "rename" } }; omit name, or if sent it must equal the file name in the path. PUT bytes to uploadUrl with Content-Range and no Authorization header; each request is under 60 MiB, intermediate chunks return 202, and the final size must match. upload-file-content is a base64 PUT of at most 250 MB that overwrites; pass driveItemId the same unencoded way and do not add /content.',
  ];
  if (opts.readOnly) parts.push('This server is read-only; write operations are disabled.');
  if (opts.multiAccount)
    parts.push('Multiple accounts: pass the account parameter when required (see list-accounts).');
  if (!opts.orgMode)
    parts.push('Work/school-only tools require starting the server with --org-mode.');
  return parts.join(' ');
}

const DISCOVERY_MODE_INSTRUCTIONS_ADDON =
  'DISCOVERY MODE ADD-ON: Graph is reached via search-tools → get-tool-schema → execute-tool (plus auth helpers). ' +
  'Workflow: (1) call search-tools with short natural-language keywords (BM25-ranked); ' +
  '(2) call get-tool-schema(tool_name) to see the parameters, required fields, and enum values; ' +
  '(3) call execute-tool with tool_name exactly as returned and parameters shaped per the schema. ' +
  'Skipping get-tool-schema is the leading cause of Graph 400 errors here. ' +
  'If search-tools returns no matches, retry with shorter or different keywords.';

/**
 * Full MCP `initialize.instructions` string: general guidance for every mode, plus a discovery-only suffix when applicable.
 */
export function buildMcpServerInstructions(
  opts: McpInstructionsContext & { discovery: boolean; extraInstructions?: string }
): string {
  const general = buildGeneralMcpInstructions(opts);
  const base = opts.discovery ? `${general} ${DISCOVERY_MODE_INSTRUCTIONS_ADDON}` : general;
  return appendExtraInstructions(base, opts.extraInstructions);
}
