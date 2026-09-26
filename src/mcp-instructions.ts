import { readFileSync } from 'node:fs';
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
 * Cut `text` to at most `maxBytes` UTF-8 bytes without splitting a code point.
 */
function truncateUtf8(text: string, maxBytes: number): { text: string; truncated: boolean } {
  const buf = Buffer.from(text, 'utf8');
  if (buf.length <= maxBytes) return { text, truncated: false };
  let end = maxBytes;
  // buf[end] is the first excluded byte. Walk back over continuation bytes so
  // a multibyte character that straddles the cap is dropped whole.
  while (end > 0) {
    const byte = buf[end];
    if (byte === undefined || (byte & 0xc0) !== 0x80) break;
    end--;
  }
  return { text: buf.subarray(0, end).toString('utf8'), truncated: true };
}

/**
 * Read operator instructions from a UTF-8 file. Missing or unreadable files
 * warn and return '' so startup continues with the built-in text. Over-long
 * files are truncated to {@link EXTRA_INSTRUCTIONS_MAX_BYTES}.
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

  let text: string;
  try {
    const buffer = readFileSync(trimmedPath);
    text = buffer.toString('utf8');
  } catch (error) {
    const detail = error instanceof Error ? error.message : String(error);
    warn(
      `Could not read extra MCP instructions file "${trimmedPath}": ${detail}. Continuing with default instructions.`
    );
    return '';
  }

  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
  const capped = truncateUtf8(text, maxBytes);
  if (capped.truncated) {
    warn(
      `Extra MCP instructions file "${trimmedPath}" exceeds ${maxBytes} bytes; truncating to ${maxBytes} bytes.`
    );
  }
  return capped.text;
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
    'Files / binary content: for large drive/SharePoint file content, prefer get-download-url to resolve a pre-authenticated URL for out-of-band download. Use download-bytes for authenticated byte reads such as mail attachments, profile photos, Teams hosted content, and meeting recordings. In stdio mode (or HTTP with --http-local-file-tools), download-bytes-to-file writes those same authenticated bytes straight to a local absolute path instead of returning base64 — the only out-of-band option for large mail attachments and meeting recordings, which get-download-url cannot handle. These tools take relative Microsoft Graph paths, not absolute URLs. For uploads, prefer create-upload-session with driveItemId root:/folder/name.ext: (trailing colon; the tool builds /drives/{drive-id}/items/root:/folder/name.ext:/createUploadSession), body { item: { "@microsoft.graph.conflictBehavior": "rename" } } only, then PUT the bytes to uploadUrl with Content-Range and no Authorization header and confirm the response size equals the byte count. upload-file-content is base64 in the tool argument (a truncated string is stored broken, and it overwrites); its driveItemId is root:/name.ext: so the tool builds /drives/{drive-id}/items/root:/name.ext:/content.',
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
