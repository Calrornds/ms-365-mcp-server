import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, describe, expect, it } from 'vitest';
import {
  buildMcpServerInstructions,
  EXTRA_INSTRUCTIONS_MAX_BYTES,
  loadExtraInstructions,
} from '../src/mcp-instructions.js';

const baseCtx = { orgMode: true, readOnly: false, multiAccount: false };

describe('buildMcpServerInstructions', () => {
  it('includes general Graph guidance for standard mode', () => {
    const s = buildMcpServerInstructions({ ...baseCtx, discovery: false });
    expect(s).toContain('Microsoft Graph');
    expect(s).toContain('$filter');
    expect(s).not.toContain('DISCOVERY MODE ADD-ON');
  });

  it('appends discovery addon when discovery is true', () => {
    const s = buildMcpServerInstructions({ ...baseCtx, discovery: true });
    expect(s).toContain('DISCOVERY MODE ADD-ON');
    expect(s).toContain('search-tools');
    expect(s).toContain('$filter');
  });

  it('adds read-only line when readOnly', () => {
    const s = buildMcpServerInstructions({ ...baseCtx, discovery: false, readOnly: true });
    expect(s).toContain('read-only');
  });

  it('does not suggest account switching when multiAccount is false', () => {
    const s = buildMcpServerInstructions({ ...baseCtx, discovery: false, multiAccount: false });
    expect(s).not.toContain('Multiple accounts');
    expect(s).not.toContain('account parameter');
  });

  it('routes drive file downloads to get-download-url and authenticated byte reads to download-bytes', () => {
    const s = buildMcpServerInstructions({ ...baseCtx, discovery: false });
    expect(s).toContain('large drive/SharePoint file content');
    expect(s).toContain('prefer get-download-url');
    expect(s).toContain('download-bytes for authenticated byte reads');
    expect(s).toContain(
      'mail attachments, profile photos, Teams hosted content, and meeting recordings'
    );
    expect(s).toContain('relative Microsoft Graph paths, not absolute URLs');
  });

  it('describes the OneDrive upload path, rename body, and uploadUrl PUT', () => {
    const s = buildMcpServerInstructions({ ...baseCtx, discovery: false });
    expect(s).toContain('root:/folder/name.ext:');
    expect(s).toContain('create-upload-session');
    expect(s).toContain('@microsoft.graph.conflictBehavior');
    expect(s).toContain('rename');
    expect(s).toContain('Content-Range');
    expect(s).toContain('no Authorization');
    expect(s).toContain('root:/name.ext:');
    expect(s).not.toContain('/items/root:/path/to/file.txt:/content');
  });
});

describe('loadExtraInstructions', () => {
  const dirs: string[] = [];

  afterEach(() => {
    for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true });
  });

  function tempDir(): string {
    const dir = mkdtempSync(join(tmpdir(), 'ms365-extra-'));
    dirs.push(dir);
    return dir;
  }

  it('appends a set file after a blank line, including discovery mode', () => {
    const file = join(tempDir(), 'extra.md');
    writeFileSync(file, 'Custom upload rule.\n', 'utf8');
    const warnings: string[] = [];
    const extra = loadExtraInstructions(file, { warn: (message) => warnings.push(message) });

    expect(warnings).toEqual([]);
    const base = buildMcpServerInstructions({ ...baseCtx, discovery: false });
    expect(
      buildMcpServerInstructions({ ...baseCtx, discovery: false, extraInstructions: extra })
    ).toBe(`${base}\n\nCustom upload rule.`);

    const discovery = buildMcpServerInstructions({
      ...baseCtx,
      discovery: true,
      extraInstructions: extra,
    });
    expect(discovery.indexOf('Custom upload rule.')).toBeGreaterThan(
      discovery.indexOf('DISCOVERY MODE ADD-ON')
    );
    expect(discovery.endsWith('\n\nCustom upload rule.')).toBe(true);
  });

  it('warns and keeps default instructions when the file is missing', () => {
    const missing = join(tempDir(), 'missing.txt');
    const warnings: string[] = [];
    const extra = loadExtraInstructions(missing, { warn: (message) => warnings.push(message) });

    expect(extra).toBe('');
    expect(warnings).toHaveLength(1);
    expect(warnings[0]).toContain(missing);
    expect(warnings[0]).toContain('Continuing with default instructions');
    expect(
      buildMcpServerInstructions({ ...baseCtx, discovery: false, extraInstructions: extra })
    ).toBe(buildMcpServerInstructions({ ...baseCtx, discovery: false }));
  });

  it('truncates an oversize file and warns', () => {
    const file = join(tempDir(), 'big.txt');
    writeFileSync(file, 'x'.repeat(EXTRA_INSTRUCTIONS_MAX_BYTES + 20), 'utf8');
    const warnings: string[] = [];
    const extra = loadExtraInstructions(file, { warn: (message) => warnings.push(message) });

    expect(Buffer.byteLength(extra, 'utf8')).toBe(EXTRA_INSTRUCTIONS_MAX_BYTES);
    expect(extra).toBe('x'.repeat(EXTRA_INSTRUCTIONS_MAX_BYTES));
    expect(warnings).toHaveLength(1);
    expect(warnings[0]).toContain(file);
    expect(warnings[0]).toMatch(/truncat/i);
    expect(warnings[0]).toContain(String(EXTRA_INSTRUCTIONS_MAX_BYTES));
  });

  it('does not split a multibyte character at the cap', () => {
    const file = join(tempDir(), 'utf8.txt');
    writeFileSync(file, 'áá', 'utf8');
    const warnings: string[] = [];
    const extra = loadExtraInstructions(file, {
      warn: (message) => warnings.push(message),
      maxBytes: 3,
    });

    expect(extra).toBe('á');
    expect(Buffer.byteLength(extra, 'utf8')).toBeLessThanOrEqual(3);
    expect(warnings[0]).toMatch(/truncat/i);
  });
});

describe('OneDrive upload llmTips', () => {
  const endpoints = JSON.parse(
    readFileSync(fileURLToPath(new URL('../src/endpoints.json', import.meta.url)), 'utf8')
  ) as Array<{ toolName: string; pathPattern: string; llmTip?: string }>;

  function tip(toolName: string): { pathPattern: string; llmTip: string } {
    const endpoint = endpoints.find((entry) => entry.toolName === toolName);
    if (!endpoint?.llmTip) throw new Error(`missing llmTip for ${toolName}`);
    return { pathPattern: endpoint.pathPattern, llmTip: endpoint.llmTip };
  }

  it('tells create-upload-session to pass root:/folder/name.ext: and PUT with Content-Range', () => {
    const { pathPattern, llmTip } = tip('create-upload-session');
    expect(pathPattern).toBe('/drives/{drive-id}/items/{driveItem-id}/createUploadSession');
    // driveItemId is one path segment. Substituting root:/folder/name.ext: yields a
    // Graph path-addressed URL; the runtime percent-encodes that segment (preserving '=').
    expect(llmTip).toContain('/drives/{drive-id}/items/root:/folder/name.ext:/createUploadSession');
    expect(llmTip).toContain('root:/folder/name.ext:');
    expect(llmTip).not.toContain('{parentId}');
    expect(llmTip).toContain('rename');
    expect(llmTip).toContain('adding name');
    expect(llmTip).toContain('Content-Range');
    expect(llmTip).toContain('bytes 0-(N-1)/N');
    expect(llmTip).toContain('no Authorization');
    expect(llmTip).toContain('327680');
    expect(llmTip).toContain('size');
  });

  it('tells upload-file-content not to include /content in driveItemId', () => {
    const { pathPattern, llmTip } = tip('upload-file-content');
    expect(pathPattern).toBe('/drives/{drive-id}/items/{driveItem-id}/content');
    expect(llmTip).toContain('/drives/{drive-id}/items/root:/folder/name.ext:/content');
    expect(llmTip).toContain('root:/folder/name.ext:');
    expect(llmTip).toContain('root:/name.ext:/content');
    expect(llmTip).not.toContain('/items/root:/path/to/file.txt:/content');
  });
});
