# OneDrive uploads

Verified 2026-09-26 (Word file, 43878 bytes, 201 Created, size matched).

Pass this file with `--extra-instructions-file` or `MS365_MCP_EXTRA_INSTRUCTIONS_FILE` (absolute path).

## Rule

- Do not use `upload-file-content` for real files. The base64 travels inside the tool argument and, if it is truncated, a broken file is stored with no error. It also overwrites without warning. Graph allows up to 250 MB on that PUT. If you still call it, pass `driveItemId` unencoded (for example `root:/folder/my file.docx:`). The server percent-encodes that one segment and PUTs `/drives/{drive-id}/items/root%3A%2Ffolder%2Fmy%20file.docx%3A/content`. Do not percent-encode the value yourself (a space would become `%2520`). Do not add `/content`; the tool appends it.
- Always use an upload session, then PUT the bytes from your sandbox.

## Steps

1. Call `create-upload-session`. `driveId` comes from `list-drives` (your drive or a SharePoint site drive). Pass `driveItemId` unencoded as `root:/folder/my file.docx:` (colon at the end, spaces as written, no `/content` and no `/createUploadSession`). The server encodes it and POSTs `/drives/{drive-id}/items/root%3A%2Ffolder%2Fmy%20file.docx%3A/createUploadSession`. Do not percent-encode `driveItemId` yourself.
   Body: `{ "item": { "@microsoft.graph.conflictBehavior": "rename" } }`.
   Omit `name`. If you send `name`, it must equal the file name in the path.
2. Save `uploadUrl`. Do not show it in the chat (it is already an authorized URL).
3. From the sandbox: `PUT` the file bytes to `uploadUrl` with `Content-Length: <N>` and `Content-Range: bytes 0-<N-1>/<N>`. No `Authorization` header.
   Each request must be under 60 MiB. Larger files: chunks in multiples of 320 KiB (327680 bytes); the last chunk may be shorter. Each chunk has its own `Content-Range`. An intermediate chunk returns 202.
4. Success is the final 201 or 200 with `size` equal to N. If `size` does not match, say so; do not treat the upload as complete.
5. Confirm to the user with the final name (it may change because of `rename`) and the `webUrl` link.
