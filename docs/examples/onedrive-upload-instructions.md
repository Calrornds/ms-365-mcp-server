# OneDrive uploads

Verified 2026-09-26 (Word file, 43878 bytes, 201 Created, size matched).

Pass this file with `--extra-instructions-file` or `MS365_MCP_EXTRA_INSTRUCTIONS_FILE` (absolute path).

## Rule

- Do not use `upload-file-content` for real files. The base64 travels inside the tool argument and, if it is truncated, a broken file is stored with no error. It also overwrites without warning. If you still call it, pass `driveItemId` as `root:/name.ext:` so the tool builds `PUT /drives/{drive-id}/items/root:/name.ext:/content`. Passing `root:/name.ext:/content` returns 400 because the tool already appends `/content`.
- Always use an upload session, then PUT the bytes from your sandbox.

## Steps

1. Call `create-upload-session`. `driveId` comes from `list-drives`. Pass `driveItemId` as `root:/folder/name.ext:` (colon at the end, no `/content` and no `/createUploadSession`; the tool appends `/createUploadSession`). That builds `POST /drives/{drive-id}/items/root:/folder/name.ext:/createUploadSession`.
   Body: `{ "item": { "@microsoft.graph.conflictBehavior": "rename" } }`.
   Do not send `name` together with `rename` in the body: that returns 400.
2. Save `uploadUrl`. Do not show it in the chat (it is already an authorized URL).
3. From the sandbox: `PUT` the file bytes to `uploadUrl` with `Content-Length: <N>` and `Content-Range: bytes 0-<N-1>/<N>`. No `Authorization` header.
   Up to about 60 MB in a single PUT. Larger files: chunks in multiples of 320 KiB (327680 bytes), each with its own `Content-Range`.
4. Success is 201 or 200 with `size` equal to N. If `size` does not match, say so; do not treat the upload as complete.
5. Confirm to the user with the final name (it may change because of `rename`) and the `webUrl` link.
