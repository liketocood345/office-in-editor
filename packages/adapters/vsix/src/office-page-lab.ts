import { randomBytes } from "node:crypto";
import { realpath } from "node:fs/promises";
import { createServer } from "node:http";
import * as path from "node:path";

export interface OfficeLabPage {
  url: string;
  origin: string;
  port: number;
  requested: Promise<boolean>;
  close(): Promise<void>;
}

/** The isolated SSH experiment hands a single preselected fixture to the IDE. */
export async function openOfficeLabPage(document: string): Promise<OfficeLabPage> {
  const root = process.env.OIE_NATIVE_LAB_ROOT ?? "";
  if (!/^editor-slot-[a-f0-9]{32}$/.test(path.basename(root)) ||
      !/^synthetic\.(docx|xlsx|pptx)$/.test(path.basename(document)) ||
      path.dirname(document).toLowerCase() !== root.toLowerCase() ||
      (await realpath(document)).toLowerCase() !== document.toLowerCase()) {
    throw new Error("unexpected_lab_fixture");
  }
  const requestPath = "/" + randomBytes(24).toString("hex");
  const nonce = randomBytes(18).toString("base64");
  let authority = "";
  let origin = "";
  let markRequested!: (value: boolean) => void;
  const requested = new Promise<boolean>(resolve => { markRequested = resolve; });
  const server = createServer({ requestTimeout: 3000, headersTimeout: 3000 }, (request, response) => {
    response.setHeader("Cache-Control", "no-store");
    response.setHeader("Referrer-Policy", "no-referrer");
    response.setHeader("X-Content-Type-Options", "nosniff");
    if (request.headers.host !== authority) { response.writeHead(403).end(); return; }
    if (request.method === "GET" && request.url === requestPath) {
      response.writeHead(200, {
        "Content-Type": "text/html; charset=utf-8",
        "Content-Security-Policy": `default-src 'none'; script-src 'nonce-${nonce}'; connect-src 'self'; base-uri 'none'; form-action 'none'`,
      }).end(`<!doctype html><html lang="zh-CN"><meta charset="utf-8"><title>本机 Office</title>
      <main>正在连接本机 Office</main><script nonce="${nonce}">
      fetch(location.pathname + '/open', {method:'POST'}).catch(() => {});
      </script></html>`);
      return;
    }
    if (request.method === "POST" && request.url === requestPath + "/open" && request.headers.origin === origin) {
      response.writeHead(204).end();
      markRequested(true);
      return;
    }
    response.writeHead(404).end();
  });
  server.maxConnections = 8;
  await new Promise<void>((resolve, reject) => {
    server.once("error", reject);
    server.listen(0, "127.0.0.1", resolve);
  });
  const address = server.address();
  if (!address || typeof address === "string") { server.close(); throw new Error("missing_lab_listener"); }
  const { port } = address;
  authority = "127.0.0.1:" + port;
  origin = "http://" + authority;
  return { url: origin + requestPath, origin, port, requested, close: () => {
    markRequested(false);
    return new Promise<void>(resolve => {
      server.close(() => resolve());
      server.closeAllConnections();
    });
  } };
}
