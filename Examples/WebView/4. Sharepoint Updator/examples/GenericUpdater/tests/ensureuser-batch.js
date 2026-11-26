// Run in the developer console of a signed-in page on the target SharePoint site.
(async () => {
  const siteInput = prompt("SharePoint site URL", window._spPageContextInfo?.webAbsoluteUrl || location.origin);
  if (!siteInput) return;
  const site = new URL(siteInput);
  if (site.origin !== location.origin || site.search || site.hash) throw Error("Enter a site URL on this origin.");
  const api = site.href.replace(/\/$/, "") + "/_api";
  const logins = [prompt("First test user's email or login"), prompt("Second test user's email or login")];
  if (logins.some(value => !value?.trim())) return;
  const json = "application/json;odata=verbose";
  async function request(path, options) {
    const response = await fetch(api + path, options);
    const text = await response.text();
    if (!response.ok) throw Error(path + " HTTP " + response.status + " " + text.slice(0, 500));
    return text;
  }
  const context = JSON.parse(await request("/contextinfo", { method: "POST", headers: { Accept: json } }));
  const digest = context.d.GetContextWebInformation.FormDigestValue;
  const batch = "batch_" + crypto.randomUUID();
  const changes = "changeset_" + crypto.randomUUID();
  const payloads = logins.map(login => JSON.stringify({ logonName: login.trim() }));
  const lines = ["--" + batch, "Content-Type: multipart/mixed; boundary=" + changes, ""];
  for (const payload of payloads) {
    lines.push("--" + changes, "Content-Type: application/http", "Content-Transfer-Encoding: binary", "",
      "POST " + api + "/web/ensureuser HTTP/1.1", "Content-Type: " + json, "Accept: " + json, "", payload, "");
  }
  lines.push("--" + changes + "--", "--" + batch + "--", "");
  const body = await request("/$batch", { method: "POST",
    headers: { Accept: json, "X-RequestDigest": digest, "Content-Type": "multipart/mixed; boundary=" + batch },
    body: lines.join("\r\n") });
  console.log("Raw batch response", body);
  const parts = [...body.replace(/\r\n/g, "\n").matchAll(/^HTTP\/\d\.\d (\d{3})[^\n]*\n([\s\S]*?)(?=\n--)/gm)];
  const rows = [];
  for (const [index, payload] of payloads.entries()) {
    const part = parts[index];
    const status = Number(part?.[1]);
    let batchId;
    if (status >= 200 && status < 300) {
      const message = part[0];
      const content = message.slice(message.indexOf("\n\n") + 2);
      batchId = JSON.parse(content).d.Id;
    }
    const control = JSON.parse(await request("/web/ensureuser", { method: "POST",
      headers: { Accept: json, "Content-Type": json, "X-RequestDigest": digest }, body: payload }));
    rows.push({ login: logins[index], status, batchId, individualId: control.d.Id });
  }
  console.table(rows);
  if (parts.length !== 2 || rows[0].individualId === rows[1].individualId ||
      rows.some(row => !Number.isInteger(row.batchId) || row.batchId <= 0 || row.batchId !== row.individualId)) {
    throw Error("Inconclusive or failed. Check the raw responses and use two distinct valid accounts.");
  }
  console.log("PASS: both batched ensureuser calls returned the same IDs as the individual calls.");
})().catch(error => console.error(error));
