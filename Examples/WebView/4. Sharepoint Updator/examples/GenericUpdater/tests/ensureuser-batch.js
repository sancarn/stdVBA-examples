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
  const payloads = logins.map(login => JSON.stringify({ logonName: login.trim() }));
  const lines = [];
  for (const payload of payloads) {
    const changes = "changeset_" + crypto.randomUUID();
    lines.push("--" + batch, "Content-Type: multipart/mixed; boundary=" + changes, "", "--" + changes, "Content-Type: application/http", "Content-Transfer-Encoding: binary", "",
      "POST " + api + "/web/ensureuser HTTP/1.1", "Content-Type: " + json, "Accept: " + json, "", payload, "", "--" + changes + "--");
  }
  lines.push("--" + batch + "--", "");
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
    const control = await fetch(api + "/web/ensureuser", { method: "POST",
      headers: { Accept: json, "Content-Type": json, "X-RequestDigest": digest }, body: payload });
    const individualId = control.ok ? (await control.json()).d.Id : undefined;
    rows.push({ login: logins[index], status, batchId, individualStatus: control.status, individualId });
  }
  console.table(rows);
  const successful = rows.filter(row => row.status >= 200 && row.status < 300);
  if (parts.length !== 2 || !successful.length ||
      rows.some(row => row.status !== row.individualStatus) ||
      successful.some(row => !Number.isInteger(row.batchId) || row.batchId <= 0 || row.batchId !== row.individualId) ||
      (successful.length === 2 && successful[0].batchId === successful[1].batchId)) {
    throw Error("Inconclusive or failed. Check the raw responses and use distinct accounts, including one valid account.");
  }
  console.log("PASS: both batch statuses and all returned user IDs match the individual calls.");
})().catch(error => console.error(error));
