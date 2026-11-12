#!/usr/bin/env python3
"""Build index.html (the Microsoft 365 hub) from the Claude artifact version of the hub.

The hub is developed as a Claude artifact. To publish a new version here:
  1. Save the artifact's HTML as src/hub-artifact.html
  2. python3 tools/build.py
  3. Commit src/hub-artifact.html and index.html

The changes below are small and targeted; the build stops if any of them no longer applies,
so a change in the artifact can't silently break the Microsoft 365 version.
"""
import pathlib, re, sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
src = (ROOT / 'src' / 'hub-artifact.html').read_text(encoding='utf-8')
s = src


def rep(old, new, count=1):
    global s
    n = s.count(old)
    if n != count:
        sys.exit(f'build: expected {count} of {old[:70]!r}, found {n}')
    s = s.replace(old, new)


# 1. Load MSAL, the tenant config and the Microsoft 365 runtime before the hub script.
MAIN = '\n<script>\n/* ---------- utils'
if s.count(MAIN) != 1:
    # find the main inline script (the last <script> without src)
    i = s.rfind('<script>')
    if i < 0:
        sys.exit('build: hub script not found')
    s = s[:i] + '<!--M365-->' + s[i:]
    s = s.replace('<!--M365-->', '<script src="vendor/msal-browser-3.28.1.min.js"></script>\n<script src="hub-config.js"></script>\n<script src="m365.js"></script>\n', 1)
else:
    rep(MAIN, '\n<script src="vendor/msal-browser-3.28.1.min.js"></script>\n<script src="hub-config.js"></script>\n<script src="m365.js"></script>' + MAIN)

# 2. Stored documents live in SharePoint: fetch them through Graph.
rep('await fetch(fileUrl(f))', 'await (f.sp && window.fetchHubFile ? fetchHubFile(f) : fetch(fileUrl(f)))', 2)

# 3. Uploaded documents keep their SharePoint location.
rep("JSON.stringify({id: r.id, name: file.name, size: r.sizeBytes, type: r.contentType})",
    "JSON.stringify({id: r.id, name: file.name, size: r.sizeBytes, type: r.contentType, ...(r.sp ? {sp: true, driveId: r.driveId, itemId: r.itemId, webUrl: r.webUrl} : {})})")

# 4. The Exchange rule adds the signature to every email sent from Microsoft 365,
#    so the hub must not add a second one to replies and new emails.
rep("function withSignature(text, on){\n  if (!on) return {body: text};",
    "function withSignature(text, on){\n  if (!on || window.HUB_M365) return {body: text};")
rep("${signatureHtml()}</div></div>`;", "${window.HUB_M365 ? '' : signatureHtml()}</div></div>`;")
rep("${s.signoff ? s.signoff + '\\n' : ''}${signatureText()}`;", "${s.signoff ? s.signoff + '\\n' : ''}${window.HUB_M365 ? '' : signatureText()}`;")

# 5. Mail is Outlook here.
rep("bytes > 700 * 1024", "bytes > 3 * 1024 * 1024")
rep("Gmail via the hub takes about 700 KB per email", "Email from the hub takes about 3 MB of attachments")
rep("{k:'gmailUrl',l:'Gmail link',t:'url'}", "{k:'gmailUrl',l:'Email link',t:'url'}")
rep("lnk(item.gmailUrl,'Open in Gmail')", "lnk(item.gmailUrl,'Open email')")
rep("'Attachments are in Gmail: '", "'Attachments are on the email: '")
s = re.sub(r"Gmail's default Send mail as address", 'Outlook mailbox', s)
s = s.replace('Save as Gmail draft', 'Save as Outlook draft').replace('Gmail drafts', 'Outlook drafts')
s = s.replace("'Gmail inbox'", "'Inbox'").replace('Your business mail through Gmail.', 'Your business mail in Outlook.')
s = s.replace(" in Gmail yet.'", " in Outlook yet.'").replace('>Open in Gmail</a>', '>Open in Outlook</a>')
s = s.replace('Check your Sent folder in Gmail before trying again.', 'Check your Sent Items in Outlook before trying again.')
s = s.replace("'Gmail didn’t", "'Outlook didn’t").replace("? 'Gmail didn’t", "? 'Outlook didn’t")
s = s.replace("open in Gmail to view", "open in Outlook to view")
# forwarding steps don't apply: this mailbox is the business mailbox
rep("function setupNote(){", "function setupNote(){\n  if (window.HUB_M365) return '';")

# 5b. The document viewer needs the SharePoint location, not just the old file id.
rep("data-viewfile=\"${esc(JSON.stringify({id: f.id, name: f.name || '', type: f.type || '', size: f.size || 0}))}\"",
    "data-viewfile=\"${esc(JSON.stringify({id: f.id, name: f.name || '', type: f.type || '', size: f.size || 0, ...(f.sp ? {sp: true, driveId: f.driveId, itemId: f.itemId} : {})}))}\"")
rep('placeholder="Write your reply, or jot notes and let Claude draft it"', 'placeholder="Write your reply"')

# 6. Sign-in screen and signed-in footer styles.
rep('</style>', '''#m365gate{position:fixed;inset:0;z-index:9999;display:grid;place-items:center;background:var(--bg,#f4f4f4);padding:16px}
#m365gate .m365card{max-width:420px;width:100%;background:var(--panel,#fff);border:1px solid var(--line,#ddd);border-radius:10px;padding:28px;text-align:center}
#m365gate h1{font-size:1.4rem;margin:0 0 8px}#m365gate p{margin:0 0 20px;color:var(--muted,#666)}
#m365who{margin:12px 0 0;word-break:break-all}
html.m365 .reply-sig,html.m365 .sigline,html.m365 .sigprev{display:none}
</style>''')

(ROOT / 'index.html').write_text(s, encoding='utf-8')
left = [m.start() for m in re.finditer('Gmail', s)]
print(f'index.html written ({len(s):,} bytes). Remaining "Gmail" mentions: {len(left)}')
