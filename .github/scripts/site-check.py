#!/usr/bin/env python3
"""Static site checks: every local link and asset in the HTML exists, and every JSON file parses.
Usage: site-check.py <site root>"""
import html.parser, json, pathlib, re, sys, urllib.parse

root = pathlib.Path(sys.argv[1]).resolve()
problems = []


class Links(html.parser.HTMLParser):
    def __init__(self):
        super().__init__(); self.refs = []
    def handle_starttag(self, tag, attrs):
        for k, v in attrs:
            if k in ('href', 'src') and v: self.refs.append(v)


for page in root.rglob('*.html'):
    if '.git' in page.parts or 'node_modules' in page.parts: continue
    p = Links(); p.feed(page.read_text(encoding='utf-8', errors='replace'))
    for ref in p.refs:
        if re.match(r'^(https?:|mailto:|tel:|data:|javascript:|#|//|\$\{|\{\{|blob:)', ref) or '${' in ref: continue
        path = urllib.parse.unquote(urllib.parse.urlsplit(ref).path)
        if not path: continue
        target = (root / path.lstrip('/')) if path.startswith('/') else (page.parent / path)
        if target.is_dir(): target = target / 'index.html'
        if not target.exists():
            problems.append(f'{page.relative_to(root)}: missing {ref}')

for f in root.rglob('*.json'):
    if '.git' in f.parts or 'node_modules' in f.parts: continue
    try: json.loads(f.read_text(encoding='utf-8'))
    except Exception as e: problems.append(f'{f.relative_to(root)}: invalid JSON ({e})')

for p in problems: print(f'::error::{p}')
print(f'{len(problems)} problem(s) in {root}')
sys.exit(1 if problems else 0)
