"""Loads the hub signed out: the Microsoft sign-in screen must appear and no script may fail."""
import sys, time
from playwright.sync_api import sync_playwright

url = sys.argv[1]
with sync_playwright() as p:
    b = p.chromium.launch(); page = b.new_page()
    errors = []
    page.on('pageerror', lambda e: errors.append(str(e)))
    for _ in range(20):
        try: page.goto(url); break
        except Exception: time.sleep(0.5)
    page.wait_for_selector('#m365gate #m365signin', timeout=20000)
    assert page.is_visible('#m365gateTitle'), 'title hidden'
    page.screenshot(path='hub-signin.png')
    b.close()
if errors:
    print('::error::Page errors: ' + ' | '.join(errors)); sys.exit(1)
print('hub sign-in screen ok')
