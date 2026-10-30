from pathlib import Path
from http.server import ThreadingHTTPServer,SimpleHTTPRequestHandler
from threading import Thread
import os
from playwright.sync_api import sync_playwright
ROOT=Path(__file__).parents[1]/'dist'
os.chdir(ROOT)
server=ThreadingHTTPServer(('127.0.0.1',0),SimpleHTTPRequestHandler)
Thread(target=server.serve_forever,daemon=True).start()
url=(ROOT/'index.html').as_uri()
print('APP',url, flush=True)
with sync_playwright() as p:
 browser=p.chromium.launch(headless=True,executable_path='/usr/bin/chromium',args=['--no-sandbox'])
 page=browser.new_page(viewport={'width':1440,'height':900},device_scale_factor=1)
 errors=[]
 page.on('pageerror',lambda ex: errors.append(str(ex)))
 page.goto(url,wait_until='domcontentloaded');page.wait_for_timeout(1200)
 print('INITIAL',page.evaluate('window.CRApp?.getState()'), 'heading',page.locator('h1').inner_text(),flush=True)
 assert page.evaluate('window.CRApp.getState().records') >= 0
 assert page.locator('.metric').count()==4
 assert page.locator('#syncState').inner_text().startswith('Snapshot')
 page.screenshot(path=str(ROOT.parent/'desktop_preview.png'),full_page=True)
 # Filter, sheet switching, recon scope.
 page.locator('#workstream').select_option('Cost CR');page.wait_for_timeout(100)
 assert page.evaluate('window.CRApp.getState().sheet')=='Cost CR'
 page.locator('#workstream').select_option('CR Ultimate')
 page.locator('#search').fill('Auto Cancel Order');page.wait_for_timeout(70)
 assert page.evaluate('window.CRApp.getState().filtered')>=1
 page.locator('tbody button.open-row').first.click();page.wait_for_timeout(150)
 assert page.locator('#detailDrawer').is_visible()
 print('DRAWER',page.locator('#drawerTitle').inner_text(),flush=True)
 original=page.locator('#editAction').input_value()
 page.locator('#editAction').fill(original+' | QA draft')
 page.locator('#saveEdit').click();page.wait_for_timeout(130)
 assert page.locator('#detailDrawer').is_hidden()
 assert page.evaluate('window.CRApp.getState().drafts')>=1
 assert page.locator('#syncState').inner_text().startswith('Snapshot')
 page.reload();page.wait_for_timeout(300)
 assert page.evaluate('window.CRApp.getState().drafts')>=1
 page.locator('#settingsButton').click();page.wait_for_timeout(70)
 assert page.locator('#settingsDialog').is_visible()
 assert 'belum terhubung' in page.locator('#connTitle').inner_text()
 page.keyboard.press('Escape');assert page.locator('#settingsDialog').is_hidden()
 # Small device visual; no horizontal viewport overflow.
 mobile=browser.new_page(viewport={'width':390,'height':844},device_scale_factor=1)
 mobile.on('pageerror',lambda ex: errors.append(str(ex)))
 mobile.goto(url,wait_until='domcontentloaded');mobile.wait_for_timeout(450)
 assert mobile.locator('#menuToggle').is_visible()
 mobile.screenshot(path=str(ROOT.parent/'mobile_preview.png'),full_page=True)
 width=mobile.evaluate('document.documentElement.scrollWidth'),mobile.evaluate('window.innerWidth')
 print('MOBILE WIDTH',width,flush=True)
 assert width[0] <= width[1]+1
 print('ERRORS',errors, flush=True)
 assert not errors
 browser.close()
server.shutdown()
print('UI PASS',flush=True)
