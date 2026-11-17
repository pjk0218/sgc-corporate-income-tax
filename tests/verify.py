"""Check the web app against the Excel working, and check shared-folder saving and history.
Usage: python3 tests/verify.py <Corporate_Income_Tax_Calculation.xlsx> [screenshot_dir]
- Excel totals on the Normal Guarantee sheet carry manual rounding (+/-0.01..0.03) inside the SUM formulas.
  The app uses exact sums, so the expected values below remove that manual adjustment.
- The "shared folder" in this test is the browser's private file area (same folder API as a network folder)."""
import sys, pathlib, threading, functools, http.server, openpyxl
from playwright.sync_api import sync_playwright
xlsx=pathlib.Path(sys.argv[1]).resolve(); shots=pathlib.Path(sys.argv[2]) if len(sys.argv)>2 else None
root=pathlib.Path(__file__).resolve().parent.parent
wb=openpyxl.load_workbook(xlsx,data_only=True)
cit=[w for w in wb.worksheets if w.title.startswith('Corporate Income Tax')][0]; dt=wb['Deferred Tax']; ppe=wb['2.PPE- Depreciation']; g=wb['3.Normal Guarantee']
class H(http.server.SimpleHTTPRequestHandler):
    def log_message(self,*a): pass
srv=http.server.ThreadingHTTPServer(('127.0.0.1',0),functools.partial(H,directory=str(root))); threading.Thread(target=srv.serve_forever,daemon=True).start()
url=f'http://127.0.0.1:{srv.server_address[1]}/index.html'
# The test browser crashes when it re-opens a remembered private-area handle, so the handle is not remembered here
# and each page connects again. With a real network folder the app remembers the folder (same as the payroll app).
CONNECT="async()=>{ await window.__app.connect(await navigator.storage.getDirectory()); await new Promise((res,rej)=>{const r=indexedDB.open('sgc-cit-app',1); r.onsuccess=()=>{const q=r.result.transaction('kv','readwrite').objectStore('kv').delete('dirHandle'); q.onsuccess=res; q.onerror=rej;}; }); }"
fails=[]
def check(name,cond):
    print(('  OK   ' if cond else '  FAIL ')+name); 
    if not cond: fails.append(name)
with sync_playwright() as p:
    b=p.chromium.launch(); ctx=b.new_context(viewport={'width':1300,'height':900}); pg=ctx.new_page(); errs=[]
    pg.on('pageerror',lambda e:errs.append(str(e))); pg.on('console',lambda m:errs.append(m.text) if m.type=='error' else None)
    pg.goto(url); pg.wait_for_function('window.__ready')
    print('== Start =='); check('asks for the fiscal year first',pg.locator('#newYear').count()==1)
    pg.evaluate(CONNECT); pg.wait_for_function("window.__app.App.store.kind==='folder'")
    pg.fill('#newYear','2025'); pg.click('#btnStartYear'); pg.wait_for_selector('#fileIn',state='attached')
    check('6 missing-input notices before upload',pg.evaluate("window.__notices.filter(n=>n.lv==='err').length")==6)
    if shots: pg.screenshot(path=str(shots/'0_empty.png'),full_page=True)
    pg.set_input_files('#fileIn',str(xlsx)); pg.wait_for_function("window.__R && window.__R.tb && window.__R.guar")
    pg.fill('#k-fsFA',str(dt['B28'].value)); pg.press('#k-fsFA','Tab')
    pg.fill('#k-priorDT',str(dt['B21'].value)); pg.press('#k-priorDT','Tab')
    R=pg.evaluate("JSON.parse(JSON.stringify(window.__R))")
    gm=0.03   # manual rounding inside Excel's guarantee total (L16 = SUM(...) - 0.03), not applied by the app
    exp=[('Total income',R['tb']['income'],cit['F525'].value),('Total expenses',R['tb']['expense'],cit['F530'].value),('Net profit',R['tb']['netProfit'],cit['F532'].value),
     ('Guarantee disallowed',R['cit']['adj'][0]['amt'],cit['B535'].value+gm),('Entertainment',R['cit']['adj'][1]['amt'],cit['B536'].value),('Tax & dues',R['cit']['adj'][2]['amt'],cit['B537'].value),
     ('Provisions made',R['cit']['adj'][3]['amt'],cit['B538'].value),('Provisions utilized',R['cit']['adj'][4]['amt'],cit['B539'].value),('Depreciation difference',R['cit']['adj'][5]['amt'],cit['B540'].value),
     ('Total adjustments',R['cit']['totalAdj'],cit['B542'].value+gm),('Taxable profit',R['cit']['taxable'],cit['F537'].value+gm),('Income tax',R['cit']['tax'],cit['F538'].value+gm*0.2),
     ('Prepaid tax',R['cit']['prepaid'],cit['H536'].value),('Due to ZATCA',R['cit']['due'],cit['H537'].value+gm*0.2),
     ('PPE dep ZATCA',R['ppe']['tot']['dep'],ppe['I19'].value),('PPE carried fwd',R['ppe']['tot']['cf'],ppe['I21'].value),('Guarantee fee total',R['guar']['fee'],g['I16'].value+0.01),('Guarantee market total',R['guar']['mv'],g['K16'].value-0.02),
     ('DT tax-base FA',R['dt']['taxFA'],dt['B27'].value),('DT FA diff',R['dt']['diff'],dt['B29'].value),('DTL (A)',R['dt']['dtl'],dt['B12'].value),('DTA (B)',R['dt']['dta'],dt['B17'].value),
     ('Net DT (C)',R['dt']['net'],dt['B19'].value),('DT movement (E)',R['dt']['movement'],dt['B23'].value)]
    print(f'\n== App vs Excel (manual rounding removed) ==\n{"Line":28}{"App":>22}{"Excel":>22}{"Diff":>12}'); bad=0
    for n,a,e in exp:
        d=a-e; ok=abs(d)<0.005; bad+=not ok; print(f'{n:28}{a:22,.4f}{e:22,.4f}{d:12.4f} {"OK" if ok else "<<"}')
    check('all 24 lines match',bad==0)
    # manual adjustment key-in: entering Excel's -0.03 must reproduce the filed Excel figures exactly
    pg.click('#tab-inputs'); pg.select_option('#manTarget','A1'); pg.click('#btnAddMan')
    check('manual adjustment without a reason is refused',pg.evaluate("window.__app.state().manual.length")==0)
    pg.fill('#manDesc','Rounding in Excel total'); pg.fill('#manAmt','-0.03'); pg.click('#btnAddMan')
    M=pg.evaluate("JSON.parse(JSON.stringify(window.__R))")
    check('with manual -0.03: taxable profit, tax and due equal the Excel as filed',all(abs(x-y)<0.0005 for x,y in [(M['cit']['adj'][0]['amt'],cit['B535'].value),(M['cit']['taxable'],cit['F537'].value),(M['cit']['tax'],cit['F538'].value),(M['cit']['due'],cit['H537'].value)]))
    check('manual adjustment is listed in Notices',any('manual adjustment' in n['msg'] for n in pg.evaluate("window.__notices")))
    if shots: pg.screenshot(path=str(shots/'manual_inputs.png'),full_page=True); pg.click('#tab-cit'); pg.screenshot(path=str(shots/'manual_cit.png'),full_page=True); pg.click('#tab-inputs')
    pg.click('[data-rmman="0"]'); pg.wait_for_timeout(200)
    check('removing it restores the exact-sum result',abs(pg.evaluate("window.__R.cit.taxable")-R['cit']['taxable'])<0.0005)
    print(f'\nvs Excel as filed: income tax app {R["cit"]["tax"]:,.2f} / Excel {cit["F538"].value:,.2f}; due to ZATCA app {R["cit"]["due"]:,.2f} / Excel {cit["H537"].value:,.2f}')
    print('\nNotices:'); [print(' ',n['lv'].upper(),'-',n['msg'][:150]) for n in pg.evaluate("window.__notices")]
    if shots:
        for t in ['inputs','cit','dt','notices']:
            pg.click(f'#tab-{t}'); pg.wait_for_timeout(100); pg.screenshot(path=str(shots/f'{t}.png'),full_page=True)
    print('\n== Hover highlighting (F2-style) ==')
    pg.click('#tab-cit'); pg.hover('#main section:not([style*="none"]) tr[data-id="TP"] td.n')
    prec=pg.evaluate("[...document.querySelectorAll('.hl-prec')].map(e=>e.dataset.id).sort()")
    check('CIT: pointing at TP highlights NP and A: '+str(prec),prec==['A','NP'])
    check('CIT: formula bar under the tabs lists the precedent amounts','73,200,091.72' in pg.inner_text('#fbar') and 'NP + A' in pg.inner_text('#fbar'))
    if shots: pg.screenshot(path=str(shots/'hover_cit_tp.png'))
    pg.hover('#main section:not([style*="none"]) tr[data-id="A"] td.n')
    check('CIT: pointing at A highlights A1..A6',pg.evaluate("document.querySelectorAll('.hl-prec').length")==6)
    if shots: pg.screenshot(path=str(shots/'hover_cit_a.png'))
    pg.click('#tab-dt'); pg.hover('#main section:not([style*="none"]) tr[data-id="C"] td.n')
    prec=pg.evaluate("[...document.querySelectorAll('.hl-prec')].map(e=>e.dataset.id).sort()")
    check('DT: pointing at C highlights A, B and SH: '+str(prec),prec==['A','B','SH'])
    if shots: pg.screenshot(path=str(shots/'hover_dt_c.png'))
    pg.hover('#main section:not([style*="none"]) .kpi[data-hl="E"]')
    check('DT: summary box E highlights C and D',sorted(pg.evaluate("[...document.querySelectorAll('.hl-prec')].map(e=>e.dataset.id)"))==['C','D'])
    pg.mouse.move(5,5); check('moving away clears the highlight',pg.evaluate("document.querySelectorAll('.hl-prec,.hl-focus').length")==0)
    print('\n== Split-screen source panel ==')
    V='#main section:not([style*="none"]) '
    hlc=lambda: pg.evaluate("[...document.querySelectorAll('#side .srccell')].map(e=>e.textContent)")
    vis=lambda: pg.evaluate("(()=>{const c=document.querySelector('#side .srccell'); if(!c) return false; const r=c.getBoundingClientRect(), s=document.getElementById('sideBody').getBoundingClientRect(); return r.top>=s.top&&r.bottom<=s.bottom+1&&r.left>=s.left&&r.right<=s.right+1;})()")
    pg.click('#tab-dt'); pg.click(V+'tr[data-id="N1a"] td.n')
    check('DT N1a: PPE table opens beside the report with the carried-forward total marked: '+str(hlc()),pg.evaluate("document.body.classList.contains('split')") and hlc()==['2,434,805.41'] and 'PPE' in pg.inner_text('#side .smeta'))
    check('marked cell is scrolled into view',vis())
    if shots: pg.screenshot(path=str(shots/'side_dt_n1a.png'))
    pg.click(V+'tr[data-id="B1"] td.n'); check('DT B1: provision ending balance total marked',hlc()==['1,066,147.00'])
    pg.keyboard.press('Escape'); check('Esc closes the panel',not pg.evaluate("document.body.classList.contains('split')"))
    pg.click(V+'tr[data-id="N1e"] td.n'); check('DT N1e (key-in): jumps to the key-in field on Inputs',pg.evaluate("document.activeElement.id")=='k-fsFA')
    pg.click('#tab-cit'); pg.click(V+'tr[data-id="A2"] .srcbtn')
    check('CIT A2: both entertainment accounts marked in the trial balance: '+str(hlc()),len(hlc())==2 and vis())
    check('location text names the row and column',"CM Cost (Entertainment)" in pg.inner_text('#sideWhere') and 'Dr (YTD)' in pg.inner_text('#sideWhere'))
    pg.click('#sideNext'); check('Next button moves to the second marked cell',"G&A" in pg.inner_text('#sideWhere') or "Entertainment" in pg.inner_text('#sideWhere') and '20,948.67' in pg.inner_text('#sideWhere'))
    pg.click('#sidePrev')
    if shots: pg.screenshot(path=str(shots/'side_cit_a2.png'))
    pg.click(V+'tr[data-id="A6"] td.n'); check('CIT A6: depreciation difference, GL and ZATCA amounts marked',set(['130,154.65','896,587.35']).issubset(set(hlc())))
    pg.click(V+'tr[data-id="A1"] td.n'); check('CIT A1: guarantee disallowed total marked',hlc()==['7,737,531.59'] and vis())
    pg.click(V+'tr[data-id="P"] td.n'); check('CIT P: prepaid tax credits marked',hlc()==['2,947,146.24'])
    if shots: pg.screenshot(path=str(shots/'side_cit_p.png'))
    pg.hover(V+'tr[data-id="TP"] td.n'); check('hover highlighting still works with the panel open',pg.evaluate("document.querySelectorAll('.hl-prec').length")==2)
    print('\n== Hover on source cells ==')
    S='#sideBody '
    prec=lambda: sorted(pg.evaluate("[...document.querySelectorAll('#side .hl-prec')].map(e=>e.dataset.cid)"))
    tip=lambda: pg.inner_text('#sideF')
    pg.click(V+'tr[data-id="A1"] td.n'); pg.hover(S+'td[data-cid="guar:0:dis"]')
    check('Guarantee: disallowed = fee - market value: '+str(prec()),prec()==['guar:0:fee','guar:0:mv'] and 'Guarantee fee − Market value' in tip())
    check('explanation sits in the panel header bar and does not overlap the table',pg.evaluate("(()=>{const f=document.getElementById('sideF').getBoundingClientRect(),b=document.getElementById('sideBody').getBoundingClientRect();return f.bottom<=b.top+1&&f.height>60;})()"))
    pg.hover(S+'td[data-cid="guar:0:mv"]'); check('Guarantee: market value uses principal, days and market rate',prec()==['guar:0:days','guar:0:principal','guar:mr'])
    if shots: pg.screenshot(path=str(shots/'side_hover_guar.png'))
    pg.hover(S+'td[data-cid="guar:T:dis"]'); check('Guarantee: total = sum of all lines',len(prec())==pg.evaluate("window.__R.guar.items.length"))
    pg.click(V+'tr[data-id="A4"] td.n'); pg.hover(S+'td[data-cid="prov:0:end"]')
    check('Provision: ending = opening + provided - utilized + adjustments',len(prec())==4 and 'Opening + Provided' in tip())
    pg.click(V+'tr[data-id="A6"] td.n'); pg.hover(S+'td[data-cid="ppe:depDiff"]'); check('PPE: difference = GL - ZATCA',prec()==['ppe:depZ','ppe:gl'])
    pg.hover(S+'td[data-cid="ppe:dep:3"]'); check('PPE: depreciation = ending balance x rate',prec()==['ppe:end:3','ppe:rate:3'])
    pg.hover(S+'td[data-cid="ppe:cf:T"]'); check('PPE: carried-forward total = sum of groups',len(prec())==6)
    if shots: pg.screenshot(path=str(shots/'side_hover_ppe.png'))
    pg.click(V+'tr[data-id="E1"] td.n'); pg.hover(S+'td[data-cid="tb:002-004:ytdDr"]')
    n=pg.evaluate("document.querySelectorAll('#side .hl-prec').length"); tot=pg.evaluate("[...document.querySelectorAll('#side .hl-prec')].reduce((a,e)=>a+Number(e.textContent.replace(/,/g,'')),0)")
    check(f'Trial balance: Projects Cost Expenses = sum of {n} accounts below ({tot:,.2f})',n>5 and abs(tot-693000287.98)<0.01)
    if shots: pg.screenshot(path=str(shots/'side_hover_tb.png'))
    pg.click('#sideOnly'); check('"Show marked rows only" leaves just the source account',pg.locator('#sideBody tbody tr').count()==1)
    pg.click('#tab-data'); pg.click('[data-dtab=ppe]'); pg.hover('#main td[data-cid="ppe:end:3"]')
    check('Source data tab: same hover works there',sorted(pg.evaluate("[...document.querySelectorAll('#main .hl-prec')].map(e=>e.dataset.cid)"))==['ppe:halfAdd:3','ppe:halfCons:3','ppe:open:3'])
    pg.click('#tab-cit'); pg.click(V+'tr[data-id="P"] td.n')
    pg.click('#sideClose'); check('Close button closes the panel',pg.evaluate("document.getElementById('side').innerHTML")=='')
    print('\n== Shared folder and history ==')
    pg.evaluate("window.__app.App && new Promise(r=>setTimeout(r,400))")
    files=pg.evaluate("(async()=>{const r=await navigator.storage.getDirectory(); const d=await r.getDirectoryHandle('FY2025'); const o=[]; for await(const [n] of d.entries()) o.push(n); const s=await d.getDirectoryHandle('source'); for await(const [n] of s.entries()) o.push('source/'+n); return o.sort();})()")
    check('FY2025/data.json and source copy written to the folder: '+str(files),'data.json' in files and any(f.startswith('source/') for f in files))
    pg2=ctx.new_page(); pg2.goto(url); pg2.wait_for_function('window.__ready'); pg2.evaluate(CONNECT); pg2.wait_for_function('window.__R && window.__R.cit')     # a colleague opens the app
    check('colleague sees FY2025 with the same tax',abs(pg2.evaluate("window.__R.cit?window.__R.cit.tax:0")-R['cit']['tax'])<0.005)
    # colleague changes a key-in; first user's next change must not overwrite it silently
    pg2.click('#tab-inputs'); pg2.fill('#k-fsFA','3112612.40'); pg2.press('#k-fsFA','Tab'); pg2.wait_for_timeout(400)
    pg.click('#tab-inputs'); pg.fill('#k-priorDT','10931'); pg.press('#k-priorDT','Tab'); pg.wait_for_selector('#btnReload')
    check('conflict detected instead of overwriting',True); pg.click('#btnReload'); pg.wait_for_timeout(300)
    check('latest data loaded after conflict',pg.evaluate("window.__app.state().keyins.fsFA")==3112612.40)
    pg.fill('#k-fsFA',str(dt['B28'].value)); pg.press('#k-fsFA','Tab'); pg.wait_for_timeout(300)
    # wrong-year trial balance is refused in a new year; prior-year suggestion appears
    pg.select_option('#yearSel','new'); pg.fill('#newYear','2026'); pg.click('#btnStartYear'); pg.wait_for_selector('#fileIn',state='attached')
    pg.set_input_files('#fileIn',str(xlsx)); pg.wait_for_function("document.getElementById('toastMsg').textContent.includes('NOT loaded')")
    check('FY2025 trial balance refused in FY2026',pg.evaluate("window.__app.state().tb")is None)
    check('prior-year deferred tax suggestion offered',pg.locator('#btnUsePrior').count()==1)
    pg.click('#btnUsePrior'); check('suggestion fills key-in',abs(pg.evaluate("window.__app.state().keyins.priorDT")-round(R['dt']['net'],2))<0.001)
    pg.wait_for_timeout(300); pg.click('#tab-history'); pg.wait_for_timeout(100)
    check('history lists FY2026 and FY2025',pg.locator('#main table tbody tr').filter(has_text='FY20').count()>=2)
    if shots: pg.screenshot(path=str(shots/'history.png'),full_page=True)
    pg.click('[data-open="2025"]'); pg.wait_for_function("window.__app.App.year===2025")
    pg.click('#tab-inputs'); pg.click('#btnFinal'); pg.click('#modalOk'); pg.wait_for_timeout(300)
    check('final year is read-only',pg.locator('#fileIn').count()==0 and pg.locator('#k-fsFA').is_disabled())
    pg.reload(); pg.wait_for_function('window.__ready'); pg.evaluate(CONNECT)
    check('after reopening: both years are still in the folder, FY2025 still final',pg.evaluate("window.__app.App.hist[2025].final")is True and pg.evaluate("window.__app.App.store.kind")=='folder' and pg.evaluate("window.__app.App.years.length")==2)
    if shots: pg.click('#tab-inputs'); pg.screenshot(path=str(shots/'final.png'),full_page=True)
    check('no JavaScript errors: '+str(errs),not errs)
    print('\nRESULT:','PASS' if not fails else 'FAIL '+str(fails)); b.close()
