import sys, time, subprocess, os, tempfile
sys.path.insert(0, "/root/.claude/skills/synced/06c47b33-d640-428c-a3a5-a42ae98094c8_2ae10ec8-84d2-43dd-8c58-6c8a9c68ff68/pptx/scripts")
from office.soffice import get_soffice_env
import uno
from com.sun.star.beans import PropertyValue
src, dst = map(os.path.abspath, sys.argv[1:3])
prof = tempfile.mkdtemp()
p = subprocess.Popen(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--invisible", "--norestore",
    "--accept=socket,host=127.0.0.1,port=2099;urp;"], env=get_soffice_env())
ctx = uno.getComponentContext()
res = ctx.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", ctx)
for _ in range(60):
    try:
        rc = res.resolve("uno:socket,host=127.0.0.1,port=2099;urp;StarOffice.ComponentContext"); break
    except Exception: time.sleep(1)
desk = rc.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", rc)
def pv(n, v): x = PropertyValue(); x.Name = n; x.Value = v; return x
doc = desk.loadComponentFromURL(uno.systemPathToFileUrl(src), "_blank", 0, (pv("Hidden", True),))
cnt = 0
def fix(shape):
    global cnt
    if shape.supportsService("com.sun.star.drawing.GroupShape"):
        for i in range(shape.getCount()): fix(shape.getByIndex(i))
        return
    if shape.supportsService("com.sun.star.drawing.TableShape"):
        m = shape.Model
        for r in range(m.Rows.Count):
            for c in range(m.Columns.Count):
                fixtext(m.getCellByPosition(c, r))
        return
    if hasattr(shape, "getText"):
        try: fixtext(shape)
        except Exception: pass
def fixtext(t):
    global cnt
    e = t.getText().createEnumeration()
    while e.hasMoreElements():
        para = e.nextElement()
        try: para.ParaIsCharacterDistance = False; cnt += 1
        except Exception: pass
pages = doc.DrawPages
for i in range(pages.Count):
    pg = pages.getByIndex(i)
    for j in range(pg.Count): fix(pg.getByIndex(j))
mp = doc.MasterPages
for i in range(mp.Count):
    pg = mp.getByIndex(i)
    for j in range(pg.Count): fix(pg.getByIndex(j))
doc.storeToURL(uno.systemPathToFileUrl(dst), (pv("FilterName", "impress_pdf_Export"),))
doc.close(True)
try: desk.terminate()
except Exception: pass
p.wait(timeout=30)
print("paras fixed", cnt)
