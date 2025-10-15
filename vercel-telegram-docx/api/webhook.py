
from fastapi import FastAPI, Request, Response
import os, re
from io import BytesIO
from docx import Document as Docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import httpx

TOKEN = os.environ["TELEGRAM_TOKEN"]
API = f"https://api.telegram.org/bot{TOKEN}"

app = FastAPI()

def rtl_para(p, align="right"):
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if align == "center" else WD_ALIGN_PARAGRAPH.RIGHT
    pf = p.paragraph_format
    pf.space_before = Pt(0); pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p._p.get_or_add_pPr().append(OxmlElement("w:rtl"))

def force_normal(doc):
    s = doc.styles["Normal"]; s.font.name = "Arial"; s.font.size = Pt(12)
    rPr = s._element.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts")) or OxmlElement("w:rFonts")
    if rf not in rPr: rPr.append(rf)
    for k in ["w:ascii","w:hAnsi","w:cs","w:eastAsia"]: rf.set(qn(k),"Arial")

def add_border(section):
    sectPr = section._sectPr
    b = OxmlElement("w:pgBorders"); b.set(qn("w:offsetFrom"),"page")
    for side in ["top","left","bottom","right"]:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"),"double"); el.set(qn("w:sz"),"12"); el.set(qn("w:space"),"12"); el.set(qn("w:color"),"auto")
        b.append(el)
    sectPr.append(b)

def is_bullet(s): 
    t = s.lstrip("\u200f ").lstrip()
    return t.startswith("•") or t.startswith("▪")

def is_num(s):
    t = s.lstrip("\u200f ").lstrip()
    return re.match(r"^\d+[.)]\s", t) or re.match(r"^[۰-۹]+[.)]\s", t)

@app.get("/health")
async def health():
    return {"ok": True}

def build_docx(raw: str) -> BytesIO:
    doc = Docx()
    for s in doc.sections:
        s.left_margin = s.right_margin = s.top_margin = s.bottom_margin = Cm(1.5)
        add_border(s)
    force_normal(doc)
    for line in raw.splitlines():
        p = doc.add_paragraph(); rtl_para(p, "right")
        if is_bullet(line) or is_num(line):
            pf = p.paragraph_format; pf.right_indent = Cm(0.75); pf.first_line_indent = -Cm(0.5)
        r = p.add_run(line); r.font.name = "Arial"; r.font.size = Pt(12)
        r._element.get_or_add_rPr().append(OxmlElement("w:rtl"))
    buf = BytesIO(); doc.save(buf); buf.seek(0); return buf

@app.post("/webhook")
async def webhook(req: Request):
    update = await req.json()
    msg = update.get("message") or update.get("edited_message")
    if not msg:
        return Response()
    chat_id = msg["chat"]["id"]
    text = msg.get("text")
    async with httpx.AsyncClient() as client:
        if text:
            doc = build_docx(text)
            files = {"document": ("خروجی_قالب_RTL_Arial.docx", doc.getvalue(),
                      "application/vnd.openxmlformats-officedocument.wordprocessingml.document")}
            await client.post(f"{API}/sendDocument", data={"chat_id": chat_id}, files=files)
        else:
            await client.post(f"{API}/sendMessage", data={"chat_id": chat_id,
                                    "text": "فقط متن بفرستید یا فایل txt (در این نسخه فقط متن پشتیبانی می‌شود)."})
    return Response()
