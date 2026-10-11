# -*- coding: utf-8 -*-
"""
CET-4 每日精读 · 云端自动生成并邮件推送
平台：腾讯云 CloudBase 云函数（Python 3.9+），纯标准库实现，无需第三方依赖。

流程：抓取 China Daily 最新文章(优先中华文化/科技前沿) -> 调用大模型生成精读 Markdown
      -> 用标准库生成 .docx -> SMTP 发送到指定邮箱。

环境变量：
  LLM_BASE_URL  例如 https://api.deepseek.com/v1   （兼容 OpenAI 接口）
  LLM_API_KEY   大模型 API Key
  LLM_MODEL     例如 deepseek-chat / glm-4-flash / qwen-turbo
  SMTP_HOST     例如 smtp.qq.com
  SMTP_PORT     465（SSL）
  SMTP_USER     发信邮箱
  SMTP_PASS     邮箱 SMTP 授权码
  MAIL_TO       收件邮箱（默认 2026200703@ruc.edu.cn）
  MAIL_FROM_NAME 发件人显示名（可选）
"""

import os
import re
import ssl
import json
import time
import zipfile
import smtplib
import datetime
import io
import urllib.request
import urllib.error
from html import escape, unescape
from email.message import EmailMessage

CHINA_DAILY_HOME = "https://www.chinadaily.com.cn/"
PREFER_KEYWORDS = ["culture", "cultural", "heritage", "museum", "tradition", "traditional",
                   "festival", "art", "artisan", "craft", "opera", "cuisine", "food", "tea",
                   "ancient", "dynasty", "calligraphy", "intangible", "tourism", "travel",
                   "tech", "technology", "AI", "artificial", "space", "rocket", "satellite",
                   "innovation", "science", "scientist", "robot", "digital", "smart",
                   "green", "energy", "medicine", "health", "education"]
# 命中即排除：政治/涉台/军事/灾祸/犯罪等不适合做学习材料的题材
EXCLUDE_KEYWORDS = ["taiwan", "lai", "tsai", "defense", "defence", "military", "army",
                    "navy", "missile", "war", "ukraine", "russia", "putin", "trump",
                    "sanction", "shooting", "shot", "killed", "dead", "death", "attack",
                    "crime", "murder", "court", "sentenced", "bribe", "graft", "quake",
                    "flood", "accident", "protest", "election", "senate", "congress",
                    "terror", "hostage", "drug", "corruption", "lawsuit"]


# ---------------------------------------------------------------- HTTP 工具
def http_get(url, timeout=25):
    req = urllib.request.Request(url, headers={
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/120.0 Safari/537.36",
        "Accept-Language": "en-US,en;q=0.9",
    })
    with urllib.request.urlopen(req, timeout=timeout) as r:
        raw = r.read()
    for enc in ("utf-8", "gb18030", "latin-1"):
        try:
            return raw.decode(enc)
        except Exception:
            continue
    return raw.decode("utf-8", "ignore")


def http_post_json(url, payload, headers, timeout=180):
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data, method="POST", headers=headers)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))


# ---------------------------------------------------------------- 抓文章
def fetch_candidates():
    """返回 [(url, title, date_str), ...] 按当天优先。"""
    html = http_get(CHINA_DAILY_HOME)
    today = datetime.datetime.now().strftime("%Y%m%d")
    # 注意：China Daily 首页用的是协议相对链接（//www...），正则需兼容
    pat = re.compile(
        r'<a\b[^>]*?href="((?:https?:)?//[^"]*?/a/(\d{6})/(\d{2})/(WS[0-9a-f]+)\.html)"[^>]*>(.*?)</a>',
        re.S | re.I)
    seen, items = set(), []
    for m in pat.finditer(html):
        url = m.group(1)
        if url.startswith("//"):
            url = "https:" + url
        if url in seen:
            continue
        title = unescape(re.sub(r"<[^>]+>", " ", m.group(5)))
        title = re.sub(r"\s+", " ", title).strip()
        if len(title) < 6:
            continue
        seen.add(url)
        items.append({"url": url, "title": title[:140], "date": m.group(2) + m.group(3)})
    # 当天优先
    items.sort(key=lambda x: (x["date"] != today,))
    return today, items


def score(item):
    t = item["title"].lower()
    for k in EXCLUDE_KEYWORDS:
        if k in t:
            return -1000          # 命中排除词：直接淘汰
    s = 0
    for k in PREFER_KEYWORDS:
        if k.lower() in t:
            s += 2
    return s


def rank_candidates(today, items, n=5):
    """当天优先 + 排除敏感/灾祸题材 + 主题打分，返回前 n 篇候选。"""
    pool = [i for i in items if i["date"] == today] or items
    pool = [i for i in pool if score(i) > -1000]
    pref = [i for i in pool if score(i) > 0]
    if pref:                      # 有中华文化/科技前沿类文章就优先这类
        pool = pref
    pool = sorted(pool, key=score, reverse=True)
    return pool[:n]


def pick_article(today, items):
    pool = [i for i in items if i["date"] == today] or items
    pool = sorted(pool, key=score, reverse=True)
    return pool[0] if pool else None


def extract_article(url):
    html = http_get(url)
    # 标题
    m = re.search(r"<title>(.*?)</title>", html, re.S | re.I)
    title = unescape(re.sub(r"\s+", " ", m.group(1))).strip() if m else ""
    title = re.sub(r"\s*[-|]\s*(China Daily|Chinadaily\.com\.cn|chinadaily\.com\.cn).*$", "", title, flags=re.I)
    # 作者与时间
    ma = re.search(r"By\s+([A-Z][A-Za-z\.\s]{2,40})\s*\|", html)
    author = re.sub(r"\s+", " ", ma.group(1)).strip() if ma else ""
    mt = re.search(r"Updated?:\s*([0-9]{4}-[0-9]{2}-[0-9]{2}[^<]{0,12})", html)
    pubdate = mt.group(1).strip() if mt else ""
    # 正文段落
    body = ""
    mc = re.search(r'<div[^>]*id="Content"[^>]*>(.*?)</div>\s*</div>', html, re.S | re.I)
    if not mc:
        mc = re.search(r'<div[^>]*class="[^"]*article[^"]*"[^>]*>(.*?)</div>\s*</div>', html, re.S | re.I)
    seg = mc.group(1) if mc else html
    paras = re.findall(r"<p[^>]*>(.*?)</p>", seg, re.S | re.I)
    clean = []
    for p in paras:
        txt = re.sub(r"<[^>]+>", "", p)
        txt = unescape(re.sub(r"\s+", " ", txt)).strip()
        if len(txt) > 40 and not txt.lower().startswith(("share", "click", "download", "photo by")):
            clean.append(txt)
    body = "\n\n".join(clean)
    return {"title": title, "author": author, "date": pubdate, "url": url, "body": body}


# ---------------------------------------------------------------- 大模型
SYSTEM_PROMPT = """你是 CET-4 英语助教。请把给定的 China Daily 英文文章加工成一份精读学习材料，\
严格输出 Markdown，只包含以下五部分，不要任何额外说明文字：

# CET-4 每日精读 · China Daily

> 整理日期：<今天日期>　|　适用水平：CET-4（难度适中偏高）　|　生词：精选 30 个高频词

## 一、文章信息

| 项目 | 内容 |
| --- | --- |
| 标题 | <英文标题> |
| 来源 | China Daily（中国日报英文网） |
| 作者 | <作者> |
| 发布时间 | <发布时间> |
| 网页链接 | <URL> |
| 体裁 | 新闻报道 / 特写 |
| 主题 | <一句话中文概括> |

## 二、原文（含生词 + 语法双重标注）

> **标注说明**
> - **加粗**单词后的 **①—㉚**：第三部分「CET-4 高频核心生词」（带圈数字）。
> - **加粗并加下划线**短语后的 **【1】—【6】**：第四部分「重点语法与长难句」（方头括号）。

**[1]** <段落原文，生词写成 **word**① 形式，语法短语写成 **<u>phrase</u>**【1】形式>
（逐段编号 [1][2]…，全文照录不改写）

## 三、CET-4 高频核心生词（精选 30 词）

> 序号为带圈数字 ①—㉚，与原文词末序号一一对应。

### ① <word> `/<音标>/`
- **词性**：<n./v./adj./adv.>
- **释义**：<中文释义>
- **例句**：<原文例句，目标词加粗>
- **语法与搭配**：<常用固定搭配与语法点>
- **词形变化**：<复数/时态/派生词>

（依次写满 30 个词）

## 四、重点语法与长难句（逐句标注）

**【1】出处：第 X 段第 Y 句**
- **原句**：<原句>
- **语法点**：<语法名称>
- **讲解**：<简明讲解>

（写满 6 条）

## 五、学习建议

1. …
2. …

要求：原文必须逐段照录、不得改写；生词序号 ①—㉚ 与原文标注一一对应；语法序号【1】—【6】与原文句末标注一一对应。"""


def gen_markdown(article):
    base = (os.environ.get("LLM_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
    key = os.environ.get("LLM_API_KEY") or ""
    model = os.environ.get("LLM_MODEL") or "deepseek-chat"
    today = datetime.datetime.now().strftime("%Y-%m-%d")
    user = ("【今天日期】{d}\n【文章标题】{t}\n【作者】{a}\n【发布时间】{p}\n【网页链接】{u}\n\n"
            "【原文】\n{b}").format(d=today, t=article["title"], a=article["author"],
                                   p=article["date"], u=article["url"], b=article["body"])
    payload = {
        "model": model,
        "messages": [{"role": "system", "content": SYSTEM_PROMPT},
                     {"role": "user", "content": user}],
        "temperature": 0.3,
        "max_tokens": 8000,
    }
    headers = {"Content-Type": "application/json", "Authorization": "Bearer " + key}
    resp = http_post_json(base + "/chat/completions", payload, headers)
    return resp["choices"][0]["message"]["content"]


CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳㉑㉒㉓㉔㉕㉖㉗㉘㉙㉚"
_FW = str.maketrans("０１２３４５６７８９", "0123456789")


def validate_md(md):
    """发送前自检：返回问题列表（空列表 = 通过）。"""
    problems = []
    norm = md.translate(_FW)
    miss_w = [c for c in CIRCLED if c not in md]
    if miss_w:
        problems.append("缺少生词序号 " + "".join(miss_w))
    miss_g = [str(i) for i in range(1, 7) if ("【%s】" % i) not in norm]
    if miss_g:
        problems.append("缺少语法序号 " + "、".join(miss_g))
    for sec in ("## 一、", "## 二、", "## 三、", "## 四、", "## 五、"):
        if sec not in md:
            problems.append("缺少标题 " + sec)
    n_word = sum(1 for c in CIRCLED if c in md)
    if 0 < n_word < 30:
        problems.append("生词数不足 30（当前 %d）" % n_word)
    return problems


# ---------------------------------------------------------------- Markdown -> docx（纯标准库）
def _run(text, bold=False, underline=False, italic=False, code=False, sz=None):
    rpr = "<w:rPr>"
    if code:
        rpr += '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:eastAsia="Consolas"/>'
    else:
        rpr += '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="微软雅黑"/>'
    if bold:
        rpr += "<w:b/>"
    if underline:
        rpr += '<w:u w:val="single"/>'
    if italic:
        rpr += "<w:i/>"
    if sz:
        rpr += '<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (sz, sz)
    rpr += "</w:rPr>"
    return '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, escape(text))


TOKEN = re.compile(r"\*\*<u>(.+?)</u>\*\*|\*\*(.+?)\*\*|<u>(.+?)</u>|`([^`]+)`|\*([^*\n]+)\*")


def _para(text, bold=False, sz=None, indent=0, quote=False, hr=False):
    ppr = "<w:pPr>"
    if indent:
        ppr += '<w:ind w:left="%d"/>' % indent
    if sz:
        ppr += '<w:rPr><w:sz w:val="%d"/></w:rPr>' % sz
    if hr:
        ppr += '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="BFBFBF"/></w:pBdr>'
    ppr += "</w:pPr>"
    runs, pos = [], 0
    for m in TOKEN.finditer(text):
        if m.start() > pos:
            runs.append(_run(text[pos:m.start()], bold=bold, sz=sz))
        if m.group(1) is not None:
            runs.append(_run(m.group(1), bold=True, underline=True, sz=sz))
        elif m.group(2) is not None:
            runs.append(_run(m.group(2), bold=True, sz=sz))
        elif m.group(3) is not None:
            runs.append(_run(m.group(3), underline=True, sz=sz))
        elif m.group(4) is not None:
            runs.append(_run(m.group(4), code=True, sz=sz))
        elif m.group(5) is not None:
            runs.append(_run(m.group(5), italic=True, sz=sz))
        pos = m.end()
    if pos < len(text):
        runs.append(_run(text[pos:], bold=bold, sz=sz))
    return "<w:p>%s%s</w:p>" % (ppr, "".join(runs))


def _cell(text, bold):
    return ('<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>'
            + _para(text, bold=bold, sz=20) + "</w:tc>")


def _table(rows):
    borders = ("<w:tblBorders>" + "".join(
        '<w:%s w:val="single" w:sz="4" w:space="0" w:color="808080"/>' % b
        for b in ("top", "left", "bottom", "right", "insideH", "insideV")) + "</w:tblBorders>")
    cols = len(rows[0]) if rows else 1
    grid = "<w:tblGrid>" + "".join('<w:gridCol w:w="%d"/>' % int(9000 / max(cols, 1))
                                   for _ in range(cols)) + "</w:tblGrid>"
    xml = ["<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/>" + borders + "</w:tblPr>", grid]
    for ri, row in enumerate(rows):
        xml.append("<w:tr>")
        for c in row:
            xml.append(_cell(c, ri == 0))
        xml.append("</w:tr>")
    xml.append("</w:tbl>")
    return "".join(xml)


HEAD = {1: 36, 2: 30, 3: 25, 4: 23}


def md_to_docx(md):
    body = []
    lines = md.replace("\r\n", "\n").split("\n")
    i, n = 0, len(lines)
    while i < n:
        s = lines[i].strip()
        if s == "":
            i += 1
            continue
        if s == "---":
            body.append(_para("", hr=True))
            i += 1
            continue
        if s.startswith("|"):
            rows = []
            while i < n and lines[i].strip().startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                rows.append(cells)
                i += 1
            rows = [r for r in rows if not all(set(c) <= set("-: ") for c in r)]
            if rows:
                cols = max(len(r) for r in rows)
                rows = [r + [""] * (cols - len(r)) for r in rows]
                body.append(_table(rows))
            continue
        m = re.match(r"^(#{1,4})\s+(.*)$", s)
        if m:
            body.append(_para(m.group(2), bold=True, sz=HEAD.get(len(m.group(1)), 22)))
            i += 1
            continue
        if s.startswith(">"):
            inner = s[1:].strip()
            bullet = inner.startswith("- ")
            if bullet:
                inner = inner[2:].strip()
            body.append(_para(("• " if bullet else "") + inner, sz=19, indent=(360 if bullet else 200)))
            i += 1
            continue
        if s.startswith("- "):
            body.append(_para("• " + s[2:].strip(), indent=360))
            i += 1
            continue
        if re.match(r"^\d+\.\s+", s):
            body.append(_para(s, indent=360))
            i += 1
            continue
        body.append(_para(s))
        i += 1

    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1247" w:right="1134" w:bottom="1247" w:left="1134" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
    document = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                "<w:body>" + "".join(body) + sect + "</w:body></w:document>")
    return document


CONTENT_TYPES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                 '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                 '<Default Extension="xml" ContentType="application/xml"/>'
                 '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                 "</Types>")

ROOT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
             "</Relationships>")


def build_docx(md):
    document = md_to_docx(md)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/document.xml", document)
    return buf.getvalue()


# ---------------------------------------------------------------- 发邮件
def send_mail(docx_bytes, filename, subject, body_text):
    host = os.environ.get("SMTP_HOST") or "smtp.qq.com"
    port = int(os.environ.get("SMTP_PORT") or "465")
    user = os.environ.get("SMTP_USER") or ""
    pwd = os.environ.get("SMTP_PASS") or ""
    if not user or not pwd:
        raise RuntimeError("SMTP_USER / SMTP_PASS not configured")
    to = os.environ.get("MAIL_TO") or "2026200703@ruc.edu.cn"
    from_name = os.environ.get("MAIL_FROM_NAME") or "CET-4 助教"

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = "%s <%s>" % (from_name, user)
    msg["To"] = to
    msg.set_content(body_text)
    msg.add_attachment(docx_bytes,
                       maintype="application",
                       subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
                       filename=filename)
    ctx = ssl.create_default_context()
    with smtplib.SMTP_SSL(host, port, context=ctx, timeout=60) as s:
        s.login(user, pwd)
        s.send_message(msg)


# ---------------------------------------------------------------- 入口
def main_handler(event=None, context=None):
    try:
        today = datetime.datetime.now().strftime("%Y-%m-%d")
        d, items = fetch_candidates()
        cands = rank_candidates(d, items, 4)
        if not cands:
            return {"ok": False, "msg": "no article found"}
        full = None
        for c in cands:  # 在候选里挑正文最长的一篇，内容更充实
            try:
                f = extract_article(c["url"])
            except Exception:
                continue
            if len(f["body"]) >= 400 and (full is None or len(f["body"]) > len(full["body"])):
                full = f
        if full is None:
            return {"ok": False, "msg": "no usable article (bodies too short)"}
        md = gen_markdown(full)
        warns = validate_md(md)
        if warns:  # 自检不通过 -> 重生成一次，取更完整的一版
            try:
                md2 = gen_markdown(full)
                w2 = validate_md(md2)
                if len(w2) < len(warns):
                    md, warns = md2, w2
            except Exception:
                pass
        docx = build_docx(md)
        fname = "CET4_ChinaDaily_%s.docx" % today
        try:  # 额外落盘一份（GitHub Actions 会作为 artifact 上传）
            outdir = os.path.join(os.getcwd(), "output")
            os.makedirs(outdir, exist_ok=True)
            with open(os.path.join(outdir, fname), "wb") as f:
                f.write(docx)
        except Exception:
            pass
        subject = "CET-4 每日精读 · China Daily（%s）| %s" % (today, full["title"])
        body = ("今日 CET-4 精读材料见附件（Word）。\n\n文章：%s\n来源：China Daily\n链接：%s\n"
                % (full["title"], full["url"]))
        if warns:
            body += "\n（自检提示：%s）\n" % "；".join(warns)
        send_mail(docx, fname, subject, body)
        return {"ok": True, "article": full["title"], "url": full["url"],
                "file": fname, "warnings": warns}
    except Exception as e:
        return {"ok": False, "error": repr(e)}


if __name__ == "__main__":
    import sys
    _res = main_handler()
    print(json.dumps(_res, ensure_ascii=False))
    sys.exit(0 if _res.get("ok") else 1)
