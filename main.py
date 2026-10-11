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
# 选题偏好（加权）：教育 / 科技 / 社会 / 文化 / 环境 优先，其中教育与社会权重更高
PREFER_KEYWORDS = {
    # —— 教育（权重 4）
    "education": 4, "school": 4, "student": 4, "university": 4, "college": 4,
    "teacher": 4, "campus": 4, "curriculum": 4, "classroom": 4, "exam": 4,
    "literacy": 4, "kindergarten": 4, "scholarship": 4, "lecture": 4, "tuition": 4,
    # —— 社会（权重 4）
    "society": 4, "social": 4, "community": 4, "volunteer": 4, "charity": 4,
    "welfare": 4, "elderly": 4, "youth": 4, "family": 4, "rural": 4,
    "city": 4, "urban": 4, "public": 4, "civic": 4, "poverty": 4,
    # —— 文化（权重 3）
    "culture": 3, "cultural": 3, "heritage": 3, "museum": 3, "tradition": 3,
    "traditional": 3, "festival": 3, "art": 3, "artisan": 3, "craft": 3,
    "opera": 3, "cuisine": 3, "ancient": 3, "dynasty": 3, "calligraphy": 3,
    "intangible": 3, "tourism": 3, "travel": 3, "archaeolog": 3, "folklore": 3,
    # —— 科技（权重 3）
    "tech": 3, "technology": 3, "ai": 3, "artificial": 3, "space": 3,
    "rocket": 3, "satellite": 3, "innovation": 3, "science": 3, "scientist": 3,
    "robot": 3, "digital": 3, "smart": 3, "computer": 3, "internet": 3,
    "data": 3, "algorithm": 3, "chip": 3, "startup": 3, "invention": 3,
    # —— 环境（权重 3）
    "environment": 3, "climate": 3, "green": 3, "energy": 3, "solar": 3,
    "wind power": 3, "wildlife": 3, "ecolog": 3, "pollution": 3, "conservation": 3,
    "sustainability": 3, "biodiversity": 3, "carbon": 3, "recycl": 3, "forest": 3,
    # —— 其他常见软新闻（权重 2）
    "medicine": 2, "health": 2, "sport": 2, "food": 2, "tea": 2,
    "panda": 2, "bird": 2, "agriculture": 2, "farming": 2, "village": 2,
}
# 命中即排除：政治/涉台/军事/灾祸/犯罪等不适合做学习材料的题材
EXCLUDE_KEYWORDS = ["taiwan", "lai", "tsai", "defense", "defence", "military", "army",
                    "navy", "missile", "war", "ukraine", "russia", "putin", "trump",
                    "sanction", "shooting", "shot", "killed", "dead", "death", "attack",
                    "crime", "murder", "court", "sentenced", "bribe", "graft", "quake",
                    "flood", "accident", "protest", "election", "senate", "congress",
                    "terror", "hostage", "drug", "corruption", "lawsuit", "tariff",
                    "visa", "embassy", "summit", "official says", "spokesperson"]



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


UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/122.0 Safari/537.36")


def http_post_json(url, payload, headers, timeout=180):
    data = json.dumps(payload).encode("utf-8")
    # 必须带浏览器 UA：Resend 等 API 前置 Cloudflare，会拦掉程序默认 UA（返回 403 / 1010）
    hdrs = {"User-Agent": UA, "Accept": "application/json"}
    hdrs.update(headers or {})
    req = urllib.request.Request(url, data=data, method="POST", headers=hdrs)
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
    for k, w in PREFER_KEYWORDS.items():
        if k in t:
            s += w                # 教育/社会/文化/科技/环境类加权
    return s


def rank_candidates(today, items, n=5):
    """当天优先 + 排除敏感/灾祸题材 + 主题打分，返回前 n 篇候选。"""
    pool = [i for i in items if i["date"] == today] or items
    pool = [i for i in pool if score(i) > -1000]
    pref = [i for i in pool if score(i) >= 3]
    if pref:                      # 有教育/科技/社会/文化/环境类文章就优先这类
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
WORD_PROMPT = """你是 CET-4 / 六级英语助教。请从给定的英文原文中，精选 30 个值得学习的核心词汇。
只输出一个 JSON 对象，不要任何解释、不要 Markdown 代码块。格式：
{"words":[{"w":"原文形式","base":"原形","lvl":"CET-4 核心","phon":"/音标/","pos":"n./v./adj./adv.","cn":"中文释义","eg":"取自原文的例句，目标词用 **加粗**","col":"常用固定搭配","der":"词形变化与派生词"}]}

选词要求：
1. **话题优先**：优先选取与「社会、教育、科技、文化、环境」相关的**话题词汇**（例如 heritage、curriculum、
   innovation、sustainability、urban、literacy、ecosystem 这一类），体现文章的领域表达；
2. **不限于四级考纲**：不必拘泥于 CET-4 词表。请包含 **8—12 个「拓展词」**（六级/雅思/托福高频词、
   学术词或地道表达），帮助扩大词汇量；其余为四级高频核心词；
3. `lvl` 字段标注级别，取值**只能是** `CET-4 核心` 或 `拓展`；
4. 必须正好 30 个、互不相同；`w` 必须与原文拼写**完全一致**（含时态、复数形式，如 scattered、embracing）；
5. `der` 写复数、时态形式与同根派生词；`col` 写 1—2 个常用搭配。
"""

GRAMMAR_PROMPT = """你是英语语法讲师。请从给定英文原文中挑出 6 处最适合 CET-4 学生学习的重点语法现象或长难句。
只输出一个 JSON 对象，不要任何解释、不要 Markdown 代码块。格式：
{"topic_cn":"一句话中文概括全文主题","grammar":[{"phrase":"原文中逐字出现的片段","point":"语法点名称","note":"简明中文讲解"}],"tips":["学习建议1","学习建议2","学习建议3"]}

硬性要求：
1. 必须正好 6 条，对应 6 个**不同**的句子，覆盖不同语法点（分词、定语从句、名词性从句、并列结构、时态等）；
2. `phrase` 必须是原文中**连续、逐字、一字不差**出现的片段（程序会用它在原文里定位，并自动加粗加下划线），
   长度 40—120 个字符，不要改写、不要用省略号、不要加引号；
3. `note` 用中文，1—2 句说清语法点与句子主干；
4. `tips` 给 3—4 条中文学习建议。
"""

CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳㉑㉒㉓㉔㉕㉖㉗㉘㉙㉚"


def llm_chat(system, user, temperature=0.3, max_tokens=8000, json_mode=False):
    base = (os.environ.get("LLM_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
    key = os.environ.get("LLM_API_KEY") or ""
    model = os.environ.get("LLM_MODEL") or "deepseek-chat"
    payload = {
        "model": model,
        "messages": [{"role": "system", "content": system},
                     {"role": "user", "content": user}],
        "temperature": temperature,
        "max_tokens": max_tokens,
    }
    if json_mode:
        payload["response_format"] = {"type": "json_object"}
    headers = {"Content-Type": "application/json", "Authorization": "Bearer " + key}
    resp = http_post_json(base + "/chat/completions", payload, headers)
    return resp["choices"][0]["message"]["content"]


def _strip_fence(s):
    s = (s or "").strip()
    s = re.sub(r"^```[a-zA-Z]*\s*", "", s)
    s = re.sub(r"\s*```$", "", s)
    return s.strip()


def llm_json(system, user, temperature=0.3, max_tokens=8000):
    txt = _strip_fence(llm_chat(system, user, temperature, max_tokens, json_mode=True))
    try:
        return json.loads(txt)
    except Exception:
        m = re.search(r"\{.*\}", txt, re.S)
        if m:
            return json.loads(m.group(0))
        raise


# ---------------------------------------------------------------- 标注（由代码精确写入，不依赖模型排版）
_SENT_END = re.compile(r"[.!?](?=[\"')\]]?(?:\s|$))")
_ABBR = re.compile(r"(?:Mr|Mrs|Ms|Dr|St|No|U\.S|U\.K|Co|Inc|Ltd|vs|etc)\.$", re.I)


def locate_phrase(paras, phrase):
    """在段落里定位模型给出的短语（容忍空白/大小写差异），返回 (段号, 起, 止)。"""
    toks = (phrase or "").split()
    if len(toks) < 3:
        return None
    for cut in (len(toks), int(len(toks) * 0.7), int(len(toks) * 0.5)):
        if cut < 3:
            continue
        rx = re.compile(r"\s+".join(re.escape(w) for w in toks[:cut]), re.I)
        for i, p in enumerate(paras):
            m = rx.search(p)
            if m:
                return i, m.start(), m.end()
    return None


def sentence_end(text, start):
    """从 start 向后找句子结尾，返回「句末注释序号」的插入位置。"""
    for m in _SENT_END.finditer(text, start):
        if _ABBR.search(text[max(0, m.start() - 6):m.start() + 1]):
            continue
        pos = m.end()
        while pos < len(text) and text[pos] in "\"')]":
            pos += 1
        return pos
    return len(text)


def sentence_span(text, pos):
    """返回 pos 所在句子的 (起, 止)。"""
    st = 0
    for m in _SENT_END.finditer(text, 0, pos):
        st = m.end()
    return st, sentence_end(text, pos)


def sentence_no(text, pos):
    """pos 所在句子在本段中的序号（1 起）。"""
    return sum(1 for _ in _SENT_END.finditer(text, 0, pos)) + 1


def mark_word(text, word, n):
    """把第一个独立出现的 word 标成 **word**①。"""
    if not word or n < 1 or n > len(CIRCLED):
        return text
    m = re.search(r"(?<![A-Za-z])" + re.escape(word) + r"(?![A-Za-z])", text)
    if not m:
        return text
    return text[:m.start()] + "**" + m.group(0) + "**" + CIRCLED[n - 1] + text[m.end():]


def build_section2(paras, words, grammar):
    """生成第二部分：语法短语 **<u>…</u>** + 句末【n】；生词 **word**①。"""
    spots = {}
    infos = []
    for n, g in enumerate(grammar, 1):
        loc = locate_phrase(paras, g.get("phrase", ""))
        if not loc:
            infos.append({"n": n, "g": g, "found": False})
            continue
        i, s, e = loc
        spots.setdefault(i, []).append((s, e, n))
        st, en = sentence_span(paras[i], s)
        infos.append({"n": n, "g": g, "found": True, "para": i + 1,
                      "sent": sentence_no(paras[i], s),
                      "sentence": paras[i][st:en].strip()})
    marked = []
    for i, p in enumerate(paras):
        for (s, e, n) in sorted(spots.get(i, []), key=lambda x: -x[0]):
            pos = max(sentence_end(p, e), e)
            p = p[:pos] + "【%d】" % n + p[pos:]
            p = p[:s] + "**<u>" + p[s:e] + "</u>**" + p[e:]
        marked.append(p)
    for n, w in enumerate(words[:len(CIRCLED)], 1):
        wd = (w.get("w") or w.get("base") or "").strip()
        if not wd:
            continue
        for i in range(len(marked)):
            m = re.search(r"(?<![A-Za-z])" + re.escape(wd) + r"(?![A-Za-z])", marked[i])
            if m:
                marked[i] = mark_word(marked[i], wd, n)
                break
    return marked, infos


def build_markdown(article, today, words, gj):
    paras = [p.strip() for p in article["body"].split("\n\n") if p.strip()]
    grammar = [g for g in (gj.get("grammar") or []) if isinstance(g, dict) and g.get("phrase")][:6]
    marked, infos = build_section2(paras, words, grammar)
    sec2 = "\n\n".join("**[%d]** %s" % (i + 1, p) for i, p in enumerate(marked))

    # 第一部分
    head = ["# CET-4 每日精读 · China Daily", "",
            "> 整理日期：%s　|　适用水平：CET-4（含适度拓展）　|　词汇：30 个（四级核心 + 拓展词）" % today,
            "", "## 一、文章信息", "",
            "| 项目 | 内容 |", "| --- | --- |",
            "| 标题 | %s |" % article["title"],
            "| 来源 | China Daily（中国日报英文网） |",
            "| 作者 | %s |" % (article["author"] or "China Daily"),
            "| 发布时间 | %s |" % article["date"],
            "| 网页链接 | %s |" % article["url"],
            "| 体裁 | 新闻报道 / 特写 |",
            "| 主题 | %s |" % (gj.get("topic_cn") or ""),
            "", "## 二、原文（含生词 + 语法双重标注）", "",
            "> **标注说明**",
            "> - **加粗**单词后的 **①—㉚**：第三部分「核心词汇与拓展词」（带圈数字）。",
            "> - **加粗并加下划线**短语后的 **【1】—【6】**：第四部分「重点语法与长难句」（方头括号）。",
            "", sec2, ""]

    # 第三部分（每个生词独占一块，字段分行，便于阅读）
    sec3 = ["## 三、核心词汇与拓展词（精选 30 词）", "",
            "> 序号为带圈数字 ①—㉚，与原文词末序号一一对应。",
            "> 选词侧重社会 / 教育 / 科技 / 文化 / 环境等话题词汇；标 **〔拓展〕** 者为超出四级考纲的拓展词。", ""]
    for n, w in enumerate(words[:len(CIRCLED)], 1):
        wd = (w.get("w") or w.get("base") or "").strip()
        if not wd:
            continue
        base = (w.get("base") or wd).strip()
        phon = (w.get("phon") or "").strip()
        tag = "  **〔拓展〕**" if ("拓展" in (w.get("lvl") or "")) else ""
        sec3.append("**%s %s**%s%s" % (CIRCLED[n - 1], wd, ("  `" + phon + "`") if phon else "", tag))
        if base and base.lower() != wd.lower():
            sec3.append("- **原形**：%s" % base)
        sec3.append("- **词性**：%s" % (w.get("pos") or ""))
        sec3.append("- **级别**：%s" % ((w.get("lvl") or "CET-4 核心").strip()))
        sec3.append("- **释义**：%s" % (w.get("cn") or ""))
        sec3.append("- **例句**：%s" % (w.get("eg") or ""))
        sec3.append("- **语法与搭配**：%s" % (w.get("col") or ""))
        sec3.append("- **词形变化**：%s" % (w.get("der") or ""))
        sec3.append("")

    # 第四部分
    sec4 = ["## 四、重点语法与长难句（逐句标注）", ""]
    for it in infos:
        g = it["g"]
        if it["found"]:
            sec4.append("**【%d】出处：第 %d 段第 %d 句**" % (it["n"], it["para"], it["sent"]))
            sec4.append("- **原句**：%s" % it["sentence"])
        else:
            sec4.append("**【%d】出处：见原文**" % it["n"])
            sec4.append("- **原句**：%s" % (g.get("phrase") or ""))
        sec4.append("- **语法点**：%s" % (g.get("point") or ""))
        sec4.append("- **讲解**：%s" % (g.get("note") or ""))
        sec4.append("")

    # 第五部分
    sec5 = ["## 五、学习建议", ""]
    for i, t in enumerate((gj.get("tips") or [])[:5], 1):
        sec5.append("%d. %s" % (i, t))
    sec5.append("")

    md = "\n".join(head + sec3 + sec4 + sec5)
    miss_w = [n for n in range(1, len(words[:30]) + 1) if CIRCLED[n - 1] not in sec2]
    miss_g = [n for n in range(1, 7) if ("【%d】" % n) not in sec2]
    return md, miss_w, miss_g


def gen_markdown(article):
    """先让模型给出「生词表 + 语法清单」（JSON），再由代码把标注精确写进原文。"""
    today = datetime.datetime.now().strftime("%Y-%m-%d")
    paras = [p.strip() for p in article["body"].split("\n\n") if p.strip()]
    numbered = "\n\n".join("【第%d段】%s" % (i + 1, p) for i, p in enumerate(paras))
    base_user = ("【今天日期】{d}\n【文章标题】{t}\n【作者】{a}\n【发布时间】{p}\n【网页链接】{u}\n\n"
                 "【原文（分段编号）】\n{b}").format(
        d=today, t=article["title"], a=article["author"], p=article["date"],
        u=article["url"], b=numbered)
    words = []
    try:
        wj = llm_json(WORD_PROMPT, base_user, temperature=0.3)
        words = [w for w in (wj.get("words") or []) if isinstance(w, dict) and (w.get("w") or w.get("base"))]
        words = words[:len(CIRCLED)]          # 严格截断到 30 个，防越界
    except Exception:
        words = []
    best = None
    for attempt in (1, 2):
        try:
            gj = llm_json(GRAMMAR_PROMPT, base_user, temperature=0.3)
        except Exception:
            gj = {}
        gj.setdefault("tips", [])
        md, miss_w, miss_g = build_markdown(article, today, words, gj)
        score = len(miss_g) + len(miss_w)
        if best is None or score < best[0]:
            best = (score, md, miss_w, miss_g)
        if score == 0:
            break
    return best[1]



_FW = str.maketrans("０１２３４５６７８９", "0123456789")


def split_sections(md):
    """把 Markdown 按五个部分切分，返回 {'一':..., '二':...} 形式的字典。"""
    idx = {}
    for key, tag in (("一", "## 一、"), ("二", "## 二、"), ("三", "## 三、"),
                     ("四", "## 四、"), ("五", "## 五、")):
        idx[key] = md.find(tag)
    sec = {}
    for key in ("一", "二", "三", "四", "五"):
        start = idx[key]
        if start < 0:
            sec[key] = ""
            continue
        later = [v for v in idx.values() if v > start]
        sec[key] = md[start:(min(later) if later else len(md))]
    return sec


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

    # ---- 重点：原文（第二部分）里必须"就地"标注，不能只在后面几部分讲 ----
    sec2 = split_sections(md).get("二", "").translate(_FW)
    # 去掉「> 标注说明」等引用/说明行，避免其中的字样干扰判断
    sec2 = "\n".join(l for l in sec2.split("\n") if not l.strip().startswith(">"))
    if sec2:
        miss_g2 = [str(i) for i in range(1, 7) if ("【%s】" % i) not in sec2]
        if miss_g2:
            problems.append("原文缺语法标注【%s】" % "】【".join(miss_g2))
        n_u = sec2.count("<u>") + sec2.count("<ins>")
        if n_u < 6:
            problems.append("原文加下划线短语仅 %d 处（应≥6）" % n_u)
        n_w2 = sum(1 for c in CIRCLED if c in sec2)
        if n_w2 < 30:
            problems.append("原文里生词标注仅 %d 个（应=30）" % n_w2)
    else:
        problems.append("未找到第二部分正文")
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
            # 语法标注：加粗 + 下划线；内部若嵌了 **word**① 去掉多余星号，保留带圈序号
            runs.append(_run(m.group(1).replace("**", ""), bold=True, underline=True, sz=sz))
        elif m.group(2) is not None:
            runs.append(_run(m.group(2), bold=True, sz=sz))
        elif m.group(3) is not None:
            runs.append(_run(m.group(3).replace("**", ""), underline=True, sz=sz))
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
def _build_msg(user, from_name, to, subject, body_text, docx_bytes, filename):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = "%s <%s>" % (from_name, user)
    msg["To"] = to
    msg.set_content(body_text)
    msg.add_attachment(docx_bytes,
                       maintype="application",
                       subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
                       filename=filename)
    return msg


def _smtp_send(host, port, mode, user, pwd, msg, ctx):
    if mode == "ssl":
        with smtplib.SMTP_SSL(host, port, context=ctx, timeout=60) as s:
            s.login(user, pwd)
            s.send_message(msg)
        return
    s = smtplib.SMTP(host, port, timeout=60)
    try:
        s.ehlo()
        s.starttls(context=ctx)
        s.ehlo()
        s.login(user, pwd)
        s.send_message(msg)
    finally:
        try:
            s.quit()
        except Exception:
            pass


def _send_via_resend(user, from_name, to, subject, body_text, docx_bytes, filename):
    """可选：配置 RESEND_API_KEY 后改用 Resend HTTP API 发信（海外服务器不被邮箱风控）。

    未验证域名时，Resend 只允许：from 用 onboarding@resend.dev，且只能发到「注册 Resend 的那个邮箱」。
    可用 RESEND_FROM 覆盖发件地址（已验证域名时填写）。
    """
    key = os.environ.get("RESEND_API_KEY")
    if not key:
        return False
    import base64
    from_addr = os.environ.get("RESEND_FROM") or "onboarding@resend.dev"
    payload = {
        "from": "%s <%s>" % (from_name, from_addr),
        "to": [to],
        "subject": subject,
        "text": body_text,
        "attachments": [{"filename": filename,
                         "content": base64.b64encode(docx_bytes).decode()}],
    }
    http_post_json("https://api.resend.com/emails", payload,
                   {"Content-Type": "application/json", "Authorization": "Bearer " + key})
    return True


def send_mail(docx_bytes, filename, subject, body_text):
    host = os.environ.get("SMTP_HOST") or "smtp.qq.com"
    user = os.environ.get("SMTP_USER") or ""
    pwd = os.environ.get("SMTP_PASS") or ""
    to = os.environ.get("MAIL_TO") or "2026200703@ruc.edu.cn"
    from_name = os.environ.get("MAIL_FROM_NAME") or "CET-4 助教"

    # 若配置了 Resend，优先用它（不受海外云 IP 风控影响）
    try:
        if _send_via_resend(user, from_name, to, subject, body_text, docx_bytes, filename):
            return
    except Exception as e:
        print("resend failed:", repr(e))

    if not user or not pwd:
        raise RuntimeError("SMTP_USER / SMTP_PASS not configured")

    msg = _build_msg(user, from_name, to, subject, body_text, docx_bytes, filename)
    ctx = ssl.create_default_context()
    port_env = os.environ.get("SMTP_PORT") or ""
    if port_env:
        p = int(port_env)
        plan = [(host, p, "ssl" if p == 465 else "starttls")]
    else:
        # 依次尝试：465(SSL) → 587(STARTTLS) → 465(SSL)，每个再重试一次
        plan = [(host, 465, "ssl"), (host, 587, "starttls"), (host, 465, "ssl")]

    errs = []
    for h, p, mode in plan:
        for attempt in (1, 2):
            try:
                _smtp_send(h, p, mode, user, pwd, msg, ctx)
                return
            except Exception as e:
                errs.append("%s:%d/%s try%d -> %r" % (h, p, mode, attempt, e))
                time.sleep(2)
    raise RuntimeError("mail send failed: " + " | ".join(errs))


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
        if warns:  # 自检不通过 -> 整篇重生成一次，取更完整的一版
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
        import traceback
        return {"ok": False, "error": repr(e),
                "trace": traceback.format_exc()[-1800:]}


if __name__ == "__main__":
    import sys
    _res = main_handler()
    print(json.dumps(_res, ensure_ascii=False))
    sys.exit(0 if _res.get("ok") else 1)
