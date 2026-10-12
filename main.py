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
# 选题优先级（用户指定，权重即优先级）：科技 > 文化 > 体育健康 > 生态 > 社会 > 教育 > 其他
PREFER_KEYWORDS = {
    # —— ① 科技（权重 6）：AI、新发现发明、新技术、自然科学
    "tech": 6, "technology": 6, "technological": 6, "ai": 6, "artificial intelligence": 6,
    "robot": 6, "algorithm": 6, "chip": 6, "semiconductor": 6, "quantum": 6,
    "space": 6, "rocket": 6, "satellite": 6, "astronaut": 6, "lunar": 6, "mars": 6,
    "science": 6, "scientist": 6, "research": 6, "breakthrough": 6, "discovery": 6,
    "invention": 6, "innovate": 6, "innovation": 6, "laboratory": 6, "experiment": 6,
    "digital": 6, "internet": 6, "computer": 6, "smart": 6, "startup": 6,
    "physics": 6, "chemistry": 6, "biology": 6, "gene": 6, "vaccine": 6,
    # —— ② 文化（权重 5）：中华文化、中外文化交流、世界文化
    "culture": 5, "cultural": 5, "heritage": 5, "civilization": 5, "tradition": 5,
    "traditional": 5, "museum": 5, "archaeolog": 5, "relic": 5, "calligraphy": 5,
    "opera": 5, "folk": 5, "folklore": 5, "intangible": 5, "dynasty": 5,
    "ancient": 5, "art": 5, "artisan": 5, "craft": 5, "cuisine": 5,
    "exchange": 5, "festival": 5, "tourism": 5, "travel": 5, "peony": 5,
    # —— ③ 体育健康（权重 4）
    "sport": 4, "athlete": 4, "olympic": 4, "paralympic": 4, "tournament": 4,
    "championship": 4, "medal": 4, "fitness": 4, "coach": 4, "stadium": 4,
    "football": 4, "basketball": 4, "swimming": 4, "marathon": 4, "youth": 4,
    "health": 4, "healthy": 4, "exercise": 4, "nutrition": 4, "wellness": 4,
    # —— ③b 体育细化偏好（用户偏好：青年男运动员 / 足球·篮球·田径 / 体能与健身）
    "footballer": 4, "football team": 4, "basketball player": 4, "volleyball": 4,
    "track and field": 4, "sprinter": 4, "hurdler": 4, "athletics": 4,
    "running": 4, "sprint": 4, "jogging": 4, "student athlete": 4,
    "young athlete": 4, "university team": 4, "sports meeting": 4, "university games": 4,
    "training": 4, "workout": 4, "gym": 4, "strength": 4, "physique": 4,
    # —— ④ 生态（权重 4）：环境保护
    "environment": 4, "environmental": 4, "climate": 4, "green": 4, "ecology": 4,
    "ecolog": 4, "wildlife": 4, "biodiversity": 4, "conservation": 4, "sustainability": 4,
    "recycl": 4, "carbon": 4, "emission": 4, "pollution": 4, "forest": 4,
    "energy": 4, "solar": 4, "wind power": 4, "wetland": 4, "migratory": 4,
    # —— ⑤ 社会（权重 3）：文史哲、社会科学、调查研究
    "society": 3, "social": 3, "community": 3, "volunteer": 3, "charity": 3,
    "welfare": 3, "elderly": 3, "survey": 3, "study": 3, "study finds": 3,
    "philosophy": 3, "history": 3, "historian": 3, "literature": 3, "rural": 3,
    "urban": 3, "poverty": 3, "census": 3, "demographic": 3, "employment": 3,
    # —— ⑥ 教育（权重 3）
    "education": 3, "school": 3, "student": 3, "university": 3, "college": 3,
    "teacher": 3, "campus": 3, "curriculum": 3, "classroom": 3, "exam": 3,
    "literacy": 3, "kindergarten": 3, "scholarship": 3, "lecture": 3, "tuition": 3,
    # —— 其他常见软新闻（权重 1）
    "medicine": 1, "food": 1, "tea": 1, "panda": 1, "bird": 1,
    "agriculture": 1, "farming": 1, "village": 1, "city": 1, "public": 1,
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


def rank_candidates(today, items, n=10):
    """当天优先 + 排除敏感/灾祸题材 + 主题打分，返回前 n 篇候选。"""
    pool = [i for i in items if i["date"] == today] or items
    pool = [i for i in pool if score(i) > -1000]
    pool = sorted(pool, key=score, reverse=True)
    pref = [i for i in pool if score(i) >= 3]
    if len(pref) >= 4:            # 目标题材够多时只看目标题材
        pool = pref
    return pool[:n]


def pick_article(today, items):
    pool = [i for i in items if i["date"] == today] or items
    pool = sorted(pool, key=score, reverse=True)
    return pool[0] if pool else None


# 代码/CSS/JS 特征：命中即判定"不是正文散文"
_JUNK_MARK = ("{", "}", "function(", "var ", "margin:", "padding:", "color:", "width:",
              "height:", "@media", "px;", "display:", "background:", "document.",
              "getElementById", "://", "www.", "cookie", "jQuery")


def _clean_html(html):
    """去掉脚本 / 样式 / 注释，避免把 CSS、JS 当成正文抓进来。"""
    h = re.sub(r"<!--.*?-->", " ", html, flags=re.S)
    for tag in ("script", "style", "noscript", "template", "iframe", "svg"):
        h = re.sub(r"<%s[^>]*>.*?</%s>" % (tag, tag), " ", h, flags=re.S | re.I)
    return h


def _is_prose(txt):
    """判断一段文本是否像"正文散文"（而非代码、导航、广告）。"""
    if not (40 <= len(txt) <= 4000):
        return False
    if any(k in txt for k in _JUNK_MARK):
        return False
    letters = sum(1 for c in txt if c.isalpha() or c in " ,.'\"-;:!?()")
    if letters / float(len(txt)) < 0.85:
        return False
    low = txt.lower()
    return not low.startswith(("share", "click", "download", "photo by", "related",
                               "copyright", "advertisement", "subscribe", "read more"))


def _paragraphs_from(seg):
    out = []
    for p in re.findall(r"<p[^>]*>(.*?)</p>", seg, re.S | re.I):
        txt = re.sub(r"<[^>]+>", "", p)
        txt = unescape(re.sub(r"\s+", " ", txt)).strip()
        if _is_prose(txt):
            out.append(txt)
    return out


def _title_overlap(title, body):
    """标题里的实词有多少出现在正文中（用于确认"抽到的正文=这篇文章"）。
    用词干前缀匹配，避免 accessibility/accessible 这类词形变化造成误判。"""
    words = [w.lower() for w in re.findall(r"[A-Za-z]{4,}", title)]
    if not words:
        return 1.0
    low = body.lower()
    hit = 0
    for w in words:
        stem = w[:6] if len(w) >= 6 else w
        if w in low or stem in low:
            hit += 1
    return hit / float(len(words))


def _meta_desc(html):
    """取页面 meta description / og:description（用于校验正文是否对得上这篇文章）。"""
    for pat in (r'<meta[^>]+property=["\']og:description["\'][^>]+content=["\'](.*?)["\']',
                r'<meta[^>]+name=["\']description["\'][^>]+content=["\'](.*?)["\']',
                r'<meta[^>]+content=["\'](.*?)["\'][^>]+name=["\']description["\']'):
        m = re.search(pat, html, re.S | re.I)
        if m:
            d = unescape(re.sub(r"\s+", " ", m.group(1))).strip()
            if len(d) >= 60:
                return d
    return ""


def _desc_overlap(desc, body):
    """正文与 meta description 的吻合度。"""
    words = [w.lower() for w in re.findall(r"[A-Za-z]{5,}", desc)]
    if len(words) < 6:
        return -1.0
    low = body.lower()
    return sum(1 for w in words if w[:6] in low) / float(len(words))


def _photo_info(html, seg=""):
    """提取文章配图的客观信息：优先正文内的图注 / alt，其次页面级 og:image 说明。"""
    def clean(t):
        return unescape(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t))).strip()

    cap, alt, src = "", "", "content"
    # ① 正文内：figcaption / caption 段落 / 图片 alt
    if seg:
        for pat in (r'<figcaption[^>]*>(.*?)</figcaption>',
                    r'<p[^>]*class="[^"]*(?:caption|photo-des|desc)[^"]*"[^>]*>(.*?)</p>',
                    r'<span[^>]*class="[^"]*(?:caption|photo-des)[^"]*"[^>]*>(.*?)</span>'):
            m = re.search(pat, seg, re.S | re.I)
            if m:
                cap = clean(m.group(1))
                if len(cap) >= 8:
                    break
                cap = ""
        mi = re.search(r'<img[^>]+alt=["\']([^"\']{6,220})["\']', seg, re.I)
        alt = clean(mi.group(1)) if mi else ""
    # ② 页面级兜底：og:image:alt / 首个 caption
    if not (cap or alt):
        src = "page"
        for pat in (r'<meta[^>]+property=["\']og:image:alt["\'][^>]+content=["\'](.*?)["\']',
                    r'<meta[^>]+name=["\']twitter:image:alt["\'][^>]+content=["\'](.*?)["\']'):
            m = re.search(pat, html, re.S | re.I)
            if m:
                alt = clean(m.group(1))
                if len(alt) >= 5:
                    break
                alt = ""
        for pat in (r'<figcaption[^>]*>(.*?)</figcaption>',
                    r'<p[^>]*class="[^"]*(?:caption|photo-des)[^"]*"[^>]*>(.*?)</p>'):
            m = re.search(pat, html, re.S | re.I)
            if m:
                cap = clean(m.group(1))
                if len(cap) >= 8:
                    break
                cap = ""
    has = bool(re.search(r'<meta[^>]+property=["\']og:image["\']', html, re.I)) or bool(cap or alt)
    return {"has_image": has, "photo_alt": alt[:300], "photo_caption": cap[:300], "photo_src": src}


def extract_article(url):
    raw = http_get(url)
    html = _clean_html(raw)
    # 标题
    m = re.search(r"<title>(.*?)</title>", html, re.S | re.I)
    title = unescape(re.sub(r"\s+", " ", m.group(1))).strip() if m else ""
    title = re.sub(r"\s*[-|]\s*(China Daily|Chinadaily\.com\.cn|chinadaily\.com\.cn).*$", "", title, flags=re.I)
    # 作者与时间
    ma = re.search(r"By\s+([A-Z][A-Za-z\.\s]{2,40})\s*\|", html)
    author = re.sub(r"\s+", " ", ma.group(1)).strip() if ma else ""
    mt = re.search(r"Updated?:\s*([0-9]{4}-[0-9]{2}-[0-9]{2}[^<]{0,12})", html)
    pubdate = mt.group(1).strip() if mt else ""

    # 正文：依次尝试多种容器，选"与标题最吻合、且内容最充实"的那一份
    segs = []
    for pat in (r'<div[^>]*id="Content"[^>]*>(.*?)</div>\s*</div>',
                r'<div[^>]*class="[^"]*article[^"]*"[^>]*>(.*?)</div>\s*</div>',
                r'<div[^>]*id="Content"[^>]*>(.*)',
                r'<div[^>]*class="[^"]*article[^"]*"[^>]*>(.*)'):
        mm = re.search(pat, html, re.S | re.I)
        if mm:
            segs.append(mm.group(1)[:300000])
    segs.append(html)
    desc = _meta_desc(html)
    best, best_key, best_seg = [], (-1, -1), ""
    for seg in segs:
        ps = _paragraphs_from(seg)
        if len(ps) < 2:
            continue
        text = " ".join(ps)
        ov = _title_overlap(title, text)
        dov = _desc_overlap(desc, text)
        good = 1 if (ov >= 0.5 or dov >= 0.6) else 0
        key = (good, len(text))
        if key > best_key:
            best, best_key, best_seg = ps, key, seg
    body = "\n\n".join(best)
    ph = _photo_info(raw, best_seg)
    return {"title": title, "author": author, "date": pubdate, "url": url,
            "body": body, "desc": desc, "photo_alt": ph["photo_alt"],
            "photo_caption": ph["photo_caption"], "has_image": ph["has_image"],
            "match": _title_overlap(title, body),
            "desc_match": _desc_overlap(desc, body)}



def _pub_age_days(pubdate, today):
    """文章发布日距今天数（用于时效性把关）；解析失败返回 999。"""
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", pubdate or "")
    if not m:
        return 999
    try:
        d = datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        t = datetime.datetime.strptime(today, "%Y-%m-%d").date()
    except Exception:
        return 999
    return max(0, (t - d).days)


def pick_best_article(cands, today=None):
    """挑"题材优先级最高 + 正文够长（难度更高）+ 与标题吻合 + 时效新"的一篇。"""
    today = today or datetime.datetime.now().strftime("%Y-%m-%d")
    notes = []
    cache = {}

    def get(c):
        if c["url"] not in cache:
            try:
                cache[c["url"]] = extract_article(c["url"])
            except Exception as e:
                cache[c["url"]] = None
                notes.append("%s -> 抓取失败 %r" % (c["title"][:28], e))
        return cache[c["url"]]

    for minlen in (1200, 800):        # 先要"够长够充实"的，没有再放宽
        best, key = None, None
        for c in cands[:8]:
            f = get(c)
            if not f:
                continue
            age = _pub_age_days(f.get("date", ""), today)
            if age > 3:               # 时效性：超过 3 天的一律不用
                notes.append("%s -> 时效性不足(发布 %d 天前)" % (c["title"][:28], age))
                continue
            if len(f["body"]) < minlen:
                notes.append("%s -> 正文过短(%d)" % (c["title"][:28], len(f["body"])))
                continue
            if not (f.get("match", 1.0) >= 0.5 or f.get("desc_match", -1.0) >= 0.6):
                notes.append("%s -> 正文与标题/摘要不吻合" % c["title"][:28])
                continue
            f["age_days"] = age
            k = (score(c), len(f["body"]), -age)   # 主题优先级 > 篇幅 > 越新越优先
            if key is None or k > key:
                key, best = k, f
        if best is not None:
            return best, notes
    best = None
    for c in cands[:8]:               # 兜底：取最长且不乱的正文（仍要求 7 天内）
        f = get(c)
        if not f or len(f["body"]) < 500:
            continue
        age = _pub_age_days(f.get("date", ""), today)
        if age > 7:
            continue
        f["age_days"] = age
        if best is None or len(f["body"]) > len(best["body"]):
            best = f
    return best, notes


# ---------------------------------------------------------------- 大模型
WORD_PROMPT = """你是 CET-4 高分 / 六级英语助教。请从给定的英文原文中，精选 30 个**有难度、值得精学**的核心词汇。
只输出一个 JSON 对象，不要任何解释、不要 Markdown 代码块。格式：
{"words":[{"w":"原文形式","base":"原形","lvl":"CET-4 核心","phon":"/音标/","pos":"n./v./adj./adv.","cn":"中文释义","eg":"取自原文的例句，目标词用 **加粗**","col":"常用固定搭配","der":"词形变化与派生词"}]}

选词要求：
1. **难度要够**：面向**四级高分 / 六级水平**，优先选**不易掌握**的词（抽象名词、学术动词、多义动词、
   形容词化名词、地道搭配）。**不要**选过于基础的常见词（如 city、public、family、school、good、make、
   people、year、work、way 这类初中/基础词汇）；
2. **话题优先**：优先与「文化、科技、教育、社会、体育、环境、经济」相关的**话题领域词**
   （例如 heritage、curriculum、innovation、sustainability、infrastructure、urbanization、
   cognitive、collaborative、resilience、stamina 这一类）；
3. **不限于四级考纲**：请包含 **12—20 个「拓展词」**（六级 / 雅思 / 托福高频词或学术词），
   其余为四级中偏难的词（**四级核心词请保留 8 个以上**）；`lvl` 取值**只能是** `CET-4 核心` 或 `拓展`；
4. 必须正好 30 个、互不相同；`w` 必须与原文拼写**完全一致**（含时态、复数形式，如 scattered、embracing）；
5. `der` 写复数、时态形式与同根派生词；`col` 写 1—2 个常用搭配。
"""

GRAMMAR_PROMPT = """你是英语语法讲师。请从给定英文原文中挑出 6 处**有难度、最具学习价值**的重点语法现象或长难句，
并对文章配图给出一句客观描述。只输出一个 JSON 对象，不要任何解释、不要 Markdown 代码块。格式：
{"topic_cn":"一句话中文概括全文主题","photo_cn":"文章配图的客观描述","grammar":[{"phrase":"原文中逐字出现的片段","point":"语法点名称","note":"简明中文讲解"}],"tips":["学习建议1","学习建议2","学习建议3"]}

硬性要求：
1. **`photo_cn`（配图描述）**：只依据用户消息里给出的「配图信息」（图片 alt 文本 / 图注），做一句**客观、事实性**的说明，
   例如「一名运动员在田径赛场上奔跑」「施工现场的工程车辆」。**禁止**任何主观评价、外貌评价或臆测
   （不得出现"帅气/好看/身材/阳光"等词）；若未提供配图信息，写「原文未提供配图说明」。
2. **`grammar`**：必须正好 6 条，对应 6 个**不同**的句子；**优先**挑对四级考生偏难的考点：非谓语动词（分词/不定式/
   动名词的复杂用法）、独立主格、with 复合结构、复杂定语从句、名词性从句、虚拟语气、倒装、强调句、
   过去完成时、被动语态与情态动词连用、插入语与长句主干识别等；
3. `phrase` 必须是原文中**连续、逐字、一字不差**出现的片段（程序会用它在原文里定位，并自动加粗加下划线），
   长度 40—120 个字符，不要改写、不要用省略号、不要加引号；
4. `note` 用中文，1—2 句说清语法点并**划出句子主干**；
5. `tips` 给 3—4 条中文学习建议。
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
    """生成第二部分：语法短语 **<u>…</u>** + 句末【n】；生词 **word**①。
    语法编号按"在原文中出现的先后"重排，保证【1】—【6】严格按顺序、不错标。"""
    found, missed = [], []
    for g in grammar:
        loc = locate_phrase(paras, g.get("phrase", ""))
        if loc:
            i, s, e = loc
            found.append((i, s, e, g))
        else:
            missed.append(g)
    found.sort(key=lambda x: (x[0], x[1]))          # 按原文位置排序后再统一编号

    spots, infos = {}, []
    for n, (i, s, e, g) in enumerate(found, 1):
        spots.setdefault(i, []).append((s, e, n))
        st, en = sentence_span(paras[i], s)
        infos.append({"n": n, "g": g, "found": True, "para": i + 1,
                      "sent": sentence_no(paras[i], s),
                      "sentence": paras[i][st:en].strip()})
    for k, g in enumerate(missed, len(found) + 1):
        infos.append({"n": k, "g": g, "found": False})
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
    # 配图描述：优先用模型给的客观描述；缺失则直接采用原文图注 / alt（客观文本）
    photo = (gj.get("photo_cn") or "").strip()
    if len(photo) < 4:
        photo = (article.get("photo_caption") or article.get("photo_alt") or "").strip()
        if photo:
            mark = "（正文配图说明）" if article.get("photo_src") == "content" else "（页面配图说明，以网页为准）"
            photo = mark + photo
        else:
            photo = "原文未提供配图说明（可到原文网页查看）"
    head = ["# CET-4 每日精读 · China Daily", "",
            "> 整理日期：%s　|　适用水平：CET-4 高分 / 六级预备（难度偏高）　|　词汇：30 个（含拓展词）" % today,
            "", "## 一、文章信息", "",
            "| 项目 | 内容 |", "| --- | --- |",
            "| 标题 | %s |" % article["title"],
            "| 来源 | China Daily（中国日报英文网） |",
            "| 作者 | %s |" % (article["author"] or "China Daily"),
            "| 发布时间 | %s |" % article["date"],
            "| 网页链接 | %s |" % article["url"],
            "| 体裁 | 新闻报道 / 特写 |",
            "| 主题 | %s |" % (gj.get("topic_cn") or ""),
            "| 配图 | %s |" % photo,
            "", "## 二、原文（含生词 + 语法双重标注）", "",
            "> **标注说明**",
            "> - **加粗**单词后的 **①—㉚**：第三部分「核心词汇与拓展词」（带圈数字）。",
            "> - **加粗并加下划线**短语后的 **【1】—【6】**：第四部分「重点语法与长难句」（方头括号）。",
            "", sec2, ""]

    # 第三部分（每个生词独占一块，字段分行，便于阅读）
    sec3 = ["## 三、核心词汇与拓展词（精选 30 词）", "",
            "> 序号为带圈数字 ①—㉚，与原文词末序号一一对应。",
            "> 选词侧重文化 / 科技 / 教育 / 社会 / 体育 / 环境等话题领域词，难度面向四级高分与六级预备；标 **〔拓展〕** 者超出四级考纲。", ""]
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
    ph_bits = []
    if article.get("photo_caption"):
        ph_bits.append("图注：" + article["photo_caption"])
    if article.get("photo_alt"):
        ph_bits.append("alt：" + article["photo_alt"])
    ph = "　".join(ph_bits) or "无"
    base_user = ("【今天日期】{d}\n【文章标题】{t}\n【作者】{a}\n【发布时间】{p}\n【网页链接】{u}\n"
                 "【配图信息】{ph}\n\n【原文（分段编号）】\n{b}").format(
        d=today, t=article["title"], a=article["author"], p=article["date"],
        u=article["url"], b=numbered, ph=ph)
    words = []
    try:
        wj = llm_json(WORD_PROMPT, base_user, temperature=0.3)
        words = [w for w in (wj.get("words") or []) if isinstance(w, dict) and (w.get("w") or w.get("base"))]
        words = words[:len(CIRCLED)]          # 严格截断到 30 个，防越界
    except Exception:
        words = []
    # 按在原文中首次出现的位置排序 -> 保证 ①—㉚ 的顺序与原文完全一致
    alltext = "\n\n".join(paras)
    def _firstpos(w):
        wd = (w.get("w") or w.get("base") or "").strip()
        if not wd:
            return 10 ** 9
        m = re.search(r"(?<![A-Za-z])" + re.escape(wd) + r"(?![A-Za-z])", alltext)
        return m.start() if m else 10 ** 9
    words.sort(key=_firstpos)
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


# 过于基础的高频词：命中过多说明选词太容易，需要重新生成
BASIC_WORDS = set("""city public family school good make people year work way day time
place thing life hand part group area level order form food water book word world
house room friend country number night point home mother father child woman man
week month long small large little young old right left high low big new now just
only also very well more most many much some any each other after before again
get give take come go see know think want need use find tell ask feel try leave
call move turn show hear play run open close start stop help talk look""".split())


def _heads_from(sec3):
    """从第三部分提取 30 个词头（①—㉚ 后面的单词）。"""
    heads = []
    for c in CIRCLED:
        m = re.search(r"\*\*" + re.escape(c) + r"\s+([A-Za-z][A-Za-z\-']*)", sec3)
        if m:
            heads.append(m.group(1))
    return heads


def validate_md(md, article=None):
    """STEP 7 · 发送前的严格自审：返回问题列表（空列表 = 全部通过，可发送）。"""
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

    sec_all = split_sections(md)
    sec1 = sec_all.get("一", "")
    if "| 配图 |" not in sec1:
        problems.append("文章信息缺「配图」一行")
    # 去掉「> 标注说明」等引用行，避免其中的字样干扰判断
    sec2 = "\n".join(l for l in sec_all.get("二", "").translate(_FW).split("\n")
                     if not l.strip().startswith(">"))
    sec3 = sec_all.get("三", "")
    sec4 = sec_all.get("四", "")
    if not sec2:
        problems.append("未找到第二部分正文")
        return problems

    # ---- ① 原文标注：齐全、不重复、顺序不乱 ----
    n_w2 = sum(1 for c in CIRCLED if c in sec2)
    if n_w2 < 30:
        problems.append("原文里生词标注仅 %d 个（应=30）" % n_w2)
    dup_w = [c for c in CIRCLED if sec2.count(c) > 1]
    if dup_w:
        problems.append("原文生词序号重复：" + "".join(dup_w))
    pos_w = [sec2.find(c) for c in CIRCLED if c in sec2]
    if pos_w != sorted(pos_w):
        problems.append("原文生词序号未按 ①—㉚ 顺序出现")
    miss_g2 = [str(i) for i in range(1, 7) if ("【%s】" % i) not in sec2]
    if miss_g2:
        problems.append("原文缺语法标注【%s】" % "】【".join(miss_g2))
    dup_g = [i for i in range(1, 7) if sec2.count("【%d】" % i) > 1]
    if dup_g:
        problems.append("原文语法序号重复：" + "、".join(str(i) for i in dup_g))
    pos_g = [sec2.find("【%d】" % i) for i in range(1, 7) if ("【%d】" % i) in sec2]
    if pos_g != sorted(pos_g):
        problems.append("原文语法序号未按【1】—【6】顺序出现")
    n_u = sec2.count("<u>")
    if n_u != 6:
        problems.append("原文加粗下划线短语 %d 处（应=6）" % n_u)
    junk = [k for k in ("{", "}", "px;", "function(", "margin:", "var ", "://") if k in sec2]
    if junk:
        problems.append("原文疑似混入代码/乱码：" + " ".join(junk))

    # ---- ② 第三部分：词数、级别、字段、难度 ----
    heads = _heads_from(sec3)
    if len(heads) < 30:
        problems.append("第三部分词条仅 %d 个（应=30）" % len(heads))
    n_ext = sec3.count("〔拓展〕")
    n_core = sec3.count("CET-4 核心")
    if n_ext < 12:
        problems.append("拓展词仅 %d 个（应≥12）" % n_ext)
    if n_core < 5:
        problems.append("四级核心词仅 %d 个（应≥5）" % n_core)
    hits = [w for w in (h.lower() for h in heads) if w in BASIC_WORDS]
    if len(hits) > 5:
        problems.append("选词偏简单（命中基础词 %d 个：%s）" % (len(hits), " ".join(hits[:6])))
    for f in ("词性", "级别", "释义", "例句", "语法与搭配", "词形变化"):
        if ("**%s**" % f) not in sec3:
            problems.append("第三部分缺字段「%s」" % f)

    # 逐词块校验：例句必须包含目标词、音标必须用 /…/ 包裹
    blocks, cur = [], None
    for ln in sec3.split("\n"):
        if re.match(r"^\*\*[" + re.escape(CIRCLED) + r"]", ln):
            if cur:
                blocks.append(cur)
            cur = [ln]
        elif cur is not None:
            cur.append(ln)
    if cur:
        blocks.append(cur)
    for b in blocks:
        hm = re.match(r"^\*\*[" + re.escape(CIRCLED) + r"]\s+([A-Za-z][A-Za-z\-']*)", b[0])
        if not hm:
            continue
        wd = hm.group(1)
        text = "\n".join(b)
        m_eg = re.search(r"\*\*例句\*\*：(.+)", text)
        stem = wd.lower()[:6] if len(wd) >= 6 else wd.lower()
        if m_eg and stem not in m_eg.group(1).lower():
            problems.append("例句未包含目标词：「%s」" % wd)
        m_ph = re.search(r"`([^`]+)`", b[0])
        if m_ph and not (m_ph.group(1).startswith("/") and m_ph.group(1).endswith("/")):
            problems.append("音标格式异常：「%s」" % wd)

    # ---- ③ 第四部分：条数、字段 ----
    if sec4.count("出处：") < 6:
        problems.append("第四部分仅 %d 条（应=6）" % sec4.count("出处："))
    for f in ("原句", "语法点", "讲解"):
        if ("**%s**" % f) not in sec4:
            problems.append("第四部分缺字段「%s」" % f)

    # ---- ④ 原文一致性（忠实于新闻原文）----
    if article:
        ov = _title_overlap(article.get("title", ""), sec2)
        dv = _desc_overlap(article.get("desc", ""), sec2)
        if ov < 0.5 and dv < 0.6:
            problems.append("原文与标题/摘要不吻合（标题%.2f 摘要%.2f）" % (ov, dv))
    return problems


def review_report(md, article):

    """生成审核报告（写入邮件正文，便于每天核对）。"""
    sec = split_sections(md)
    sec2 = sec.get("二", "").translate(_FW)
    sec2 = "\n".join(l for l in sec2.split("\n") if not l.strip().startswith(">"))
    sec3 = sec.get("三", "")
    return {
        "标题": article.get("title", ""),
        "链接": article.get("url", ""),
        "正文字数": len(article.get("body", "")),
        "标题吻合": round(_title_overlap(article.get("title", ""), sec2), 2),
        "摘要吻合": round(_desc_overlap(article.get("desc", ""), sec2), 2),
        "生词数": sum(1 for c in CIRCLED if c in sec2),
        "语法点数": sum(1 for i in range(1, 7) if ("【%d】" % i) in sec2),
        "下划线": sec2.count("<u>"),
        "拓展词": sec3.count("〔拓展〕"),
        "配图": (article.get("photo_caption") or article.get("photo_alt") or "原文未提供配图说明")[:60],
        "总字数": len(md),
    }


def docx_text(docx_bytes):
    """把 .docx 里的可见文本抽出来（用于成品级检查）。"""
    try:
        with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
            xml = z.read("word/document.xml").decode("utf-8", "ignore")
    except Exception:
        return ""
    xml = re.sub(r"</w:p>", "\n", xml)
    xml = re.sub(r"<[^>]+>", "", xml)
    return unescape(xml)


def audit_docx(docx_bytes):
    """成品级审核：Word 成品里不得残留任何 Markdown 记号（标注只以"加粗/下划线"呈现）。"""
    t = docx_text(docx_bytes)
    bad = [s for s in ("**", "`", "<u>", "</u>", "## ") if s in t]
    return ["成品文档残留 Markdown 记号：" + " ".join(bad)] if bad else []



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
        cands = rank_candidates(d, items, 10)
        if not cands:
            return {"ok": False, "msg": "no article found"}
        full, notes = pick_best_article(cands, today)
        if full is None:
            return {"ok": False, "msg": "no usable article", "notes": notes[:5]}
        print("[选文] %s | %s | 发布 %s | 正文 %d 字 | 标题吻合 %.2f | 摘要吻合 %.2f"
              % (full["title"], full["url"], full["date"], len(full["body"]),
                 full.get("match", -1), full.get("desc_match", -1)))

        # STEP 7 · 生成后严格自审（内容清单 + 成品级检查）：不合格就重新生成（最多 3 次）
        md, docx, warns = None, None, None
        for attempt in range(3):
            try:
                md_try = gen_markdown(full)
                docx_try = build_docx(md_try)
            except Exception as e:
                notes.append("第%d次生成失败 %r" % (attempt + 1, e))
                continue
            w_try = validate_md(md_try, full) + audit_docx(docx_try)
            print("[自审] 第 %d 次：问题 %d 个 %s" % (attempt + 1, len(w_try), w_try))
            if md is None or len(w_try) < len(warns):
                md, docx, warns = md_try, docx_try, w_try
            if not w_try:                     # 一次过审
                break
        if md is None:
            return {"ok": False, "msg": "generate failed", "notes": notes[:5]}

        fname = "CET4_ChinaDaily_%s.docx" % today
        try:  # 额外落盘一份（GitHub Actions 会作为 artifact 上传）
            outdir = os.path.join(os.getcwd(), "output")
            os.makedirs(outdir, exist_ok=True)
            with open(os.path.join(outdir, fname), "wb") as f:
                f.write(docx)
        except Exception:
            pass
        rep = review_report(md, full)          # 只写入运行日志（不再随邮件发送）
        print("[审核报告] " + json.dumps(rep, ensure_ascii=False))

        subject = "CET-4 每日精读 · China Daily（%s）| %s" % (today, full["title"])
        body = ("今日 CET-4 精读材料见附件（Word）。\n\n"
                "文章：%s\n来源：China Daily\n发布时间：%s\n链接：%s\n"
                % (full["title"], full["date"], full["url"]))
        if warns:                              # 仅在有问题时提示（正常情况下保持简洁）
            body += "\n（本次自检提示：%s）\n" % "；".join(warns)
        send_mail(docx, fname, subject, body)
        return {"ok": True, "article": full["title"], "url": full["url"],
                "file": fname, "warnings": warns, "report": rep}
    except Exception as e:
        import traceback
        return {"ok": False, "error": repr(e),
                "trace": traceback.format_exc()[-1800:]}


if __name__ == "__main__":
    import sys
    _res = main_handler()
    print(json.dumps(_res, ensure_ascii=False))
    sys.exit(0 if _res.get("ok") else 1)
