#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
云上通城（https://www.hbtctv.cn/）新闻「全量」抓取脚本（全类型版）
==================================================================

用途
----
抓取站点全部历史内容（从最新 ID 一直扫到最早的 ID≈24850）。
全量扫描会经过全站每一个 ID，因此产出的「类型普查」也是最完整的。

收录全部 6 种类型（由页面 shareVariable 的 aid 字段判定，实测）
----------------------------------------------------------------
    aid = 1  图文类      正文 div.article-content                              → 正文文字 + [图片]/[视频]
    aid = 2  组图类      div.picture-content，img[data-img] + data-caption     → 图集内所有图片地址 + 图注
    aid = 3  转外链类    window.onload 跳转到其它网站                          → [外链] 跳转目标地址
    aid = 4  视频类      div.video-content 内 iframe 播放器                    → [视频] mp4 直链（调取流接口）+ [视频封面]
    aid = 11 直播类      div.body-live，直播图文流走 /index/live/query 接口    → [直播页] + 封面 + 摘要 + 直播图文（文字/图/视频）
    aid = 12 报名/表单类 signup 表单页                                        → [报名页] 页面链接

说明：aid=11 直播页在直播结束后页面上没有主播放器（取不到 ts/m3u8 流），
脚本给出直播页链接 + 封面 + 摘要 + 通过 /index/live/query 接口抓到的全部直播图文内容
（其中视频帖会解析成 mp4 直链）。
aid=3（转外链）/12（报名页）页面本身无标题/发布时间，「发布时间」列为空。

输出字段
--------
    序号 | 发布时间 | 标题 | 内容 | 类型 | 类型ID | 原文链接
「内容」内嵌媒体链接格式：
    [图片] https://...      [视频] https://....mp4      [视频封面] https://...
    [外链] https://...      [直播页] https://...        [报名页] https://...

输出文件
--------
    <out>/news_all.html
    <out>/news_all.xlsx                         （需 openpyxl，缺失时退化为 .csv）
    <out>/news_all.records.jsonl                （断点续跑用）
    <out>/news_all.state.json                   （断点续跑用）
    <out>/news_all.aid_census.json              （类型普查结果）

用法示例
--------
    python fetch_all_articles.py                       # 全量抓取（约 9 万个 ID，需要较长时间）
    python fetch_all_articles.py --workers 24          # 提高并发
    python fetch_all_articles.py --report-only         # 只用已有 jsonl 重新生成表格
    中断后重跑同一条命令即可续跑（state 记录扫描前沿）。

注意：脚本默认绕过系统代理直连。robots.txt 建议普通爬虫 Crawl-delay: 2，
若要严格遵守请用 --workers 1 --delay 2。
"""

from __future__ import annotations

import argparse
import csv
import gzip
import html as html_mod
import json
import os
import re
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from html.parser import HTMLParser

# --------------------------------------------------------------------------------------
# 站点常量
# --------------------------------------------------------------------------------------
BASE = "https://www.hbtctv.cn"
ART_URL = BASE + "/p/{aid}.html"
HOME_URL = BASE + "/"
LIVE_QUERY = BASE + "/index/live/query?liveid={liveid}&iscustomerlive=&pass_time=0&contentid={cid}&offset={offset}"
VIDEO_API = "https://app.cjyun.org.cn/video/player/video?sid={sid}&vid={vid}&type=video"
DEFAULT_SID = "10132"
MIN_VALID_ID = 24850            # 实测站点最早的有效文章 ID

# aid → 类型名（实测：全站共 6 种取值）
AID_TYPES = {"1": "图文", "2": "组图", "3": "转外链", "4": "视频",
             "11": "直播", "12": "报名/表单"}

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")

# 优先用 requests（长连接复用，快很多）；缺失时退回 urllib
try:
    import requests as _requests
except ImportError:
    _requests = None

_tls = threading.local()
_print_lock = threading.Lock()


def log(*args):
    with _print_lock:
        print(*args, flush=True)


def _session():
    """线程本地 requests.Session（不走系统代理）。"""
    if _requests is None:
        return None
    s = getattr(_tls, "s", None)
    if s is None:
        s = _requests.Session()
        s.trust_env = False
        s.headers.update({"User-Agent": UA,
                          "Accept-Language": "zh-CN,zh;q=0.9",
                          "Accept": "text/html,application/xhtml+xml,application/json,*/*;q=0.8"})
        _tls.s = s
    return s


# --------------------------------------------------------------------------------------
# HTTP
# --------------------------------------------------------------------------------------
def http_get(url: str, timeout: int = 25, retries: int = 3):
    """返回响应体 bytes；404 返回 None；其它错误重试后返回 b''。"""
    s = _session()
    last = None
    for attempt in range(retries):
        try:
            if s is not None:
                r = s.get(url, timeout=timeout)
                if r.status_code == 404:
                    return None
                if r.status_code != 200:
                    raise RuntimeError("HTTP %d" % r.status_code)
                return r.content
            req = urllib.request.Request(url, headers={
                "User-Agent": UA,
                "Accept": "text/html,application/xhtml+xml,application/json,*/*;q=0.8",
                "Accept-Encoding": "gzip",
            })
            with urllib.request.build_opener(urllib.request.ProxyHandler({})).open(req, timeout=timeout) as resp:
                raw = resp.read()
                if resp.headers.get("Content-Encoding", "").lower() == "gzip":
                    try:
                        raw = gzip.decompress(raw)
                    except Exception:  # noqa: BLE001
                        pass
                return raw
        except urllib.error.HTTPError as e:
            if e.code == 404:
                return None
            last = e
        except Exception as e:  # noqa: BLE001
            last = e
        time.sleep(0.4 * (attempt + 1))
    if last is not None:
        log("  ! 请求失败 %s (%s)" % (url, last))
    return b""


# --------------------------------------------------------------------------------------
# 类型普查（抓取过程中统计所有出现过的 aid 取值）
# --------------------------------------------------------------------------------------
_CENSUS: dict = {}                 # aid -> {"count": int, "samples":[{url,pub,title}...]}
_CENSUS_LOCK = threading.Lock()
CENSUS_SAMPLE_MAX = 60             # 每个 aid 最多留多少个样本链接


def census_add(aid: str, url: str, pub: str = "", title: str = ""):
    aid = str(aid or "?").strip() or "?"
    with _CENSUS_LOCK:
        d = _CENSUS.get(aid)
        if d is None:
            d = {"count": 0, "samples": [], "_urls": set()}
            _CENSUS[aid] = d
        d["count"] += 1
        if title:
            for s in d["samples"]:
                if s["url"] == url and not s.get("title"):
                    s["title"] = title
                    break
        if url not in d["_urls"] and len(d["samples"]) < CENSUS_SAMPLE_MAX:
            d["_urls"].add(url)
            d["samples"].append({"url": url, "pub": pub, "title": title})


def census_load(path: str):
    if not os.path.exists(path):
        return 0
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:  # noqa: BLE001
        return 0
    n = 0
    with _CENSUS_LOCK:
        for aid, d in (data.get("aids") or {}).items():
            cur = _CENSUS.setdefault(aid, {"count": 0, "samples": [], "_urls": set()})
            cur["count"] += int(d.get("count") or 0)
            for s in (d.get("samples") or []):
                if len(cur["samples"]) >= CENSUS_SAMPLE_MAX:
                    break
                if s.get("url") not in cur["_urls"]:
                    cur["_urls"].add(s.get("url"))
                    cur["samples"].append(s)
            n += 1
    return n


def census_save(path: str, extra: dict):
    with _CENSUS_LOCK:
        aids = {}
        for aid, d in sorted(_CENSUS.items(), key=lambda kv: kv[0]):
            aids[aid] = {"count": d["count"],
                         "type": AID_TYPES.get(aid, "未知"),
                         "samples": d["samples"]}
    out = dict(extra)
    out["aids"] = aids
    with open(path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    return out


def census_log_summary():
    with _CENSUS_LOCK:
        items = sorted(_CENSUS.items(), key=lambda kv: kv[0])
    if not items:
        return
    total = sum(d["count"] for _, d in items)
    log("")
    log("-" * 78)
    log("类型普查（抓取过程中统计到 %d 个正文页）" % total)
    log("-" * 78)
    for aid, d in items:
        log("   aid=%-4s %-7s %6d 篇   %s"
            % (aid, AID_TYPES.get(aid, "未知"), d["count"], "✔ 收录" if aid in AID_TYPES else "✘ 未知取值"))
    odd = [(a, d) for a, d in items if a not in AID_TYPES]
    if odd:
        for aid, d in odd:
            log("   → 发现 AID_TYPES 之外的新取值 aid=%s，共 %d 篇，示例：" % (aid, d["count"]))
            for s in d["samples"][:20]:
                log("        %s   %s" % (s["url"], (s.get("title") or "")[:40]))
    log("")


# --------------------------------------------------------------------------------------
# 正文抽取
# --------------------------------------------------------------------------------------
BLOCK_TAGS = {"p", "div", "br", "li", "h1", "h2", "h3", "h4", "h5",
              "tr", "td", "section", "figure", "blockquote", "article"}


class BodyParser(HTMLParser):
    """在正文 HTML 片段中按文档顺序收集 文字 / 图片 / 视频。"""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.events = []
        self._skip = 0

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag in ("script", "style"):
            self._skip += 1
            return
        if self._skip:
            return
        if tag == "img":
            url = (a.get("data-origin-src") or a.get("data-img") or a.get("src") or "").strip()
            if url:
                self.events.append(("img", url))
            cap = (a.get("data-caption") or "").strip()
            if cap:
                self.events.append(("text", cap))
                self.events.append(("nl", ""))
        elif tag in ("iframe", "video", "audio"):
            url = (a.get("src") or "").strip()
            if url:
                self.events.append(("media", url))
        elif tag in BLOCK_TAGS:
            self.events.append(("nl", ""))

    def handle_startendtag(self, tag, attrs):
        self.handle_starttag(tag, attrs)

    def handle_endtag(self, tag):
        if tag in ("script", "style"):
            if self._skip:
                self._skip -= 1
            return
        if self._skip:
            return
        if tag in BLOCK_TAGS:
            self.events.append(("nl", ""))

    def handle_data(self, data):
        if self._skip:
            return
        if data and data.strip():
            self.events.append(("text", data))


def extract_body_html(page: str):
    """截出正文容器（article-content / video-content）所在片段。"""
    i = page.find('class="article-content"')
    if i < 0:
        i = page.find('class="video-content"')
    if i < 0:
        return None
    start = page.rfind("<div", 0, i)
    if start < 0:
        start = i
    end = page.find("</article>", i)
    if end < 0:
        for marker in ('<div class="sharebox"', '<div class="article">', "</body>"):
            end = page.find(marker, i)
            if end > 0:
                break
    if end <= start:
        end = min(len(page), start + 300000)
    return page[start:end]


def _attr(tag: str, name: str) -> str:
    m = (re.search(r'\b' + name + r'\s*=\s*"([^"]*)"', tag, re.I)
         or re.search(r"\b" + name + r"\s*=\s*'([^']*)'", tag, re.I))
    return html_mod.unescape(m.group(1)).strip() if m else ""


LOADING_GIF = "loading.gif"


def extract_gallery_events(page: str):
    """组图页（aid=2）：按顺序取 大图(+图注)。"""
    k = page.find('id="big-list"')
    if k < 0:
        k = page.find('class="big-images"')
    if k < 0:
        k = page.find('class="picture-content"')
    if k < 0:
        return []
    j = page.find("</ul>", k)
    seg = page[k:(j if j > 0 else k + 200000)]

    events = []
    for tag in re.findall(r"<img\b[^>]*>", seg, re.I):
        url = _attr(tag, "data-img") or _attr(tag, "data-origin-src") or _attr(tag, "src")
        if not url or LOADING_GIF in url:
            continue
        events.append(("img", url))
        cap = _attr(tag, "data-caption")
        if cap:
            events.append(("text", cap))
            events.append(("nl", ""))
    return events


TITLE_RE = re.compile(
    r'<h1[^>]*class="[^"]*(?:article-title|video-title|picture-title)[^"]*"[^>]*>(.*?)</h1>', re.S)
LIVE_TITLE_RE = re.compile(r'<h1[^>]*class="title"[^>]*>(.*?)</h1>', re.S)
GENERIC_H1_RE = re.compile(r"<h1(?![^>]*class=\"lo)[^>]*>(.*?)</h1>", re.S)
TIME_RE = re.compile(r'<time[^>]*class="[^"]*date[^"]*"[^>]*>(.*?)</time>', re.S)
TIME_ANY_RE = re.compile(r"<time[^>]*>(.*?)</time>", re.S)
DOCTITLE_RE = re.compile(r"<title>(.*?)</title>", re.S)
AID_RE = re.compile(r"aid\s*:\s*'(\d+)'")
SID_RE = re.compile(r"sid\s*:\s*'(\d+)'")
LIVEID_RE = re.compile(r"liveid\s*:\s*'(\d+)'")
LIVE_COVER_RE = re.compile(r'section-cover"\s+data-image="([^"]+)"', re.S)
LIVE_SUMMARY_RE = re.compile(
    r'<time[^>]*class="[^"]*date[^"]*"[^>]*>.*?</time>\s*<p[^>]*>(.*?)</p>', re.S)
REDIRECT_RE = re.compile(r'location\.href\s*=\s*["\']([^"\']+)["\']')
DATE_RE = re.compile(r"(\d{4}-\d{2}-\d{2})(?:\s+(\d{2}:\d{2}))?")


def strip_tags(s: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", s or "")).strip()


def extract_title(page: str) -> str:
    m = TITLE_RE.search(page) or LIVE_TITLE_RE.search(page)
    if m:
        t = strip_tags(m.group(1))
        if t:
            return t
    for hm in GENERIC_H1_RE.finditer(page):
        t = strip_tags(hm.group(1))
        if t:
            return t
    m = DOCTITLE_RE.search(page)
    if m:
        return re.sub(r"[\s_\-|]*云上通城\s*$", "", strip_tags(m.group(1))).strip()
    return ""


def extract_pub(page: str) -> str:
    m = TIME_RE.search(page) or TIME_ANY_RE.search(page)
    pub = strip_tags(m.group(1)) if m else ""
    if not re.match(r"\d{4}-\d{2}-\d{2}", pub):
        m = DATE_RE.search(strip_tags(page[:40000]))
        pub = m.group(0) if m else ""
    return pub


def fetch_live_posts(liveid: str, cid: int) -> list:
    """直播页（aid=11）：通过 /index/live/query 拉取直播图文条目。

    实测该接口固定返回最新一批（约 22 条），offset / pass_time 翻页参数均无效，
    故只请求一次并按帖子 id 去重。
    每条含 文字 / images[].url / video(player或url) / audio(url)。
    返回 events：("text", ...)/("img", ...)/("media", ...)。
    """
    events = []
    if not liveid:
        return events
    raw = http_get(LIVE_QUERY.format(liveid=liveid, cid=cid, offset=0))
    if not raw:
        return events
    try:
        data = json.loads(raw.decode("utf-8", "replace"))
    except Exception:  # noqa: BLE001
        return events
    posts = data.get("data") or []
    if not isinstance(posts, list):
        return events
    seen = set()
    for p in posts:
        if not isinstance(p, dict):
            continue
        pid = str(p.get("id") or "")
        if pid and pid in seen:
            continue
        seen.add(pid)
        dt = str(p.get("datetime") or "").strip()
        txt = strip_tags(str(p.get("text") or ""))
        events.append(("text", ("[%s] " % dt if dt else "") + txt))
        for im in (p.get("images") or []):
            u = (im or {}).get("url") or ""
            if u:
                events.append(("img", u))
        v = p.get("video") or {}
        vu = (v.get("url") or v.get("player") or "").strip()
        if vu:
            events.append(("media", html_mod.unescape(vu)))
        a = p.get("audio") or {}
        au = (a.get("url") or "").strip()
        if au:
            events.append(("media", html_mod.unescape(au)))
    return events


def parse_article(page: str, art_id: int):
    """解析详情页 → 记录 dict；无效页 / 未知类型且无正文返回 None。

    收录 aid ∈ {1,2,3,4,11,12}；其余未知 aid 一律跳过。
    """
    if not page:
        return None

    aid_m = AID_RE.search(page)
    aid = aid_m.group(1) if aid_m else ""
    if aid and aid not in AID_TYPES:
        return None                              # 未知类型，跳过（普查已另行记录）

    has_article = 'class="article-content"' in page
    has_video = 'class="video-content"' in page
    has_picture = 'class="picture-content"' in page

    title = extract_title(page)
    pub = extract_pub(page)
    events = []

    if aid == "3":                               # 转外链：取跳转目标
        m = REDIRECT_RE.search(page)
        if not m:
            return None
        events = [("link", html_mod.unescape(m.group(1)).strip())]

    elif aid == "11":                            # 直播：直播页 + 封面 + 摘要 + 直播图文流
        events = [("link", ART_URL.format(aid=art_id))]
        m = LIVE_COVER_RE.search(page)
        if m:
            events.append(("img", m.group(1).strip()))
        m = LIVE_SUMMARY_RE.search(page)
        if m:
            t = strip_tags(m.group(1))
            if t:
                events.append(("text", t))
                events.append(("nl", ""))
        lm = LIVEID_RE.search(page)
        events += fetch_live_posts(lm.group(1) if lm else "", art_id)
        if len(events) <= 1:
            return None                          # 连直播页都拿不到内容意义的空壳

    elif aid == "12":                            # 报名/表单：给出报名页链接
        if not title:
            return None
        events = [("link", ART_URL.format(aid=art_id))]

    elif has_picture or aid == "2":              # 组图
        events = extract_gallery_events(page)
        if not events:
            body_html = extract_body_html(page)
            if body_html:
                bp = BodyParser()
                try:
                    bp.feed(body_html)
                except Exception:  # noqa: BLE001
                    pass
                events = bp.events
        if not events:
            return None

    else:                                        # 图文 / 视频（含未知但有正文的情况）
        body_html = extract_body_html(page)
        if body_html:
            bp = BodyParser()
            try:
                bp.feed(body_html)
            except Exception:  # noqa: BLE001
                pass
            events = bp.events
        if not events:
            return None

    if not title and aid != "3":                 # 转外链页无标题，允许为空
        return None

    type_name = AID_TYPES.get(aid)
    if type_name is None:
        type_name = "组图" if has_picture else ("视频" if has_video else "图文")

    sid_m = SID_RE.search(page)
    return {
        "id": art_id,
        "aid": aid or {"组图": "2", "视频": "4"}.get(type_name, "1"),
        "sid": sid_m.group(1) if sid_m else DEFAULT_SID,
        "type": type_name,
        "pub": pub,
        "title": title,
        "events": events,
        "url": ART_URL.format(aid=art_id),
    }


# --------------------------------------------------------------------------------------
# 组装「内容」文本
# --------------------------------------------------------------------------------------
VID_RE = re.compile(r"[?&]vid=(\d+)")
_video_cache: dict = {}
_video_lock = threading.Lock()
DROP_IMG_KEYWORDS: list = []

LINK_LABEL = {"转外链": "外链", "报名/表单": "报名页", "直播": "直播页"}


def resolve_video(iframe_url: str, sid: str, use_api: bool) -> dict:
    """把播放页 iframe 换成 mp4 直链 + 封面图。带缓存。"""
    out = {"page": iframe_url, "mp4": "", "cover": ""}
    m = VID_RE.search(iframe_url)
    if not m:
        return out
    vid = m.group(1)
    with _video_lock:
        if vid in _video_cache:
            return _video_cache[vid]

    qs = urllib.parse.parse_qs(urllib.parse.urlparse(html_mod.unescape(iframe_url)).query)
    if qs.get("thumb"):
        out["cover"] = urllib.parse.unquote(qs["thumb"][0])

    if use_api:
        raw = http_get(VIDEO_API.format(sid=sid or DEFAULT_SID, vid=vid))
        if raw:
            try:
                data = json.loads(raw.decode("utf-8", "replace"))
                f = data.get("file") or {}
                mp4 = f.get("sd") or f.get("hd") or f.get("ed") or f.get("url") or ""
                if mp4:
                    out["mp4"] = mp4
                if not out["cover"] and data.get("image"):
                    img = data["image"]
                    out["cover"] = ("https://img.cjyun.org" + img) if img.startswith("/") else img
            except Exception:  # noqa: BLE001
                pass

    with _video_lock:
        _video_cache[vid] = out
    return out


def build_content(events, sid: str, use_video_api: bool, type_name: str = "") -> str:
    parts, buf = [], []

    def flush():
        if not buf:
            return
        raw = "".join(buf)
        buf.clear()
        for line in raw.split("\n"):
            line = re.sub(r"[ \t\u00a0\u3000]+", " ", line).strip()
            if line:
                parts.append(("text", line))

    for kind, val in events:
        if kind == "text":
            buf.append(val)
        elif kind == "nl":
            buf.append("\n")
        elif kind == "img":
            flush()
            parts.append(("img", val))
        elif kind == "media":
            flush()
            parts.append(("video", val))
        elif kind == "link":
            flush()
            parts.append(("link", val))
    flush()

    lines = []
    for kind, val in parts:
        if kind == "text":
            lines.append(val)
        elif kind == "img":
            if any(k and k in val for k in DROP_IMG_KEYWORDS):
                continue
            lines.append("[图片] " + val)
        elif kind == "link":
            lines.append("[%s] %s" % (LINK_LABEL.get(type_name, "链接"), val))
        else:
            info = resolve_video(val, sid, use_video_api)
            lines.append("[视频] " + (info["mp4"] or info["page"]))
            if info.get("cover"):
                lines.append("[视频封面] " + info["cover"])
    out = "\n".join(lines).strip()
    if len(out) > 32000:                         # Excel 单元格上限 32767 字符
        out = out[:32000].rstrip() + "\n……（内容过长已截断）"
    return out


# --------------------------------------------------------------------------------------
# ID 探测
# --------------------------------------------------------------------------------------
def discover_max_id(gap: int = 3) -> int:
    seed = 0
    raw = http_get(HOME_URL)
    if raw:
        ids = [int(x) for x in re.findall(r"/p/(\d+)\.html", raw.decode("utf-8", "replace"))]
        if ids:
            seed = max(ids)
    if not seed:
        seed = 110000
    log("首页最新文章 ID：%d" % seed)
    cur, miss = seed, 0
    while miss < gap and cur < seed + 3000:
        cur += 1
        if http_get(ART_URL.format(aid=cur)) is not None:
            seed, miss = cur, 0
        else:
            miss += 1
    return seed


# --------------------------------------------------------------------------------------
# 主流程
# --------------------------------------------------------------------------------------
def run(args):
    out_dir = os.path.abspath(args.out)
    os.makedirs(out_dir, exist_ok=True)

    prefix = "news_all"
    records_path = os.path.join(out_dir, prefix + ".records.jsonl")
    state_path = os.path.join(out_dir, prefix + ".state.json")
    census_path = os.path.join(out_dir, prefix + ".aid_census.json")

    done_ids, records = set(), []
    if os.path.exists(records_path):
        with open(records_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    rec = json.loads(line)
                except Exception:  # noqa: BLE001
                    continue
                records.append(rec)
                done_ids.add(rec["id"])
        log("断点续跑：已缓存 %d 篇" % len(records))

    n_old = census_load(census_path)
    if n_old:
        log("已并入上次的类型普查结果（%d 种 aid）" % n_old)

    if not args.report_only:
        state = {}
        if os.path.exists(state_path) and not args.rescan:
            try:
                with open(state_path, "r", encoding="utf-8") as f:
                    state = json.load(f) or {}
            except Exception:  # noqa: BLE001
                state = {}

        max_id = args.start_id or discover_max_id()
        stats = {"scan": 0, "in": 0, "err": 0, "nodate": 0}
        type_stat: dict = {}
        jsonl_f = open(records_path, "a", encoding="utf-8")
        jsonl_lock = threading.Lock()
        failed = set(int(x) for x in state.get("failed", []) if str(x).isdigit())
        t0 = time.time()

        def scan_batch(ids):
            ids = [i for i in ids if i not in done_ids]
            if not ids:
                return
            results = []
            with ThreadPoolExecutor(max_workers=args.workers) as ex:
                futs = {ex.submit(work_one, i, args): i for i in ids}
                for fut in as_completed(futs):
                    results.append(fut.result())
                    if args.delay:
                        time.sleep(args.delay / max(args.workers, 1))

            for kind, art_id, rec, aid, pub in results:
                stats["scan"] += 1
                if kind == "miss":
                    continue
                if kind == "err":
                    stats["err"] += 1
                    failed.add(art_id)
                    continue
                if kind == "skip" or rec is None:
                    continue
                failed.discard(art_id)
                type_stat[rec["type"]] = type_stat.get(rec["type"], 0) + 1
                if not rec.get("pub"):
                    stats["nodate"] += 1
                with jsonl_lock:
                    records.append(rec)
                    done_ids.add(rec["id"])
                    jsonl_f.write(json.dumps(rec, ensure_ascii=False) + "\n")
                    if stats["in"] % 200 == 0:
                        jsonl_f.flush()
                stats["in"] += 1

        def save_state(frontier):
            old = int(state.get("frontier") or 0) or None
            fr = frontier if old is None else min(old, frontier)
            cur_state = {"frontier": fr, "max_id": max(int(state.get("max_id") or 0), max_id),
                         "min_id": args.min_id, "failed": sorted(failed)}
            with open(state_path, "w", encoding="utf-8") as f:
                json.dump(cur_state, f, ensure_ascii=False)
            return cur_state

        log("开始全量抓取：workers=%d 批次=%d ID 范围 %d → %d"
            % (args.workers, args.batch, max_id, args.min_id))

        # —— 优先补抓：上次运行之后新发布的 ID + 上次出错的 ID ——
        pending = []
        if state:
            last_max = int(state.get("max_id") or 0)
            if max_id > last_max:
                pending += list(range(max_id, last_max, -1))
            pending += sorted(failed, reverse=True)
        pending = [i for i in dict.fromkeys(pending) if i not in done_ids]
        if pending:
            log("补抓新增/失败重试：%d 个 ID" % len(pending))
            scan_batch(pending)

        # —— 从上次断点或最新 ID 向下，一直扫到最早 ID ——
        if state and state.get("frontier") and not args.rescan:
            cur = int(state["frontier"]) - 1
            log("从上次断点 ID=%d 继续向下扫描" % int(state["frontier"]))
        else:
            cur = max_id

        while cur >= args.min_id:
            low = max(cur - args.batch + 1, args.min_id)
            scan_batch(list(range(cur, low - 1, -1)))
            save_state(low - 1)

            el = max(time.time() - t0, 0.001)
            eta = (cur - args.min_id) / max(stats["scan"] / el, 0.1) / 60.0
            log("  已扫描 %d 个 ID | 命中 %d 篇 | 无时间已收 %d | 出错 %d"
                " | %.1f 个/秒 | 当前 ID=%d | 剩余约 %.0f 分钟"
                % (stats["scan"], stats["in"], stats["nodate"], stats["err"],
                   stats["scan"] / el, cur, eta))
            cur -= args.batch

        jsonl_f.close()
        log("全量抓取结束：扫描 %d 个 ID，命中 %d 篇。" % (stats["scan"], stats["in"]))
        if type_stat:
            log("  类型分布：" + "、".join("%s %d" % (k, v) for k, v in sorted(type_stat.items())))
        if stats["err"]:
            log("  提示：%d 个 ID 请求出错，可重新运行本脚本补抓。" % stats["err"])

    census_log_summary()
    census_save(census_path, {
        "range": "%d ~ latest" % args.min_id,
        "generated": time.strftime("%Y-%m-%d %H:%M:%S"),
    })
    log("  ✓ 类型普查 → %s" % census_path)

    render(records, out_dir, prefix, args)


def work_one(art_id: int, args):
    raw = http_get(ART_URL.format(aid=art_id))
    if raw is None:
        return "miss", art_id, None, None, ""
    if not raw:
        return "err", art_id, None, None, ""
    page = raw.decode("utf-8", "replace")
    url = ART_URL.format(aid=art_id)

    aid_m = AID_RE.search(page)
    aid = aid_m.group(1) if aid_m else "?"
    pub = extract_pub(page)

    rec = parse_article(page, art_id)
    census_add(aid, url, pub, rec["title"] if rec else "")
    if rec is None:
        return "skip", art_id, None, aid, pub
    rec["content"] = build_content(rec.pop("events"), rec["sid"], not args.no_video_api, rec["type"])
    return "ok", art_id, rec, aid, pub


# --------------------------------------------------------------------------------------
# 输出
# --------------------------------------------------------------------------------------
HEADERS = ["序号", "发布时间", "标题", "内容", "类型", "类型ID", "原文链接"]
MEDIA_RE = re.compile(r"^\[([^\]]{1,4})\]\s*(https?://\S+)\s*$")
MEDIA_CLS = {"图片": "m-img", "视频": "m-video", "视频封面": "m-cover",
             "外链": "m-link", "直播页": "m-link", "报名页": "m-link"}
TYPE_CLS = {"图文": "t-text", "组图": "t-gallery", "视频": "t-video",
            "转外链": "t-link", "直播": "t-live", "报名/表单": "t-signup"}
TYPE_ORDER = ["图文", "组图", "视频", "转外链", "直播", "报名/表单"]


def content_to_html(content: str) -> str:
    out = []
    for line in (content or "").split("\n"):
        m = MEDIA_RE.match(line.strip())
        if m:
            label, url = m.group(1), m.group(2)
            cls = MEDIA_CLS.get(label, "")
            out.append('<div class="media %s">[%s] <a href="%s" target="_blank" rel="noopener">'
                       "%s</a></div>" % (cls, html_mod.escape(label), html_mod.escape(url, quote=True),
                                         html_mod.escape(url)))
        else:
            out.append("<p>%s</p>" % html_mod.escape(line))
    return "\n".join(out)


def render(records, out_dir, prefix, args):
    records = sorted(records,
                     key=lambda r: (r.get("pub") or "0000-00-00 00:00", r.get("id") or 0),
                     reverse=(args.sort == "desc"))
    for i, r in enumerate(records, 1):
        r["seq"] = i
    log("共 %d 条记录，生成输出文件…" % len(records))

    rows = []
    for r in records:
        tname = r.get("type") or ""
        rows.append(
            "<tr>"
            '<td class="c-seq">%d</td>'
            '<td class="c-time">%s</td>'
            '<td class="c-title">%s</td>'
            '<td class="c-content">%s</td>'
            '<td class="c-type"><span class="tag %s">%s</span></td>'
            '<td class="c-aid">%s</td>'
            '<td class="c-url"><a href="%s" target="_blank" rel="noopener">原文</a></td>'
            "</tr>" % (
                r["seq"], html_mod.escape(r.get("pub") or ""),
                html_mod.escape(r.get("title") or ""), content_to_html(r.get("content") or ""),
                TYPE_CLS.get(tname, "t-text"), html_mod.escape(tname),
                html_mod.escape(str(r.get("aid") or "")),
                html_mod.escape(r.get("url") or "", quote=True)))

    type_cnt = {}
    for r in records:
        type_cnt[r.get("type") or "未知"] = type_cnt.get(r.get("type") or "未知", 0) + 1
    stat_txt = "、".join("%s <b>%d</b>" % (t, type_cnt[t]) for t in TYPE_ORDER if type_cnt.get(t))
    pub_times = [r.get("pub") or "" for r in records if r.get("pub")]
    span = "%s ~ %s" % (min(pub_times), max(pub_times)) if pub_times else "—"

    html_path = os.path.join(out_dir, prefix + ".html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(HTML_TMPL.format(
            n=len(records), stats=stat_txt, span=html_mod.escape(span),
            now=time.strftime("%Y-%m-%d %H:%M:%S"),
            order="倒序（最新在前）" if args.sort == "desc" else "正序（最早在前）",
            heads="".join("<th>%s</th>" % h for h in HEADERS),
            rows="\n".join(rows)))
    log("  ✓ HTML  → %s (%.1f MB)" % (html_path, os.path.getsize(html_path) / 1048576.0))

    write_sheet(records, out_dir, prefix)


HTML_TMPL = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>云上通城 · 全站新闻（{n} 条）</title>
<style>
  :root {{ --line:#e3e6eb; --head:#1a4b8c; --bg:#f5f7fa; }}
  * {{ box-sizing:border-box; }}
  body {{ margin:0; background:var(--bg); color:#1c2430;
         font:14px/1.7 "Microsoft YaHei","PingFang SC",system-ui,sans-serif; }}
  header {{ position:sticky; top:0; z-index:10; background:#fff;
            border-bottom:1px solid var(--line); padding:12px 20px; }}
  h1 {{ margin:0 0 5px; font-size:18px; color:var(--head); }}
  .meta {{ color:#5b6675; font-size:13px; }}
  .meta b {{ color:#1c2430; }}
  main {{ padding:16px 20px 60px; }}
  table {{ width:100%; border-collapse:collapse; background:#fff; border:1px solid var(--line); }}
  thead th {{ position:sticky; top:70px; z-index:9; background:#eef2f8; color:var(--head);
              font-size:13px; text-align:left; padding:10px 12px;
              border-bottom:1px solid var(--line); white-space:nowrap; }}
  td {{ padding:10px 12px; border-bottom:1px solid #eef1f5; vertical-align:top; }}
  tr:hover td {{ background:#fafcff; }}
  .c-seq {{ width:56px; color:#8492a6; }}
  .c-time {{ width:132px; white-space:nowrap; color:#42506a; }}
  .c-title {{ width:280px; font-weight:600; }}
  .c-content p {{ margin:0 0 6px; }}
  .media {{ font-size:12.5px; word-break:break-all; }}
  .media a {{ color:#1a6fd4; text-decoration:none; }}
  .media a:hover {{ text-decoration:underline; }}
  .m-img {{ color:#7a5c1e; }} .m-video {{ color:#b03a2e; }} .m-cover {{ color:#6c757d; }}
  .m-link {{ color:#1a6fd4; }}
  .c-type {{ width:72px; }}
  .tag {{ display:inline-block; padding:1px 8px; border-radius:10px; font-size:12px; white-space:nowrap; }}
  .t-text {{ background:#e8f0fb; color:#1a4b8c; }}
  .t-gallery {{ background:#eaf6ec; color:#1f7a3d; }}
  .t-video {{ background:#fdecea; color:#b03a2e; }}
  .t-link {{ background:#ede7f6; color:#5e35b1; }}
  .t-live {{ background:#fff3e0; color:#c05621; }}
  .t-signup {{ background:#e0f2f1; color:#00695c; }}
  .c-aid {{ width:62px; text-align:center; color:#5b6675; }}
  .c-url {{ width:64px; white-space:nowrap; }}
  .c-url a {{ color:#1a6fd4; text-decoration:none; }}
</style>
</head>
<body>
<header>
  <h1>云上通城（hbtctv.cn）全站新闻数据</h1>
  <div class="meta">共 <b>{n}</b> 条 &nbsp;|&nbsp; {stats} &nbsp;|&nbsp; 时间跨度 {span} &nbsp;|&nbsp; 生成时间 {now} &nbsp;|&nbsp; 排序：{order}</div>
</header>
<main>
<table>
<thead><tr>{heads}</tr></thead>
<tbody>
{rows}
</tbody>
</table>
</main>
</body>
</html>
"""


def write_sheet(records, out_dir, prefix):
    xlsx_path = os.path.join(out_dir, prefix + ".xlsx")
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError:
        csv_path = os.path.join(out_dir, prefix + ".csv")
        with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(HEADERS)
            for r in records:
                w.writerow(row_values(r))
        log("  ! 未安装 openpyxl，已改为输出 CSV → %s" % csv_path)
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "新闻数据"
    head_font = Font(bold=True, color="FFFFFF", size=11)
    head_fill = PatternFill("solid", fgColor="1A4B8C")
    head_align = Alignment(horizontal="center", vertical="center")
    for c, name in enumerate(HEADERS, 1):
        cell = ws.cell(row=1, column=c, value=name)
        cell.font = head_font
        cell.fill = head_fill
        cell.alignment = head_align

    for r in records:
        ws.append(row_values(r))

    wrap = Alignment(wrap_text=True, vertical="top")
    ctr = Alignment(horizontal="center", vertical="top")
    link_col = HEADERS.index("原文链接") + 1
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(HEADERS)):
        for cell in row:
            col = cell.column
            if col == HEADERS.index("内容") + 1:
                cell.alignment = wrap
            elif col in (1, 2, HEADERS.index("类型") + 1, HEADERS.index("类型ID") + 1):
                cell.alignment = ctr
            else:
                cell.alignment = Alignment(vertical="top")
            if col == link_col and cell.value:
                cell.hyperlink = cell.value
                cell.font = Font(color="1A6FD4", underline="single", size=10)

    for col, width in zip(range(1, len(HEADERS) + 1), [6, 18, 42, 90, 10, 8, 26]):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(HEADERS)), ws.max_row)

    write_census_sheets(wb, head_font, head_fill, head_align, get_column_letter)

    wb.save(xlsx_path)
    log("  ✓ Excel → %s (%.1f MB)" % (xlsx_path, os.path.getsize(xlsx_path) / 1048576.0))


_ILLEGAL_XLSX_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")


def _xlsx_clean(v):
    """Excel 单元格不允许出现 XML 非法控制字符，写入前清洗。"""
    if isinstance(v, str):
        return _ILLEGAL_XLSX_RE.sub("", v)
    return v


def row_values(r):
    return [_xlsx_clean(v) for v in
            [r["seq"], r.get("pub", ""), r.get("title", ""), r.get("content", ""),
             r.get("type", ""), r.get("aid", ""), r.get("url", "")]]


def write_census_sheets(wb, head_font, head_fill, head_align, get_column_letter):
    """把类型普查结果写进 Excel：「类型普查」工作表。"""
    from openpyxl.styles import Alignment, Font
    with _CENSUS_LOCK:
        data = {a: {"count": d["count"], "samples": list(d["samples"])} for a, d in _CENSUS.items()}
    if not data:
        return

    ws = wb.create_sheet("类型普查")
    heads = ["类型ID", "类型名称", "出现次数", "内容提取方式", "页面示例（前 20 个链接）"]
    for c, name in enumerate(heads, 1):
        cell = ws.cell(row=1, column=c, value=name)
        cell.font = head_font
        cell.fill = head_fill
        cell.alignment = head_align
    HOW = {"1": "正文文字+图片/视频", "2": "图集全部图片+图注", "3": "跳转目标外链",
           "4": "取流接口 mp4 直链+封面", "11": "直播页+封面+摘要+直播图文流",
           "12": "报名页链接"}
    for aid in sorted(data, key=lambda x: (len(x), x)):
        d = data[aid]
        ws.append([aid, AID_TYPES.get(aid, "未知"), d["count"], HOW.get(aid, ""),
                   "\n".join(s["url"] for s in d["samples"][:20])])
    for col, width in zip(range(1, 6), [10, 12, 12, 30, 70]):
        ws.column_dimensions[get_column_letter(col)].width = width
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=5):
        row[3].alignment = Alignment(wrap_text=True, vertical="top")
        row[4].alignment = Alignment(wrap_text=True, vertical="top")
        for cell in row[:4]:
            if cell.value:
                cell.font = Font(size=10)


# --------------------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(
        description="云上通城（hbtctv.cn）新闻全量抓取：收录全部 6 种类型"
                    "（图文/组图/视频/转外链/直播/报名表单）",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default="output", help="输出目录（默认 output）")
    ap.add_argument("--start-id", type=int, default=0, help="起始文章 ID（默认自动探测站点最新 ID）")
    ap.add_argument("--min-id", type=int, default=MIN_VALID_ID,
                    help="扫到的最早 ID（默认 %d，即站点最早有效文章）" % MIN_VALID_ID)
    ap.add_argument("--workers", type=int, default=8, help="并发线程数（默认 8）")
    ap.add_argument("--batch", type=int, default=300, help="每批处理的 ID 个数（默认 300）")
    ap.add_argument("--max-scan", type=int, default=0, help="最多扫描多少个 ID（0=不限，试跑用）")
    ap.add_argument("--delay", type=float, default=0.02, help="请求间隔秒数（默认 0.02）")
    ap.add_argument("--no-video-api", action="store_true", help="不查取流接口，视频只给播放页地址")
    ap.add_argument("--drop-img-keyword", action="append", default=[], metavar="KW",
                    help="图片 URL 含该关键字则丢弃（可重复）")
    ap.add_argument("--sort", choices=["desc", "asc"], default="desc", help="排序：desc 最新在前（默认）")
    ap.add_argument("--report-only", action="store_true", help="只用已有 jsonl 重新生成 html/xlsx")
    ap.add_argument("--rescan", action="store_true", help="忽略断点状态，从头重新扫描")
    args = ap.parse_args()

    DROP_IMG_KEYWORDS.extend(args.drop_img_keyword)

    log("=" * 78)
    log("云上通城 · 新闻全量抓取（全类型版）")
    log("=" * 78)
    run(args)
    log("完成。")


if __name__ == "__main__":
    main()
