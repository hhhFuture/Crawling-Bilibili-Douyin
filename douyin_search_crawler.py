# -*- coding: utf-8 -*-
r"""
抖音关键词搜索爬虫（独立命令行版，不依赖 Streamlit）

功能：模拟人工在搜索框输入关键词，点击“筛选”并选择“最多点赞”，保留平台
      搜索结果且不做相关性过滤，按点赞数降序生成点赞榜/播放榜，默认仅保存 Excel，
      并支持可选视频下载、扫码登录及浏览器登录状态复用。
      遇到安全验证时需在打开的浏览器中人工完成。

首次使用：
  python -m pip install playwright openpyxl yt-dlp imageio-ffmpeg
默认优先使用本机 Chrome，不可用时自动尝试 Microsoft Edge；
若将 BROWSER_CHANNEL 改为 None，则需再执行：
  python -m playwright install chromium

命令行参数：
  -k, --keyword       搜索关键词；默认读取配置区 KEYWORD。
  -o, --output        Excel 保存目录，也是默认视频根目录；默认 F:\Nabby。
  -n, --top-n         点赞榜、播放榜分别保留前 N 条；必须是正整数，默认 30。
  --max-scrolls       最多点赞结果页的滚动轮数；必须是正整数，数值越大覆盖越广。
  --login             清除旧登录状态并强制重新扫码；不能与 --no-login、--headless 同用。
  --no-login          不询问是否扫码，直接复用当前浏览器会话；不会主动清除已有 Cookie。
  --headless          使用无界面浏览器；出现扫码或安全验证时需要去掉此参数。
  --download          下载点赞榜和播放榜合并去重后的视频；不传则只导出表格。
  --video-output      自定义视频保存目录；仅配合 --download 使用。

功能测试命令（先进入脚本所在目录，并执行 conda activate py310）：
  python douyin_search_crawler.py --help
  python douyin_search_crawler.py -k "钥匙扣相机" --no-login --max-scrolls 1
  python douyin_search_crawler.py -k "钥匙扣相机" -n 5 -o "F:\Nabby\测试输出" --no-login --max-scrolls 1
  python douyin_search_crawler.py -k "钥匙扣相机" --login --max-scrolls 1
  python douyin_search_crawler.py -k "钥匙扣相机" -n 1 --no-login --max-scrolls 1 --download

登录凭证保存在当前操作系统用户的私有目录，不写入搜索结果目录。
"""

import argparse
import csv
import html
import io
import json
import os
import random
import re
import sys
import time
import urllib.parse
from datetime import datetime
from pathlib import Path

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


# ---------------------- 配置区（按需修改） ----------------------
KEYWORD = "钥匙扣相机"
OUTPUT_DIR = r"F:\Nabby"          # 默认表格和视频根目录，也可用 --output 修改
TOP_N = 30
SEARCH_MAX_SCROLLS = 15          # “最多点赞”结果页的滚动轮数，每轮通常加载约10条
PAGE_DELAY = (2.0, 3.5)          # 每次滚动后的随机等待秒数
MAX_SCROLLS = None               # 命令行默认滚动轮数；None 表示使用 SEARCH_MAX_SCROLLS
HEADLESS = False                 # 首次运行或需要扫码/验证时不要设为 True
BROWSER_CHANNEL = "chrome"       # 优先 Chrome，失败时自动尝试 Edge；None 使用 Playwright Chromium
LOGIN_WAIT_SECONDS = 300
VERIFY_WAIT_SECONDS = 180
REQUEST_TIMEOUT = 60_000         # 页面打开超时（毫秒）
USE_SYSTEM_PROXY = False         # yt-dlp 下载是否使用系统代理；浏览器使用系统默认网络
# ================================================================

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36")
LOGIN_COOKIES = {"sessionid", "sessionid_ss", "sid_tt", "uid_tt"}
HEADERS = ["排名", "视频标题", "笔记链接(视频)", "点赞", "收藏", "评论", "播放量",
           "发布日期", "博主名字", "博主主页链接", "博主粉丝量", "博主最新视频发布时间",
           "抖音视频ID", "上榜情况", "备注"]
COL_WIDTHS = [6, 50, 45, 10, 10, 10, 12, 18, 16, 40, 14, 22, 24, 15, 34]
LINK_COLS = (3, 10)


class CrawlStopped(Exception):
    """用户主动停止当前爬取或下载任务。"""


def _check_stop(should_stop=None):
    if should_stop is not None and should_stop():
        raise CrawlStopped("用户停止了任务")


_LOG_HOOKS = []


def add_log_hook(fn):
    if fn not in _LOG_HOOKS:
        _LOG_HOOKS.append(fn)


def remove_log_hook(fn):
    try:
        _LOG_HOOKS.remove(fn)
    except ValueError:
        pass


def log(msg):
    line = f"[{datetime.now().strftime('%H:%M:%S')}] {msg}"
    print(line, flush=True)
    for fn in list(_LOG_HOOKS):
        try:
            fn(line)
        except Exception:
            pass


def fmt_num(value):
    if value is None:
        return "-"
    try:
        value = int(value)
    except (TypeError, ValueError):
        return str(value)
    return f"{value / 10000:.1f}万" if value >= 10000 else str(value)


def parse_count(value):
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return int(value)
    text = str(value).strip().replace(",", "")
    if not text or text in ("-", "--"):
        return None
    match = re.match(r"^(\d+(?:\.\d+)?)\s*(万|亿)?", text)
    if not match:
        return None
    number = float(match.group(1))
    if match.group(2) == "万":
        number *= 1e4
    elif match.group(2) == "亿":
        number *= 1e8
    return int(number)


def clean_text(value):
    return html.unescape(re.sub(r"<[^>]+>", "", str(value or ""))).strip()


def get_profile_dir():
    base = Path(os.environ.get("LOCALAPPDATA") or (Path.home() / ".config"))
    path = base / "DouyinCrawler" / "browser_profile"
    path.mkdir(parents=True, exist_ok=True)
    return path


def iter_aweme_items(value):
    """兼容搜索接口的 data[].aweme_info 和 aweme_list 两种结构。"""
    if isinstance(value, dict):
        if (value.get("aweme_id") or value.get("group_id")) and (
                "statistics" in value or "video" in value or value.get("_dom_fallback")):
            yield value
            return
        for child in value.values():
            if isinstance(child, (dict, list)):
                yield from iter_aweme_items(child)
    elif isinstance(value, list):
        for child in value:
            yield from iter_aweme_items(child)


def parse_search_payloads(response):
    """解析普通 JSON，以及综合搜索 stream 接口的 chunked JSON 分段。"""
    body = response.body()
    if body.lstrip().startswith(b"{"):
        return [json.loads(body)]

    payloads = []
    position = 0
    while position < len(body):
        header_end = body.find(b"\r\n", position)
        if header_end < 0:
            raise ValueError("流式搜索响应缺少分段头")
        chunk_size = int(body[position:header_end].split(b";", 1)[0], 16)
        position = header_end + 2
        if chunk_size == 0:
            break
        chunk = body[position:position + chunk_size]
        if len(chunk) != chunk_size:
            raise ValueError("流式搜索响应分段不完整")
        payloads.append(json.loads(chunk))
        position += chunk_size
        if body[position:position + 2] == b"\r\n":
            position += 2
    return payloads


class DouyinBrowser:
    def __init__(self, profile_dir, headless=False):
        self.profile_dir = Path(profile_dir)
        self.headless = headless
        self.cookie_file = self.profile_dir.parent / "douyin_cookies.txt"
        self._playwright = None
        self.context = None
        self.page = None
        self._pending_responses = []

    def __enter__(self):
        try:
            from playwright.sync_api import sync_playwright
        except ImportError as exc:
            raise RuntimeError(
                "缺少 Playwright，请执行：python -m pip install playwright") from exc
        self._playwright = sync_playwright().start()
        channels = [BROWSER_CHANNEL or None]
        if BROWSER_CHANNEL == "chrome":
            channels.append("msedge")
        errors = []
        for channel in channels:
            try:
                self.context = self._playwright.chromium.launch_persistent_context(
                    user_data_dir=str(self.profile_dir),
                    headless=self.headless,
                    channel=channel,
                    viewport={"width": 1440, "height": 1000},
                    locale="zh-CN",
                    timezone_id="Asia/Shanghai",
                    accept_downloads=False,
                )
                if channel == "msedge" and BROWSER_CHANNEL == "chrome":
                    log("[提示] Google Chrome 不可用，已自动改用 Microsoft Edge")
                break
            except Exception as exc:
                errors.append(f"{channel or 'chromium'}：{exc}")
        else:
            self._playwright.stop()
            self._playwright = None
            raise RuntimeError(
                "浏览器启动失败：\n" + "\n".join(errors) +
                "\n默认优先使用 Chrome，不可用时自动尝试 Microsoft Edge；也可将 "
                "BROWSER_CHANNEL 改为 None 并执行：python -m playwright install chromium")
        self.context.set_default_timeout(30_000)
        self.page = self.context.pages[0] if self.context.pages else self.context.new_page()
        self.page.on("response", self._on_response)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            if self.context is not None:
                self.save_cookies()
                self.context.close()
        finally:
            if self._playwright is not None:
                self._playwright.stop()

    def _on_response(self, response):
        try:
            url = response.url.lower()
            if (response.request.resource_type in ("xhr", "fetch")
                    and "/aweme/v1/web/" in url and "search" in url):
                self._pending_responses.append(response)
        except Exception:
            pass

    def drain_search_responses(self):
        responses, self._pending_responses = self._pending_responses, []
        items = []
        for response in responses:
            try:
                # 切换筛选前发出的普通搜索响应可能稍后返回，不能混入点赞排序候选池。
                if response.status != 200 or not self._is_most_liked_url(response.url):
                    continue
                for payload in parse_search_payloads(response):
                    items.extend(iter_aweme_items(payload))
            except Exception:
                # 页面继续导航后，部分响应体可能已被浏览器释放。
                continue
        return items

    def has_login_cookie(self):
        try:
            return any(c.get("name") in LOGIN_COOKIES and c.get("value")
                       for c in self.context.cookies("https://www.douyin.com/"))
        except Exception:
            return False

    def wait_for_login(self, force=False, should_stop=None):
        _check_stop(should_stop)
        if not force and self.has_login_cookie():
            log("[提示] 检测到已保存的抖音登录状态，将直接复用")
            return True
        if self.headless:
            log("[错误] 无头模式无法扫码，请先去掉 --headless 完成登录")
            return False
        if force:
            self.context.clear_cookies()
        self.page.goto("https://www.douyin.com/", wait_until="domcontentloaded",
                       timeout=REQUEST_TIMEOUT)
        log("=" * 60)
        log("【需要你操作】请在浏览器中点击【登录】，使用抖音 App 扫码并确认")
        log(f"  最多等待 {LOGIN_WAIT_SECONDS} 秒；登录状态会自动保存，下次无需重复扫码")
        log("=" * 60)
        deadline = time.monotonic() + LOGIN_WAIT_SECONDS
        while time.monotonic() < deadline:
            _check_stop(should_stop)
            if self.has_login_cookie():
                self.page.wait_for_timeout(2500)
                _check_stop(should_stop)
                self.save_cookies()
                log("[成功] 已检测到抖音登录状态")
                return True
            self.page.wait_for_timeout(2000)
        log("[错误] 等待扫码登录超时")
        return False

    def verification_pending(self):
        selectors = ("#captcha_container", ".captcha_verify_container",
                     "iframe[src*='captcha']", "iframe[src*='verify']",
                     "#root[style*='111111']")
        for selector in selectors:
            try:
                locator = self.page.locator(selector)
                if locator.count() and locator.first.is_visible():
                    return True
            except Exception:
                continue
        try:
            # 抖音的中间验证页可能在 shadow DOM 中渲染，常规选择器无法命中。
            body_html = self.page.locator("body").inner_html(timeout=1500)
            return "captchaOptions" in body_html and "verify_data" in body_html
        except Exception:
            return False

    def wait_for_verification(self, should_stop=None):
        _check_stop(should_stop)
        if not self.verification_pending():
            return True
        if self.headless:
            raise RuntimeError("页面需要人工安全验证；请去掉 --headless 后在浏览器中完成")
        log(f"[需要你操作] 请在浏览器中完成安全验证，最多等待 {VERIFY_WAIT_SECONDS} 秒 ...")
        deadline = time.monotonic() + VERIFY_WAIT_SECONDS
        while time.monotonic() < deadline:
            _check_stop(should_stop)
            if not self.verification_pending():
                log("[提示] 安全验证已完成，继续抓取")
                self.page.wait_for_timeout(2500)
                _check_stop(should_stop)
                return True
            self.page.wait_for_timeout(2000)
        return False

    def _search_input(self, timeout=15_000):
        selector = (':is(input[data-e2e="searchbar-input"], input[data-e2e="search-input"], '
                    'input[placeholder*="搜索"], input[type="search"]):visible')
        locator = self.page.locator(selector).first
        try:
            locator.wait_for(state="visible", timeout=timeout)
            return locator
        except Exception:
            return None

    @staticmethod
    def _is_most_liked_url(url):
        query = urllib.parse.parse_qs(urllib.parse.urlparse(url).query)
        if query.get("sort_type") == ["1"]:
            return True
        try:
            filter_selected = json.loads(query.get("filter_selected", [""])[0])
            return filter_selected.get("sort_type") == "1"
        except (json.JSONDecodeError, AttributeError, TypeError):
            return False

    def search_keyword(self, keyword, should_stop=None):
        """从抖音首页通过真实搜索框提交关键词；首页异常时再尝试搜索页搜索框。"""
        _check_stop(should_stop)
        self.page.goto("https://www.douyin.com/", wait_until="domcontentloaded",
                       timeout=REQUEST_TIMEOUT)
        self.page.keyboard.press("Escape")
        # 首页会自动跳转到 /jingxuan 并重建搜索框，等待跳转和渲染稳定后再输入。
        self.page.wait_for_timeout(2500)
        if self.verification_pending() and not self.wait_for_verification(should_stop):
            raise RuntimeError("安全验证未完成，本次任务已取消")

        search_input = self._search_input()
        if search_input is None:
            log("[警告] 首页未显示搜索框，尝试在搜索结果页输入关键词")
            url = f"https://www.douyin.com/jingxuan/search/{urllib.parse.quote(keyword)}?type=general"
            self.page.goto(url, wait_until="domcontentloaded", timeout=REQUEST_TIMEOUT)
            self.page.wait_for_timeout(2500)
            search_input = self._search_input()
        if search_input is None:
            raise RuntimeError(
                f"未找到可用的抖音搜索框（当前页面：{self.page.title() or self.page.url}）")

        search_input.click()
        search_input.fill("")
        search_input.fill(keyword)
        if search_input.input_value() != keyword:
            raise RuntimeError("搜索框内容未能稳定写入，请稍后重试")
        self._pending_responses.clear()
        search_button = self.page.locator(
            ':is(button[data-e2e="searchbar-button"], button[data-e2e="search-button"]):visible'
        ).first
        try:
            if search_button.count() and search_button.is_visible():
                search_button.click()
            else:
                search_input.press("Enter")
        except Exception:
            search_input.press("Enter")
        try:
            self.page.wait_for_url(re.compile(r"/search/"), timeout=15_000)
        except Exception as exc:
            if self.verification_pending():
                if not self.wait_for_verification(should_stop=should_stop):
                    raise RuntimeError("安全验证未完成，本次任务已取消") from exc
            if "/search/" not in self.page.url:
                raise RuntimeError("关键词已输入，但点击搜索后页面未进入搜索结果页") from exc
        self.page.wait_for_timeout(2500)
        if self.verification_pending() and not self.wait_for_verification(should_stop):
            raise RuntimeError("安全验证未完成，本次任务已取消")
        log(f"[提示] 已在抖音搜索框输入关键词：{keyword}")

    def select_most_liked(self, should_stop=None):
        """保持与人工搜索相同的综合结果，点击“筛选 → 最多点赞”。"""
        _check_stop(should_stop)
        # 不切换“视频”标签：type=general 与 type=video 的点赞排序候选集不同。
        filter_text = self.page.get_by_text("筛选", exact=True).first
        filter_text.wait_for(state="visible", timeout=20_000)
        filter_control = filter_text
        parent = filter_text.locator("..")
        try:
            if parent.get_attribute("tabindex") is not None:
                filter_control = parent
        except Exception:
            pass

        # 新版页面悬停展开，点击反而收起；旧版页面则在悬停无效时使用点击展开。
        filter_control.hover()
        most_liked = self.page.get_by_text("最多点赞", exact=True).first
        try:
            most_liked.wait_for(state="visible", timeout=5000)
        except Exception:
            filter_control.click()
            most_liked.wait_for(state="visible", timeout=10_000)
        # 丢弃筛选前已完成的响应，后续候选池只接收“最多点赞”结果。
        self._pending_responses.clear()
        most_liked.click()

        deadline = time.monotonic() + 20
        selected = False
        while time.monotonic() < deadline:
            _check_stop(should_stop)
            if self._is_most_liked_url(self.page.url):
                selected = True
                break
            if any(self._is_most_liked_url(response.url)
                   for response in self._pending_responses):
                selected = True
                break
            if self.verification_pending():
                if not self.wait_for_verification(should_stop=should_stop):
                    raise RuntimeError("安全验证未完成，本次任务已取消")
                continue
            self.page.wait_for_timeout(300)
        if not selected:
            raise RuntimeError(
                "已点击“最多点赞”，但未检测到对应的点赞筛选搜索请求；"
                "为避免混入未排序结果，本次任务已取消")
        self.page.wait_for_timeout(1500)
        log("[成功] 已点击筛选并切换为“最多点赞”")

    def get_author_stats(self, sec_uid, should_stop=None):
        """打开作者主页并捕获资料/作品响应，返回粉丝数和最新作品时间。"""
        result = {"author_fans": None, "author_latest_pubdate": None}
        responses = []
        sec_uid = str(sec_uid)

        def on_response(response):
            try:
                path = urllib.parse.urlparse(response.url).path.lower()
                if (response.request.resource_type in ("xhr", "fetch")
                        and path in ("/aweme/v1/web/user/profile/other/",
                                     "/aweme/v1/web/aweme/post/")):
                    query = urllib.parse.parse_qs(urllib.parse.urlparse(response.url).query)
                    if query.get("sec_user_id") == [sec_uid]:
                        responses.append(response)
            except Exception:
                pass

        self.page.on("response", on_response)
        try:
            url = f"https://www.douyin.com/user/{urllib.parse.quote(sec_uid, safe='')}"
            self.page.goto(url, wait_until="domcontentloaded", timeout=REQUEST_TIMEOUT)
            deadline = time.monotonic() + 25
            while time.monotonic() < deadline:
                _check_stop(should_stop)
                paths = {urllib.parse.urlparse(response.url).path.lower()
                         for response in responses}
                if ("/aweme/v1/web/user/profile/other/" in paths
                        and "/aweme/v1/web/aweme/post/" in paths):
                    break
                if self.verification_pending():
                    if not self.wait_for_verification(should_stop=should_stop):
                        raise RuntimeError("作者主页安全验证未完成")
                    continue
                self.page.wait_for_timeout(300)

            for response in responses:
                if response.status != 200:
                    continue
                path = urllib.parse.urlparse(response.url).path.lower()
                payload = response.json()
                if path == "/aweme/v1/web/user/profile/other/":
                    user = payload.get("user") or (payload.get("data") or {}).get("user") or {}
                    fans = parse_count(user.get("mplatform_followers_count"))
                    if fans is None:
                        fans = parse_count(user.get("follower_count"))
                    result["author_fans"] = fans
                else:
                    items = payload.get("aweme_list") or []
                    timestamps = [parse_count(item.get("create_time")) for item in items]
                    timestamps = [value for value in timestamps if value and value > 0]
                    if timestamps:
                        # 作者主页可能置顶旧作品，最新时间必须取本页作品时间最大值。
                        latest = max(timestamps)
                        current = result["author_latest_pubdate"]
                        result["author_latest_pubdate"] = max(current or 0, latest)
        except CrawlStopped:
            raise
        except Exception as exc:
            result["error"] = str(exc)
        finally:
            self.page.remove_listener("response", on_response)

        missing = []
        if result["author_fans"] is None:
            missing.append("粉丝量")
        if result["author_latest_pubdate"] is None:
            missing.append("最新视频时间")
        if missing and not result.get("error"):
            result["error"] = "、".join(missing) + "未返回"
        return result

    def dom_video_items(self):
        """接口响应不可用时的兜底：至少保留页面中可见的视频链接和标题。"""
        try:
            cards = self.page.evaluate(r"""
                () => Array.from(document.querySelectorAll('a[href*="/video/"]')).map(a => {
                    const href = new URL(a.href, location.origin).href;
                    const match = href.match(/\/video\/(\d+)/);
                    const card = a.closest('li') || a.parentElement?.parentElement || a;
                    return {
                        id: match ? match[1] : '',
                        url: match ? `https://www.douyin.com/video/${match[1]}` : href,
                        title: (a.getAttribute('title') || a.innerText || '').trim(),
                        cardText: (card.innerText || '').trim()
                    };
                })
            """)
        except Exception:
            return []
        items = []
        for card in cards:
            if not card.get("id"):
                continue
            title = clean_text(card.get("title"))
            if not title:
                title = clean_text((card.get("cardText") or "").splitlines()[0]
                                   if card.get("cardText") else "")
            items.append({
                "aweme_id": card["id"],
                "desc": title,
                "create_time": None,
                "author": {},
                "statistics": {},
                "_dom_fallback": True,
                "_page_url": card.get("url"),
            })
        return items

    def save_cookies(self):
        """导出完整浏览器 Cookie（含 HttpOnly），供 yt-dlp 使用。"""
        if self.context is None:
            return
        lines = ["# Netscape HTTP Cookie File",
                 "# Generated by douyin_search_crawler.py; do not share this file."]
        try:
            cookies = self.context.cookies()
            for cookie in cookies:
                domain = cookie.get("domain") or ""
                name = cookie.get("name") or ""
                if not domain or not name:
                    continue
                prefix = "#HttpOnly_" if cookie.get("httpOnly") else ""
                include_subdomains = "TRUE" if domain.startswith(".") else "FALSE"
                secure = "TRUE" if cookie.get("secure") else "FALSE"
                try:
                    expires = max(0, int(cookie.get("expires", 0)))
                except (TypeError, ValueError):
                    expires = 0
                lines.append("\t".join([
                    prefix + domain, include_subdomains, cookie.get("path") or "/",
                    secure, str(expires), name, str(cookie.get("value") or ""),
                ]))
            self.cookie_file.write_text("\n".join(lines) + "\n", encoding="utf-8")
        except Exception as exc:
            log(f"[警告] Cookie 保存失败：{exc}")

    def user_agent(self):
        try:
            return self.page.evaluate("navigator.userAgent") or UA
        except Exception:
            return UA


def extract_video(item):
    video_id = str(item.get("aweme_id") or item.get("group_id") or "").strip()
    stats = item.get("statistics") or item.get("stats") or {}
    author = item.get("author") or {}
    user_id = (author.get("sec_uid") or author.get("unique_id")
               or author.get("short_id") or author.get("uid") or "")
    video_data = item.get("video") or {}
    play_addr = video_data.get("play_addr") or {}
    url_list = play_addr.get("url_list") or []
    download_url = next((str(url) for url in url_list
                         if isinstance(url, str) and url.startswith("http")), "")
    play = parse_count(stats.get("play_count") if isinstance(stats, dict) else None)
    if play is not None and play <= 0:
        play = None   # 抖音搜索接口常用 0 作为未公开播放量的占位值
    note = ""
    if item.get("_dom_fallback"):
        note = "仅获取到页面卡片，接口统计字段缺失"
    elif play is None:
        note = "本次接口未返回公开播放量"
    return {
        "video_id": video_id,
        "title": clean_text(item.get("desc") or item.get("title")),
        "desc": clean_text(item.get("desc") or item.get("title")),
        "author": clean_text(author.get("nickname") or author.get("name")),
        "author_id": user_id,
        "author_fans": None,
        "author_latest_pubdate": None,
        "url": item.get("_page_url") or f"https://www.douyin.com/video/{video_id}",
        "download_url": download_url,
        "space": f"https://www.douyin.com/user/{user_id}" if user_id else "",
        "pubdate": item.get("create_time"),
        "like": parse_count(stats.get("digg_count")) if isinstance(stats, dict) else None,
        "fav": parse_count(stats.get("collect_count")) if isinstance(stats, dict) else None,
        "reply": parse_count(stats.get("comment_count")) if isinstance(stats, dict) else None,
        "play": play,
        "note": note,
    }


def crawl_all(browser, keyword, max_scrolls, should_stop=None):
    """人工操作搜索和筛选后，滚动“最多点赞”结果页并监听浏览器签名请求。"""
    candidates = {}
    seen_raw = set()
    log(f"—— 开始搜索（最多点赞，最多滚动 {max_scrolls} 轮，不做相关性过滤）——")
    browser.search_keyword(keyword, should_stop=should_stop)
    browser.select_most_liked(should_stop=should_stop)

    no_new = 0
    round_no = 0
    while round_no < max_scrolls:
        _check_stop(should_stop)
        round_no += 1
        browser.page.wait_for_timeout(int(random.uniform(*PAGE_DELAY) * 1000))
        raw_items = browser.drain_search_responses()
        new_count = 0
        for item in raw_items:
            video = extract_video(item)
            video_id = video["video_id"]
            if not video_id or video_id in seen_raw:
                continue
            seen_raw.add(video_id)
            candidates[video_id] = video
            new_count += 1
        log(f"  第{round_no:>2}轮：新增 {new_count} 条，搜索结果累计 {len(candidates)} 条")
        no_new = no_new + 1 if new_count == 0 else 0
        if new_count == 0 and browser.verification_pending():
            if not browser.wait_for_verification(should_stop=should_stop):
                raise RuntimeError("安全验证未完成，本次任务已取消")
            no_new = 0
            round_no -= 1   # 人工验证不占用搜索滚动轮数
            continue
        if no_new >= 4:
            log("  连续4轮没有新视频，结束抓取")
            break
        if round_no >= max_scrolls:
            break
        try:
            browser.page.evaluate("window.scrollBy(0, Math.max(3000, window.innerHeight * 3))")
        except Exception as exc:
            log(f"[警告] 页面滚动失败：{exc}")
            break

    # 若没有捕获到接口数据，尝试读取已渲染的“最多点赞”页面卡片。
    if not seen_raw:
        fallback_items = browser.dom_video_items()
        for item in fallback_items:
            video = extract_video(item)
            video_id = video["video_id"]
            if not video_id or video_id in seen_raw:
                continue
            seen_raw.add(video_id)
            candidates[video_id] = video
        if candidates:
            log(f"[提示] 页面卡片兜底新增 {len(candidates)} 条；这些结果的统计数据可能不完整")
    return candidates


def fetch_author_stats(browser, candidates, should_stop=None):
    """按作者去重补充粉丝数和最新视频时间，同一作者只打开一次主页。"""
    videos_by_author = {}
    for video in candidates.values():
        author_id = video.get("author_id")
        if author_id:
            videos_by_author.setdefault(str(author_id), []).append(video)
    total = len(videos_by_author)
    log(f"—— 获取 {total} 位博主的粉丝量和最新视频时间（按博主去重）——")
    ok = 0
    for index, (author_id, videos) in enumerate(videos_by_author.items(), 1):
        _check_stop(should_stop)
        stats = browser.get_author_stats(author_id, should_stop=should_stop)
        for video in videos:
            video["author_fans"] = stats.get("author_fans")
            video["author_latest_pubdate"] = stats.get("author_latest_pubdate")
            if stats.get("error"):
                message = f"作者资料部分缺失：{stats['error']}"
                video["note"] = "；".join(filter(None, [video.get("note"), message]))
        if not stats.get("error"):
            ok += 1
        if index % 5 == 0 or index == total:
            log(f"  作者进度：{index}/{total}（完整成功 {ok}）")
        browser.page.wait_for_timeout(int(random.uniform(*PAGE_DELAY) * 1000))
    return ok


def _like_key(video):
    return video["like"] if isinstance(video.get("like"), int) else -1


def _play_key(video):
    return video["play"] if isinstance(video.get("play"), int) else -1


def build_row(rank, video, tag):
    pub = ""
    pubdate = video.get("pubdate")
    try:
        if float(pubdate) > 0:
            pub = datetime.fromtimestamp(float(pubdate)).strftime("%Y-%m-%d %H:%M")
    except (TypeError, ValueError, OSError):
        if pubdate:
            pub = str(pubdate)
    latest_pub = ""
    latest = video.get("author_latest_pubdate")
    try:
        if float(latest) > 0:
            latest_pub = datetime.fromtimestamp(float(latest)).strftime("%Y-%m-%d %H:%M")
    except (TypeError, ValueError, OSError):
        if latest:
            latest_pub = str(latest)
    return [rank, video.get("title", ""), video.get("url", ""), video.get("like"),
            video.get("fav"), video.get("reply"), video.get("play"), pub,
            video.get("author", ""), video.get("space", ""), video.get("author_fans"),
            latest_pub, video.get("video_id", ""), tag, video.get("note", "")]


def build_rankings(candidates, top_n):
    like_top = sorted(candidates.values(), key=_like_key, reverse=True)[:top_n]
    play_top = sorted(candidates.values(), key=_play_key, reverse=True)[:top_n]
    like_ids = {v["video_id"] for v in like_top}
    play_ids = {v["video_id"] for v in play_top}

    def tag_of(video):
        in_like = video["video_id"] in like_ids
        in_play = video["video_id"] in play_ids
        if in_like and in_play:
            return "点赞榜+播放榜"
        if in_like:
            return "仅点赞榜"
        if in_play:
            return "仅播放榜"
        return "未上榜(仅收录于全部)"

    like_rows = [build_row(i, v, tag_of(v)) for i, v in enumerate(like_top, 1)]
    play_rows = [build_row(i, v, tag_of(v)) for i, v in enumerate(play_top, 1)]
    all_sorted = sorted(candidates.values(), key=lambda v: (_like_key(v), _play_key(v)), reverse=True)
    all_rows = [build_row(i, v, tag_of(v)) for i, v in enumerate(all_sorted, 1)]
    return like_rows, play_rows, all_rows


def write_excel(path, sheets):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    workbook = Workbook()
    workbook.remove(workbook.active)
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="305496")
    link_font = Font(color="0563C1", underline="single")
    for name, rows in sheets:
        sheet = workbook.create_sheet(name[:31])
        sheet.append(HEADERS)
        for index, width in enumerate(COL_WIDTHS, 1):
            sheet.column_dimensions[get_column_letter(index)].width = width
        for cell in sheet[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for row in rows:
            sheet.append(row)
        for row in sheet.iter_rows(min_row=2):
            for col in LINK_COLS:
                cell = row[col - 1]
                if cell.value:
                    cell.hyperlink = str(cell.value)
                    cell.font = link_font
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
    workbook.save(str(path))


def csv_bytes(rows):
    """在内存中生成 UTF-8 BOM CSV，供 Streamlit 按需下载，不写入本地文件。"""
    buffer = io.StringIO(newline="")
    writer = csv.writer(buffer)
    writer.writerow(HEADERS)
    writer.writerows(rows)
    return buffer.getvalue().encode("utf-8-sig")


def filter_candidates_by_fans(candidates, min_fans, max_fans):
    """仅保留粉丝数已知且位于闭区间内的视频，原候选池不受影响。"""
    return {key: video for key, video in candidates.items()
            if isinstance(video.get("author_fans"), int)
            and min_fans <= video["author_fans"] <= max_fans}


def export_files(out_dir, keyword, top_n, like_rows, play_rows, all_rows,
                 filtered_like_rows=None, filtered_play_rows=None, fans_range=None):
    """默认仅保存 Excel；区间榜单存在时放在原有工作表之前。"""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    safe_keyword = re.sub(r'[\\/:*?"<>|\r\n\t ]+', "_", keyword) or "keyword"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = out_dir / f"{safe_keyword}_抖音搜索_{timestamp}.xlsx"
    sheets = []
    if fans_range is not None:
        min_fans, max_fans = fans_range
        sheets.extend([
            (f"粉丝{min_fans}-{max_fans}点赞TOP{top_n}", filtered_like_rows or []),
            (f"粉丝{min_fans}-{max_fans}播放TOP{top_n}", filtered_play_rows or []),
        ])
    sheets.extend([
        (f"点赞榜TOP{top_n}", like_rows),
        (f"播放榜TOP{top_n}", play_rows),
        ("全部搜索结果", all_rows),
    ])
    write_excel(path, sheets)
    return {"xlsx": path}


def _download_direct(session, video, file_path, user_agent, should_stop=None):
    """优先使用搜索接口返回的临时播放直链，避免额外请求视频详情接口。"""
    _check_stop(should_stop)
    if file_path.exists() and file_path.stat().st_size > 0:
        return True
    direct_url = video.get("download_url")
    if not direct_url:
        return False
    part_path = file_path.with_suffix(file_path.suffix + ".part")
    try:
        with session.get(direct_url, headers={
                "User-Agent": user_agent,
                "Referer": video.get("url") or "https://www.douyin.com/",
                "Accept": "*/*",
        }, stream=True, allow_redirects=True, timeout=(15, 60)) as response:
            response.raise_for_status()
            content_type = response.headers.get("Content-Type", "").lower()
            if "text/html" in content_type or "application/json" in content_type:
                raise RuntimeError(f"直链返回了 {content_type} 而非视频")
            with open(part_path, "wb") as file:
                for chunk in response.iter_content(chunk_size=1024 * 1024):
                    _check_stop(should_stop)
                    if chunk:
                        file.write(chunk)
        if part_path.stat().st_size < 1024:
            raise RuntimeError("直链返回的文件过小")
        part_path.replace(file_path)
        return True
    except Exception:
        part_path.unlink(missing_ok=True)
        raise


def download_videos(videos, output_dir, cookie_file, user_agent, use_system_proxy=False,
                    should_stop=None):
    try:
        import imageio_ffmpeg
        import yt_dlp
    except ImportError as exc:
        raise RuntimeError("视频下载需要 yt-dlp 和 imageio-ffmpeg，请先安装这两个包") from exc
    from http.cookiejar import MozillaCookieJar
    import requests

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    direct_session = requests.Session()
    direct_session.trust_env = use_system_proxy
    cookie_jar = MozillaCookieJar(str(cookie_file))
    cookie_jar.load(ignore_discard=True, ignore_expires=True)
    direct_session.cookies.update(cookie_jar)
    downloaded_files = []
    failures = []
    total = len(videos)
    log(f"—— 开始下载 {total} 个视频，保存目录：{output_dir} ——")
    for index, video in enumerate(videos, 1):
        _check_stop(should_stop)
        video_id = video.get("video_id") or f"video_{index}"
        title = video.get("title") or video_id
        safe_title = re.sub(r'[\\/:*?"<>|\r\n\t%]+', "_", title).strip()[:100] or video_id
        url = video.get("url") or f"https://www.douyin.com/video/{video_id}"
        log(f"[下载] {index}/{total} 开始：{title}")
        direct_path = output_dir / f"{safe_title} [{video_id}].mp4"
        try:
            if _download_direct(direct_session, video, direct_path, user_agent,
                                should_stop=should_stop):
                downloaded_files.append(str(direct_path))
                log(f"[下载] {index}/{total} 成功：{direct_path.name}"
                    f"（{direct_path.stat().st_size / 1024 / 1024:.1f} MB）")
                continue
        except CrawlStopped:
            raise
        except Exception as exc:
            log(f"[提示] 直链下载失败，将尝试 yt-dlp：{exc}")
        options = {
            "format": "bv*[width<=854]+ba/b[width<=854]/b",
            "merge_output_format": "mp4",
            "ffmpeg_location": imageio_ffmpeg.get_ffmpeg_exe(),
            "outtmpl": str(output_dir / f"{safe_title} [{video_id}].%(ext)s"),
            "cookiefile": str(cookie_file),
            "http_headers": {"User-Agent": user_agent, "Referer": "https://www.douyin.com/"},
            "noplaylist": True,
            "quiet": True,
            "no_warnings": True,
            "noprogress": True,
            "progress_hooks": [lambda _status: _check_stop(should_stop)],
            "retries": 3,
            "fragment_retries": 3,
            "socket_timeout": 30,
            "continuedl": True,
            "overwrites": False,
            "windowsfilenames": True,
        }
        if not use_system_proxy:
            options["proxy"] = ""
        try:
            with yt_dlp.YoutubeDL(options) as downloader:
                downloader.download([url])
            files = [p for p in output_dir.iterdir() if p.is_file()
                     and f"[{video_id}]" in p.name
                     and p.suffix.lower() in (".mp4", ".mkv", ".webm", ".flv")]
            if not files:
                raise RuntimeError("下载结束但未找到视频文件")
            file_path = max(files, key=lambda p: p.stat().st_mtime)
            downloaded_files.append(str(file_path))
            log(f"[下载] {index}/{total} 成功：{file_path.name}"
                f"（{file_path.stat().st_size / 1024 / 1024:.1f} MB）")
        except Exception as exc:
            _check_stop(should_stop)
            message = str(exc).strip().splitlines()[-1] if str(exc).strip() else type(exc).__name__
            failures.append({"video_id": video_id, "title": title, "error": message})
            log(f"[下载] {index}/{total} 失败：{title}；原因：{message}")
    result = {"total": total, "success_count": len(downloaded_files),
              "failed_count": len(failures), "files": downloaded_files,
              "failures": failures, "output_dir": str(output_dir)}
    log(f"[下载完成] 成功 {result['success_count']} 个，失败 {result['failed_count']} 个")
    return result


def preview(title, rows, count=5):
    log(f"—— {title}（前{min(count, len(rows))}条预览）——")
    for row in rows[:count]:
        log(f"  {row[0]:>3}. {str(row[1])[:30]} | 赞{fmt_num(row[3])} "
            f"藏{fmt_num(row[4])} 评{fmt_num(row[5])} 播{fmt_num(row[6])}")


def positive_int(value):
    number = int(value)
    if number < 1:
        raise argparse.ArgumentTypeError("必须是正整数")
    return number


def parse_args():
    parser = argparse.ArgumentParser(description="抖音关键词搜索爬虫（点赞榜/播放榜，默认保存Excel+可选下载）")
    # 搜索关键词；命令行不传时使用配置区 KEYWORD。
    parser.add_argument("-k", "--keyword", default=KEYWORD, help=f"搜索关键词（默认：{KEYWORD}）")
    # Excel 输出目录，同时作为默认视频根目录。
    parser.add_argument("-o", "--output", default=OUTPUT_DIR, help=f"输出目录（默认：{OUTPUT_DIR}）")
    # 点赞榜和播放榜分别保留的条数，只接受正整数。
    parser.add_argument("-n", "--top-n", type=positive_int, default=TOP_N, help=f"每个榜单条数（默认：{TOP_N}）")
    # “最多点赞”结果页的最大滚动轮数，用于控制搜索覆盖范围和运行时间。
    parser.add_argument("--max-scrolls", type=positive_int, default=MAX_SCROLLS,
                        help="最多点赞结果页的最大滚动轮数（调试/提速用）")
    # 清除旧登录状态并强制扫码；与 --no-login、--headless 互斥。
    parser.add_argument("--login", action="store_true", help="清除旧登录状态并强制重新扫码")
    # 不询问扫码，直接使用当前浏览器会话（已有 Cookie 仍会保留）。
    parser.add_argument("--no-login", action="store_true", help="不询问扫码，直接使用当前浏览器会话")
    # 无界面运行；需要扫码或人工安全验证时不能使用。
    parser.add_argument("--headless", action="store_true", default=HEADLESS,
                        help="后台运行浏览器（扫码或人工验证时不可用）")
    # 是否下载双榜合并去重后的视频；默认只导出表格。
    parser.add_argument("--download", action="store_true", help="下载双榜合并去重后的视频")
    # 自定义视频目录；不传时使用“输出目录/videos”。
    parser.add_argument("--video-output", default=None, help="视频目录（默认：输出目录/videos）")
    return parser.parse_args()


def main():
    started = time.time()
    args = parse_args()
    keyword = args.keyword.strip()
    if not keyword:
        raise ValueError("搜索关键词不能为空")
    if args.login and args.no_login:
        raise ValueError("--login 和 --no-login 不能同时使用")
    if args.login and args.headless:
        raise ValueError("--login 需要可见浏览器，不能与 --headless 同时使用")
    out_dir = Path(args.output).expanduser()
    video_dir = Path(args.video_output).expanduser() if args.video_output else out_dir / "videos"
    out_dir.mkdir(parents=True, exist_ok=True)
    max_scrolls = args.max_scrolls or SEARCH_MAX_SCROLLS

    log("=" * 60)
    log("抖音关键词搜索爬虫 启动")
    log(f"  关键词  ：{keyword}")
    log(f"  输出目录：{out_dir}")
    log(f"  榜单条数：TOP {args.top_n}（点赞榜 / 播放榜）")
    log("  搜索方式：筛选 → 最多点赞（不做相关性过滤）")
    log(f"  滚动轮数：{max_scrolls}")
    log("=" * 60)

    with DouyinBrowser(get_profile_dir(), headless=args.headless) as browser:
        if args.login:
            if not browser.wait_for_login(force=True):
                raise RuntimeError("未完成扫码登录，本次任务已取消")
        elif not args.no_login:
            if browser.has_login_cookie():
                log("[提示] 将复用本地浏览器登录状态")
            else:
                try:
                    answer = input(">> 是否现在扫码登录？(回车=扫码 / n=直接抓取): ").strip().lower()
                except (EOFError, KeyboardInterrupt):
                    answer = "n"
                    print()
                if answer in ("", "y", "yes"):
                    if not browser.wait_for_login():
                        raise RuntimeError("未完成扫码登录，本次任务已取消")
                else:
                    log("[提示] 使用未登录浏览器会话抓取")

        candidates = crawl_all(browser, keyword, max_scrolls)
        if not candidates:
            raise RuntimeError(
                "没有获取到搜索结果。请检查关键词和“最多点赞”筛选是否生效；"
                "若页面要求登录或验证，请使用 --login 并在浏览器中完成操作")
        fetch_author_stats(browser, candidates)
        browser.save_cookies()
        missing_play = sum(not isinstance(v.get("play"), int) for v in candidates.values())
        if missing_play:
            log(f"[提示] {missing_play}/{len(candidates)} 个结果未返回公开播放量；"
                "对应单元格留空，播放榜只能按已返回的数据排序")

        like_rows, play_rows, all_rows = build_rankings(candidates, args.top_n)
        if args.download:
            selected_ids = dict.fromkeys(row[12] for row in like_rows + play_rows)
            selected = [candidates[video_id] for video_id in selected_ids if video_id in candidates]
            download_videos(selected, video_dir, browser.cookie_file,
                            browser.user_agent(), USE_SYSTEM_PROXY)

        paths = export_files(out_dir, keyword, args.top_n, like_rows, play_rows, all_rows)

    log("=" * 60)
    log(f"抓取完成：搜索结果 {len(candidates)} 个（已去重，不做相关性过滤）")
    log(f"  点赞榜：{len(like_rows)} 条；播放榜：{len(play_rows)} 条")
    log(f"  Excel：{paths['xlsx']}")
    if args.download:
        log(f"  视频目录：{video_dir}")
    log(f"  耗时：{time.time() - started:.0f} 秒")
    log("=" * 60)
    preview(f"点赞榜 TOP{len(like_rows)}", like_rows)
    preview(f"播放榜 TOP{len(play_rows)}", play_rows)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print()
        log("已手动退出。")
        sys.exit(130)
    except Exception as exc:
        log(f"[错误] {exc}")
        sys.exit(1)
