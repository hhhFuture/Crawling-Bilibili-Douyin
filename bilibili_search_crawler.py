# -*- coding: utf-8 -*-
"""
B站(哔哩哔哩)关键词搜索爬虫
================================================================
功能：
  1) 按关键词搜索B站视频，关键词可在下方【配置区】或用 --keyword 参数修改
  2) 严格过滤相关性：只保留标题/简介中真正包含关键词的视频，
     解决B站搜索"只要命中其中一个词就展示"导致大量无关结果的问题
  3) 分别按【点赞量】与【播放量】排序，各取前 TOP_N 条；
     同一视频只请求一次详情数据，自动去重，防止重复抓取
  4) 默认仅保存 Excel，Streamlit 可按需下载 CSV；字段包括：
     笔记链接(视频)、点赞、收藏、评论、播放量、发布日期、博主名字、
     博主主页链接、博主粉丝量、博主最新视频发布时间、BV号、上榜情况、备注
  5) 支持扫码登录(Cookie持久化)、登录凭证过期自动提醒、
     风控自动重试，并在终端明确提示需要人工操作的事项(扫码等)

常用命令：
  python bilibili_search_crawler.py                          # 使用下方默认配置运行
  python bilibili_search_crawler.py --keyword 钥匙扣相机     # 临时更换关键词
  python bilibili_search_crawler.py --output D:\\mydir       # 临时更换输出目录
  python bilibili_search_crawler.py --login                  # 强制重新扫码登录(换账号用)
  python bilibili_search_crawler.py --no-login               # 跳过登录，匿名抓取(可能被限制)
  python -m streamlit run bilibili_crawler_app.py            # 可视化界面(推荐，页面输入关键词/保存路径)

依赖：requests / qrcode / pillow / openpyxl（缺失时脚本会尝试自动 pip 安装）
"""

import argparse
import csv
import hashlib
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

# 让中文/符号在各类终端(含重定向输出)下都不报编码错误
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


# ---------------------- 配置区（按需修改） ----------------------
KEYWORD = "钥匙扣相机"          # 【搜索关键词】改这里即可
OUTPUT_DIR = r"F:\Nabby"        # 【输出目录】改这里即可（也可用 --output 临时指定）
TOP_N = 30                      # 点赞榜/播放榜各取前多少条
FILTER_MODE = "phrase"          # 相关性过滤模式：
                                #   phrase - 标题或简介必须包含完整关键词（默认，最严格）
                                #   smart  - phrase 或 同义词分组全部命中（较宽松，见下）
                                #   terms  - 关键词按空格拆分，每个词都必须出现
                                #   off    - 不过滤（等同B站原始搜索结果）
SYNONYM_GROUPS = [              # smart 模式同义词组：每组内至少命中一个词才算相关
    ["钥匙扣", "钥匙串", "钥匙链", "钥匙圈", "挂件", "吊坠"],
    ["相机", "摄像机", "摄像头", "摄影机", "camera"],
]                               # 注意：更换 KEYWORD 时请同步修改此分组
SEARCH_ORDERS = [               # 搜索排序及抓取页数，每页返回数量由B站接口决定：
    ("totalrank", 15),          #   totalrank - 综合排序(相关视频集中在前面，多抓几页)
    ("click", 10),              #   click     - 按播放量排序
    # ("pubdate", 5),           #   pubdate   - 按最新发布
    # ("stow", 5),              #   stow      - 按收藏数
]
PAGE_DELAY = (1.2, 2.5)         # 搜索翻页之间的随机延时(秒)，防风控，可调大
STAT_DELAY = (0.3, 0.8)         # 获取单个视频详情之间的随机延时(秒)
REQUEST_TIMEOUT = 15            # 单次请求超时(秒)
COOKIE_FILENAME = "bilibili_cookies.json"   # Cookie 文件(保存在输出目录)
USE_SYSTEM_PROXY = False        # 是否走系统代理。默认 False=直连B站。
                                # 若本机系统代理已失效(如127.0.0.1:10809未启动)，
                                # 直连可避免 SSL 报错；仅当你的网络必须经代理才能
                                # 访问B站时才改为 True。
# ================================================================

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")

# B站 WBI 签名用的固定混淆表(公开算法)
MIXIN_KEY_ENC_TAB = [
    46, 47, 18, 2, 53, 8, 23, 32, 15, 50, 10, 31, 58, 3, 45, 35, 27, 43, 5, 49,
    33, 9, 42, 19, 29, 28, 14, 39, 12, 38, 41, 13, 37, 48, 7, 16, 24, 55, 40,
    61, 26, 17, 0, 1, 60, 51, 30, 4, 22, 25, 54, 21, 56, 59, 6, 63, 57, 62, 11,
    36, 20, 34, 44, 52,
]

ORDER_NAMES = {
    "totalrank": "综合排序",
    "click": "播放量优先",
    "pubdate": "最新发布",
    "stow": "收藏优先",
    "dm": "弹幕优先",
    "scores": "评论优先",
}

RISK_CODES = {-412, 412, -352, -799}   # 风控/频率限制相关错误码


class CrawlStopped(Exception):
    """用户主动停止当前爬取/下载任务"""


def _check_stop(should_stop=None):
    if should_stop is not None and should_stop():
        raise CrawlStopped("用户停止了任务")


def _sleep(seconds, should_stop=None):
    """可被停止按钮打断的延时，CLI 不传 should_stop 时行为不变"""
    end = time.monotonic() + seconds
    while True:
        _check_stop(should_stop)
        remaining = end - time.monotonic()
        if remaining <= 0:
            return
        time.sleep(min(0.2, remaining))


def get_user_cookie_file() -> Path:
    """登录凭证保存在当前操作系统用户的私有目录，不随代码或导出表格分享"""
    base = Path(os.environ.get("LOCALAPPDATA") or (Path.home() / ".config"))
    path = base / "BilibiliCrawler" / COOKIE_FILENAME
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


_LOG_HOOKS = []   # 日志钩子：供 Streamlit 等外部UI接收实时日志


def add_log_hook(fn):
    """注册日志回调：每条日志会以完整字符串传给 fn(line)"""
    if fn not in _LOG_HOOKS:
        _LOG_HOOKS.append(fn)


def remove_log_hook(fn):
    """移除日志回调"""
    try:
        _LOG_HOOKS.remove(fn)
    except ValueError:
        pass


def log(msg: str):
    line = f"[{datetime.now().strftime('%H:%M:%S')}] {msg}"
    print(line, flush=True)
    for fn in list(_LOG_HOOKS):
        try:
            fn(line)
        except Exception:
            pass


def fmt_num(v):
    """终端预览用的数字格式化：12345 -> 1.2万"""
    if v is None:
        return "-"
    try:
        v = int(v)
    except (TypeError, ValueError):
        return str(v)
    return f"{v / 10000:.1f}万" if v >= 10000 else str(v)


def parse_count(v):
    """把接口返回的计数(可能是 int / '1.2万' / '13亿')统一转成 int"""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return int(v)
    s = str(v).strip().replace(",", "")
    if not s or s in ("-", "--"):
        return None
    m = re.match(r"^(\d+(?:\.\d+)?)\s*(万|亿)?", s)
    if not m:
        return None
    n = float(m.group(1))
    if m.group(2) == "万":
        n *= 1e4
    elif m.group(2) == "亿":
        n *= 1e8
    return int(n)


def clean_html(s: str) -> str:
    """去掉搜索结果标题/简介里的 <em class="keyword"> 等高亮标签"""
    return html.unescape(re.sub(r"<[^>]+>", "", s or ""))


# ---------------------- 依赖自检与自动安装 ----------------------
def _ensure_packages():
    import importlib.util
    import subprocess
    required = [("requests", "requests"), ("qrcode", "qrcode"),
                ("PIL", "pillow"), ("openpyxl", "openpyxl")]
    missing = [pkg for mod, pkg in required if importlib.util.find_spec(mod) is None]
    if not missing:
        return
    print(f"[依赖] 检测到缺少: {', '.join(missing)}，正在自动安装 ...")
    try:
        subprocess.run([sys.executable, "-m", "pip", "install", *missing, "-q"], check=True)
        print("[依赖] 安装完成")
    except Exception:
        print(f"[依赖] 自动安装失败，请手动执行: pip install {' '.join(missing)}")
        sys.exit(1)


_ensure_packages()

import qrcode  # noqa: E402
import requests  # noqa: E402


# ====================== B站接口客户端 ======================
class BiliClient:
    NAV_URL = "https://api.bilibili.com/x/web-interface/nav"
    SPI_URL = "https://api.bilibili.com/x/frontend/finger/spi"
    SEARCH_URL = "https://api.bilibili.com/x/web-interface/wbi/search/type"
    STAT_URL = "https://api.bilibili.com/x/web-interface/archive/stat"
    VIEW_URL = "https://api.bilibili.com/x/web-interface/wbi/view"
    FOLLOWER_URL = "https://api.bilibili.com/x/relation/stat"
    USER_VIDEOS_URL = "https://api.bilibili.com/x/space/wbi/arc/search"
    QR_GEN_URL = "https://passport.bilibili.com/x/passport-login/web/qrcode/generate"
    QR_POLL_URL = "https://passport.bilibili.com/x/passport-login/web/qrcode/poll"

    def __init__(self, cookie_file: Path):
        self.cookie_file = cookie_file
        self.out_dir = cookie_file.parent
        self.session = requests.Session()
        self.session.trust_env = USE_SYSTEM_PROXY   # 默认直连，绕过失效的系统代理
        self.session.headers.update({
            "User-Agent": UA,
            "Referer": "https://www.bilibili.com/",
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        })
        self._img_key = None
        self._sub_key = None
        self._wbi_ts = 0.0

    # ---------------- Cookie / 登录状态 ----------------
    def load_cookies(self) -> bool:
        if not self.cookie_file.exists():
            return False
        try:
            data = json.loads(self.cookie_file.read_text(encoding="utf-8"))
            for k, v in (data.get("cookies") or {}).items():
                self.session.cookies.set(k, v, domain=".bilibili.com")
            log(f"[提示] 已加载本地Cookie(保存于 {data.get('saved_at', '未知时间')})：{self.cookie_file}")
            return True
        except Exception as e:
            log(f"[警告] Cookie 文件读取失败({e})，将忽略")
            return False

    def save_cookies(self):
        try:
            payload = {
                "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "cookies": {c.name: c.value for c in self.session.cookies},
            }
            self.cookie_file.write_text(json.dumps(payload, ensure_ascii=False, indent=2),
                                        encoding="utf-8")
        except Exception as e:
            log(f"[警告] Cookie 保存失败：{e}")

    def has_sessdata(self) -> bool:
        return any(c.name == "SESSDATA" and c.value for c in self.session.cookies)

    def ensure_buvid(self):
        """游客也需要设备标识 buvid3/buvid4，否则容易被风控拦截"""
        if any(c.name == "buvid3" for c in self.session.cookies):
            return
        try:
            j = self.session.get(self.SPI_URL, timeout=REQUEST_TIMEOUT).json()
            if j.get("code") == 0 and j.get("data"):
                self.session.cookies.set("buvid3", j["data"]["b_3"], domain=".bilibili.com")
                self.session.cookies.set("buvid4", j["data"]["b_4"], domain=".bilibili.com")
                log("[提示] 已获取游客设备标识(buvid3/buvid4)")
                return
        except Exception as e:
            log(f"[提示] 获取 buvid 失败({e})，尝试访问B站首页获取")
        try:
            self.session.get("https://www.bilibili.com/", timeout=REQUEST_TIMEOUT)
        except Exception:
            pass

    def check_login(self):
        """返回 (是否已登录, 用户名)"""
        try:
            j = self.session.get(self.NAV_URL, timeout=REQUEST_TIMEOUT).json()
        except Exception as e:
            log(f"[警告] 登录状态检查失败({e})")
            return False, None
        data = j.get("data") or {}
        wbi = data.get("wbi_img") or {}
        if wbi.get("img_url") and wbi.get("sub_url"):
            self._img_key = wbi["img_url"].rsplit("/", 1)[-1].split(".")[0]
            self._sub_key = wbi["sub_url"].rsplit("/", 1)[-1].split(".")[0]
            self._wbi_ts = time.time()
        if j.get("code") == 0 and data.get("isLogin"):
            return True, data.get("uname") or ""
        return False, None

    # ---------------- 扫码登录 ----------------
    def qr_login(self, max_rounds: int = 3, qr_callback=None, should_stop=None) -> bool:
        """
        扫码登录。只有B站返回 data.code=0 且 nav 验证已登录时才返回 True。
        qr_callback(png_path, qr_url) 用于外部UI展示二维码；should_stop 用于中途取消。
        """
        if qr_callback is None:
            print()
            log("=" * 60)
            log("【需要你操作】请使用 哔哩哔哩手机App 扫描二维码完成登录")
            log("  App路径：打开B站App → 右下角【我的】→ 右上角【扫一扫】")
            log("=" * 60)
        else:
            log("【需要你操作】请使用 哔哩哔哩手机App 扫描页面上的二维码，并在手机上确认登录")
        for _round in range(1, max_rounds + 1):
            _check_stop(should_stop)
            try:
                j = self.session.get(self.QR_GEN_URL, timeout=REQUEST_TIMEOUT).json()
            except Exception as e:
                log(f"[错误] 获取二维码失败({e})，5秒后重试 ...")
                _sleep(5, should_stop)
                continue
            if j.get("code") != 0:
                log(f"[错误] 获取二维码失败 code={j.get('code')} {j.get('message')}")
                _sleep(5, should_stop)
                continue
            qr_url = j["data"]["url"]
            qrcode_key = j["data"]["qrcode_key"]

            png_path = self.out_dir / "bilibili_login_qr.png"
            saved = False
            try:
                qrcode.make(qr_url).save(str(png_path))
                saved = True
            except Exception as e:
                log(f"[提示] 二维码图片生成失败({e})")

            if qr_callback is not None:
                qr_callback(png_path if saved else None, qr_url)
            else:
                opened = False
                if saved:
                    log(f"[提示] 二维码图片已保存：{png_path}")
                    try:
                        if hasattr(os, "startfile"):
                            os.startfile(str(png_path))
                            opened = True
                    except Exception:
                        pass
                if not opened:
                    try:
                        qr = qrcode.QRCode(border=1)
                        qr.add_data(qr_url)
                        qr.make(fit=True)
                        print()
                        qr.print_ascii(invert=True)
                        print()
                    except Exception as e:
                        log(f"[提示] 终端二维码打印失败({e})")
            log("[提示] 二维码约3分钟内有效；扫码后请选择登录环境，并在手机上完成最终确认")

            t0 = time.time()
            last_code = None
            last_beat = time.time()
            while time.time() - t0 < 175:
                _sleep(2.5, should_stop)
                if time.time() - last_beat > 30:
                    log("[提示] 仍在等待扫码/确认登录 ...")
                    last_beat = time.time()
                try:
                    j = self.session.get(self.QR_POLL_URL, params={"qrcode_key": qrcode_key},
                                         timeout=REQUEST_TIMEOUT).json()
                except Exception as e:
                    log(f"[提示] 查询扫码状态失败({e})，将继续尝试")
                    continue

                # 关键：扫码状态在 data.code；外层 code=0 只表示本次接口请求成功
                data = j.get("data") or {}
                code = data.get("code", j.get("code"))
                if code == 0:
                    log("[提示] B站已确认扫码，正在保存凭证并验证登录状态 ...")
                    url = data.get("url", "")
                    qs = urllib.parse.parse_qs(urllib.parse.urlparse(url).query)
                    for name in ("SESSDATA", "bili_jct", "DedeUserID", "DedeUserID__ckMd5"):
                        if qs.get(name):
                            self.session.cookies.set(name, qs[name][0], domain=".bilibili.com")
                    # 部分网络环境还需要访问 crossDomain 完成服务端凭证交换
                    if url:
                        try:
                            self.session.get(url, timeout=REQUEST_TIMEOUT)
                        except Exception as e:
                            log(f"[提示] crossDomain 访问未完成({e})，将直接验证已有凭证")
                    self.save_cookies()
                    ok, uname = self.check_login()
                    if ok:
                        log(f"[成功] 扫码登录成功！当前账号：{uname}")
                        log(f"[提示] 登录凭证已保存到当前用户目录：{self.cookie_file}")
                        return True
                    log("[错误] B站返回扫码成功，但登录状态验证未通过")
                    break
                if code == 86090 and last_code != 86090:
                    log("[提示] 已扫码，请在手机上选择登录环境并点击【确认登录】")
                elif code == 86038:
                    log("[提示] 二维码已过期，正在重新生成 ...")
                    break
                elif code not in (86101, 86090) and code != last_code:
                    log(f"[提示] 扫码状态：code={code} {data.get('message') or j.get('message')}")
                # 86101=未扫码，继续等待
                last_code = code
        log("[错误] 未能完成扫码登录；本次不会开始爬取")
        return False

    # ---------------- WBI 签名 ----------------
    def _fetch_wbi_keys(self):
        j = self.session.get(self.NAV_URL, timeout=REQUEST_TIMEOUT).json()
        wbi = (j.get("data") or {}).get("wbi_img") or {}
        if not (wbi.get("img_url") and wbi.get("sub_url")):
            raise RuntimeError("nav 接口未返回 WBI 密钥")
        self._img_key = wbi["img_url"].rsplit("/", 1)[-1].split(".")[0]
        self._sub_key = wbi["sub_url"].rsplit("/", 1)[-1].split(".")[0]
        self._wbi_ts = time.time()

    def wbi_sign(self, params: dict) -> str:
        if not self._img_key or time.time() - self._wbi_ts > 3600:
            self._fetch_wbi_keys()
        mixin_key = "".join((self._img_key + self._sub_key)[i] for i in MIXIN_KEY_ENC_TAB)[:32]
        p = dict(params)
        p["wts"] = int(time.time())
        p = {k: "".join(ch for ch in str(v) if ch not in "!'()*") for k, v in p.items()}
        p = dict(sorted(p.items()))
        query = urllib.parse.urlencode(p)
        return query + "&w_rid=" + hashlib.md5((query + mixin_key).encode()).hexdigest()

    # ---------------- 搜索 / 视频详情 ----------------
    def search_page(self, keyword: str, order: str, page: int, should_stop=None):
        """返回视频条目列表；[] 表示无更多结果；None 表示被风控/网络拦截"""
        params = {"search_type": "video", "keyword": keyword, "order": order, "page": page}
        try:
            url = self.SEARCH_URL + "?" + self.wbi_sign(params)
        except Exception as e:
            log(f"[错误] WBI签名失败：{e}")
            return None
        for attempt in range(1, 5):
            _check_stop(should_stop)
            try:
                r = self.session.get(url, timeout=REQUEST_TIMEOUT)
                j = r.json()
            except Exception as e:
                log(f"[警告] 搜索请求异常({e})，3秒后重试({attempt}/4) ...")
                _sleep(3, should_stop)
                continue
            code = j.get("code")
            if code == 0:
                res = (j.get("data") or {}).get("result")
                return res if isinstance(res, list) else []
            if code in RISK_CODES:
                wait = 5 * attempt
                log(f"[警告] 触发B站风控(code={code}: {j.get('message')})，等待{wait}秒后重试({attempt}/4) ...")
                log("       若反复出现：建议扫码登录、稍后再试，或增大脚本顶部 PAGE_DELAY")
                _sleep(wait, should_stop)
                continue
            log(f"[错误] 搜索失败 code={code}：{j.get('message')}")
            return []
        return None

    @staticmethod
    def _stat_params(bvid, aid):
        if bvid and str(bvid).startswith("BV"):
            return {"bvid": bvid}
        return {"aid": aid or bvid}

    def get_video_stat(self, bvid, aid=None, should_stop=None):
        """获取单个视频的 点赞/收藏/评论/播放；返回 dict，失败时含 'error'
        主通道为 wbi/view(视频页接口，数据最全)；archive/stat 作为兜底"""
        params = self._stat_params(bvid, aid)
        last_err = None
        for attempt in range(1, 4):
            _check_stop(should_stop)
            try:
                url = self.VIEW_URL + "?" + self.wbi_sign(params)
                j = self.session.get(url, timeout=REQUEST_TIMEOUT).json()
                code = j.get("code")
                if code == 0:
                    st = (j.get("data") or {}).get("stat") or {}
                    return {"like": st.get("like"), "favorite": st.get("favorite"),
                            "reply": st.get("reply"), "view": st.get("view")}
                if code not in RISK_CODES:
                    return {"error": f"code={code} {j.get('message')}"}
                last_err = f"code={code} {j.get('message')}"
            except Exception as e:
                last_err = f"网络异常({e})"
            wait = 3 * attempt
            log(f"[警告] 视频详情请求失败({last_err})，{wait}秒后重试({attempt}/3) ...")
            _sleep(wait, should_stop)
        # 兜底通道：archive/stat
        try:
            r = self.session.get(self.STAT_URL, params=params, timeout=REQUEST_TIMEOUT)
            j = r.json()
            if j.get("code") == 0:
                d = j.get("data") or {}
                return {"like": d.get("like"), "favorite": d.get("favorite"),
                        "reply": d.get("reply"), "view": d.get("view")}
            return {"error": f"code={j.get('code')} {j.get('message')}"}
        except Exception:
            return {"error": last_err or "多次请求失败"}

    def get_author_stats(self, mid, should_stop=None):
        """获取作者粉丝数和最新视频发布时间；单个字段失败不影响另一字段。"""
        _check_stop(should_stop)
        result = {"author_fans": None, "author_latest_pubdate": None}
        errors = []
        try:
            response = self.session.get(self.FOLLOWER_URL, params={"vmid": mid},
                                        timeout=REQUEST_TIMEOUT)
            payload = response.json()
            if payload.get("code") == 0:
                result["author_fans"] = parse_count((payload.get("data") or {}).get("follower"))
            else:
                errors.append(f"粉丝数 code={payload.get('code')} {payload.get('message')}")
        except Exception as exc:
            errors.append(f"粉丝数请求失败({exc})")

        _check_stop(should_stop)
        try:
            params = {"mid": mid, "pn": 1, "ps": 10, "order": "pubdate",
                      "platform": "web", "web_location": 1550101}
            url = self.USER_VIDEOS_URL + "?" + self.wbi_sign(params)
            payload = self.session.get(url, timeout=REQUEST_TIMEOUT).json()
            if payload.get("code") == 0:
                videos = (((payload.get("data") or {}).get("list") or {}).get("vlist") or [])
                timestamps = [video.get("created") for video in videos
                              if isinstance(video.get("created"), (int, float))
                              and video.get("created") > 0]
                if timestamps:
                    # 作者可能置顶旧视频，不能直接把列表第一条当成最新作品。
                    result["author_latest_pubdate"] = max(timestamps)
            else:
                errors.append(f"最新视频 code={payload.get('code')} {payload.get('message')}")
        except Exception as exc:
            errors.append(f"最新视频请求失败({exc})")
        if errors:
            result["error"] = "；".join(errors)
        return result


# ====================== 数据处理 ======================
def extract_video(item: dict) -> dict:
    bvid = (item.get("bvid") or "").strip()
    aid = item.get("aid")
    mid = item.get("mid")
    url = f"https://www.bilibili.com/video/{bvid}" if bvid else f"https://www.bilibili.com/video/av{aid}"
    return {
        "bvid": bvid or f"av{aid}",
        "aid": aid,
        "title": clean_html(item.get("title", "")),
        "desc": clean_html(item.get("description", "")),
        "author": item.get("author") or "",
        "mid": mid,
        "url": url,
        "space": f"https://space.bilibili.com/{mid}" if mid else "",
        "pubdate": item.get("pubdate"),
        "author_fans": None,
        "author_latest_pubdate": None,
        "play": parse_count(item.get("play")),
        "fav": parse_count(item.get("favorites")),
        "like": None, "reply": None, "view": None,
        "note": "",
    }


def make_filter(keyword: str, mode: str):
    """构造相关性判断函数：解决B站'命中一个词就算结果'导致的无关视频问题"""
    kw = (keyword or "").strip().lower()

    def _relevant(title: str, desc: str) -> bool:
        if not kw or mode == "off":
            return True
        text = f"{title} {desc}".lower()
        if mode == "terms":
            return all(t in text for t in kw.split() if t)
        # phrase：完整关键词必须出现在标题或简介中
        if kw in text or kw.replace(" ", "") in text.replace(" ", ""):
            return True
        # smart：同义词分组全部命中也算相关
        if mode == "smart" and SYNONYM_GROUPS:
            return all(any(str(t).lower() in text for t in group) for group in SYNONYM_GROUPS)
        return False

    return _relevant


def crawl_all(client: BiliClient, keyword: str, orders, relevant, should_stop=None):
    """按多种排序抓取搜索结果，过滤+去重后返回候选视频池"""
    candidates = {}
    blocked_orders = 0
    for order, max_pages in orders:
        _check_stop(should_stop)
        order_name = ORDER_NAMES.get(order, order)
        log(f"—— 开始搜索（排序：{order_name}，最多{max_pages}页，每页数量以接口返回为准）——")
        blocked = False
        no_new = 0
        for page in range(1, max_pages + 1):
            _check_stop(should_stop)
            items = client.search_page(keyword, order, page, should_stop=should_stop)
            if items is None:
                blocked = True
                break
            if not items:
                log(f"  第{page}页无更多结果，结束该排序")
                break
            new_unique, new_relevant = 0, 0
            for it in items:
                if not isinstance(it, dict) or not (it.get("bvid") or it.get("aid")):
                    continue
                v = extract_video(it)
                if v["bvid"] in candidates:   # 去重：不同排序/页码出现同一视频只保留一份
                    continue
                new_unique += 1
                if relevant(v["title"], v["desc"]):
                    candidates[v["bvid"]] = v
                    new_relevant += 1
            log(f"  第{page:>2}页：{len(items)}条 → 新增相关 {new_relevant} 条，相关视频累计 {len(candidates)} 条")
            if new_unique == 0:
                no_new += 1
                if no_new >= 2:
                    log("  连续2页没有新视频，结束该排序")
                    break
            else:
                no_new = 0
            _sleep(random.uniform(*PAGE_DELAY), should_stop)
        if blocked:
            blocked_orders += 1
            log(f"[警告] 「{order_name}」排序的搜索被风控拦截")
    return candidates, blocked_orders


def fetch_stats(client: BiliClient, candidates: dict, should_stop=None) -> int:
    """给每个候选视频补齐 点赞/收藏/评论/播放（每视频只请求一次，防止重复抓取）"""
    total = len(candidates)
    log(f"—— 获取 {total} 个相关视频的 点赞/收藏/评论/播放 数据（每个视频只请求一次）——")
    ok = 0
    for i, v in enumerate(candidates.values(), 1):
        _check_stop(should_stop)
        st = client.get_video_stat(v.get("bvid"), v.get("aid"), should_stop=should_stop)
        if "error" in st:
            v["note"] = f"详情获取失败：{st['error']}"
        else:
            v["like"] = st.get("like")
            v["view"] = st.get("view")
            if st.get("favorite") is not None:
                v["fav"] = st["favorite"]
            if st.get("reply") is not None:
                v["reply"] = st["reply"]
            if st.get("like") is None:
                v["note"] = "点赞数缺失"
            ok += 1
        if i % 10 == 0 or i == total:
            log(f"  进度：{i}/{total}（成功 {ok}）")
        _sleep(random.uniform(*STAT_DELAY), should_stop)
    return ok


def fetch_author_stats(client: BiliClient, candidates: dict, should_stop=None) -> int:
    """按作者去重补充粉丝数和最新视频时间，同一作者只请求一次。"""
    videos_by_mid = {}
    for video in candidates.values():
        mid = video.get("mid")
        if mid:
            videos_by_mid.setdefault(str(mid), []).append(video)
    total = len(videos_by_mid)
    log(f"—— 获取 {total} 位博主的粉丝量和最新视频时间（按博主去重）——")
    ok = 0
    for index, (mid, videos) in enumerate(videos_by_mid.items(), 1):
        _check_stop(should_stop)
        stats = client.get_author_stats(mid, should_stop=should_stop)
        for video in videos:
            video["author_fans"] = stats.get("author_fans")
            video["author_latest_pubdate"] = stats.get("author_latest_pubdate")
            if stats.get("error"):
                message = f"作者资料部分缺失：{stats['error']}"
                video["note"] = "；".join(filter(None, [video.get("note"), message]))
        if not stats.get("error"):
            ok += 1
        if index % 10 == 0 or index == total:
            log(f"  作者进度：{index}/{total}（完整成功 {ok}）")
        _sleep(random.uniform(*STAT_DELAY), should_stop)
    return ok


def _like_key(v):
    return v["like"] if isinstance(v.get("like"), int) else -1


def _play_key(v):
    for k in ("view", "play"):
        if isinstance(v.get(k), int):
            return v[k]
    return -1


def download_videos(videos, output_dir: Path, cookies=None, should_stop=None) -> dict:
    """真正下载视频文件（视频流+音频流自动合并为MP4）。
    videos 为 extract_video 产生的视频字典列表；复用当前会话 Cookie，支持断点续传。
    单个视频失败不会中断整批任务，返回成功/失败统计。"""
    try:
        import imageio_ffmpeg
        import yt_dlp
    except ImportError as e:
        raise RuntimeError("视频下载需要 yt-dlp 和 imageio-ffmpeg，请先执行："
                           "pip install yt-dlp imageio-ffmpeg") from e

    videos = list(videos)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    ffmpeg_location = imageio_ffmpeg.get_ffmpeg_exe()
    cookie_header = "; ".join(f"{c.name}={c.value}" for c in (cookies or []))
    downloaded_files = []
    failures = []
    total = len(videos)
    log(f"—— 开始下载 {total} 个视频文件，保存目录：{output_dir} ——")

    for index, video in enumerate(videos, 1):
        _check_stop(should_stop)
        title = video.get("title") or video.get("bvid") or "video"
        bvid = video.get("bvid") or f"video_{index}"
        url = video.get("url") or f"https://www.bilibili.com/video/{bvid}"
        safe_title = re.sub(r'[\\/:*?"<>|\r\n\t%]+', "_", title).strip()[:100] or bvid
        log(f"[下载] {index}/{total} 开始：{title}")

        ydl_opts = {
            # 最高取约480P（B站格式字段以画面宽度标记）；音视频分离时自动合并
            "format": "bv*[width<=854]+ba/b[width<=854]/b",
            "merge_output_format": "mp4",
            "ffmpeg_location": ffmpeg_location,
            "outtmpl": str(output_dir / f"{safe_title} [{bvid}].%(ext)s"),
            "http_headers": {
                "User-Agent": UA,
                "Referer": "https://www.bilibili.com/",
                **({"Cookie": cookie_header} if cookie_header else {}),
            },
            "proxy": "",                 # 与搜索模块一致：默认直连，避开失效系统代理
            "noplaylist": True,
            "quiet": True,
            "no_warnings": True,
            "noprogress": True,
            "progress_hooks": [lambda _status: _check_stop(should_stop)],
            "retries": 3,
            "fragment_retries": 3,
            "socket_timeout": 30,
            "continuedl": True,           # 支持断点续传
            "overwrites": False,          # 已下载的同名视频不重复下载
            "windowsfilenames": True,
            "restrictfilenames": False,
        }
        try:
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                ydl.download([url])
            files = [
                p for p in output_dir.iterdir()
                if p.is_file() and f"[{bvid}]" in p.name
                and p.suffix.lower() in (".mp4", ".mkv", ".webm", ".flv")
            ]
            if not files:
                raise RuntimeError("下载命令结束但未找到视频文件")
            file_path = max(files, key=lambda p: p.stat().st_mtime)
            downloaded_files.append(str(file_path))
            log(f"[下载] {index}/{total} 成功：{file_path.name}"
                f"（{file_path.stat().st_size / 1024 / 1024:.1f} MB）")
        except Exception as e:
            _check_stop(should_stop)
            error = str(e).strip().splitlines()[-1] if str(e).strip() else type(e).__name__
            failures.append({"bvid": bvid, "title": title, "error": error})
            log(f"[下载] {index}/{total} 失败：{title}；原因：{error}")

    result = {
        "total": total,
        "success_count": len(downloaded_files),
        "failed_count": len(failures),
        "files": downloaded_files,
        "failures": failures,
        "output_dir": str(output_dir),
    }
    log(f"[下载完成] 成功 {result['success_count']} 个，失败 {result['failed_count']} 个，"
        f"保存目录：{output_dir}")
    return result


# ====================== 导出 ======================
HEADERS = ["排名", "视频标题", "笔记链接(视频)", "点赞", "收藏", "评论", "播放量",
           "发布日期", "博主名字", "博主主页链接", "博主粉丝量", "博主最新视频发布时间",
           "BV号", "上榜情况", "备注"]
COL_WIDTHS = [6, 50, 45, 10, 10, 10, 12, 18, 16, 40, 14, 22, 22, 15, 30]
LINK_COLS = (3, 10)   # 第3列(视频链接)、第10列(博主主页)设置超链接


def build_row(rank, v, tag):
    pub = ""
    if isinstance(v.get("pubdate"), (int, float)) and v["pubdate"] > 0:
        pub = datetime.fromtimestamp(v["pubdate"]).strftime("%Y-%m-%d %H:%M")
    latest_pub = ""
    latest = v.get("author_latest_pubdate")
    if isinstance(latest, (int, float)) and latest > 0:
        latest_pub = datetime.fromtimestamp(latest).strftime("%Y-%m-%d %H:%M")
    play = v["view"] if isinstance(v.get("view"), int) else v.get("play")
    return [rank, v.get("title", ""), v.get("url", ""), v.get("like"), v.get("fav"),
            v.get("reply"), play, pub, v.get("author", ""), v.get("space", ""),
            v.get("author_fans"), latest_pub, v.get("bvid", ""), tag, v.get("note", "")]


def write_excel(path: Path, sheets):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    wb.remove(wb.active)
    hfont = Font(bold=True, color="FFFFFF")
    hfill = PatternFill("solid", fgColor="305496")
    lfont = Font(color="0563C1", underline="single")
    for name, rows in sheets:
        ws = wb.create_sheet(name[:31])
        ws.append(HEADERS)
        for idx, w in enumerate(COL_WIDTHS, 1):
            ws.column_dimensions[get_column_letter(idx)].width = w
        for cell in ws[1]:
            cell.font = hfont
            cell.fill = hfill
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for row in rows:
            ws.append(row)
        for r in ws.iter_rows(min_row=2):
            for col in LINK_COLS:
                cell = r[col - 1]
                if cell.value:
                    cell.hyperlink = str(cell.value)
                    cell.font = lfont
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    wb.save(str(path))


def csv_bytes(rows):
    """在内存中生成 UTF-8 BOM CSV，供 Streamlit 按需下载，不写入本地文件。"""
    buffer = io.StringIO(newline="")
    writer = csv.writer(buffer)
    writer.writerow(HEADERS)
    writer.writerows(rows)
    return buffer.getvalue().encode("utf-8-sig")


def preview(title, rows, n=5):
    log(f"—— {title}（前{min(n, len(rows))}条预览）——")
    for r in rows[:n]:
        log(f"  {r[0]:>3}. {str(r[1])[:30]} | 赞{fmt_num(r[3])} 藏{fmt_num(r[4])} "
            f"评{fmt_num(r[5])} 播{fmt_num(r[6])}")
    if len(rows) > n:
        log(f"  ... 其余 {len(rows) - n} 条见导出文件")


def build_rankings(candidates: dict, top_n: int):
    """对去重后的候选池排序，生成 点赞榜/播放榜/全部 三份行数据。
    两份榜单共用同一候选池，天然防重复；返回 (like_rows, play_rows, all_rows)"""
    like_top = sorted(candidates.values(), key=_like_key, reverse=True)[:top_n]
    play_top = sorted(candidates.values(), key=_play_key, reverse=True)[:top_n]
    like_ids = {v["bvid"] for v in like_top}
    play_ids = {v["bvid"] for v in play_top}

    def tag_of(v):
        a, b = v["bvid"] in like_ids, v["bvid"] in play_ids
        if a and b:
            return "点赞榜+播放榜"
        if a:
            return "仅点赞榜"
        if b:
            return "仅播放榜"
        return "未上榜(仅收录于全部)"

    like_rows = [build_row(i, v, tag_of(v)) for i, v in enumerate(like_top, 1)]
    play_rows = [build_row(i, v, tag_of(v)) for i, v in enumerate(play_top, 1)]
    all_sorted = sorted(candidates.values(), key=lambda x: (_like_key(x), _play_key(x)), reverse=True)
    all_rows = [build_row(i, v, tag_of(v)) for i, v in enumerate(all_sorted, 1)]
    return like_rows, play_rows, all_rows


def filter_candidates_by_fans(candidates: dict, min_fans: int, max_fans: int) -> dict:
    """仅保留粉丝数已知且位于闭区间内的视频，原候选池不受影响。"""
    return {key: video for key, video in candidates.items()
            if isinstance(video.get("author_fans"), int)
            and min_fans <= video["author_fans"] <= max_fans}


def export_files(out_dir: Path, keyword: str, top_n: int, like_rows, play_rows, all_rows,
                 filtered_like_rows=None, filtered_play_rows=None, fans_range=None) -> dict:
    """默认仅保存 Excel；区间榜单存在时放在原有工作表之前。"""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    safe_kw = re.sub(r'[\\/:*?"<>|\r\n\t ]+', "_", keyword) or "keyword"
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = out_dir / f"{safe_kw}_B站搜索_{ts}.xlsx"
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
        ("全部相关视频", all_rows),
    ])
    write_excel(path, sheets)
    return {"xlsx": path}


# ====================== 主流程 ======================
def parse_args():
    p = argparse.ArgumentParser(description="B站关键词搜索爬虫（点赞榜/播放榜 TOP N，默认保存 Excel）")
    p.add_argument("-k", "--keyword", default=None, help=f"搜索关键词（默认读取配置：{KEYWORD}）")
    p.add_argument("-o", "--output", default=None, help=f"输出目录（默认读取配置：{OUTPUT_DIR}）")
    p.add_argument("-n", "--top-n", type=int, default=None, help=f"每个榜单条数（默认：{TOP_N}）")
    p.add_argument("--filter-mode", choices=("phrase", "smart", "terms", "off"), default=None,
                   help=f"相关性过滤模式（默认读取配置：{FILTER_MODE}）")
    p.add_argument("--max-pages", type=int, default=None, help="每种排序最多翻页数（调试/提速用）")
    p.add_argument("--login", action="store_true", help="强制重新扫码登录（更换账号时使用）")
    p.add_argument("--no-login", action="store_true", help="跳过登录环节，直接匿名抓取")
    return p.parse_args()


def main():
    t_start = time.time()
    args = parse_args()
    keyword = (args.keyword or KEYWORD).strip()
    out_dir = Path(args.output or OUTPUT_DIR)
    top_n = args.top_n or TOP_N
    filter_mode = args.filter_mode or FILTER_MODE
    orders = list(SEARCH_ORDERS)
    if args.max_pages:
        orders = [(o, min(p, args.max_pages)) for o, p in orders]

    log("=" * 60)
    log("B站关键词搜索爬虫 启动")
    log(f"  关键词  ：{keyword}")
    log(f"  输出目录：{out_dir}")
    log(f"  榜单条数：TOP {top_n}（点赞榜 / 播放榜）")
    log(f"  过滤模式：{filter_mode}")
    log("=" * 60)

    try:
        out_dir.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        log(f"[错误] 无法创建输出目录 {out_dir}：{e}")
        sys.exit(1)

    client = BiliClient(get_user_cookie_file())
    client.load_cookies()
    client.ensure_buvid()

    # ---------- 登录状态检查 ----------
    if args.login:
        log("已指定 --login：强制重新扫码登录")
        if not client.qr_login():
            log("[警告] 未完成登录，将继续匿名抓取（可能被限制或数据不全）")
    else:
        logged, uname = client.check_login()
        if logged:
            log(f"[成功] B站账号已登录：{uname}")
        else:
            if client.has_sessdata():
                log("[警告] 检测到本地保存的登录凭证(SESSDATA)已过期，需要重新扫码登录！")
                for name in ("SESSDATA", "bili_jct", "DedeUserID", "DedeUserID__ckMd5"):
                    try:
                        client.session.cookies.set(name, "", domain=".bilibili.com")
                    except Exception:
                        pass
            else:
                log("[提示] 当前未登录B站（可匿名抓取，但可能被限制或数据不全）")
            if args.no_login:
                log("[提示] 已指定 --no-login：跳过登录，匿名抓取")
            else:
                try:
                    ans = input(">> 是否现在扫码登录？(回车=扫码登录 / n=匿名继续): ").strip().lower()
                except (EOFError, KeyboardInterrupt):
                    ans = "n"
                    print()
                if ans in ("", "y", "yes"):
                    if not client.qr_login():
                        log("[警告] 未完成登录，将继续匿名抓取（可能被限制或数据不全）")
                else:
                    log("[提示] 已选择匿名抓取；如遇风控可重新运行并选择扫码登录")
    client.save_cookies()

    # ---------- 搜索 ----------
    relevant = make_filter(keyword, filter_mode)
    candidates = {}
    blocked_orders = 0
    interrupted = False
    try:
        candidates, blocked_orders = crawl_all(client, keyword, orders, relevant)
    except KeyboardInterrupt:
        interrupted = True
        print()
        log("[警告] 搜索阶段被手动中断")

    if not candidates:
        if orders and blocked_orders == len(orders):
            log("[错误] 搜索请求全部被B站风控拦截，未获取到数据。可尝试：")
            log("   1) 等待10~30分钟后重新运行；")
            log("   2) 重新运行并选择扫码登录（或加 --login）；")
            log("   3) 增大脚本顶部 PAGE_DELAY（如 (2.0, 4.0)）降低请求频率；")
            log("   4) 更换网络环境后重试。")
        else:
            log(f"[错误] 关键词「{keyword}」按当前过滤模式({filter_mode})没有命中任何相关视频。可尝试：")
            log("   1) 将脚本顶部 FILTER_MODE 改为 \"smart\" 或 \"terms\"（放宽过滤）；")
            log("   2) 检查关键词拼写，或用 --keyword 换个说法再试。")
        sys.exit(1)

    # ---------- 视频详情（点赞/收藏/评论/播放） ----------
    if not interrupted:
        try:
            fetch_stats(client, candidates)
            fetch_author_stats(client, candidates)
        except KeyboardInterrupt:
            interrupted = True
            print()
            log("[警告] 详情抓取被手动中断，将导出已获取的数据（缺失字段为空）")

    # ---------- 排序 + 导出（候选池已按视频去重，两份榜单共用同一份数据） ----------
    like_rows, play_rows, all_rows = build_rankings(candidates, top_n)
    try:
        paths = export_files(out_dir, keyword, top_n, like_rows, play_rows, all_rows)
    except Exception as e:
        log(f"[错误] 导出失败：{e}")
        sys.exit(1)
    xlsx_path = paths["xlsx"]

    # ---------- 汇总 ----------
    log("=" * 60)
    log("抓取完成！结果汇总：")
    log(f"  关键词「{keyword}」，过滤模式 {filter_mode}，命中相关视频 {len(candidates)} 个（已去重）")
    log(f"  点赞榜：{len(like_rows)} 条    播放榜：{len(play_rows)} 条    （目标 TOP {top_n}）")
    if len(candidates) < top_n:
        log(f"  [提示] 相关视频数量不足 {top_n} 个；如需更多结果，可将 FILTER_MODE 改为 smart/terms，")
        log("         或增大 SEARCH_ORDERS 中的页数")
    log(f"  Excel：{xlsx_path}")
    log(f"  耗时：{time.time() - t_start:.0f} 秒")
    log("=" * 60)

    preview(f"点赞榜 TOP{len(like_rows)}", like_rows)
    preview(f"播放榜 TOP{len(play_rows)}", play_rows)
    if interrupted:
        log("[警告] 本次运行曾被中断，以上为部分数据")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print()
        log("已退出。")
        sys.exit(130)
