# -*- coding: utf-8 -*-
"""B站 / 抖音关键词搜索爬虫 - Streamlit 可视化界面

运行：python -m streamlit run bilibili_crawler_app.py
支持：平台切换、粉丝区间榜单、扫码登录、后台爬取、视频下载、Excel保存及CSV按需导出。
登录 Cookie 仅保存在当前操作系统用户的私有目录，不随代码或结果文件分享。
"""

import queue
import sys
import threading
import time
import traceback
from pathlib import Path

import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parent))
import bilibili_search_crawler as bili  # noqa: E402
import douyin_search_crawler as douyin  # noqa: E402

st.set_page_config(page_title="视频关键词搜索爬虫", page_icon="📺", layout="wide")

PLATFORM_BILI = "B站"
PLATFORM_DOUYIN = "抖音"
FILTER_LABELS = {
    "phrase": "严格 - 标题/简介含完整关键词（推荐）",
    "smart": "智能 - 完整关键词 或 同义词组全部命中",
    "terms": "分词 - 关键词拆词后每个词都须命中",
    "off": "不过滤 - 平台原始结果（噪音多）",
}
LOGIN_ANON = "匿名抓取（无需登录）"
LOGIN_SESSION = "当前浏览器会话（无需扫码）"
LOGIN_QR = "扫码登录（更稳定）"
COOKIE_PATH = bili.get_user_cookie_file()


class LoginFailed(Exception):
    """扫码登录未完成，禁止继续爬取。"""


# ---------------- 会话与后台任务 ----------------
def get_client() -> bili.BiliClient:
    key = f"bili_client::{COOKIE_PATH}"
    if key not in st.session_state:
        client = bili.BiliClient(COOKIE_PATH)
        client.load_cookies()
        st.session_state[key] = client
    return st.session_state[key]


def init_state():
    defaults = {
        "job_running": False,
        "job_status": "idle",
        "job_queue": None,
        "job_thread": None,
        "job_platform": None,
        "stop_event": None,
        "job_logs": [],
        "job_error": "",
        "job_qr": None,
        "job_started_at": None,
        "login_uname": None,
        "last_result": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def reset_platform_view():
    """平台切换后清空上一平台的页面状态，但保留各自的输入控件值。"""
    st.session_state.job_status = "idle"
    st.session_state.job_platform = None
    st.session_state.job_logs = []
    st.session_state.job_error = ""
    st.session_state.job_qr = None
    st.session_state.login_uname = None
    st.session_state.last_result = None


def make_result(platform, headers, config, candidates, like_rows, play_rows, all_rows,
                paths, download_summary, missing_play=0, fans_range=None,
                filtered_total=0, filtered_like_rows=None, filtered_play_rows=None):
    return {
        "platform": platform,
        "headers": headers,
        "keyword": config["keyword"],
        "filter_mode": config.get("filter_mode"),
        "total": len(candidates),
        "top_n": config["top_n"],
        "like_rows": like_rows,
        "play_rows": play_rows,
        "all_rows": all_rows,
        "fans_range": fans_range,
        "filtered_total": filtered_total,
        "filtered_like_rows": filtered_like_rows,
        "filtered_play_rows": filtered_play_rows,
        "paths": {k: str(v) for k, v in paths.items()},
        "download": download_summary,
        "missing_play": missing_play,
    }


def build_fan_rankings(module, candidates, config):
    """原榜单不变，仅为启用的粉丝区间生成额外点赞榜和播放榜。"""
    if not config.get("fans_filter_enabled"):
        return 0, None, None
    min_fans, max_fans = config["fans_min"], config["fans_max"]
    filtered = module.filter_candidates_by_fans(candidates, min_fans, max_fans)
    like_rows, play_rows, _ = module.build_rankings(filtered, config["top_n"])
    module.log(f"[提示] 粉丝区间 {min_fans}-{max_fans} 内共有 {len(filtered)} 个视频，"
               "已额外生成区间点赞榜和播放榜")
    return len(filtered), like_rows, play_rows


def csv_file_name(res, ranking_name, filtered=False):
    prefix = "粉丝区间" if filtered else ""
    return (f"{Path(res['paths']['xlsx']).stem}_{prefix}{ranking_name}"
            f"TOP{res['top_n']}.csv")


def run_bili_task(config, client, should_stop, emit):
    bili._check_stop(should_stop)
    bili.log("[提示] 正在初始化B站网络会话并检查登录状态 ...")
    client.ensure_buvid()
    logged, uname = client.check_login()

    if config["login_mode"] == LOGIN_QR:
        if config["force_relogin"] or not logged:
            if not logged and client.has_sessdata():
                bili.log("[警告] 本地登录凭证已失效，请重新扫码")

            def show_qr(png_path, _qr_url):
                emit("qr", str(png_path) if png_path else None)

            if not client.qr_login(qr_callback=show_qr, should_stop=should_stop):
                bili._check_stop(should_stop)
                raise LoginFailed("未完成扫码登录，本次已取消爬取；请重新扫码并完成手机确认")
            emit("qr", None)
            logged, uname = client.check_login()
        if not logged:
            raise LoginFailed("登录状态验证失败，本次不会开始爬取")
        bili.log(f"[成功] 登录状态有效：{uname}，即将开始爬取")
        emit("login", uname)
    else:
        bili.log("[提示] B站匿名抓取模式")

    client.save_cookies()
    bili._check_stop(should_stop)
    relevant = bili.make_filter(config["keyword"], config["filter_mode"])
    candidates, blocked_orders = bili.crawl_all(
        client, config["keyword"], config["orders"], relevant, should_stop=should_stop)
    if not candidates:
        if config["orders"] and blocked_orders == len(config["orders"]):
            raise RuntimeError("搜索请求全部被B站风控拦截，请稍后重试或改用扫码登录")
        raise RuntimeError(
            f"关键词「{config['keyword']}」在当前过滤模式下没有相关视频，"
            "可尝试 smart/terms 模式或更换关键词")

    bili.fetch_stats(client, candidates, should_stop=should_stop)
    bili.fetch_author_stats(client, candidates, should_stop=should_stop)
    like_rows, play_rows, all_rows = bili.build_rankings(candidates, config["top_n"])
    filtered_total, filtered_like_rows, filtered_play_rows = build_fan_rankings(
        bili, candidates, config)
    fans_range = ((config["fans_min"], config["fans_max"])
                  if config.get("fans_filter_enabled") else None)
    download_summary = None
    if config["download_mode"] == "是":
        selected_ids = dict.fromkeys(row[12] for row in like_rows + play_rows)
        selected_videos = [candidates[bvid] for bvid in selected_ids if bvid in candidates]
        download_summary = bili.download_videos(
            selected_videos, Path(config["video_output_dir"]),
            client.session.cookies, should_stop=should_stop)

    bili._check_stop(should_stop)
    paths = bili.export_files(
        Path(config["output_dir"]), config["keyword"], config["top_n"],
        like_rows, play_rows, all_rows, filtered_like_rows, filtered_play_rows, fans_range)
    return make_result(PLATFORM_BILI, bili.HEADERS, config, candidates,
                       like_rows, play_rows, all_rows, paths, download_summary,
                       fans_range=fans_range, filtered_total=filtered_total,
                       filtered_like_rows=filtered_like_rows,
                       filtered_play_rows=filtered_play_rows)


def run_douyin_task(config, should_stop, emit):
    douyin._check_stop(should_stop)
    with douyin.DouyinBrowser(douyin.get_profile_dir(), headless=False) as browser:
        if config["login_mode"] == LOGIN_QR:
            if not browser.wait_for_login(
                    force=config["force_relogin"], should_stop=should_stop):
                raise LoginFailed("未完成抖音扫码登录，本次已取消爬取")
            emit("login", "抖音账号")
        else:
            douyin.log("[提示] 直接使用当前抖音浏览器会话，不主动要求扫码")

        douyin._check_stop(should_stop)
        candidates = douyin.crawl_all(
            browser, config["keyword"], config["max_scrolls"], should_stop=should_stop)
        if not candidates:
            raise RuntimeError(
                f"关键词「{config['keyword']}」没有获取到搜索结果。请检查关键词和页面筛选状态，"
                "或改用扫码登录后在浏览器中完成安全验证")

        douyin.fetch_author_stats(browser, candidates, should_stop=should_stop)
        browser.save_cookies()
        missing_play = sum(not isinstance(v.get("play"), int) for v in candidates.values())
        if missing_play:
            douyin.log(f"[提示] {missing_play}/{len(candidates)} 个结果未返回公开播放量；"
                       "对应单元格留空，播放榜只能按已返回的数据排序")
        like_rows, play_rows, all_rows = douyin.build_rankings(candidates, config["top_n"])
        filtered_total, filtered_like_rows, filtered_play_rows = build_fan_rankings(
            douyin, candidates, config)
        fans_range = ((config["fans_min"], config["fans_max"])
                      if config.get("fans_filter_enabled") else None)

        download_summary = None
        if config["download_mode"] == "是":
            selected_ids = dict.fromkeys(row[12] for row in like_rows + play_rows)
            selected_videos = [candidates[video_id] for video_id in selected_ids
                               if video_id in candidates]
            download_summary = douyin.download_videos(
                selected_videos, Path(config["video_output_dir"]), browser.cookie_file,
                browser.user_agent(), use_system_proxy=douyin.USE_SYSTEM_PROXY,
                should_stop=should_stop)

        douyin._check_stop(should_stop)
        paths = douyin.export_files(
            Path(config["output_dir"]), config["keyword"], config["top_n"],
            like_rows, play_rows, all_rows, filtered_like_rows, filtered_play_rows, fans_range)
        return make_result(PLATFORM_DOUYIN, douyin.HEADERS, config, candidates,
                           like_rows, play_rows, all_rows, paths, download_summary,
                           missing_play=missing_play, fans_range=fans_range,
                           filtered_total=filtered_total,
                           filtered_like_rows=filtered_like_rows,
                           filtered_play_rows=filtered_play_rows)


def run_job(config, client, events, stop_event):
    """后台线程只通过队列与页面通信，不在子线程直接调用 Streamlit。"""
    final_state = "completed"
    error_message = ""
    should_stop = stop_event.is_set
    module = bili if config["platform"] == PLATFORM_BILI else douyin

    def emit(kind, payload=None):
        events.put((kind, payload))

    def log_hook(line):
        emit("log", line)

    module.add_log_hook(log_hook)
    try:
        if config["platform"] == PLATFORM_BILI:
            result = run_bili_task(config, client, should_stop, emit)
        else:
            result = run_douyin_task(config, should_stop, emit)
        emit("result", result)
        module.log("=" * 60)
        module.log("[完成] 已完成所有任务")
        module.log(f"[完成] Excel：{result['paths']['xlsx']}")
        download_summary = result.get("download")
        if download_summary is not None:
            module.log(f"[完成] 视频文件：成功 {download_summary['success_count']} 个，"
                       f"失败 {download_summary['failed_count']} 个；"
                       f"保存目录：{download_summary['output_dir']}")
    except (bili.CrawlStopped, douyin.CrawlStopped):
        final_state = "stopped"
        module.log("[停止] 已停止当前任务，未执行后续步骤")
    except LoginFailed as exc:
        final_state = "login_failed"
        error_message = str(exc)
        module.log(f"[错误] {error_message}")
    except Exception:
        final_state = "error"
        error_message = traceback.format_exc()
        module.log("[错误] 任务异常终止，详细错误已显示在页面上")
    finally:
        module.remove_log_hook(log_hook)
        if client is not None:
            try:
                client.save_cookies()
            except Exception:
                pass
        emit("done", {"state": final_state, "error": error_message})


def drain_events():
    events = st.session_state.job_queue
    if events is None:
        return
    while True:
        try:
            kind, payload = events.get_nowait()
        except queue.Empty:
            break
        if kind == "log":
            st.session_state.job_logs.append(payload)
            st.session_state.job_logs = st.session_state.job_logs[-500:]
        elif kind == "qr":
            st.session_state.job_qr = payload
        elif kind == "login":
            st.session_state.login_uname = payload
        elif kind == "result":
            st.session_state.last_result = payload
        elif kind == "done":
            st.session_state.job_running = False
            st.session_state.job_status = payload["state"]
            st.session_state.job_error = payload["error"]
            st.session_state.job_qr = None


init_state()

# ---------------- 标题、停止按钮与平台切换 ----------------
stop_col, title_col = st.columns([1, 6], vertical_alignment="center")
stop_clicked = stop_col.button(
    "停止", key="stop_button", width="stretch",
    disabled=not st.session_state.job_running or st.session_state.job_status == "stopping")
title_col.title("视频关键词搜索爬虫")
if stop_clicked and st.session_state.stop_event is not None:
    st.session_state.stop_event.set()
    st.session_state.job_status = "stopping"
    module = bili if st.session_state.job_platform == PLATFORM_BILI else douyin
    module.log("[停止] 收到停止请求，正在安全停止当前任务 ...")

drain_events()
platform = st.radio(
    "选择爬取平台", [PLATFORM_BILI, PLATFORM_DOUYIN], index=0, horizontal=True,
    key="platform_selector", on_change=reset_platform_view,
    disabled=st.session_state.job_running)
module = bili if platform == PLATFORM_BILI else douyin
platform_key = "bili" if platform == PLATFORM_BILI else "douyin"
if platform == PLATFORM_BILI:
    st.caption("按关键词搜索B站视频 → 过滤无关内容 → 点赞榜 / 播放榜 → 可选视频下载 → 导出表格")
else:
    st.caption("搜索框输入关键词 → 筛选“最多点赞” → 保留全部搜索结果并按点赞降序排列 → "
               "可选视频下载 → 导出表格；扫码和安全验证在弹出的浏览器中完成")

# ---------------- 侧边栏配置 ----------------
with st.sidebar:
    st.header(f"⚙️ {platform}抓取配置")
    keyword = st.text_input(
        "搜索关键词", value="", key=f"keyword_input_{platform_key}",
        placeholder="例如：拍摄设备", help="输入完整的视频主题关键词").strip()
    output_dir = st.text_input(
        "保存路径", value="", key=f"output_dir_input_{platform_key}",
        placeholder="例如：D:\\VideoResults", help="Excel 的保存目录；CSV 仅在点击下载时生成").strip()
    top_n = int(st.number_input(
        "每个榜单条数（TOP N）", min_value=1, value=1, step=1, format="%d",
        placeholder="请输入任意正整数，例如：30", key=f"top_n_input_{platform_key}",
        help="默认1，可填写任意正整数，不再限制固定档位"))
    filter_mode = None
    if platform == PLATFORM_BILI:
        filter_mode = st.selectbox(
            "相关性过滤模式", list(FILTER_LABELS), index=0, format_func=lambda k: FILTER_LABELS[k],
            placeholder="请选择过滤模式", key="filter_mode_input_bili",
            help="严格模式只保留标题或简介包含完整关键词的视频")
        crawl_label = "每种排序最多抓取页数"
        crawl_help = "每次请求返回的视频数量由B站接口决定；页数越大覆盖越广，耗时也越长"
    else:
        crawl_label = "“最多点赞”结果最多滚动轮数（每轮约10条）"
        crawl_help = "自动点击筛选并选择最多点赞；不做相关性过滤，人工验证不占用轮数"
    max_crawl_input = st.number_input(
        crawl_label, min_value=1, max_value=50, value=None, step=1, format="%d",
        placeholder="15", key=f"max_crawl_input_{platform_key}", help=crawl_help)
    max_crawl = int(max_crawl_input or 15)

    fans_filter_enabled = st.checkbox(
        "启用粉丝区间筛选（生成额外榜单）", key=f"fans_filter_enabled_{platform_key}",
        help="默认不启用；原榜单始终统计全部视频，区间榜单会排除区间外及粉丝数未知的视频")
    fans_min, fans_max = 0, 0
    if fans_filter_enabled:
        fans_col1, fans_col2 = st.columns(2)
        fans_min = int(fans_col1.number_input(
            "最少粉丝数", min_value=0, value=10000, step=1000, format="%d",
            key=f"fans_min_input_{platform_key}"))
        fans_max = int(fans_col2.number_input(
            "最多粉丝数", min_value=0, value=100000, step=1000, format="%d",
            key=f"fans_max_input_{platform_key}"))
        if fans_min > fans_max:
            st.error("最少粉丝数不能大于最多粉丝数")

    login_options = ([LOGIN_ANON, LOGIN_QR] if platform == PLATFORM_BILI
                     else [LOGIN_SESSION, LOGIN_QR])
    login_mode = st.radio(
        "登录方式", login_options, key=f"login_mode_input_{platform_key}",
        help=("扫码登录成功后才会爬取；未扫码或验证失败会立即取消任务"
              if platform == PLATFORM_BILI else
              "当前会话会复用浏览器Cookie；选择扫码时会打开浏览器，登录状态也会自动保存"))
    force_relogin = False
    if login_mode == LOGIN_QR:
        force_relogin = st.checkbox(
            "强制重新扫码登录（更换账号时勾选）", key=f"force_relogin_{platform_key}")

    download_mode = st.radio(
        "是否爬取视频", ["否", "是"], index=0, horizontal=True,
        key=f"download_mode_input_{platform_key}",
        help="选择“是”会下载双榜合并去重后的真实视频文件")
    video_output_dir = ""
    if download_mode == "是":
        video_output_dir = st.text_input(
            "视频保存路径", value=str(Path(output_dir) / "videos") if output_dir else "",
            key=f"video_output_dir_input_{platform_key}", placeholder="例如：D:\\VideoResults\\videos",
            help="真实 MP4 视频文件的保存目录").strip()

    st.divider()
    valid_config = bool(keyword and output_dir
                        and (download_mode == "否" or video_output_dir)
                        and (not fans_filter_enabled or fans_min <= fans_max))
    run_btn = st.button(
        "🚀 开始爬取", type="primary", width="stretch", key=f"run_button_{platform_key}",
        disabled=not valid_config or st.session_state.job_running)
    st.caption("ℹ️ 任务在后台运行，可点击标题左侧“停止”取消")
    st.caption("🔒 Cookie 保存在当前用户私有目录，不随代码或表格分享")

    client = None
    if platform == PLATFORM_BILI:
        client = get_client()
        if st.session_state.login_uname:
            st.success(f"已登录：{st.session_state.login_uname}")
        elif client.has_sessdata():
            st.info("已保存登录 Cookie，有效性将在爬取时检查")
        else:
            st.info("当前未登录（匿名可用）")
    else:
        douyin_cookie_file = douyin.get_profile_dir().parent / "douyin_cookies.txt"
        if st.session_state.login_uname:
            st.success("已检测到抖音登录状态")
        elif douyin_cookie_file.exists():
            st.info("检测到已保存的浏览器 Cookie，有效性将在爬取时检查")
        else:
            st.info("尚未保存抖音 Cookie；可直接抓取或选择扫码登录")

# ---------------- 启动后台任务 ----------------
if run_btn and not st.session_state.job_running:
    try:
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        if download_mode == "是":
            Path(video_output_dir).mkdir(parents=True, exist_ok=True)
    except Exception as exc:
        st.error(f"无法创建保存目录：{exc}")
    else:
        events = queue.Queue()
        stop_event = threading.Event()
        config = {
            "platform": platform,
            "keyword": keyword,
            "output_dir": output_dir,
            "top_n": top_n,
            "filter_mode": filter_mode,
            "orders": [(o, min(p, max_crawl)) for o, p in bili.SEARCH_ORDERS],
            "max_scrolls": max_crawl,
            "fans_filter_enabled": fans_filter_enabled,
            "fans_min": fans_min,
            "fans_max": fans_max,
            "login_mode": login_mode,
            "force_relogin": force_relogin,
            "download_mode": download_mode,
            "video_output_dir": video_output_dir,
        }
        st.session_state.job_queue = events
        st.session_state.stop_event = stop_event
        st.session_state.job_platform = platform
        st.session_state.job_logs = []
        st.session_state.job_error = ""
        st.session_state.job_qr = None
        st.session_state.job_status = "running"
        st.session_state.job_running = True
        st.session_state.job_started_at = time.time()
        st.session_state.login_uname = None
        st.session_state.last_result = None
        worker = threading.Thread(
            target=run_job, args=(config, client, events, stop_event),
            name=f"{platform_key}-crawl-job", daemon=True)
        st.session_state.job_thread = worker
        worker.start()

# ---------------- 当前任务状态 ----------------
if st.session_state.job_running:
    elapsed = int(time.time() - (st.session_state.job_started_at or time.time()))
    label = "正在停止任务 ..." if st.session_state.job_status == "stopping" else f"任务运行中（{elapsed} 秒）"
    st.status(label, expanded=True, state="running")
    if (st.session_state.job_platform == PLATFORM_DOUYIN):
        st.info("请在弹出的浏览器中完成抖音扫码或安全验证；脚本检测到完成后会自动继续")
    if st.session_state.job_qr and Path(st.session_state.job_qr).exists():
        st.image(st.session_state.job_qr, width=300,
                 caption="二维码固定在这里：请用B站App扫码，选择登录环境并完成最终确认")
    if st.session_state.job_logs:
        st.code("\n".join(st.session_state.job_logs[-300:]), language=None)
elif st.session_state.job_status == "stopped":
    st.warning("🛑 任务已停止，未执行后续步骤")
    if st.session_state.job_logs:
        with st.expander("查看本次任务日志"):
            st.code("\n".join(st.session_state.job_logs[-300:]), language=None)
elif st.session_state.job_status in ("login_failed", "error"):
    st.error(st.session_state.job_error or "任务未完成")
    if st.session_state.job_logs:
        with st.expander("查看本次任务日志"):
            st.code("\n".join(st.session_state.job_logs[-300:]), language=None)

# ---------------- 结果展示 ----------------
res = st.session_state.last_result
if res:
    st.divider()
    st.subheader(f"📊 {res['platform']}抓取结果：{res['keyword']}")
    c1, c2, c3, c4 = st.columns(4)
    is_douyin = res["platform"] == PLATFORM_DOUYIN
    c1.metric("搜索结果（去重后）" if is_douyin else "相关视频（去重后）", res["total"])
    c2.metric(f"点赞榜 TOP{res['top_n']}", len(res["like_rows"]))
    c3.metric(f"播放榜 TOP{res['top_n']}", len(res["play_rows"]))
    c4.metric("筛选排序" if is_douyin else "过滤模式",
              "最多点赞（不过滤）" if is_douyin else res["filter_mode"])
    if res["total"] < res["top_n"]:
        if is_douyin:
            st.info(f"搜索结果共 {res['total']} 个，不足 TOP{res['top_n']}；可增大滚动轮数")
        else:
            st.info(f"相关视频共 {res['total']} 个，不足 TOP{res['top_n']}；"
                    "可尝试更宽松的过滤模式或增大抓取页数")
    if res.get("missing_play"):
        st.info(f"本次有 {res['missing_play']} 个抖音结果未公开播放量，对应单元格已留空")
    if res.get("fans_range"):
        min_fans, max_fans = res["fans_range"]
        st.info(f"粉丝区间 {min_fans}-{max_fans} 内共 {res['filtered_total']} 个视频；"
                "区间榜单排在原榜单之前，原榜单仍统计全部视频")

    download = res.get("download")
    if download:
        message = (f"视频文件下载：成功 {download['success_count']} 个，"
                   f"失败 {download['failed_count']} 个；保存路径：{download['output_dir']}")
        if download["failed_count"]:
            st.warning(message)
            with st.expander("查看下载失败的视频"):
                st.dataframe(pd.DataFrame(download["failures"]), width="stretch", hide_index=True)
        else:
            st.success(message)

    link_cfg = {
        "笔记链接(视频)": st.column_config.LinkColumn("视频链接", display_text="打开视频"),
        "博主主页链接": st.column_config.LinkColumn("博主主页", display_text="打开主页"),
    }
    tab_specs = []
    if res.get("fans_range"):
        min_fans, max_fans = res["fans_range"]
        range_label = f"粉丝{min_fans}-{max_fans}"
        tab_specs.extend([
            (f"👍 {range_label} 点赞榜 TOP{res['top_n']}", res["filtered_like_rows"]),
            (f"▶️ {range_label} 播放榜 TOP{res['top_n']}", res["filtered_play_rows"]),
        ])
    tab_specs.extend([
        (f"👍 点赞榜 TOP{res['top_n']}", res["like_rows"]),
        (f"▶️ 播放榜 TOP{res['top_n']}", res["play_rows"]),
        (f"📚 {'全部搜索结果' if is_douyin else '全部相关视频'}（{res['total']}）",
         res["all_rows"]),
    ])
    tabs = st.tabs([label for label, _ in tab_specs])
    for tab, (_, rows) in zip(tabs, tab_specs):
        with tab:
            st.dataframe(pd.DataFrame(rows, columns=res["headers"]),
                         width="stretch", height=520, hide_index=True, column_config=link_cfg)

    st.subheader("📁 导出文件（Excel 已保存；CSV 点击下载时生成）")
    paths = res["paths"]
    result_module = douyin if is_douyin else bili
    try:
        if res.get("fans_range"):
            f1, f2 = st.columns(2)
            f1.download_button(
                "⬇️ 粉丝区间点赞榜 CSV", data=result_module.csv_bytes(res["filtered_like_rows"]),
                file_name=csv_file_name(res, "点赞榜", filtered=True),
                mime="text/csv", width="stretch")
            f2.download_button(
                "⬇️ 粉丝区间播放榜 CSV", data=result_module.csv_bytes(res["filtered_play_rows"]),
                file_name=csv_file_name(res, "播放榜", filtered=True),
                mime="text/csv", width="stretch")
        d1, d2, d3 = st.columns(3)
        sheet_count = 5 if res.get("fans_range") else 3
        d1.download_button(
            f"⬇️ Excel（{sheet_count}个工作表）", data=Path(paths["xlsx"]).read_bytes(),
            file_name=Path(paths["xlsx"]).name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch")
        d2.download_button(
            "⬇️ 点赞榜 CSV", data=result_module.csv_bytes(res["like_rows"]),
            file_name=csv_file_name(res, "点赞榜"), mime="text/csv", width="stretch")
        d3.download_button(
            "⬇️ 播放榜 CSV", data=result_module.csv_bytes(res["play_rows"]),
            file_name=csv_file_name(res, "播放榜"), mime="text/csv", width="stretch")
        st.caption("Excel 保存位置：" + paths["xlsx"])
        st.success("✅ 已完成所有任务")
    except FileNotFoundError as exc:
        st.warning(f"导出文件已被移动或删除：{exc}")
elif st.session_state.job_status == "idle":
    st.info("👈 在左侧填写配置，点击【🚀 开始爬取】")

# 后台任务运行期间仅刷新页面，不阻塞停止按钮、二维码或浏览器人工操作提示
if st.session_state.job_running:
    time.sleep(1)
    st.rerun()
