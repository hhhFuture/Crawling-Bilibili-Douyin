# B站 / 抖音关键词搜索爬虫

一个基于 Python 3.10 和 Streamlit 的关键词视频搜索与统计工具，支持 B站和抖音。程序可以生成点赞榜、播放榜、作者资料、粉丝区间榜单，并支持 Excel 保存、CSV 按需导出及可选视频下载。



![](..\推广\imgs\BilBil.png)



## 主要功能

### B站

- 按关键词搜索视频，支持综合排序和播放量优先等搜索方式。
- 提供相关性过滤：完整关键词、同义词组、分词匹配或不过滤。
- 获取视频点赞、收藏、评论、播放量、发布日期等信息。
- 获取作者粉丝量和最新视频发布时间。
- 支持匿名抓取和扫码登录。

### 抖音

- 模拟人工在搜索框输入关键词。
- 保持“综合”搜索类型，展开“筛选”并选择“最多点赞”。
- 不做标题或简介相关性过滤，保留平台返回的搜索结果。
- 获取视频点赞、收藏、评论、发布日期、作者及链接等信息。
- 自动打开作者主页，获取粉丝量和最新视频发布时间。
- 支持复用浏览器会话、扫码登录和人工安全验证。
- 优先使用 Google Chrome，无法启动时自动尝试 Microsoft Edge。

### 两平台通用

- 点赞榜和播放榜分别按对应数据降序排列。
- 按作者去重获取作者资料，避免同一作者重复请求。
- 最新视频时间取作者首页多个作品中的最大时间戳，避免将置顶旧视频误判为最新作品。
- 可选粉丝区间筛选，额外生成区间点赞榜和区间播放榜，原榜单不受影响。
- 默认仅在本地保存一个 Excel 文件，CSV 由用户在 Streamlit 页面点击下载时生成。
- 支持后台运行、随时停止、日志显示和可选视频下载。

## 项目结构

```text
├─ bilibili_crawler_app.py       # Streamlit 双平台可视化界面
├─ bilibili_search_crawler.py   # B站搜索、作者资料、排序及导出逻辑
├─ douyin_search_crawler.py     # 抖音浏览器操作、搜索、作者资料及导出逻辑
├─ requestment.txt              # Python 依赖清单
├─ 启动爬虫服务.bat              # Windows Conda 环境启动脚本
└─ README.md
```

## 环境要求

- Windows（当前代码主要在 Windows 下验证）。
- Python 3.10，推荐使用 Conda 管理环境。
- 抖音功能需要以下浏览器之一：
  - Google Chrome；
  - Microsoft Edge；
  - Playwright Chromium（需要手动安装）。
- 正常访问 B站、抖音及相关接口的网络环境。

## 安装

### 1. 克隆仓库

```bash
git clone <你的 GitHub 仓库地址>
cd <仓库目录>
```

### 2. 创建并激活 Conda 环境

```bash
conda create -n py310 python=3.10 -y
conda activate py310
```

### 3. 安装依赖

```bash
python -m pip install -r requestment.txt
```

注意：本项目的依赖文件名是 `requestment.txt`，不是常见的 `requirements.txt`。

主要依赖包括 Streamlit、Requests、Playwright、pandas、openpyxl、yt-dlp 和 imageio-ffmpeg。

### 4. 可选：安装 Playwright Chromium

只有在将 `douyin_search_crawler.py` 中的 `BROWSER_CHANNEL` 改为 `None` 时才需要执行：

```bash
python -m playwright install chromium
```

如果电脑已经安装 Chrome 或 Edge，通常不需要执行此步骤。

## 启动 Streamlit 界面

在 Conda 环境中进入项目目录，执行：

```bash
python -m streamlit run bilibili_crawler_app.py
```

启动后根据终端提示访问页面，通常为：

```text
http://localhost:8501
```

### Windows 启动脚本

也可以双击 `启动爬虫服务.bat`。

该脚本默认使用：

```text
D:\Miniconda3\envs\py310\python.exe
```

如果你的 Conda 安装目录或环境名称不同，请先编辑脚本中的：

```bat
set "CONDA_ROOT=D:\Miniconda3"
set "ENV_NAME=py310"
```

## Streamlit 使用说明

### 1. 基础配置

1. 选择“B站”或“抖音”。
2. 输入搜索关键词和 Excel 保存路径。
3. 设置每个榜单的 TOP N 条数。
4. 设置最大抓取页数或滚动轮数。未填写时默认按 15 处理，输入框中的 15 仅为占位提示。
5. 根据需要选择登录方式、是否下载视频等配置。
6. 点击“开始爬取”。任务运行期间可以查看日志或点击“停止”。

搜索关键词和保存路径默认留空，示例文字只是占位提示，不会作为实际配置提交。

### 2. B站相关性过滤

| 模式 | 说明 |
| --- | --- |
| `phrase` | 标题或简介包含完整关键词，过滤最严格 |
| `smart` | 完整关键词命中，或配置的同义词组全部命中 |
| `terms` | 关键词按空格拆分后，每个分词都必须命中 |
| `off` | 不进行相关性过滤，保留B站原始搜索结果 |

更换关键词时，如果使用 `smart` 模式，请同步调整 `bilibili_search_crawler.py` 中的 `SYNONYM_GROUPS`。

B站的“抓取页数”指搜索接口请求页数，不是浏览器中肉眼看到的页面卡片数量。每次请求实际返回多少视频由B站接口决定，浏览器缩放不会影响爬取结果。

### 3. 抖音搜索方式

抖音流程为：

```text
输入关键词 → 综合搜索 → 筛选 → 最多点赞 → 滚动采集 → 获取作者资料 → 排序导出
```

不会自动切换到“视频”标签，因为“综合 + 最多点赞”和“视频 + 最多点赞”可能返回不同的候选集。

如果出现扫码或安全验证，请在弹出的浏览器中手动完成，脚本检测到完成后会自动继续。需要人工操作时请不要使用无头模式。

### 4. 粉丝区间筛选

“启用粉丝区间筛选”默认关闭。启用后可设置最少粉丝数和最多粉丝数，区间包含边界值。

例如设置为 10000–100000 后：

- 原点赞榜、播放榜和全部视频仍统计所有作者；
- 额外生成该粉丝区间内的点赞榜和播放榜；
- 区间外以及粉丝数未知的视频只从区间榜单中排除；
- 区间榜单在 Streamlit 标签页和 Excel 工作表中排在原榜单之前；
- 最小粉丝数大于最大粉丝数时不能启动任务。

粉丝区间配置目前仅在 Streamlit 界面中提供。

## 导出规则

### 默认保存

任务完成后只在指定目录保存一个 Excel 文件：

```text
关键词_平台搜索_年月日_时分秒.xlsx
```

不会自动保存 CSV 文件。

### Excel 工作表

未启用粉丝区间时包含 3 个工作表：

1. 点赞榜 TOP N；
2. 播放榜 TOP N；
3. 全部相关视频（B站）或全部搜索结果（抖音）。

启用粉丝区间后包含 5 个工作表：

1. 粉丝区间点赞榜 TOP N；
2. 粉丝区间播放榜 TOP N；
3. 原点赞榜 TOP N；
4. 原播放榜 TOP N；
5. 全部相关视频或全部搜索结果。

### CSV 按需下载

在 Streamlit 结果页面可以分别下载：

- 粉丝区间点赞榜 CSV；
- 粉丝区间播放榜 CSV；
- 原点赞榜 CSV；
- 原播放榜 CSV。

CSV 在内存中生成，只有点击对应下载按钮时才导出，不会自动写入爬取输出目录。

### 导出字段

两平台均包含以下主要字段：

| 字段 | 说明 |
| --- | --- |
| 排名 | 当前榜单中的排序 |
| 视频标题 | 视频标题或简介 |
| 笔记链接(视频) | 视频访问链接 |
| 点赞 / 收藏 / 评论 / 播放量 | 视频互动数据；平台未公开时留空 |
| 发布日期 | 当前视频发布时间 |
| 博主名字 | 作者昵称 |
| 博主主页链接 | 作者主页 |
| 博主粉丝量 | 获取作者资料时的粉丝数 |
| 博主最新视频发布时间 | 作者最新作品的发布时间，已考虑置顶作品 |
| BV号 / 抖音视频ID | 平台视频唯一标识 |
| 上榜情况 | 视频所属的榜单 |
| 备注 | 字段缺失或请求异常等说明 |

## 命令行使用

Streamlit 是推荐的使用方式，两个核心爬虫也可以独立运行。命令行模式的默认关键词、输出目录和 TOP N 可在脚本顶部配置区修改。

### B站

```bash
python bilibili_search_crawler.py -k "拍摄设备" -n 10 --max-pages 3 --filter-mode phrase --no-login -o "D:\VideoResults"
```

常用参数：

| 参数 | 说明 |
| --- | --- |
| `-k, --keyword` | 搜索关键词 |
| `-o, --output` | Excel 输出目录 |
| `-n, --top-n` | 每个榜单保留的条数 |
| `--filter-mode` | `phrase`、`smart`、`terms` 或 `off` |
| `--max-pages` | 每种排序最多抓取的页数 |
| `--login` | 强制重新扫码登录，适合更换账号 |
| `--no-login` | 跳过登录环节，使用匿名抓取 |

### 抖音

```bash
python douyin_search_crawler.py -k "拍摄设备" -n 10 --max-scrolls 2 --no-login -o "D:\VideoResults"
```

常用参数：

| 参数 | 说明 |
| --- | --- |
| `-k, --keyword` | 搜索关键词 |
| `-o, --output` | Excel 输出目录 |
| `-n, --top-n` | 每个榜单保留的条数 |
| `--max-scrolls` | 最多点赞结果页的滚动轮数 |
| `--login` | 清除旧登录状态并强制重新扫码 |
| `--no-login` | 直接复用当前浏览器会话 |
| `--headless` | 使用无界面浏览器；需要扫码或验证时不能使用 |
| `--download` | 下载双榜合并去重后的视频 |
| `--video-output` | 自定义视频保存目录 |

查看完整参数：

```bash
python bilibili_search_crawler.py --help
python douyin_search_crawler.py --help
```

## 浏览器说明

抖音使用 Playwright 启动持久化浏览器上下文，默认顺序为：

```text
Google Chrome → Microsoft Edge
```

只有 Chrome 或 Edge 其中一种即可使用。两者都不可用时，可以将 `BROWSER_CHANNEL` 改为 `None`，然后安装 Playwright Chromium。

浏览器 Cookie 和登录状态会复用，但账号有效性仍以平台实际状态为准。

## 登录凭证与隐私

登录凭证默认保存在当前操作系统用户的私有目录，不会写入搜索结果 Excel，也不应上传到 GitHub 或分享给他人。

Windows 下的主要位置为：

```text
%LOCALAPPDATA%\BilibiliCrawler\bilibili_cookies.json
%LOCALAPPDATA%\DouyinCrawler\browser_profile\
%LOCALAPPDATA%\DouyinCrawler\douyin_cookies.txt
```

Cookie 属于敏感信息，泄露后可能导致账号被他人使用。更换公共电脑或不再需要登录状态时，请妥善清理。

## 常见问题

### 抖音提示需要安全验证

使用可见浏览器运行，并在弹出的页面中手动完成验证。无头模式无法处理扫码或人工验证。

### 只有 Microsoft Edge，没有 Chrome

可以直接运行。Chrome 启动失败后会自动尝试 Edge，无需修改代码。

### B站请求被风控或数据不完整

可以尝试扫码登录、降低抓取频率、稍后重试，或增大脚本中的请求延时。

### 搜索结果数量不足 TOP N

可以增大B站抓取页数或抖音滚动轮数，也可以检查关键词及相关性过滤模式。抖音接口能够返回的结果数量还受平台限制。

### 抖音播放量为什么为空

部分搜索接口不公开播放量。此时对应单元格留空，不影响点赞、评论、作者粉丝量等其他字段。

### 为什么获取作者资料耗时较长

程序会按作者去重处理，但B站仍需请求作者接口，抖音还需要逐个打开不同作者的主页。候选作者越多，耗时越长，这是为了避免重复请求并保证作者字段完整。

## 使用声明

本项目仅用于学习、研究和合法的数据整理。使用时请遵守目标平台的服务条款、访问频率限制及相关法律法规，不要进行高频请求、绕过安全控制或侵犯他人隐私和知识产权。
