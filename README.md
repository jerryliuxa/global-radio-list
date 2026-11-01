# 📻 环球收听时钟 Global Radio Guide

一个纯静态的广播节目表网站：**不同国家/城市的听友打开网页，即可按自己当地时间查看此刻正在播出的短波（SW，含 DRM 数字广播）与 FM 节目和频率**，并可一键在 KiwiSDR / OpenWebRX / WebSDR 上收听。

**中英双语**：右上角 `EN / 中文` 按钮切换界面语言（首次访问按浏览器语言自动选择，选择保存在 localStorage）；地图页随链接语言参数显示。

数据来源（2 个）：

| 数据源 | 内容 | 条目数 |
|---|---|---|
| `aoki` | **HFCC B26 频率表（青木格式，用户提供的 xlsx）**——短波广播 2485~26040 kHz 共 3115 条；STATION/Location 为全名，L/L 列含发射台坐标；DRM 由文件内"HFCC原始数据"表 MOD="N" 判定（经 REQ# 连接）；已按 FDATE/TDATE 做播出季有效性过滤（当前季 B26：2026-10-25 ~ 2027-03-27） | ~3060 |
| `fm` | **radiomap.eu 54 个城市（欧洲为主）的 FM 电台频率**，全天播出，按所选城市显示 | ~2860 |

合并后约 5900 条。本表为短波广播表，不含长波 / 中波 / 授时台条目（筛选器保留这些频段，选择时给出说明）。

## 本地运行

任意一种方式均可：

```bash
# 方式一：本地 HTTP 服务（推荐，serve.py 带 no-cache 头，避免浏览器缓存旧数据）
cd radio-guide
python serve.py 8420
# 浏览器打开 http://127.0.0.1:8420

# 方式二：直接双击 index.html（数据以 <script> 形式加载，file:// 协议同样可用）
```

## 功能

- **收听地点**：内置 220+ 世界城市（IANA 时区，自动处理夏令时），默认跟随本机时区；另提供**任意地点输入框**——可输入城市名（中英文模糊匹配）、任意 IANA 时区（如 `Pacific/Chatham`，候选列表覆盖浏览器支持的全部时区）或 UTC 偏移（如 `UTC+3`）。选择保存在 localStorage。
- **正在播出**：按当前时刻实时判定，按频段分组、按频率聚合（同频多路发射机并列展示），带播出进度条、本地时段与 UTC 时段对照、跨日标注、全天（24h）识别。
- **即将播出**：未来 3 小时内开播的节目，按开播时间排序。
- **全部检索**：10300+ 条目全量搜索（电台名/节目名/频率/语言/台站/备注），可跨城市搜到其他城市的 FM 电台。
- **筛选**：频段（短波 / DRM 数字 / FM 等）、数据源（青木 B26 / FM 城市）、语言（中英双语分组）。
- **界面语言**：中文 / English 一键切换，日期与时间格式随语言本地化。
- **FM 城市归属**：FM 条目按所选城市显示——选中城市有 FM 数据时精确匹配该城市（如柏林 42 台），否则显示同时区全部 FM；无数据城市会给出提示。
- **发射台/接收机位置**：顶栏显示接收机位置坐标（城市表内置经纬度）；有坐标的发射台在卡片上显示 `📡 经纬度`（HFCC 来自 site.txt 参考表，精选表来自附件 L/L 列，WRMI 为固定坐标），并自动计算 `↔ 大圆距离 · 在你 xx° 方向`（从接收位置指向发射台，可作天线对准参考，可与卡片上的天线"方位"对照判断信号是否朝你打来）。
- **🔍 地图**：卡片上点击"🔍 地图"打开 `map.html`（Leaflet + OpenStreetMap 瓦片），画出接收机↔发射台的大圆连线，标注中点距离与两端名称，自动缩放到两点。瓦片源按网络情况自动探测降级（OSM 官方 → OSM France → OSM Germany，均为 OSM 数据）。
- **🎧 SDR 收听联动**：每个频率卡片带"在接收机上打开"按钮，自动调谐到该频率；同频多台接收机时显示"备用"菜单直接跳转。支持三类软件接收机，链接格式：
  - KiwiSDR：`http://主机:8073/?f=9570am-z8`
  - OpenWebRX：`https://主机/#freq=9570000,mod=am`（自动把 kHz 换算为 Hz）
  - WebSDR（Twente 系）：`http://主机/?tune=9570am`

  点顶栏「🎧 SDR 收听」可添加多台接收机（类型自动识别，也可手动指定）、设置默认、删除，配置保存在 localStorage。DRM 数字信号建议在接收机中切 USB 并配合 Dream 等解码软件收听。
- 页面每 20 秒自动刷新"正在播出"状态。

## 数据说明与来源链接

- HFCC（国际短波协调联合会）公开季表：<https://new.hfcc.org/data/>（B26 标准格式包 `b26allx2.zip`，含 site/broadcas/admin 参考表）
- radiomap.eu 城市频率表：<https://radiomap.eu/>（每城一页，`{国家}/{城市}.htm`）
- FMScan.org / FMLIST.org 是更全的全球 FM/MW 数据库，但存在 Cloudflare 防抓取（403），如需接入可在浏览器手动导出数据后扩展解析器
- 在播判定完全基于 UTC 时段 + 起始日（1=周日 … 7=周六），跨日时段自动归属正确的星期；显示层用 `Intl.DateTimeFormat` 按所选时区格式化，自动处理夏令时

## 更新数据

```bash
python -m pip install openpyxl

# 重建 data.js（--schedule 指向青木格式频率表 xlsx；--radiomap 可选，指向 radiomap 城市页目录）
python tools/build_data.py --schedule "C:/path/to/HFCC_B26_青木格式频率表.xlsx" --out data --radiomap /path/to/radiomap_pages
```

青木 xlsx 要求：主表含 FRE / STATION / UTC / Su-W-Sa / Language / Pow / Azi / Location / ADM / L/L / Remarks / CIRAF Zones / REQ# / FDATE / TDATE 列，另含"HFCC原始数据"表（REQ# + MOD 列，MOD="N" 判定 DRM）。新季度（A27 等）拿到同格式文件后替换 `--schedule` 路径即可。

## 文件结构

```
radio-guide/
├── index.html          # 页面骨架
├── map.html            # 🔍 两点连线地图（Leaflet + OSM，显示大圆连线与直线距离）
├── css/style.css       # 深色收音机风格样式
├── js/app.js           # 全部逻辑（i18n/时区换算/在播判定/FM 城市归属/SDR 联动/距离方位）
├── js/cities.js        # 220+ 城市与 IANA 时区表（含经纬度）
├── data/data.js        # 由 build_data.py 生成的合并数据（勿手工编辑）
├── serve.py            # 本地开发服务器（带 no-cache 头）
└── tools/build_data.py # 青木 xlsx / radiomap → data.js 解析脚本（含国家/语言映射表）
```
