# # LLM 政策文本指标构建
#
# 目标：将广东数字经济政策文本转为可复核的政策级指标。主线：读取数据 -> 清洗样本 -> 网页补全 -> LLM 编码 -> 指标输出。

# ## 项目处理逻辑
#
# 核心原则：先准备文本，再分阶段调用模型，最后把模型输出标准化为论文变量。

# ## API 与成本控制
#
# 默认不自动全量调用 API。正式运行前设置 `DEEPSEEK_API_KEY`，并显式打开 `RUN_FULL_BATCH=True`。

# ## JSON 输出约束
#
# LLM 输出必须为 JSON；代码内置 JSON 清洗、容错解析和失败记录。

# ### 环境准备与依赖导入
#
# 检查依赖并导入数据处理、网页抓取和 API 调用所需库。

# 环境准备与依赖导入
# 主要功能：检查当前 Python 环境是否缺少必要依赖；若缺少则自动安装，然后导入后续流程所需库。
import importlib.util
import subprocess
import sys

REQUIRED_PACKAGES = {
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "openai": "openai",
    "tqdm": "tqdm",
    "requests": "requests",
    "bs4": "beautifulsoup4",
}

missing_packages = [pkg for module, pkg in REQUIRED_PACKAGES.items() if importlib.util.find_spec(module) is None]
if missing_packages:
    print("正在安装缺失依赖：", missing_packages)
    subprocess.check_call([sys.executable, "-m", "pip", "install", *missing_packages])

import ast
import hashlib
import json
import os
import random
import re
import shutil
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openai import OpenAI
from tqdm import tqdm

try:
    from IPython.display import display
except Exception:
    display = print

# ### 路径、输出文件和研究口径配置
#
# 集中设置输入、输出、城市口径、工具分类和运行开关。

# 路径、输出文件和研究口径配置
# 主要功能：统一管理输入表、缓存文件、中间结果、最终结果路径，以及广东 21 市和工具分类等全局口径。
PROJECT_ROOT = Path.cwd()  # 项目根目录；脚本应从仓库根目录运行
DATA_FILE = PROJECT_ROOT / "gd21_digital_policies_2017_2022.xlsx"  # 原始政策文本 Excel
SHEET_NAME = "gd21_digital_policies_2017_2022"  # Excel 中待读取的工作表
OUTPUT_DIR = PROJECT_ROOT / "outputs" / "chapter3_llm_policy_indicators"  # LLM 和指标输出目录
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

RESULT_JSONL = OUTPUT_DIR / "gd21_policy_llm_results.jsonl"  # 逐条追加的 LLM 主结果
INTERMEDIATE_CSV = OUTPUT_DIR / "gd21_policy_llm_results_intermediate.csv"  # 批处理中间快照 CSV
INTERMEDIATE_XLSX = OUTPUT_DIR / "gd21_policy_llm_results_intermediate.xlsx"  # 批处理中间快照 Excel
WEB_TEXT_CACHE_JSONL = OUTPUT_DIR / "gd21_policy_web_fulltext_cache.jsonl"  # URL 网页全文缓存
WEB_TEXT_REVIEW_CSV = OUTPUT_DIR / "gd21_policy_web_fulltext_review.csv"  # 网页补全复核表
FINAL_CSV = PROJECT_ROOT / "gd21_policy_indicators_2017_2022.csv"  # 全量完成后的主结果 CSV
FINAL_XLSX = PROJECT_ROOT / "gd21_policy_indicators_2017_2022.xlsx"  # 全量完成后的主结果 Excel
PARTIAL_CSV = OUTPUT_DIR / "gd21_policy_indicators_partial_preview.csv"  # 部分结果预览 CSV
PARTIAL_XLSX = OUTPUT_DIR / "gd21_policy_indicators_partial_preview.xlsx"  # 部分结果预览 Excel
FAILED_CSV = OUTPUT_DIR / "gd21_policy_llm_failed_records.csv"  # API 或解析失败记录
EXCLUDED_CSV = OUTPUT_DIR / "gd21_policy_llm_excluded_records.csv"  # 被筛除样本记录

VALID_GD_CITIES = [
    "广州", "深圳", "珠海", "汕头", "佛山", "韶关", "河源", "梅州", "惠州", "汕尾",
    "东莞", "中山", "江门", "阳江", "湛江", "茂名", "肇庆", "清远", "潮州", "揭阳", "云浮",
]

# 多个单元都会用到这些常量，提前定义可避免分单元运行时出现未定义引用。
TOOL_CATEGORIES = ["fiscal", "financial", "supply", "demand", "regulation"]  # 五类政策工具口径
ALLOWED_TONES = {"supportive", "regulatory", "neutral"}  # 允许的政策基调标签
WORKFLOW_VERSION = "multistage_digital_policy_v1"  # 多阶段 LLM 工作流版本号

# 逻辑控制项：
# 1) 同一城市-同一标题-同一年重复出现通常是搜索结果重复，默认去重，减少API成本和样本重复加权。
# 2) 如果研究口径严格限定为21个地级市，把 KEEP_ONLY_21_CITIES 改为 True。
DROP_DUPLICATE_POLICIES = True  # 是否按城市-标题-年份去重，避免重复政策进入编码
KEEP_ONLY_21_CITIES = False  # 是否只保留广东 21 个地级市记录；False 时仅标记口径

# ### JSONL 与中间结果读写工具
#
# 逐条保存结果，支持中断后继续处理。

# JSONL 与中间结果读写工具
# 主要功能：把逐条 LLM 结果安全写入 JSONL，并按 policy_id 覆盖更新，便于中断后继续处理和导出中间表。
def jsonable(value):
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    if not isinstance(value, (list, dict)):
        try:
            if pd.isna(value):
                return None
        except (TypeError, ValueError):
            pass
    if hasattr(value, "item"):
        try:
            return value.item()
        except Exception:
            pass
    return value


def append_jsonl(record, path=RESULT_JSONL):
    with path.open("a", encoding="utf-8") as file:
        file.write(json.dumps(record, ensure_ascii=False) + "\n")


def load_jsonl(path=RESULT_JSONL):
    if not path.exists():
        return pd.DataFrame()
    records = []
    with path.open("r", encoding="utf-8") as file:
        for line_no, line in enumerate(file, start=1):
            line = line.strip()
            if not line:
                continue
            try:
                records.append(json.loads(line))
            except json.JSONDecodeError as exc:
                print(f"跳过 JSONL 第 {line_no} 行，解析失败：{exc}")
    return pd.DataFrame(records)


def upsert_jsonl_records(new_records, path=RESULT_JSONL, key="policy_id"):
    """按指定 key 覆盖写入，避免重复记录。"""
    new_records = [record for record in new_records if record]
    if not new_records:
        return load_jsonl(path)

    existing_records = load_jsonl(path).to_dict("records")
    replaced_keys = {record.get(key) for record in new_records}
    kept_records = [record for record in existing_records if record.get(key) not in replaced_keys]
    all_records = kept_records + new_records

    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as file:
        for record in all_records:
            file.write(json.dumps(record, ensure_ascii=False) + "\n")
    return pd.DataFrame(all_records)


def save_intermediate_snapshot():
    current = load_jsonl(RESULT_JSONL)
    if current.empty:
        return current
    current.to_csv(INTERMEDIATE_CSV, index=False, encoding="utf-8-sig")
    current.to_excel(INTERMEDIATE_XLSX, index=False)
    return current

# ### 读取 Excel 并校验基础字段
#
# 读取政策表，并检查 city/title/date/url/content 字段。

# 读取 Excel 并校验基础字段
# 主要功能：读取政策文本工作表，统一列名，检查 city/title/date/url/content 等必要字段是否存在。
if not DATA_FILE.exists():
    raise FileNotFoundError(f"未找到数据文件：{DATA_FILE}")

try:
    df_raw = pd.read_excel(DATA_FILE, sheet_name=SHEET_NAME, engine="openpyxl")
except ValueError as exc:
    available_sheets = pd.ExcelFile(DATA_FILE, engine="openpyxl").sheet_names
    raise ValueError(f"未找到工作表 {SHEET_NAME!r}，当前文件包含：{available_sheets}") from exc

df = df_raw.copy()
df.columns = [str(col).strip().lower() for col in df.columns]

required_columns = ["city", "title", "date", "url", "content"]
missing_columns = [col for col in required_columns if col not in df.columns]
if missing_columns:
    raise ValueError(f"表格缺少必要字段：{missing_columns}；当前字段为：{list(df.columns)}")

if "keyword" not in df.columns:
    df["keyword"] = "数字经济"

if "year" not in df.columns:
    df["year"] = pd.NA

# ### 日期、文本清洗与政策 ID 工具
#
# 统一日期、文本格式，并生成可回溯 policy_id。

# 日期、文本清洗与政策 ID 工具
# 主要功能：兼容不同日期格式，清洗正文和标题，为每条政策生成可回溯的 policy_id，并整理写入结果所需元数据。
def parse_mixed_date(value):
    """兼容 Excel 日期序列值、字符串日期和空值。"""
    if pd.isna(value):
        return pd.NaT
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if 20000 <= float(value) <= 60000:
            return pd.to_datetime(value, unit="D", origin="1899-12-30", errors="coerce")
        if 1900 <= float(value) <= 2100:
            return pd.to_datetime(str(int(value)), format="%Y", errors="coerce")
    return pd.to_datetime(str(value), errors="coerce")


def clean_text(value):
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def replace_double_quotes_with_single(value):
    """送入 LLM 前统一替换双引号，降低 evidence quote 破坏 JSON 的概率。"""
    text = clean_text(value)
    quote_translation = str.maketrans({
        '"': "'",
        "“": "'",
        "”": "'",
        "„": "'",
        "‟": "'",
        "＂": "'",
        "「": "'",
        "」": "'",
        "『": "'",
        "』": "'",
    })
    return text.translate(quote_translation)


def normalize_title(value):
    return re.sub(r"\s+", "", clean_text(value))


def text_hash(value):
    return hashlib.sha256(clean_text(value).encode("utf-8")).hexdigest()[:16]


def make_policy_id(row):
    key = "|".join([
        str(row.get("url", "")),
        str(row.get("title", "")),
        str(row.get("date", "")),
        str(row.get("source_row", "")),
    ])
    digest = hashlib.md5(key.encode("utf-8")).hexdigest()[:12]
    return f"{int(row['source_row']):04d}_{digest}"


def row_metadata(row):
    return {
        "policy_id": row["policy_id"],
        "source_row": int(row["source_row"]),
        "keyword": jsonable(row.get("keyword", "数字经济")),
        "city": jsonable(row.get("city", "")),
        "title": jsonable(row.get("title", "")),
        "date": jsonable(row.get("date", "")),
        "year": jsonable(row.get("year", None)),
        "url": jsonable(row.get("url", "")),
        "content_chars": jsonable(row.get("content_chars", None)),
        "table_content_chars": jsonable(row.get("table_content_chars", None)),
        "web_text_chars": jsonable(row.get("web_text_chars", None)),
        "llm_text_source": jsonable(row.get("llm_text_source", None)),
        "fetch_status": jsonable(row.get("fetch_status", None)),
        "fetch_error": jsonable(row.get("fetch_error", None)),
        "analysis_text_hash": jsonable(row.get("analysis_text_hash", None)),
        "content_incomplete_flag": jsonable(row.get("content_incomplete_flag", None)),
        "is_21city_policy": jsonable(row.get("is_21city_policy", None)),
        "scope_review_flag": jsonable(row.get("scope_review_flag", None)),
    }

# ### 生成待 LLM 编码的样本表
#
# 生成模型输入文本、去重字段和样本状态标记。

# 生成待 LLM 编码的样本表
# 主要功能：拼接标题与正文、识别疑似正文不完整样本、标记城市口径和复核口径，并按设置去重或过滤。
# source_row 保留 Excel 原始行号，便于后续从结果回查原始记录。
df["source_row"] = df.index + 2  # Excel 第 1 行是表头
raw_row_count = len(df)
df["title"] = df["title"].map(clean_text)
df["title_norm"] = df["title"].map(normalize_title)
df["content"] = df["content"].map(clean_text)
df["city"] = df["city"].map(clean_text)
df["url"] = df["url"].map(clean_text)
df["date_parsed"] = df["date"].map(parse_mixed_date)
df["date"] = df["date_parsed"].dt.strftime("%Y-%m-%d").fillna(df["date"].astype(str))
df["year"] = pd.to_numeric(df["year"], errors="coerce").fillna(df["date_parsed"].dt.year).astype("Int64")
df["table_policy_text"] = (df["title"].fillna("") + "\n\n" + df["content"].fillna("")).str.strip()
df["table_content_chars"] = df["table_policy_text"].str.len()
df["web_text"] = ""
df["web_text_chars"] = 0
df["llm_text_source"] = "table"
df["fetch_status"] = "not_run"
df["fetch_error"] = ""
df["policy_text"] = df["table_policy_text"]
df["content_chars"] = df["policy_text"].str.len()
df["content_incomplete_flag"] = (
    (df["content_chars"] < 800)
    & df["title"].str.contains(r"规划|方案|措施|办法|意见|通知|实施细则", na=False)
    & df["policy_text"].str.contains(r"附件|\.doc|\.docx|\.pdf|现印发", na=False)
)
df["analysis_text_hash"] = df["policy_text"].map(text_hash)
df["is_21city_policy"] = df["city"].isin(VALID_GD_CITIES)
df["scope_review_flag"] = (
    ~df["is_21city_policy"]
    | df["title"].str.contains(r"国务院|国家|广东省", na=False)
)
df["policy_id"] = df.apply(make_policy_id, axis=1)

df = df[df["table_policy_text"].str.len() > 0].copy()
nonempty_row_count = len(df)

if DROP_DUPLICATE_POLICIES:
    before_dedup = len(df)
    df = (
        df.sort_values(["city", "year", "title_norm", "content_chars"], ascending=[True, True, True, False])
        .drop_duplicates(subset=["city", "title_norm", "year"], keep="first")
        .sort_values("source_row")
    )
    print(f"去重：删除 {before_dedup - len(df)} 条城市-标题-年份重复记录。")

if KEEP_ONLY_21_CITIES:
    before_city_filter = len(df)
    df = df[df["is_21city_policy"]].copy()
    print(f"城市口径过滤：删除 {before_city_filter - len(df)} 条非21市记录。")

df = df.reset_index(drop=True)

print(f"原始记录 {raw_row_count} 条；非空文本 {nonempty_row_count} 条；当前用于LLM编码 {len(df)} 条。")
print(f"年份范围：{df['year'].min()} - {df['year'].max()}；城市数：{df['city'].nunique()}")
print(f"非21市/需人工确认口径记录：{int((~df['is_21city_policy']).sum())} 条；标题含国务院/国家/广东省等需复核记录：{int(df['scope_review_flag'].sum())} 条。")
display(df[["policy_id", "keyword", "city", "is_21city_policy", "scope_review_flag", "content_incomplete_flag", "title", "date", "year", "url", "content_chars"]].head())

# ## 网页全文补全与文本替换
#
# 表格 `content` 可能只是片段；正式调用 API 前按 URL 回溯网页全文，并用长度增益规则决定是否替换。

# ### 网页全文抓取参数与正文定位规则
#
# 设置网页回溯、正文选择器和文本替换阈值。

# 网页全文抓取参数与正文定位规则
# 主要功能：设置是否在 API 前补抓网页全文、抓取超时、替换阈值、正文 CSS 选择器和噪声文本规则。
FETCH_WEB_TEXT_BEFORE_API = True  # 调用 API 前是否按 URL 回溯网页全文
FORCE_REFETCH_WEB_TEXT = False  # 是否忽略缓存并强制重新抓取网页
WEB_FETCH_SLEEP_SECONDS = 0.8  # 每次网页抓取后的间隔，降低对政府网站压力
WEB_FETCH_TIMEOUT = 15  # 单次网页请求超时时间，单位秒
WEB_TEXT_MIN_GAIN_CHARS = 300  # 网页文本至少多出多少字符才替代表格文本
WEB_TEXT_MIN_GAIN_RATIO = 1.10  # 网页文本/表格文本比例达到该阈值才替换
SKIP_INCOMPLETE_TEXT_BEFORE_API = True  # 正文仍疑似不完整时是否跳过 LLM 调用

WEB_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
    "Connection": "close",
}

CONTENT_SELECTORS = [
    "#UCAP-CONTENT", "#zoomcon", "#Zoom", "#article_content", "#content", "#mainContent",
    ".TRS_Editor", ".content", ".detail-content", ".article-content", ".article_content",
    ".zw-content", ".view-content", ".main-content", ".pages_content", ".Custom_UnionStyle",
    "article", "main",
]

NOISE_PATTERNS = [
    r"分享到：.*?打印", r"【字体：.*?】", r"网站地图", r"主办：", r"ICP备", r"公网安备",
    r"无障碍", r"适老版", r"Language\s+EN", r"当前位置：", r"浏览次数：",
]

PREPARED_TEXT_COLUMNS = [
    "web_text", "web_text_chars", "fetch_status", "fetch_error", "llm_text_source",
    "policy_text", "content_chars", "analysis_text_hash", "content_incomplete_flag",
]

_WEB_TEXT_CACHE_BY_URL = None

# ### 网页清洗、HTML 请求和正文抽取函数
#
# 使用 requests/curl 抓网页，并抽取正文。

# 网页清洗、HTML 请求和正文抽取函数
# 主要功能：优先用 requests 抓取政府网页；失败时用 curl 兜底，并从 HTML 中抽取最长的候选正文。
def clean_web_text(text):
    text = clean_text(text)
    for pattern in NOISE_PATTERNS:
        text = re.sub(pattern, " ", text)
    return clean_text(text)


def extract_main_text_from_html(html):
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "noscript", "iframe", "form", "svg"]):
        tag.decompose()
    for tag in soup.select("header, footer, nav, .nav, .footer, .header, .breadcrumb, .share, .toolbar"):
        tag.decompose()

    candidates = []
    for selector in CONTENT_SELECTORS:
        for element in soup.select(selector):
            text = clean_web_text(element.get_text(" ", strip=True))
            if len(text) >= 80:
                candidates.append(text)

    if candidates:
        return max(candidates, key=len)

    body = soup.body or soup
    return clean_web_text(body.get_text(" ", strip=True))


def fetch_html_with_requests(url):
    with requests.Session() as session:
        # 不读取系统代理，避免本机代理端口不可用导致政府网站请求失败。
        session.trust_env = False
        response = session.get(url, headers=WEB_HEADERS, timeout=WEB_FETCH_TIMEOUT)
        response.raise_for_status()
        response.encoding = response.apparent_encoding or response.encoding
        return response.text


def curl_env_without_proxy():
    env = os.environ.copy()
    for key in list(env):
        if key.lower() in {"http_proxy", "https_proxy", "all_proxy", "ftp_proxy"}:
            env.pop(key, None)
    env["NO_PROXY"] = "*"
    env["no_proxy"] = "*"
    return env


def fetch_html_with_curl(url, insecure=False):
    curl_path = shutil.which("curl")
    if not curl_path:
        raise RuntimeError("系统未找到 curl，无法使用 curl 兜底抓取")

    cmd = [
        curl_path,
        "-L",
        "--silent",
        "--show-error",
        "--fail",
        "--http1.1",
        "--compressed",
        "--connect-timeout",
        "10",
        "--max-time",
        str(WEB_FETCH_TIMEOUT),
        "-A",
        WEB_HEADERS["User-Agent"],
        "-H",
        f"Accept: {WEB_HEADERS['Accept']}",
        "-H",
        f"Accept-Language: {WEB_HEADERS['Accept-Language']}",
        url,
    ]
    if insecure:
        cmd.insert(1, "--insecure")

    result = subprocess.run(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        timeout=WEB_FETCH_TIMEOUT + 5,
        env=curl_env_without_proxy(),
    )
    if result.returncode != 0:
        error = result.stderr.decode("utf-8", errors="ignore").strip()
        raise RuntimeError(error or f"curl退出码：{result.returncode}")
    if not result.stdout:
        raise RuntimeError("curl返回空内容")
    return result.stdout.decode("utf-8", errors="ignore")


def fetch_web_fulltext(url, max_retries=1):
    if not url or not str(url).startswith(("http://", "https://")):
        return {"fetch_status": "invalid_url", "fetch_error": "URL为空或不是HTTP链接", "web_text": ""}

    errors = []
    for attempt in range(max_retries):
        try:
            html = fetch_html_with_requests(url)
            web_text = extract_main_text_from_html(html)
            return {"fetch_status": "ok_requests", "fetch_error": "", "web_text": web_text}
        except Exception as exc:
            errors.append(f"requests: {exc}")
            if attempt < max_retries - 1:
                time.sleep((2 ** attempt) + random.uniform(0, 0.5))

    for insecure in [False, True]:
        try:
            html = fetch_html_with_curl(url, insecure=insecure)
            web_text = extract_main_text_from_html(html)
            status = "ok_curl_insecure" if insecure else "ok_curl"
            return {"fetch_status": status, "fetch_error": "", "web_text": web_text}
        except Exception as exc:
            errors.append(f"curl{' --insecure' if insecure else ''}: {exc}")

    return {"fetch_status": "failed", "fetch_error": " | ".join(errors)[-1000:], "web_text": ""}


def should_use_web_text(table_text, web_text):
    table_len = len(clean_text(table_text))
    web_len = len(clean_text(web_text))
    if web_len <= 0:
        return False
    if web_len >= table_len + WEB_TEXT_MIN_GAIN_CHARS:
        return True
    return table_len > 0 and web_len / table_len >= WEB_TEXT_MIN_GAIN_RATIO and web_len >= 800

# ### 网页缓存、复核表与 LLM 输入文本准备
#
# 读取缓存、更新复核表，并确定最终送入 LLM 的文本。

# 网页缓存、复核表与 LLM 输入文本准备
# 主要功能：复用网页全文缓存，决定是否用网页正文替代表格正文，并把每条政策最终送入 LLM 的文本状态写入复核表。
def load_web_text_cache():
    global _WEB_TEXT_CACHE_BY_URL
    if _WEB_TEXT_CACHE_BY_URL is None:
        cache_df = load_jsonl(WEB_TEXT_CACHE_JSONL)
        if cache_df.empty or "url" not in cache_df.columns:
            _WEB_TEXT_CACHE_BY_URL = {}
        else:
            _WEB_TEXT_CACHE_BY_URL = cache_df.drop_duplicates("url", keep="last").set_index("url").to_dict("index")
    return _WEB_TEXT_CACHE_BY_URL


def cache_web_text_record(record):
    global _WEB_TEXT_CACHE_BY_URL
    upsert_jsonl_records([record], WEB_TEXT_CACHE_JSONL, key="url")
    cache = load_web_text_cache()
    cache[record["url"]] = record
    _WEB_TEXT_CACHE_BY_URL = cache


def upsert_web_text_review(rows):
    if not rows:
        return pd.DataFrame()

    records = []
    for row in rows:
        if hasattr(row, "to_dict"):
            records.append(row.to_dict())
        else:
            records.append(dict(row))

    review_cols = [
        "policy_id", "city", "title", "url", "llm_text_source", "fetch_status", "fetch_error",
        "table_content_chars", "web_text_chars", "content_chars", "content_incomplete_flag", "analysis_text_hash",
    ]
    new_df = pd.DataFrame(records)
    new_df = new_df[[col for col in review_cols if col in new_df.columns]]

    if WEB_TEXT_REVIEW_CSV.exists():
        existing_df = pd.read_csv(WEB_TEXT_REVIEW_CSV)
        combined_df = pd.concat([existing_df, new_df], ignore_index=True)
    else:
        combined_df = new_df

    if "policy_id" in combined_df.columns:
        combined_df = combined_df.drop_duplicates("policy_id", keep="last")
    combined_df.to_csv(WEB_TEXT_REVIEW_CSV, index=False, encoding="utf-8-sig")
    return combined_df


def prepare_policy_row_for_llm(row, fetch_web_text=FETCH_WEB_TEXT_BEFORE_API):
    """在调用 LLM 前准备单条政策文本：按需抓网页、选择更完整文本、计算hash。"""
    prepared = row.copy()
    table_text = prepared.get("table_policy_text", prepared.get("policy_text", ""))
    url = prepared.get("url", "")
    web_text = clean_text(prepared.get("web_text", ""))
    fetch_status = prepared.get("fetch_status", "not_run")
    fetch_error = prepared.get("fetch_error", "")

    if fetch_web_text:
        cache = load_web_text_cache()
        if not FORCE_REFETCH_WEB_TEXT and url in cache:
            record = cache[url]
        else:
            fetched = fetch_web_fulltext(url)
            record = {
                "policy_id": prepared.get("policy_id", ""),
                "url": url,
                "title": prepared.get("title", ""),
                "fetch_status": fetched["fetch_status"],
                "fetch_error": fetched["fetch_error"],
                "web_text": fetched["web_text"],
                "web_text_chars": len(clean_text(fetched["web_text"])),
                "fetched_at": datetime.now().isoformat(timespec="seconds"),
            }
            cache_web_text_record(record)
            time.sleep(WEB_FETCH_SLEEP_SECONDS)

        web_text = clean_text(record.get("web_text", ""))
        fetch_status = record.get("fetch_status", "cache")
        fetch_error = record.get("fetch_error", "")

    use_web = fetch_web_text and should_use_web_text(table_text, web_text)
    table_text_for_llm = replace_double_quotes_with_single(table_text)
    web_text_for_llm = replace_double_quotes_with_single(web_text)
    final_text = web_text_for_llm if use_web else table_text_for_llm

    prepared["web_text"] = web_text_for_llm
    prepared["web_text_chars"] = len(web_text_for_llm)
    prepared["fetch_status"] = fetch_status
    prepared["fetch_error"] = fetch_error
    prepared["llm_text_source"] = "web" if use_web else "table"
    prepared["policy_text"] = clean_text(final_text)
    prepared["content_chars"] = len(prepared["policy_text"])
    prepared["analysis_text_hash"] = text_hash(prepared["policy_text"])
    prepared["content_incomplete_flag"] = bool(
        prepared["content_chars"] < 800
        and bool(re.search(r"规划|方案|措施|办法|意见|通知|实施细则", str(prepared.get("title", ""))))
        and bool(re.search(r"附件|\.doc|\.docx|\.pdf|现印发", prepared["policy_text"]))
    )

    upsert_web_text_review([prepared])
    return prepared


def sync_prepared_row(source_df, idx, prepared_row):
    """把按条准备后的文本字段同步回 df，供后续质量检查和最终输出使用。"""
    for col in PREPARED_TEXT_COLUMNS:
        if col in source_df.columns:
            source_df.at[idx, col] = prepared_row.get(col)
    return source_df


def make_skipped_incomplete_record(prepared_row):
    result = {
        "policy_tone": None,
        "target_scope": [],
        "tool_count": {category: 0 for category in TOOL_CATEGORIES},
        "target_specificity": None,
        "coordination_breadth": None,
        "_status": "skipped_incomplete_text",
        "_error": "正文疑似不完整：网页或表格文本只包含通知壳、附件链接或正文过短，未送入LLM。",
        "_raw_output": "",
        "_model": globals().get("MODEL", ""),
        "_base_url": globals().get("BASE_URL", ""),
        "_response_mode": "not_called",
        "_prompt_chars": 0,
        "_completed_at": datetime.now().isoformat(timespec="seconds"),
    }
    return {**row_metadata(prepared_row), **result}


print("网页全文抓取函数已就绪：不会在本单元全表抓取；全量API调用前才按条抓取。")
print(f"网页抓取缓存：{WEB_TEXT_CACHE_JSONL}")
print(f"网页抓取复核表：{WEB_TEXT_REVIEW_CSV}")

# ## 多阶段 Prompt 设计
#
# 流程：证据抽取 -> 标题反向复核 -> 基于证据编码 -> 可选复核修正。目标是降低一次性编码带来的幻觉和误判。

# ### Prompt 全局参数、JSON 示例和共用说明
#
# 设置筛选阈值、JSON 示例和政策定义。

# Prompt 全局参数、JSON 示例和共用说明
# 主要功能：定义多阶段工作流的阈值、输出 JSON 示例、文本截断规则和数字经济产业政策判定标准。
TOOL_CATEGORIES = ["fiscal", "financial", "supply", "demand", "regulation"]  # 五类政策工具口径
ALLOWED_TONES = {"supportive", "regulatory", "neutral"}  # 允许的政策基调标签
WORKFLOW_VERSION = "multistage_digital_policy_v1"  # 多阶段 LLM 工作流版本号

# 使用多轮判断和分数阈值降低误分，避免边界样本直接进入指标表。
DIGITAL_POLICY_SCORE_MIN = 60  # 数字经济产业政策正向评分最低阈值
TITLE_NOT_POLICY_SCORE_MAX = 60  # 标题反向复核中“明显非政策”最高允许分
SCORE_MARGIN_MIN = 15  # 正向分与反向分的最小差值，避免边界样本误入
ENABLE_TITLE_COUNTERCHECK = True  # 是否启用标题反向复核
ENABLE_REVIEW_ROUND = True  # 是否启用最终复核修正，质量更高但调用更多

EVIDENCE_JSON_EXAMPLE = {
    "is_digital_industrial_policy": 1,
    "digital_policy_score": 85,
    "non_policy_reason": "",
    "evidence": {
        "policy_subject": ["广州市工业和信息化局印发实施方案"],
        "target_scope": [{"keyword": "人工智能", "quote": "推动人工智能与实体经济深度融合"}],
        "policy_tone": [{"label_hint": "supportive", "quote": "对符合条件的企业给予奖励"}],
        "tools": {
            "fiscal": [{"tool": "奖励资金", "quote": "给予最高500万元奖励"}],
            "financial": [{"tool": "融资支持", "quote": "鼓励金融机构加大信贷支持"}],
            "supply": [{"tool": "产业园区建设", "quote": "建设数字经济产业园"}],
            "demand": [{"tool": "应用示范", "quote": "支持开放应用场景"}],
            "regulation": [{"tool": "数据安全", "quote": "加强数据安全管理"}]
        },
        "target_specificity": [{"quote": "到2025年培育10家龙头企业"}],
        "coordination_breadth": [{"quote": "市工业和信息化局、市发展改革委联合印发"}]
    },
    "confidence": 85
}

TITLE_CHECK_JSON_EXAMPLE = {
    "definitely_not_digital_policy_score": 10,
    "title_based_reason": "标题显示为正式数字经济政策文件，不是明显的解读、公示或名单。",
    "obvious_non_policy_type": ""
}

CODING_JSON_EXAMPLE = {
    "is_digital_industrial_policy": 1,
    "digital_policy_score": 85,
    "policy_tone": "supportive",
    "target_scope": ["人工智能", "工业互联网"],
    "tool_count": {"fiscal": 1, "financial": 1, "supply": 1, "demand": 1, "regulation": 0},
    "target_specificity": 1,
    "coordination_breadth": 1,
    "confidence": {
        "overall": 85,
        "policy_tone": 90,
        "target_scope": 85,
        "tool_count": 80,
        "target_specificity": 90,
        "coordination_breadth": 95
    },
    "coding_notes": "依据证据中明确出现的奖励、融资、园区和应用场景条款编码。",
    "evidence_used": {
        "policy_tone": ["对符合条件的企业给予奖励"],
        "tools": {"fiscal": ["给予最高500万元奖励"]}
    }
}


def truncate_policy_text(text, max_chars=12000):
    """控制单篇文本长度；过长时保留头尾，降低遗漏发文单位和目标条款的风险。"""
    text = clean_text(text)
    if len(text) <= max_chars:
        return text
    head_chars = max_chars // 2
    tail_chars = max_chars - head_chars
    return text[:head_chars] + "\n\n[中间内容因长度控制省略]\n\n" + text[-tail_chars:]


def compact_json(obj):
    return json.dumps(obj, ensure_ascii=False, indent=2)


def digital_policy_definition_text():
    return """
请判断给定文本是否属于“数字经济产业政策”。

你必须根据文本正文进行判断，不得仅凭标题、关键词或发布网站判断。

判断时应同时满足以下四项条件；任一条件不满足，均判为“否”。

【定义】

本文所称“数字经济产业政策”，是指由政府或准政府公共管理主体发布，包含明确政策措施，并直接面向数字经济相关产业、技术、业态或活动，旨在影响产业中长期发展、产业结构、企业行为或资源配置的正式政策文本。

【判定条件】

条件1：发布主体合格

文本发布主体应为以下之一：

- 各级政府；

- 政府组成部门、直属机构、派出机构；

- 开发区、高新区、经开区、自贸区等管委会；

- 上述主体的正式下属机构或授权单位。

若发布主体为媒体、协会、企业、研究机构、公众号转载平台等，且不能确认其代表政府正式发布，则判为“否”。

条件2：文本具有正式政策属性

文本应包含具体政策内容或政策措施，例如：

- 支持方向；

- 发展目标；

- 财政、税收、金融、土地、人才、创新、平台建设等支持措施；

- 产业培育、项目建设、企业扶持、试点示范、园区建设、基础设施建设、标准规范等安排；

- 明确的组织实施、保障机制或考核要求。

以下类型通常不属于正式政策文本，除非正文中完整包含正式政策内容：

- 新闻报道；

- 政策解读；

- “一图读懂”；

- 音频/视频解读；

- 会议通知；

- 工作动态；

- 信息转载；

- 意见征集稿；

- 政策问答；

- 简短通知壳；

- 名单公示；

- 项目入库、资金分配、奖励名单、无异议说明等执行结果披露。

条件3：政策对象直接指向数字经济

政策内容应直接涉及数字经济相关产业、技术、基础设施或应用活动，包括但不限于：

- 人工智能；

- 大数据；

- 云计算；

- 区块链；

- 工业互联网；

- 物联网；

- 数据要素、数据交易、数据治理；

- 数字产业化；

- 产业数字化；

- 智能制造；

- 软件和信息服务业；

- 电子商务；

- 数字贸易；

- 数字金融；

- 数字文旅；

- 智慧城市、智慧园区、智慧物流等数字化应用；

- 数据中心、算力中心、通信网络、新型数字基础设施等。

若文本仅笼统提到“信息化”“互联网+”“数字化转型”等词语，但没有明确产业、技术、企业、项目或政策措施指向，应谨慎判断，通常判为“否”。

条件4：政策目的具有产业政策属性

文本的政策目的应是影响产业中长期发展、产业结构、企业行为或资源配置，例如：

- 培育数字经济产业集群；

- 推动企业数字化转型；

- 促进数字技术研发和应用；

- 支持数字经济企业发展；

- 引导资本、人才、土地、数据、算力等资源配置；

- 建设数字基础设施；

- 推动产业链、创新链、供应链升级；

- 制定数字经济发展规划或行动方案。

若文本仅处理短期行政事务、会议安排、材料报送、活动组织、检查通知、统计填报、项目验收等事项，即使涉及数字经济主题，也应判为“否”。

【重要反例规则】

1. 只有解读材料而无正式政策正文

仅包含“一图读懂”“政策解读”“音频解读”“媒体报道”等内容，若未附正式政策全文，应判为“否”。

2. 名单、公示、资金拨付类文本

资金名单、项目名单、试点名单、入库名单、无异议公示、奖励结果公示等，通常只是政策执行结果披露，不应编码为新的政策文本。

3. 缺少附件正文的通知壳

若文本只有“现印发给你们，请认真贯彻执行”等表述，但未提供附件中的正式政策正文，不能直接判为实质性政策，应判为“否”。

4. 上级政策文件的处理

国务院、省级政府或省级部门文件如果确实包含数字经济产业政策内容，可以判为“是”；但在进行城市层面政策分析时，应标注其行政层级，并另行复核是否适合纳入城市口径。

5. 转发类文件的处理

地方政府转发上级政策时，若仅为简单转发且没有本地化措施，通常不作为该地新的实质性政策；若转发文件同时提出本地实施方案、任务分工、支持措施或落实细则，则可判为“是”。

【判定原则】

- 宁可保守，不要把新闻、解读、公示、通知壳误判为正式政策。

- 必须依据文本正文判断，不得仅凭标题中的“数字经济”“人工智能”“产业”等词语判断。

- 若文本信息不足以确认是否包含正式政策措施，应输出“需人工复核”。

- 若文本同时具有政策正文和解读材料，应以正式政策正文为准。

【输出格式】

请严格按照以下 JSON 格式输出，不要输出多余文字：

{

  "judgement": "是/否/需人工复核",

  "reason": {

    "发布主体": "",

    "文本属性": "",

    "数字经济相关性": "",

    "产业政策属性": "",

    "需要注意的问题": ""

  }

}

【待判断文本】

{text}
""".strip()

# ### 第一阶段 Prompt：全文证据抽取
#
# 抽取原文证据并初判是否属于研究样本。

# 第一阶段 Prompt：全文证据抽取
# 主要功能：要求 LLM 阅读标题和正文，先判断是否属于数字经济产业政策，并抽取后续编码需要的原文证据。
def build_evidence_prompt(row, max_chars=12000):
    title = replace_double_quotes_with_single(row.get("title", ""))
    city = replace_double_quotes_with_single(row.get("city", ""))
    date = replace_double_quotes_with_single(row.get("date", ""))
    policy_text = truncate_policy_text(replace_double_quotes_with_single(row.get("policy_text", row.get("content", ""))), max_chars=max_chars)

    return f"""
你是一位严谨的中国产业政策文本编码员。

你的任务不是生成最终回归指标，也不是判断政策强度，而是进行多阶段文本分析思路，先完成“证据抽取”。

请只抽取文本中能够直接支持后续编码的原文证据。

你必须严格遵守以下规则：

1. 只依据给定的政策标题和政策正文进行判断。

2. 不得使用外部知识。

3. 不得根据地区、部门名称、政策背景或常识推测文本未明确表达的内容。

4. 所有 quote 必须尽量来自原文中的连续片段。

5. 不要改写、概括或翻译 quote。

6. 如果某一类没有明确证据，返回空列表 []。

7. 不要因为标题中出现关键词就抽取证据，除非标题或正文中确有明确表达。

8. 不要急于做“是否属于数字经济产业政策”的最终判断；本阶段只抽取证据。

9. 输出必须是合法 JSON，不要输出 Markdown、解释文字或代码块。

{digital_policy_definition_text()}

【需要抽取的证据】

1. policy_subject

抽取能够证明发布主体是政府、政府部门、开发区管委会或其正式下属机构的原文片段。

证据可来自标题、发文机关、正文开头、落款等。

如果仅能看到转载平台、新闻媒体或不明来源，返回空列表。

2. target_scope

抽取数字经济相关目标产业、技术、基础设施或应用领域。

每项包含：

- keyword：数字经济相关关键词；

- quote：包含该关键词或其明确上下文的原文连续片段。

可包括但不限于：

人工智能、工业互联网、大数据、云计算、区块链、数据要素、数据交易、数字产业化、产业数字化、智能制造、数字贸易、电子商务、软件和信息服务、数据中心、算力、数字基础设施、智慧城市、智慧园区、智慧物流等。

3. policy_tone

抽取能够体现政策取向的原文片段。

每项包含：

- label_hint：只能填 "supportive"、"regulatory" 或 "neutral"；

- quote：对应原文片段。

判断提示：

- supportive：扶持、鼓励、促进、支持、培育、奖励、补助、推动发展等；

- regulatory：规范、监管、整治、限制、准入、审查、处罚、风险防控等；

- neutral：规划、通知、组织实施、任务安排、统计监测等较中性的表述。

4. tools

抽取五类政策工具的证据。

每项包含：

- tool：只能填 "fiscal"、"financial"、"supply"、"demand" 或 "regulation"；

- quote：对应原文连续片段。

工具分类标准如下：

- fiscal：

  财政补贴、税收优惠、政府采购、专项资金、专项基金、奖励资金、贷款贴息、费用减免等。

- financial：

  信贷支持、融资担保、上市培育、产业投资基金、融资租赁、股权投资、保险支持、金融机构服务等。

- supply：

  土地供应、基础设施建设、技术研发支持、算力中心、数据中心、产业园区、创新平台、公共服务平台、人才引进、人才培训等。

- demand：

  场景开放、市场推广、消费补贴、以旧换新、应用示范、首购首用、供需对接、政府或国企开放应用场景等。

- regulation：

  准入标准、数据安全、网络安全、合规审查、产能管控、环保要求、质量监管、监督检查、信用惩戒、风险防控等。

如果某类工具没有明确原文证据，不要补充，不要推断。

5. target_specificity

抽取明确量化发展目标的原文证据。

只抽取面向产业发展结果的量化目标，例如：

- 到某年数字经济规模达到多少；

- 培育多少家企业；

- 建成多少个平台、园区、项目；

- 产业增加值、营收、产值、投资额、企业数量、专利数量等达到明确数值。

不要把以下内容误判为量化发展目标：

- 文件发布日期；

- 政策有效期；

- 申报截止日期；

- 补贴金额；

- 奖励标准；

- 项目资助额度；

- 工作完成时间节点；

- 联系电话、文号、附件编号。

每项包含：

- metric：目标指标名称；

- quote：对应原文连续片段。

6. coordination_breadth

抽取多部门联合发文证据。

只看标题、发文机关、正文开头或落款中的联合发布、联合印发、多个发文机关并列等证据。

不要根据正文中“协调”“配合”“会同”等执行性表述推断为联合发文。

每项包含：

- agencies：联合发文主体，尽量使用原文表述；

- quote：对应原文连续片段。

【待分析文本】

政策标题：{title}

政策城市：{city}

政策日期：{date}

政策正文：

\"\"\"

{policy_text}

\"\"\"

【输出要求】

请严格按照以下 JSON 结构输出。

所有键必须存在。

如果没有证据，对应字段返回空列表 []。

不要输出 JSON 之外的任何文字。

{compact_json(EVIDENCE_JSON_EXAMPLE)}
""".strip()

# ### 第二阶段 Prompt：标题反向复核
#
# 识别解读、公示、名单、通知壳等非政策文本。

# 第二阶段 Prompt：标题反向复核
# 主要功能：只依据标题和基础信息识别明显非政策文本，降低解读、名单、公示、通知壳等样本误入指标表的概率。
def build_title_countercheck_prompt(row, evidence_result):
    title = replace_double_quotes_with_single(row.get("title", ""))
    city = replace_double_quotes_with_single(row.get("city", ""))
    date = replace_double_quotes_with_single(row.get("date", ""))
    url = clean_text(row.get("url", ""))
    first_score = evidence_result.get("digital_policy_score", 0)
    non_policy_reason = replace_double_quotes_with_single(evidence_result.get("non_policy_reason", ""))

    return f"""
你是一位严谨的中国政策文本复核员。

你的任务是进行第二轮“反例筛查”，只根据以下信息判断该记录是否“明显不是可编码的数字经济产业政策文本”：

- 政策标题；

- 城市字段；

- 日期；

- URL；

- 第一轮 digital_policy_score；

- 第一轮 non_policy_reason。

本轮不是完整政策编码，也不是最终判断数字经济产业政策是否成立。

你的唯一任务是识别明显应排除的记录。

请严格遵守以下规则：

1. 只依据本 prompt 中给定的信息判断。

2. 不得使用外部知识。

3. 不得根据 URL 域名、城市层级或行政级别做过度推断。

4. 不要因为文件是国家级、省级、市级、区县级就直接判为反例。

5. 不要因为城市字段为空、城市字段与标题不完全一致，就直接判为反例。

6. 不要因为标题较宽泛、较宏观，就直接判为反例。

7. 只有在标题、URL 或第一轮理由已经明确显示其属于明显反例时，才给出高分。

8. 如果信息不足，宁可给低分或中低分，不要武断排除。

【数字经济产业政策定义】

{digital_policy_definition_text()}

【明显反例类型】

如果标题、URL 或第一轮 non_policy_reason 明确显示以下类型，可判定为“明显不是可编码的数字经济产业政策文本”：

1. 政策解读类

例如：

- 政策解读；

- 文字解读；

- 图文解读；

- 专家解读；

- 媒体解读；

- 问答解读；

- 新闻发布会解读；

- 政策吹风会材料。

2. 可视化或音视频解读类

例如：

- 一图读懂；

- 图解；

- 长图；

- 动漫解读；

- 音频解读；

- 视频解读。

3. 新闻报道或工作动态类

例如：

- 工作动态；

- 部门动态；

- 政务新闻；

- 新闻报道；

- 会议召开；

- 领导调研；

- 活动报道；

- 媒体转载；

- 信息转载。

4. 名单、公示、结果披露类

例如：

- 名单公示；

- 入选名单；

- 拟认定名单；

- 拟奖励名单；

- 资金分配公示；

- 项目入库名单；

- 试点示范名单；

- 无异议说明；

- 公示结果；

- 验收结果；

- 申报结果。

5. 征求意见或草案类

例如：

- 征求意见稿；

- 公开征求意见；

- 意见反馈；

- 草案；

- 征求社会公众意见。

注意：如果标题显示“正式印发”“实施方案”“若干措施”等，不应仅因曾有征求意见环节而排除。

6. 纯通知壳或附件缺失类

例如：

- 标题或第一轮理由表明只有“现印发给你们，请认真贯彻执行”，但缺少附件正文；

- 第一轮理由明确说明“附件缺失”“无正文”“仅通知壳”“正文不可见”。

注意：如果标题本身是“关于印发……的通知”，但无法确认附件缺失，不要直接判为反例。

7. 纯事务性通知类

例如：

- 会议通知；

- 培训通知；

- 报送材料通知；

- 申报通知；

- 检查通知；

- 调研通知；

- 统计填报通知；

- 活动报名通知；

- 评审通知。

注意：如果通知正文可能包含完整政策措施，但当前信息无法确认，不要给过高分。

8. 非政策文本或非正式来源

例如：

- 招标公告；

- 采购公告；

- 中标公告；

- 招聘公告；

- 办事指南；

- 操作手册；

- 服务指南；

- 下载页面；

- 机构介绍；

- 转载页；

- 仅网页导航或索引页。

【不应直接排除的情形】

以下情形本身不是明显反例，不能仅凭这些信息给高分：

1. 国家级或省级政策文件；

2. 上级政策转发文件；

3. 标题中出现“通知”“意见”“方案”“规划”“措施”“办法”“细则”；

4. 标题没有直接出现“数字经济”，但出现人工智能、工业互联网、大数据、云计算、区块链、智能制造、电子商务、软件、算力、数据中心等相关对象；

5. 城市字段与政策层级不完全一致；

6. URL 看起来像政府网站、栏目页或附件页，但不能确认正文类型；

7. 第一轮 digital_policy_score 较低，但 non_policy_reason 不明确。

【评分标准】

请输出 definitely_not_digital_policy_score，表示“明显不是可编码数字经济产业政策文本”的程度。

- 0-20：完全不像明显反例；应保留进入后续正文编码。

- 21-40：有轻微信号，但证据不足；不建议直接排除。

- 41-60：存在一定反例迹象，需要人工复核。

- 61-80：较明显是反例，建议排除或重点复核。

- 81-100：确定是明显反例，可直接排除。

评分时请特别保守：

- 只有当标题、URL 或第一轮理由明确指向反例类型时，才给 80 分以上。

- 如果只是“不确定是否为政策”，不要给高分。

- 如果只是“数字经济相关性不强”，但不是新闻、解读、公示、通知壳等明显反例，不要给高分。

- 本轮识别的是“明显不是可编码政策文本”，不是“数字经济相关性不足”。

【待复核记录】

标题：{title}

城市字段：{city}

日期：{date}

URL：{url}

第一轮 digital_policy_score：{first_score}

第一轮 non_policy_reason：{non_policy_reason}

【输出要求】

请严格按照以下 JSON 结构输出。

不要输出 Markdown、解释文字或代码块。

所有键必须存在。

{compact_json(TITLE_CHECK_JSON_EXAMPLE)}
""".strip()

# ### 第三阶段 Prompt：基于证据结构化编码
#
# 仅基于证据生成政策工具和核心变量。

# 第三阶段 Prompt：基于证据结构化编码
# 主要功能：只使用第一阶段抽取出的证据，对政策基调、目标领域、五类政策工具、目标明确性和协调广度编码。
def build_coding_prompt(row, evidence_result):
    title = replace_double_quotes_with_single(row.get("title", ""))
    city = replace_double_quotes_with_single(row.get("city", ""))
    date = replace_double_quotes_with_single(row.get("date", ""))
    evidence_json = compact_json(evidence_result)

    return f"""
你是一位严谨的中国产业政策编码员。

现在请进入第二阶段：根据第一阶段已经抽取出的证据 JSON，生成最终结构化编码指标。

重要原则：

1. 你只能依据“第一阶段证据 JSON”进行编码。

2. 不得回到原始政策正文中重新寻找证据。

3. 不得依据常识、政策背景、城市信息、部门名称或标题自行补充证据。

4. 如果第一阶段证据中没有支持某一指标的信息，该指标应按缺失或否定处理。

5. 本阶段的任务是“证据到指标”的结构化映射，不是重新判断文本内容。

6. 输出必须是合法 JSON，不要输出 Markdown、解释文字或代码块。

【编码规则】

1. policy_tone

policy_tone 只能取以下三个值之一：

- "supportive"

- "regulatory"

- "neutral"

判断规则：

- supportive：

  主导目的为扶持、促进、鼓励、培育、奖励、补贴、优化营商环境、扩大应用、开放场景、推动产业发展、支持企业成长等。

- regulatory：

  主导目的为监管、规范、限制、准入、审查、执法、风险防控、数据安全、网络安全、质量监管、合规约束、信用惩戒等。

- neutral：

  主要是信息传达、名单公示、一般通知、组织安排，或第一阶段证据中缺少明确政策工具和明确政策取向。

若 supportive 和 regulatory 证据同时存在：

- 按政策主导目的判断；

- 如果扶持与监管强度接近，或无法从证据判断主导目的，则取 "neutral"；

- 不要简单按照证据条数多少机械判断，应看证据反映的政策核心目的。

2. target_scope

target_scope 只列出第一阶段 target_scope 中有明确证据支持的数字经济目标领域关键词。

规则：

- 最多保留 8 个；

- 应优先保留更具体的领域，而不是过于宽泛的词；

- 例如有“人工智能”“大数据”“云计算”时，不必再额外列“数字经济”，除非“数字经济”本身是政策核心对象；

- 不要加入证据中没有出现或没有明确支持的领域；

- 不要把一般性的“创新”“产业发展”“高质量发展”编码为数字经济目标领域。

3. tool_count

tool_count 用于统计五类政策工具中“不同工具/措施类型”的数量。

工具类别只能包括：

- fiscal

- financial

- supply

- demand

- regulation

计数规则：

- 按“不同工具/措施类型”计数，不按词语出现次数计数；

- 同一类工具中，相同或高度相似措施重复出现，只计 1 次；

- 同一句中出现多个不同措施，可以分别计数；

- 如果只有泛泛表述，如“加大支持力度”“强化保障”“优化服务”，但没有明确工具类型，不能计入；

- 如果第一阶段 tools 中某类为空，则该类计数为 0；

- 不得根据常识推断工具存在。

工具类别说明：

- fiscal：

  财政补贴、税收优惠、政府采购、专项资金、专项基金、奖励资金、贷款贴息、费用减免等。

- financial：

  信贷支持、融资担保、上市培育、产业投资基金、融资租赁、股权投资、保险支持、金融机构服务等。

- supply：

  土地供应、基础设施建设、技术研发支持、算力中心、数据中心、产业园区、创新平台、公共服务平台、人才引进、人才培训等。

- demand：

  场景开放、市场推广、消费补贴、以旧换新、应用示范、首购首用、供需对接、政府或国企开放应用场景等。

- regulation：

  准入标准、数据安全、网络安全、合规审查、产能管控、环保要求、质量监管、监督检查、信用惩戒、风险防控等。

4. target_specificity

target_specificity 只能取 0 或 1。

取 1 的条件：

- 第一阶段 target_specificity 中存在明确量化发展目标；

- 该目标必须指向产业发展结果、企业培育、平台建设、项目建设、产业规模、营收、产值、投资额、增加值、专利、人才、园区、应用场景等发展性结果。

取 0 的情形：

- 没有量化发展目标；

- 只有文件日期、政策有效期、申报截止日期；

- 只有补贴金额、奖励标准、资助额度；

- 只有文号、附件编号、联系电话；

- 只有工作完成时间节点，但没有发展结果指标；

- 只有模糊目标，如“明显提升”“大幅增长”“持续优化”，但无明确数值。

5. coordination_breadth

coordination_breadth 只能取 0 或 1。

取 1 的条件：

- 第一阶段 coordination_breadth 中存在多部门联合发文证据；

- 证据必须来自标题、发文机关、正文开头或落款中的联合发布、联合印发、多个发文机关并列等信息。

取 0 的情形：

- 只有正文中要求部门协同、加强配合、建立机制；

- 只有领导小组、联席会议、任务分工；

- 只有一个发文机关；

- 第一阶段没有明确联合发文证据。

6. evidence_used

evidence_used 中应保留本阶段用于判断的关键证据，便于人工抽查。

要求：

- 只引用第一阶段证据 JSON 中已有的 quote；

- 不要新增原文片段；

- 每个核心指标尽量保留 1-3 条最关键证据；

- 如果某指标没有证据，填空列表 []；

- quote 应保持原样，不要改写。

【待编码记录】

政策标题：{title}

政策城市：{city}

政策日期：{date}

第一阶段证据 JSON：

{evidence_json}

【输出要求】

请严格按照以下 JSON 示例输出。

所有键必须存在。

不要添加 JSON 之外的任何解释、Markdown 或代码块。

{compact_json(CODING_JSON_EXAMPLE)}
""".strip()

# ### 第四阶段 Prompt：复核修正
#
# 检查工具、量化目标和联合发文是否误判。

# 第四阶段 Prompt：复核修正
# 主要功能：检查初始编码和证据是否一致，修正明显的工具误判、量化目标误判或联合发文误判。
def build_review_prompt(row, evidence_result, coding_result):
    title = replace_double_quotes_with_single(row.get("title", ""))
    evidence_json = compact_json(evidence_result)
    coding_json = compact_json(coding_result)

    return f"""
你是一位负责复核 LLM 政策编码结果的审稿人。

现在请进入第三阶段：根据“证据 JSON”复核“初始编码 JSON”是否存在明显错误，并在必要时进行修正。

你的任务不是重新阅读原文，也不是重新做第一阶段证据抽取。

你只能依据以下两类输入判断：

1. 证据 JSON；

2. 初始编码 JSON。

如果初始编码与证据一致，请原样输出初始编码。

如果初始编码存在明确错误，请在不改变 JSON 结构的前提下修正错误字段。

请严格遵守以下原则：

1. 只依据证据 JSON 进行复核。

2. 不得使用外部知识。

3. 不得根据政策标题、城市、部门名称或常识补充证据。

4. 不得新增证据 JSON 中不存在的 quote。

5. 没有明确错误时，不要为了“更完美”而改写初始编码。

6. 输出格式必须与初始编码 JSON 完全一致。

7. 输出必须是合法 JSON，不要输出 Markdown、解释文字或代码块。

【重点复核问题】

请重点检查以下五类常见错误：

1. 政策工具误判或计数过高

检查 tool_count 是否只统计了证据 JSON 中明确支持的工具。

需要修正的情形包括：

- 证据 JSON 中没有某类工具，但初始编码将该类计为 1 或更高；

- 把“加强支持”“完善服务”“强化保障”等泛泛表述误判为具体工具；

- 同一工具或高度相似措施重复出现，却被重复计数；

- 把同一个措施拆成多个工具，导致计数过高；

- 把不属于该类别的工具放入错误类别。

计数原则：

- 统计“不同工具/措施类型”的数量；

- 不按词语出现次数累加；

- 同类、同义或高度相似措施只计 1 次；

- 没有明确证据则计 0。

2. 量化发展目标误判

检查 target_specificity 是否被误设为 1。

只有以下情况才能取 1：

- 证据 JSON 中存在明确的量化发展目标；

- 该目标指向产业发展结果、企业培育、产业规模、产值、营收、投资额、增加值、平台数量、项目数量、园区建设、人才规模、应用场景数量等发展性结果。

需要修正为 0 的情形包括：

- 把文件日期误判为量化目标；

- 把政策有效期误判为量化目标；

- 把申报截止日期误判为量化目标；

- 把补贴金额、奖励金额、资助额度、贷款贴息比例误判为量化发展目标；

- 把文号、附件编号、联系电话误判为量化目标；

- 只有“明显提升”“不断增强”“持续优化”等无具体数值的目标。

3. 多部门联合发文误判

检查 coordination_breadth 是否被误设为 1。

只有以下情况才能取 1：

- 证据 JSON 中存在多部门联合发文证据；

- 证据来自标题、发文机关、正文开头或落款；

- 能看出多个发文机关联合发布、联合印发或并列署名。

需要修正为 0 的情形包括：

- 只是正文中写“加强部门协同”“各部门配合”“建立联动机制”；

- 只是任务分工中列出多个责任单位；

- 只是成立领导小组或联席会议；

- 只有一个发文机关；

- 证据 JSON 中没有明确联合发文证据。

4. 非正式政策文本误判

检查初始编码是否把明显非正式政策文本当成正式可编码政策。

如果证据 JSON 或初始编码中的 evidence_used 明确显示文本属于以下类型，应将编码修正为与示例结构一致的“非政策/中性/低工具”状态：

- 政策解读；

- 一图读懂；

- 音频解读；

- 视频解读；

- 新闻报道；

- 工作动态；

- 名单公示；

- 资金公示；

- 项目名单；

- 无异议说明；

- 征求意见稿；

- 意见征集；

- 纯通知壳；

- 附件缺失；

- 转载信息；

- 会议通知；

- 培训通知；

- 申报通知；

- 统计填报通知。

注意：

- 不能仅因为标题包含“通知”就判为非政策；

- 不能仅因为文件是国家级、省级或上级文件就判为非政策；

- 只有证据明确显示其为上述反例类型时才修正。

5. policy_tone 与证据主导目的不一致

检查 policy_tone 是否与证据 JSON 中的主导目的一致。

取值只能为：

- "supportive"

- "regulatory"

- "neutral"

修正规则：

- 如果证据主要体现扶持、促进、鼓励、培育、奖励、补贴、优化环境、扩大应用、开放场景、推动产业发展，应为 "supportive"；

- 如果证据主要体现监管、规范、限制、准入、审查、执法、风险防控、数据安全、质量监管、合规约束，应为 "regulatory"；

- 如果证据主要是信息传达、名单公示、一般通知、组织安排，或缺少明确政策工具和取向，应为 "neutral"；

- 如果 supportive 和 regulatory 证据都存在，按主导目的判断；

- 如果二者强度接近或无法判断主导目的，应为 "neutral"。

【修正要求】

如果需要修正：

1. 只修正存在明确错误的字段；

2. 保持初始编码 JSON 的所有键不变；

3. 不要新增初始编码 JSON 中没有的键；

4. evidence_used 只能保留或删除证据 JSON 中已有 quote，不得新增或改写 quote；

5. 如果某项指标修正为 0、空列表或 neutral，应同步检查 evidence_used 是否仍然合理；

6. 如果初始编码正确，请完全原样输出。

【待复核记录】

政策标题：{title}

证据 JSON：

{evidence_json}

初始编码 JSON：

{coding_json}

【输出要求】

请严格按照以下 JSON 示例输出。

输出格式必须与初始编码 JSON 一致。

不要添加 JSON 之外的任何解释、Markdown 或代码块。

{compact_json(CODING_JSON_EXAMPLE)}
""".strip()

# ## API 客户端配置
#
# 从环境变量读取 `DEEPSEEK_API_KEY`、`DEEPSEEK_BASE_URL`、`DEEPSEEK_MODEL`。真实 Key 不写入公开仓库。

# ### API 客户端配置
#
# 读取环境变量并创建 OpenAI 兼容客户端。

# API 客户端配置
# 主要功能：从环境变量读取 API 配置，并创建 OpenAI 兼容 client；未检测到密钥时不触发 API 调用。

API_KEY = os.getenv("DEEPSEEK_API_KEY", "").strip()  # API Key；不得写入公开仓库
BASE_URL = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com").strip()  # OpenAI 兼容接口地址
MODEL = os.getenv("DEEPSEEK_MODEL", "deepseek-chat").strip()  # 使用的模型名称

if not API_KEY:
    client = None
    print("未检测到 DEEPSEEK_API_KEY；后续不会调用 API。请设置环境变量后再运行全量批处理。")
else:
    client = OpenAI(api_key=API_KEY, base_url=BASE_URL)
    print(f"已配置 API client：model={MODEL}, base_url={BASE_URL}")

# ### JSON 输出清洗与容错解析
#
# 清理代码块、提取 JSON，并处理轻微格式错误。

# JSON 输出清洗与容错解析
# 主要功能：清理模型可能返回的 Markdown 代码块，提取第一个 JSON 对象，并尝试修复少量缺失的闭合括号。
def strip_code_fence(text):
    text = (text or "").strip()
    text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\s*```$", "", text)
    return text.strip()


def extract_first_json_object(text):
    text = strip_code_fence(text)
    start = text.find("{")
    if start < 0:
        return text

    depth = 0
    in_string = False
    escape = False
    for idx in range(start, len(text)):
        char = text[idx]
        if in_string:
            if escape:
                escape = False
            elif char == "\\":
                escape = True
            elif char == '"':
                in_string = False
        else:
            if char == '"':
                in_string = True
            elif char == "{":
                depth += 1
            elif char == "}":
                depth -= 1
                if depth == 0:
                    return text[start:idx + 1]
    return text[start:]


def add_missing_json_closers(text):
    stack = []
    in_string = False
    escape = False
    for char in text:
        if in_string:
            if escape:
                escape = False
            elif char == "\\":
                escape = True
            elif char == '"':
                in_string = False
            continue
        if char == '"':
            in_string = True
        elif char in "[{":
            stack.append(char)
        elif char == "]" and stack and stack[-1] == "[":
            stack.pop()
        elif char == "}" and stack and stack[-1] == "{":
            stack.pop()

    closers = {"{": "}", "[": "]"}
    return text + "".join(closers[ch] for ch in reversed(stack))


def safe_json_loads(raw_output):
    if not raw_output or not str(raw_output).strip():
        raise ValueError("模型返回空内容")

    candidate = extract_first_json_object(str(raw_output))
    candidate = re.sub(r",\s*([}\]])", r"\1", candidate)

    try:
        return json.loads(candidate)
    except json.JSONDecodeError:
        repaired = add_missing_json_closers(candidate)
        repaired = re.sub(r",\s*([}\]])", r"\1", repaired)
        return json.loads(repaired)

# ### 基础字段标准化函数
#
# 统一分数、二元变量、政策基调和工具计数。

# 基础字段标准化函数
# 主要功能：把 LLM 输出中的布尔值、分数、政策基调、目标领域和工具计数统一转换为后续制表可用的数据类型。
def coerce_binary(value):
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)) and not pd.isna(value):
        return 1 if int(value) == 1 else 0
    text = str(value).strip().lower()
    if text in {"1", "yes", "true", "y", "有", "是", "多部门", "联合"}:
        return 1
    return 0


def coerce_score(value):
    try:
        score = float(value)
    except Exception:
        score = 0
    return int(max(0, min(100, round(score))))


def coerce_tone(value):
    text = str(value or "").strip().lower()
    if text in ALLOWED_TONES:
        return text
    if any(token in text for token in ["support", "扶持", "支持", "鼓励", "奖励", "补贴"]):
        return "supportive"
    if any(token in text for token in ["regulat", "监管", "规范", "限制", "管控", "合规"]):
        return "regulatory"
    return "neutral"


def coerce_scope(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return []
    if isinstance(value, list):
        items = value
    elif isinstance(value, str):
        value = value.strip()
        if value.startswith("["):
            try:
                parsed = ast.literal_eval(value)
                items = parsed if isinstance(parsed, list) else [value]
            except Exception:
                items = re.split(r"[、,，;；/\s]+", value)
        else:
            items = re.split(r"[、,，;；/\s]+", value)
    else:
        items = [str(value)]

    cleaned = []
    for item in items:
        if isinstance(item, dict):
            item = item.get("keyword") or item.get("name") or item.get("scope") or ""
        item = clean_text(item)
        if item and item not in cleaned:
            cleaned.append(item)
    return cleaned[:8]


def coerce_tool_count(value):
    if isinstance(value, str):
        try:
            value = json.loads(value)
        except Exception:
            try:
                value = ast.literal_eval(value)
            except Exception:
                value = {}
    if not isinstance(value, dict):
        value = {}

    normalized = {}
    for category in TOOL_CATEGORIES:
        raw_value = value.get(category, 0)
        try:
            count = int(float(raw_value))
        except Exception:
            count = 0
        normalized[category] = max(count, 0)
    return normalized


def ensure_list(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return []
    if isinstance(value, list):
        return value
    return [value]

# ### 阶段结果标准化、证据校验与筛选规则
#
# 规范化模型结果，并检查证据是否可在原文中回查。

# 阶段结果标准化、证据校验与筛选规则
# 主要功能：规范化各阶段 JSON 结果，检查证据片段能否在原文中找到，并按分数阈值判定是否进入正式编码。
def normalize_evidence_dict(evidence):
    if not isinstance(evidence, dict):
        evidence = {}
    normalized = {
        "policy_subject": ensure_list(evidence.get("policy_subject")),
        "target_scope": ensure_list(evidence.get("target_scope")),
        "policy_tone": ensure_list(evidence.get("policy_tone")),
        "tools": {},
        "target_specificity": ensure_list(evidence.get("target_specificity")),
        "coordination_breadth": ensure_list(evidence.get("coordination_breadth")),
    }
    tools = evidence.get("tools", {})
    if not isinstance(tools, dict):
        tools = {}
    for category in TOOL_CATEGORIES:
        normalized["tools"][category] = ensure_list(tools.get(category))
    return normalized


def normalize_evidence_result(result):
    if not isinstance(result, dict):
        raise ValueError(f"证据 JSON 根对象不是 dict：{type(result)}")
    score = coerce_score(result.get("digital_policy_score", 0))
    is_policy = result.get("is_digital_industrial_policy")
    if is_policy is None:
        is_policy = 1 if score >= DIGITAL_POLICY_SCORE_MIN else 0
    else:
        is_policy = coerce_binary(is_policy)
    return {
        "is_digital_industrial_policy": is_policy,
        "digital_policy_score": score,
        "non_policy_reason": clean_text(result.get("non_policy_reason", "")),
        "evidence": normalize_evidence_dict(result.get("evidence", {})),
        "confidence": coerce_score(result.get("confidence", score)),
    }


def normalize_title_check_result(result):
    if not isinstance(result, dict):
        raise ValueError(f"标题复核 JSON 根对象不是 dict：{type(result)}")
    return {
        "definitely_not_digital_policy_score": coerce_score(result.get("definitely_not_digital_policy_score", 0)),
        "title_based_reason": clean_text(result.get("title_based_reason", "")),
        "obvious_non_policy_type": clean_text(result.get("obvious_non_policy_type", "")),
    }


def normalize_confidence(value):
    if isinstance(value, dict):
        return {str(k): coerce_score(v) for k, v in value.items()}
    score = coerce_score(value)
    return {"overall": score}


def normalize_final_coding_result(result, evidence_result=None):
    if not isinstance(result, dict):
        raise ValueError(f"最终编码 JSON 根对象不是 dict：{type(result)}")
    evidence_result = evidence_result or {}
    score = coerce_score(result.get("digital_policy_score", evidence_result.get("digital_policy_score", 0)))
    is_policy = result.get("is_digital_industrial_policy", evidence_result.get("is_digital_industrial_policy", 0))
    normalized = {
        "is_digital_industrial_policy": coerce_binary(is_policy),
        "digital_policy_score": score,
        "policy_tone": coerce_tone(result.get("policy_tone")),
        "target_scope": coerce_scope(result.get("target_scope")),
        "tool_count": coerce_tool_count(result.get("tool_count")),
        "target_specificity": coerce_binary(result.get("target_specificity")),
        "coordination_breadth": coerce_binary(result.get("coordination_breadth")),
        "confidence": normalize_confidence(result.get("confidence", score)),
        "coding_notes": clean_text(result.get("coding_notes", "")),
        "evidence_used": result.get("evidence_used", {}),
    }
    return normalized


# 兼容旧函数名；新流程中最终编码仍通过该函数口径标准化。
def normalize_llm_result(result):
    return normalize_final_coding_result(result)


def normalize_for_match(text):
    return re.sub(r"\s+", "", clean_text(text))


def collect_evidence_quotes(obj):
    quotes = []
    if isinstance(obj, dict):
        for key, value in obj.items():
            if key in {"quote", "evidence", "evidence_text", "text", "原文", "原文证据"} and isinstance(value, str):
                q = clean_text(value)
                if len(q) >= 6:
                    quotes.append(q)
            else:
                quotes.extend(collect_evidence_quotes(value))
    elif isinstance(obj, list):
        for item in obj:
            quotes.extend(collect_evidence_quotes(item))
    return quotes


def verify_evidence_against_text(policy_text, evidence):
    text_norm = normalize_for_match(policy_text)
    quotes = collect_evidence_quotes(evidence)
    verified = 0
    for quote in quotes:
        quote_norm = normalize_for_match(quote).replace("……", "").replace("...", "")
        if quote_norm and quote_norm in text_norm:
            verified += 1
    total = len(quotes)
    rate = verified / total if total else 1.0
    return {"evidence_quote_count": total, "evidence_verified_count": verified, "evidence_verification_rate": round(rate, 4)}


def passes_digital_policy_filter(evidence_result, title_check_result=None):
    positive_score = coerce_score(evidence_result.get("digital_policy_score", 0))
    is_policy = coerce_binary(evidence_result.get("is_digital_industrial_policy", 0))
    if title_check_result is None:
        not_score = 0
    else:
        not_score = coerce_score(title_check_result.get("definitely_not_digital_policy_score", 0))
    return bool(
        is_policy == 1
        and positive_score >= DIGITAL_POLICY_SCORE_MIN
        and not_score < TITLE_NOT_POLICY_SCORE_MAX
        and positive_score - not_score >= SCORE_MARGIN_MIN
    )

# ### LLM 调用封装与失败结果格式
#
# 封装模型调用、重试和失败记录。

# LLM 调用封装与失败结果格式
# 主要功能：统一调用 JSON 模式、失败重试和普通 JSON 兜底，并在异常时返回结构一致的失败记录。
USE_NON_JSON_FALLBACK_ON_LAST_RETRY = True


def call_llm_json(prompt, task_name, max_retries=3, max_tokens=4096):
    if client is None:
        raise RuntimeError("未配置 API Key，无法调用 LLM。请先设置 DEEPSEEK_API_KEY 环境变量。")

    last_error = None
    for attempt in range(max_retries):
        use_json_mode = not (USE_NON_JSON_FALLBACK_ON_LAST_RETRY and attempt == max_retries - 1)
        try:
            request_kwargs = {
                "model": MODEL,
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.0,
                "max_tokens": max_tokens,
                "timeout": 90,
            }
            if use_json_mode:
                request_kwargs["response_format"] = {"type": "json_object"}

            response = client.chat.completions.create(**request_kwargs)
            raw_output = response.choices[0].message.content or ""
            if not raw_output.strip():
                raise ValueError("模型返回空内容")
            return {
                "parsed": safe_json_loads(raw_output),
                "raw_output": raw_output,
                "response_mode": "json_object" if use_json_mode else "plain_json_fallback",
                "prompt_chars": len(prompt),
            }
        except Exception as exc:
            last_error = exc
            mode_text = "JSON模式" if use_json_mode else "普通JSON兜底"
            print(f"{task_name} 第 {attempt + 1}/{max_retries} 次调用失败（{mode_text}）: {exc}")
            if attempt < max_retries - 1:
                wait_seconds = (2 ** attempt) + random.uniform(0, 0.5)
                time.sleep(wait_seconds)
    raise RuntimeError(f"{task_name} 达到最大重试次数：{last_error}")


def default_tool_count():
    return {category: 0 for category in TOOL_CATEGORIES}


def make_failed_analysis_result(error, prompt_chars=0):
    return {
        "is_digital_industrial_policy": None,
        "digital_policy_score": None,
        "title_not_policy_score": None,
        "policy_tone": None,
        "target_scope": [],
        "tool_count": default_tool_count(),
        "target_specificity": None,
        "coordination_breadth": None,
        "confidence": {},
        "evidence": {},
        "evidence_used": {},
        "coding_notes": "",
        "evidence_quote_count": 0,
        "evidence_verified_count": 0,
        "evidence_verification_rate": None,
        "_status": "failed",
        "_error": str(error),
        "_raw_output": "",
        "_evidence_raw_output": "",
        "_title_check_raw_output": "",
        "_coding_raw_output": "",
        "_review_raw_output": "",
        "_model": MODEL,
        "_base_url": BASE_URL,
        "_workflow_version": WORKFLOW_VERSION,
        "_response_mode": "failed",
        "_prompt_chars": prompt_chars,
        "_completed_at": datetime.now().isoformat(timespec="seconds"),
    }

# ### 单条政策的多阶段 LLM 分析
#
# 按证据抽取、标题复核、编码和复核处理单条政策。

# 单条政策的多阶段 LLM 分析
# 主要功能：串联证据抽取、标题反向复核、基于证据编码和可选复核修正，并汇总证据匹配率、原始输出和状态字段。
def analyze_policy(row, max_retries=3, max_prompt_chars=12000, evidence_max_tokens=6144, coding_max_tokens=4096):
    """按多阶段流程执行：全文证据抽取 -> 标题反向复核 -> 基于证据编码 -> 可选复核。"""
    evidence_raw = title_raw = coding_raw = review_raw = ""
    prompt_chars_total = 0
    try:
        evidence_prompt = build_evidence_prompt(row, max_chars=max_prompt_chars)
        evidence_call = call_llm_json(evidence_prompt, "证据抽取", max_retries=max_retries, max_tokens=evidence_max_tokens)
        evidence_raw = evidence_call["raw_output"]
        prompt_chars_total += evidence_call["prompt_chars"]
        evidence_result = normalize_evidence_result(evidence_call["parsed"])
        evidence_check = verify_evidence_against_text(row.get("policy_text", row.get("content", "")), evidence_result.get("evidence", {}))

        if ENABLE_TITLE_COUNTERCHECK:
            title_prompt = build_title_countercheck_prompt(row, evidence_result)
            title_call = call_llm_json(title_prompt, "标题反向复核", max_retries=max_retries, max_tokens=2048)
            title_raw = title_call["raw_output"]
            prompt_chars_total += title_call["prompt_chars"]
            title_check = normalize_title_check_result(title_call["parsed"])
            title_response_mode = title_call["response_mode"]
        else:
            title_response_mode = "not_run"
            title_check = {"definitely_not_digital_policy_score": 0, "title_based_reason": "未启用标题反向复核", "obvious_non_policy_type": ""}

        if not passes_digital_policy_filter(evidence_result, title_check):
            return {
                "is_digital_industrial_policy": 0,
                "digital_policy_score": evidence_result["digital_policy_score"],
                "title_not_policy_score": title_check["definitely_not_digital_policy_score"],
                "policy_tone": None,
                "target_scope": [],
                "tool_count": default_tool_count(),
                "target_specificity": None,
                "coordination_breadth": None,
                "confidence": {"evidence": evidence_result.get("confidence", 0)},
                "evidence": evidence_result.get("evidence", {}),
                "evidence_used": {},
                "coding_notes": evidence_result.get("non_policy_reason") or title_check.get("title_based_reason", ""),
                **evidence_check,
                "_status": "excluded_non_digital_policy",
                "_error": title_check.get("title_based_reason", "") or evidence_result.get("non_policy_reason", ""),
                "_raw_output": "",
                "_evidence_raw_output": evidence_raw,
                "_title_check_raw_output": title_raw,
                "_coding_raw_output": "",
                "_review_raw_output": "",
                "_model": MODEL,
                "_base_url": BASE_URL,
                "_workflow_version": WORKFLOW_VERSION,
                "_response_mode": f"evidence:{evidence_call['response_mode']};title:{title_response_mode}",
                "_prompt_chars": prompt_chars_total,
                "_completed_at": datetime.now().isoformat(timespec="seconds"),
            }

        coding_prompt = build_coding_prompt(row, evidence_result)
        coding_call = call_llm_json(coding_prompt, "基于证据编码", max_retries=max_retries, max_tokens=coding_max_tokens)
        coding_raw = coding_call["raw_output"]
        prompt_chars_total += coding_call["prompt_chars"]
        coding_result = normalize_final_coding_result(coding_call["parsed"], evidence_result=evidence_result)

        final_result = coding_result
        review_mode = "not_run"
        if ENABLE_REVIEW_ROUND:
            review_prompt = build_review_prompt(row, evidence_result, coding_result)
            review_call = call_llm_json(review_prompt, "复核修正", max_retries=max_retries, max_tokens=coding_max_tokens)
            review_raw = review_call["raw_output"]
            prompt_chars_total += review_call["prompt_chars"]
            final_result = normalize_final_coding_result(review_call["parsed"], evidence_result=evidence_result)
            review_mode = review_call["response_mode"]

        status = "ok" if final_result["is_digital_industrial_policy"] == 1 else "excluded_non_digital_policy"
        final_result.update({
            "title_not_policy_score": title_check["definitely_not_digital_policy_score"],
            "evidence": evidence_result.get("evidence", {}),
            **evidence_check,
            "_status": status,
            "_error": "" if status == "ok" else "最终复核判定为非数字经济产业政策",
            "_raw_output": review_raw or coding_raw,
            "_evidence_raw_output": evidence_raw,
            "_title_check_raw_output": title_raw,
            "_coding_raw_output": coding_raw,
            "_review_raw_output": review_raw,
            "_model": MODEL,
            "_base_url": BASE_URL,
            "_workflow_version": WORKFLOW_VERSION,
            "_response_mode": f"evidence:{evidence_call['response_mode']};title:{title_response_mode};coding:{coding_call['response_mode']};review:{review_mode}",
            "_prompt_chars": prompt_chars_total,
            "_completed_at": datetime.now().isoformat(timespec="seconds"),
        })
        return final_result
    except Exception as exc:
        failed = make_failed_analysis_result(exc, prompt_chars=prompt_chars_total)
        failed.update({
            "_evidence_raw_output": evidence_raw,
            "_title_check_raw_output": title_raw,
            "_coding_raw_output": coding_raw,
            "_review_raw_output": review_raw,
        })
        return failed

# ## 全量批处理与结果缓存
#
# 批处理逐条准备文本、调用多阶段 LLM，并即时写入 JSONL。默认 `RUN_FULL_BATCH=False`，避免误触发费用。

# ### 全量批处理开关与旧结果清理
#
# 控制是否全量调用 API，以及是否清理旧结果。

# 全量批处理开关与旧结果清理
# 主要功能：控制是否正式调用 API、是否限制调试条数、每批保存频率，以及正式重跑前清理哪些旧输出文件。
RESET_RESULT_FILES_BEFORE_RUN = True  # 全量重跑前是否清理旧结果文件
RUN_FULL_BATCH = True  # 是否正式调用 API；确认成本和 Key 后再改为 True
LOAD_EXISTING_RESULTS_WHEN_NOT_RUNNING = True  # 不调用 API 时是否加载已有 JSONL 结果
ROW_LIMIT = None  # 调试时限制处理条数；正式全量保持 None
BATCH_SIZE = 50  # 每处理多少条保存一次中间快照
SLEEP_SECONDS = 0.5  # 两次 LLM 调用之间的间隔，单位秒


def clear_llm_result_files():
    result_files = [
        RESULT_JSONL,
        INTERMEDIATE_CSV,
        INTERMEDIATE_XLSX,
        PARTIAL_CSV,
        PARTIAL_XLSX,
        FAILED_CSV,
        EXCLUDED_CSV,
        FINAL_CSV,
        FINAL_XLSX,
    ]
    for path in result_files:
        path = Path(path)
        if path.exists():
            path.unlink()

# ### 批处理执行与已有结果加载
#
# 逐条处理样本，或在未运行时加载旧结果。

# 批处理执行与已有结果加载
# 主要功能：逐条准备文本并调用多阶段 LLM；若不开启全量运行，则读取已有 JSONL 结果用于后续指标制表。
def run_llm_batch(source_df, batch_size=50, sleep_seconds=0.5, row_limit=None, reset_results=True):
    if client is None:
        raise RuntimeError("未配置 API Key，无法进行全量批处理。")

    if reset_results:
        clear_llm_result_files()
        print("已清理旧的 LLM 编码结果文件，本次将重新生成。")

    work_df = source_df.copy()
    if row_limit is not None:
        work_df = work_df.head(row_limit)

    print(f"本次候选处理 {len(work_df)} 条。网页文本会在每条调用 API 前按需获取；不会读取旧 LLM 结果。")
    print(f"工作流：{WORKFLOW_VERSION}；标题反向复核={ENABLE_TITLE_COUNTERCHECK}；复核修正={ENABLE_REVIEW_ROUND}")

    processed_since_snapshot = 0
    api_called_candidates = 0
    records = []
    for idx, row in tqdm(work_df.iterrows(), total=len(work_df), desc="Multistage LLM coding"):
        prepared_row = prepare_policy_row_for_llm(row, fetch_web_text=FETCH_WEB_TEXT_BEFORE_API)
        sync_prepared_row(source_df, idx, prepared_row)

        if SKIP_INCOMPLETE_TEXT_BEFORE_API and bool(prepared_row.get("content_incomplete_flag", False)):
            record = make_skipped_incomplete_record(prepared_row)
        else:
            result = analyze_policy(prepared_row)
            record = {**row_metadata(prepared_row), **result}
            api_called_candidates += 1

        append_jsonl(record, RESULT_JSONL)
        records.append(record)
        processed_since_snapshot += 1

        if processed_since_snapshot >= batch_size:
            snapshot = save_intermediate_snapshot()
            print(f"已保存中间结果：{len(snapshot)} 条 -> {INTERMEDIATE_CSV.name}")
            processed_since_snapshot = 0

        time.sleep(sleep_seconds)

    snapshot = save_intermediate_snapshot()
    print(f"批处理完成。进入多阶段 LLM 的候选 {api_called_candidates} 条；当前累计结果：{len(snapshot)} 条。")
    return pd.DataFrame(records)


if RUN_FULL_BATCH:
    llm_results_df = run_llm_batch(
        df,
        batch_size=BATCH_SIZE,
        sleep_seconds=SLEEP_SECONDS,
        row_limit=ROW_LIMIT,
        reset_results=RESET_RESULT_FILES_BEFORE_RUN,
    )
elif LOAD_EXISTING_RESULTS_WHEN_NOT_RUNNING and RESULT_JSONL.exists():
    llm_results_df = load_jsonl(RESULT_JSONL)
    print(f"RUN_FULL_BATCH=False，已加载已有 LLM 结果 {len(llm_results_df)} 条：{RESULT_JSONL}")
else:
    llm_results_df = pd.DataFrame()
    print("RUN_FULL_BATCH=False，未执行全量 API 调用，且未找到可加载的旧结果。确认后改为 True 正式运行多阶段 LLM 流程。")

# ## 核心解释变量构造
#
# 仅将通过筛选的政策转为指标表；排除、失败和跳过记录保留在中间结果中供复核。

# ### 输出序列化与指标列生成
#
# 把复杂 JSON 字段序列化，并生成政策工具变量。

# 输出序列化与指标列生成
# 主要功能：把 LLM 的复杂 JSON 字段转成可保存格式，并生成五类政策工具虚拟变量、政策基调和其他核心解释变量。
FINAL_CSV = PROJECT_ROOT / "gd21_policy_indicators_2017_2022.csv"  # 全量主结果 CSV 路径
FINAL_XLSX = PROJECT_ROOT / "gd21_policy_indicators_2017_2022.xlsx"  # 全量主结果 Excel 路径
PARTIAL_CSV = OUTPUT_DIR / "gd21_policy_indicators_partial_preview.csv"  # 部分结果预览 CSV 路径
PARTIAL_XLSX = OUTPUT_DIR / "gd21_policy_indicators_partial_preview.xlsx"  # 部分结果预览 Excel 路径
FAILED_CSV = OUTPUT_DIR / "gd21_policy_llm_failed_records.csv"  # 失败记录输出路径
EXCLUDED_CSV = OUTPUT_DIR / "gd21_policy_llm_excluded_records.csv"  # 被排除记录输出路径


def parse_scope_for_output(value):
    scope = coerce_scope(value)
    return "、".join(scope)


def json_dumps_for_output(value):
    if isinstance(value, str):
        return value
    try:
        return json.dumps(value, ensure_ascii=False)
    except TypeError:
        return json.dumps(jsonable(value), ensure_ascii=False)


def add_output_indicator_columns(source_df):
    """补齐输出表需要的规范化指标列，但不再构造工具强度合计指标。"""
    output_df = source_df.copy()
    if "tool_count" in output_df.columns:
        tool_counts = output_df["tool_count"].apply(coerce_tool_count)
    else:
        tool_counts = pd.Series([default_tool_count() for _ in range(len(output_df))], index=output_df.index)

    tool_df = pd.DataFrame(tool_counts.tolist(), index=output_df.index).reindex(columns=TOOL_CATEGORIES).fillna(0).astype(int)
    tool_binary_df = (tool_df > 0).astype(int)
    for category in TOOL_CATEGORIES:
        output_df[f"{category}_count"] = tool_df[category]
        output_df[f"{category}_tool"] = tool_binary_df[category]

    output_df["tone_dummy"] = output_df.get("policy_tone", pd.Series(index=output_df.index, dtype="object")).map({"supportive": 1, "neutral": 0, "regulatory": -1})
    output_df["target_specificity"] = output_df.get("target_specificity", pd.Series(index=output_df.index, dtype="object")).map(coerce_binary).astype(int)
    output_df["coordination_breadth"] = output_df.get("coordination_breadth", pd.Series(index=output_df.index, dtype="object")).map(coerce_binary).astype(int)
    if "target_scope" in output_df.columns:
        output_df["target_scope"] = output_df["target_scope"].map(parse_scope_for_output)
    output_df["tool_count_json"] = tool_counts.map(lambda item: json.dumps(item, ensure_ascii=False))

    for json_col in ["confidence", "evidence", "evidence_used", "tool_count"]:
        if json_col in output_df.columns:
            output_df[f"{json_col}_json"] = output_df[json_col].map(json_dumps_for_output)
    return output_df


def serialize_complex_columns(output_df):
    serialized_df = output_df.copy()
    for col in serialized_df.columns:
        if serialized_df[col].map(lambda value: isinstance(value, (dict, list))).any():
            serialized_df[col] = serialized_df[col].map(json_dumps_for_output)
    return serialized_df

# ### 生成最终指标表并写入文件
#
# 写出主结果、中间结果和状态记录。

# 生成最终指标表并写入文件
# 主要功能：仅保留成功通过筛选的数字经济产业政策，输出主结果表、完整中间表、失败记录和被排除记录。
def build_indicator_table(results_df, expected_count=None):
    if results_df.empty:
        print("暂无 LLM 结果。请先将 RUN_FULL_BATCH=True 完成全量 API 调用，或确认 RESULT_JSONL 指向已有结果文件。")
        return pd.DataFrame(), pd.DataFrame()

    results_df = results_df.copy()
    if "_status" not in results_df.columns:
        results_df["_status"] = "ok"

    complete_df = add_output_indicator_columns(results_df)
    excluded_count = int(complete_df["_status"].eq("excluded_non_digital_policy").sum())
    failed_df = results_df[~results_df["_status"].isin(["ok", "excluded_non_digital_policy"])].copy()

    success_df = complete_df[complete_df["_status"].eq("ok")].copy()
    if "is_digital_industrial_policy" in success_df.columns:
        success_df = success_df[success_df["is_digital_industrial_policy"].map(coerce_binary).eq(1)].copy()

    if success_df.empty:
        print("没有通过筛选并成功解析的数字经济产业政策，无法构造指标。")
        return complete_df, pd.DataFrame()

    if globals().get("DROP_DUPLICATE_POLICIES", True) and {"city", "title", "year"}.issubset(success_df.columns):
        before_dedup = len(success_df)
        success_df["_title_norm_for_dedup"] = success_df["title"].map(normalize_title)
        success_df = (
            success_df.sort_values("content_chars", ascending=False, na_position="last")
            .drop_duplicates(subset=["city", "_title_norm_for_dedup", "year"], keep="first")
        )
        if before_dedup > len(success_df):
            print(f"最终表去重：删除 {before_dedup - len(success_df)} 条重复记录。")

    main_df = success_df.copy()
    main_df["topic"] = main_df["keyword"] if "keyword" in main_df.columns else ""
    main_columns = [
        "topic", "city", "year", "date", "title", "url", "tone_dummy",
        "fiscal_tool", "financial_tool", "supply_tool", "demand_tool", "regulation_tool",
        "target_specificity", "coordination_breadth",
    ]
    for col in main_columns:
        if col not in main_df.columns:
            main_df[col] = pd.NA
    main_df = main_df[main_columns].sort_values(["year", "city", "date", "title"], na_position="last")
    complete_df = serialize_complex_columns(complete_df).sort_values(["year", "city", "date", "title"], na_position="last")

    if expected_count is None and "df" in globals():
        expected_count = len(df)
    # 多阶段流程会正常排除非政策文本；这不应被视为 partial。
    # 只有候选未全部处理，或存在正文跳过/API失败等阻塞记录时，才保存为部分结果。
    processed_count = len(results_df)
    blocking_failure_count = len(failed_df)
    is_partial = bool(expected_count) and (processed_count < expected_count or blocking_failure_count > 0)

    output_csv = PARTIAL_CSV if is_partial else FINAL_CSV
    output_xlsx = PARTIAL_XLSX if is_partial else FINAL_XLSX
    for stale_path in [PARTIAL_CSV, FINAL_CSV, FAILED_CSV, EXCLUDED_CSV, INTERMEDIATE_CSV, INTERMEDIATE_XLSX]:
        stale_path = Path(stale_path)
        if stale_path.exists():
            stale_path.unlink()

    main_df.to_csv(output_csv, index=False, encoding="utf-8-sig")

    if not failed_df.empty:
        serialize_complex_columns(failed_df).to_csv(FAILED_CSV, index=False, encoding="utf-8-sig")

    excluded_df = complete_df[complete_df["_status"].eq("excluded_non_digital_policy")].copy()
    if not excluded_df.empty:
        excluded_df.to_csv(EXCLUDED_CSV, index=False, encoding="utf-8-sig")

    with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
        complete_df.to_excel(writer, sheet_name="complete_intermediate", index=False)
        main_df.to_excel(writer, sheet_name="main_results", index=False)

    if is_partial:
        print(f"当前候选处理 {processed_count}/{expected_count} 条，排除 {excluded_count} 条，阻塞失败/跳过 {blocking_failure_count} 条；主结果表 {len(main_df)} 条。已保存为部分结果预览。")
    else:
        print(f"候选已全部处理；排除非政策文本后，主结果表 {len(main_df)} 条。")
    print(f"CSV：{output_csv}")
    print(f"Excel：{output_xlsx}")
    print("Sheet：complete_intermediate（完整中间输出），main_results（主要结果）")
    return complete_df, main_df


llm_results_df = globals().get("llm_results_df", pd.DataFrame())
expected_policy_count = len(df) if "df" in globals() else None
complete_intermediate_df, final_df = build_indicator_table(llm_results_df, expected_count=expected_policy_count)
if not final_df.empty:
    display(final_df.head())

# ## 指标定义
#
# 核心变量：五类政策工具、量化目标、联合发文和政策基调。工具变量由 `tool_count > 0` 转为 0/1。

# ### 运行质量检查与分布统计
#
# 输出状态分布、证据匹配和人工复核线索。

# 运行质量检查与分布统计
# 主要功能：查看 LLM 状态分布、核心指标分布、证据片段匹配率和城市-年份样本量，辅助正式回归前复核。
if "llm_results_df" in globals() and not llm_results_df.empty and "_status" in llm_results_df.columns:
    print("LLM 全流程状态分布：")
    display(llm_results_df["_status"].value_counts(dropna=False).to_frame("count"))

if "final_df" not in globals() or final_df.empty:
    print("暂无本次运行生成的最终指标表，跳过质量检查。")
else:
    print("情感指标 tone_dummy 分布：")
    display(final_df["tone_dummy"].value_counts(dropna=False).sort_index().to_frame("count"))

    tool_cols = ["fiscal_tool", "financial_tool", "supply_tool", "demand_tool", "regulation_tool"]
    print("五类工具变量出现次数：")
    display(final_df[tool_cols].sum().astype(int).to_frame("count"))

    print("target_specificity / coordination_breadth 分布：")
    display(final_df[["target_specificity", "coordination_breadth"]].apply(pd.Series.value_counts).fillna(0).astype(int))

    if "complete_intermediate_df" in globals() and not complete_intermediate_df.empty and "evidence_verification_rate" in complete_intermediate_df.columns:
        print("证据片段原文匹配率概览：")
        ok_evidence_df = complete_intermediate_df[complete_intermediate_df["_status"].eq("ok")].copy()
        display(pd.to_numeric(ok_evidence_df["evidence_verification_rate"], errors="coerce").describe().to_frame("evidence_verification_rate"))
        low_evidence = ok_evidence_df[pd.to_numeric(ok_evidence_df["evidence_verification_rate"], errors="coerce").fillna(1) < 0.8]
        if not low_evidence.empty:
            print(f"证据匹配率低于 0.8 的记录 {len(low_evidence)} 条，建议人工抽查：")
            display(low_evidence[["policy_id", "city", "title", "evidence_verification_rate", "url"]].head(20))

    print("城市-年份样本量：")
    display(final_df.pivot_table(index="city", columns="year", values="title", aggfunc="count", fill_value=0))

# ## 复核重点
#
# 重点抽查：证据匹配率低、工具强度高、`target_specificity=1`、`coordination_breadth=1`、标题含解读/公示/名单/通知壳的样本。

# ## 运行顺序
#
# 依次运行：依赖与数据读取 -> 网页补全函数 -> Prompt 与解析函数 -> API 配置 -> 批处理 -> 指标构造与质量检查。
