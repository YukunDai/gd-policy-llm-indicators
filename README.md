# 广东数字经济政策文本采集与 LLM 指标构建

本仓库用于复现广东数字经济政策文本的采集、清洗、网页全文补全、LLM 多阶段识别和政策指标构建流程。

注意：GitHub 仓库只保留代码和说明文档，不上传当前本地生成的数据、缓存、模型输出、Excel 结果、PDF 汇报材料或其他中间文件。所有输出文件都需要在本地运行脚本或 notebook 后重新生成。

## 项目流程

```text
政策检索
  -> 网页爬取
  -> 文本清洗
  -> 网页全文补全
  -> LLM 多阶段识别
  -> 政策指标构建
  -> 质量检查
```

研究对象是广东 21 个地级市在 2017-2022 年间与数字经济相关的政策文本。项目目标不是只做网页爬虫，而是把政府网站上的非结构化政策文本转化为可复核、可统计、可进入后续实证分析的政策级结构化变量。

## 仓库文件

公开仓库中保留以下文件：

| 文件                                               | 说明                         |
| -------------------------------------------------- | ---------------------------- |
| `README.md`                                        | 项目说明和复现步骤           |
| `gd_digital_policy_crawler.py`                     | 广东数字经济政策文本采集脚本 |
| `chapter3_llm_policy_indicator_construction.ipynb` | LLM 政策指标构建 notebook    |

以下文件属于本地运行产物，不随 GitHub 提交：

| 本地文件或目录                            | 生成方式                                           |
| ----------------------------------------- | -------------------------------------------------- |
| `gd21_digital_policies_2017_2022.csv`     | 运行爬虫脚本后生成                                 |
| `gd21_digital_policies_2017_2022.xlsx`    | notebook 默认读取的本地输入文件，可由 CSV 转换得到 |
| `outputs/chapter3_llm_policy_indicators/` | 运行 notebook 后生成的 LLM 结果、网页缓存和复核表  |
| `gd21_policy_indicators_2017_2022.csv`    | 全量运行完成后生成的主结果表                       |
| `gd21_policy_indicators_2017_2022.xlsx`   | 全量运行完成后生成的主结果 Excel                   |

本 README 不记录本地已有输出的行数、样本数或具体结果，避免把未提交的运行状态误写成公开仓库内容。

## 运行环境

建议使用 Python 3.10 或更高版本。

主要依赖：

```bash
pip install pandas openpyxl requests beautifulsoup4 openai tqdm jupyter
```

如果使用 notebook 中的自动依赖检查，也可以先打开 notebook 运行前置单元格。

## 数据采集

数据采集由 `gd_digital_policy_crawler.py` 完成。

脚本调用广东省政府政策文件检索接口：

```python
search_api_url = "https://search.gd.gov.cn/api/search/file"
```

当前默认采集口径：

| 项目       | 当前设置                                                     |
| ---------- | ------------------------------------------------------------ |
| 检索关键词 | `数字经济`                                                   |
| 时间范围   | 2017-2022 年                                                 |
| 空间范围   | 广东 21 个地级市                                             |
| 标题排除   | 标题中包含 `已废止` 的结果会被排除                           |
| 输出字段   | `keyword`, `city`, `title`, `date`, `year`, `url`, `content` |

运行：

```bash
python gd_digital_policy_crawler.py
```

运行完成后会在本地生成：

```text
gd21_digital_policies_2017_2022.csv
```

该 CSV 是本地数据文件，不上传到 GitHub。

## 爬虫关键设计

政府网站详情页可能出现 TLS 握手失败、证书链不完整、证书过期或 Python `requests` 与目标站点兼容性差等问题。爬虫对详情页正文抓取采用三层兜底：

1. 优先使用 `requests`。
2. 如果出现 TLS 或连接异常，改用系统 `curl`。
3. 如果是证书校验异常，再尝试 `curl --insecure`。

相关函数包括：

```python
fetch_detail_html_with_requests()
fetch_detail_html_with_curl()
is_tls_compat_error()
is_certificate_verify_error()
get_detail_content()
```

脚本还会根据标题和正文中的城市名称识别政策归属城市。若没有命中广东 21 个地级市，则标记为 `广东省级/未识别`，后续可以人工复核。

## 准备 notebook 输入

`chapter3_llm_policy_indicator_construction.ipynb` 默认从项目根目录读取：

```text
gd21_digital_policies_2017_2022.xlsx
```

并默认读取工作表：

```text
gd21_digital_policies_2017_2022
```

如果先运行爬虫得到了 CSV，可以在本地转换为 notebook 默认需要的 Excel 文件：

```bash
python - <<'PY'
import pandas as pd

src = "gd21_digital_policies_2017_2022.csv"
dst = "gd21_digital_policies_2017_2022.xlsx"
sheet = "gd21_digital_policies_2017_2022"

df = pd.read_csv(src)
df.to_excel(dst, sheet_name=sheet, index=False)
PY
```

也可以直接在 notebook 中把 `DATA_FILE` 和读取方式改为 CSV。无论采用哪种方式，输入数据至少需要包含以下字段：

| 字段      | 说明                     |
| --------- | ------------------------ |
| `keyword` | 检索关键词               |
| `city`    | 城市归属                 |
| `title`   | 政策标题                 |
| `date`    | 发布日期                 |
| `year`    | 发布年份                 |
| `url`     | 政策网页链接             |
| `content` | 已抓取正文或网页文本片段 |

## LLM 指标构建

主流程位于：

```text
chapter3_llm_policy_indicator_construction.ipynb
```

notebook 会完成：

- 字段名统一、日期解析和文本清洗；
- 生成 `policy_id`；
- 按城市、标题和年份去重；
- 标记疑似不完整正文；
- 按 URL 回溯抓取网页全文；
- 使用 LLM 多阶段识别数字经济产业政策；
- 抽取证据并基于证据编码；
- 输出政策级指标表和质量检查字段。

notebook 使用 OpenAI 兼容 SDK。默认配置为 DeepSeek：

```python
API_KEY = os.getenv("DEEPSEEK_API_KEY", "").strip()
BASE_URL = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com").strip()
MODEL = os.getenv("DEEPSEEK_MODEL", "deepseek-chat").strip()
```

运行前在终端设置环境变量：

```bash
export DEEPSEEK_API_KEY="your_api_key"
export DEEPSEEK_BASE_URL="https://api.deepseek.com"
export DEEPSEEK_MODEL="deepseek-chat"
```

不要把真实 API Key 写进代码、notebook 输出或 GitHub 仓库。

## 批处理开关

notebook 默认不会自动全量调用 API：

```python
RUN_FULL_BATCH = False
```

确认 API Key、样本口径和成本后，再改为：

```python
RUN_FULL_BATCH = True
```

调试时建议先限制样本数：

```python
ROW_LIMIT = 5
```

正式全量运行时再设为：

```python
ROW_LIMIT = None
```

其他常用参数：

| 参数                              | 含义                               |
| --------------------------------- | ---------------------------------- |
| `DROP_DUPLICATE_POLICIES`         | 是否按城市、标题和年份去重         |
| `FETCH_WEB_TEXT_BEFORE_API`       | 调用 API 前是否按 URL 回溯网页全文 |
| `FORCE_REFETCH_WEB_TEXT`          | 是否忽略网页缓存并强制重抓         |
| `SKIP_INCOMPLETE_TEXT_BEFORE_API` | 正文仍不完整时是否跳过 LLM         |
| `ENABLE_REVIEW_ROUND`             | 是否启用最终复核修正               |
| `RESET_RESULT_FILES_BEFORE_RUN`   | 全量重跑前是否清理旧结果文件       |

## 运行后生成的输出

运行 notebook 后，主要输出会在本地生成，不属于 GitHub 仓库内容。

常见输出包括：

| 文件                                                         | 说明                      |
| ------------------------------------------------------------ | ------------------------- |
| `outputs/chapter3_llm_policy_indicators/gd21_policy_llm_results.jsonl` | 逐条追加保存的 LLM 主结果 |
| `outputs/chapter3_llm_policy_indicators/gd21_policy_web_fulltext_cache.jsonl` | URL 网页全文缓存          |
| `outputs/chapter3_llm_policy_indicators/gd21_policy_web_fulltext_review.csv` | 网页补全复核表            |
| `outputs/chapter3_llm_policy_indicators/gd21_policy_indicators_partial_preview.csv` | 部分结果预览 CSV          |
| `outputs/chapter3_llm_policy_indicators/gd21_policy_indicators_partial_preview.xlsx` | 部分结果预览 Excel        |
| `gd21_policy_indicators_2017_2022.csv`                       | 全量完成后的主结果 CSV    |
| `gd21_policy_indicators_2017_2022.xlsx`                      | 全量完成后的主结果 Excel  |

实际是否生成某个文件，取决于是否运行全量批处理、是否已有可加载的 JSONL 结果，以及 notebook 当前参数设置。

## 核心指标

主结果表面向政策级分析，核心变量包括：

| 指标                   | 含义                                                         | 取值                                         |
| ---------------------- | ------------------------------------------------------------ | -------------------------------------------- |
| `fiscal_tool`          | 是否使用财政支持工具，如补贴、奖励、专项资金、税费优惠       | 0/1                                          |
| `financial_tool`       | 是否使用金融支持工具，如信贷、融资、基金、担保               | 0/1                                          |
| `supply_tool`          | 是否使用供给型工具，如平台、基础设施、人才、技术、数据资源   | 0/1                                          |
| `demand_tool`          | 是否使用需求型工具，如应用场景、试点示范、政府采购、市场推广 | 0/1                                          |
| `regulation_tool`      | 是否使用监管型工具，如标准规范、数据安全、准入管理           | 0/1                                          |
| `target_specificity`   | 是否包含明确量化发展目标                                     | 0/1                                          |
| `coordination_breadth` | 是否多部门联合发文                                           | 0/1                                          |
| `tone_dummy`           | 政策基调                                                     | `supportive=1`, `neutral=0`, `regulatory=-1` |

## 质量控制

notebook 中的质量控制包括：

- JSON 输出约束和解析修复；
- API 调用失败重试；
- JSONL 逐条写入，支持断点续跑；
- 网页全文缓存，避免重复抓取；
- 证据片段回查原文；
- `_status` 字段记录成功、排除、跳过或失败原因；
- 标题反向复核，用于剔除解读、新闻、公示、名单、通知壳等非正式政策文本。

常见状态包括：

| 状态                          | 含义                                   |
| ----------------------------- | -------------------------------------- |
| `ok`                          | 成功编码并进入主结果候选               |
| `excluded_non_digital_policy` | 被判断为非数字经济产业政策或非正式政策 |
| `skipped_incomplete_text`     | 正文不完整，跳过 LLM 编码              |

## 修改研究口径

修改检索主题时，在 `gd_digital_policy_crawler.py` 中调整：

```python
keywords = ["数字经济"]
```

修改年份范围时，调整：

```python
start_year = 2017
end_year = 2022
```

如果更换研究主题，建议同步修改 notebook 中的政策定义、证据抽取 prompt、结构化编码 prompt 和指标解释，避免爬虫口径与 LLM 编码口径不一致。

## 注意事项

- 政府网页结构和接口返回可能变化，爬虫结果需要以实际运行情况为准。
- LLM 调用会产生费用，全量运行前应先用小样本测试。
- 本仓库不包含真实 API Key。
- 本仓库不包含当前本地输出，因此 README 不声称具体样本数量或结果行数。
- 若需要完全复现实证结果，应在本地保留对应版本的数据文件、JSONL 中间结果和参数设置。
