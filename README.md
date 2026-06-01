# 广东数字经济政策文本采集与 LLM 指标构建

本仓库用于复现广东数字经济政策文本的采集、清洗、网页全文补全、LLM 多阶段识别和政策指标构建流程。

公开 GitHub 仓库仅保存代码和说明文档，不上传当前本地生成的数据、缓存、模型输出、Excel 结果、PDF 汇报材料或其他中间文件。所有数据文件和结果文件都需要在本地运行脚本后重新生成。

## 仓库文件

GitHub 仓库只保留以下文件：

| 文件                                  | 说明                                               |
| ------------------------------------- | -------------------------------------------------- |
| `README.md`                           | 项目说明、运行步骤和输出说明                       |
| `gd_digital_policy_crawler.py`        | 广东数字经济政策文本采集脚本                       |
| `gd_digital_policy_llm_indicators.py` | 政策文本清洗、网页全文补全、LLM 编码和指标构建脚本 |

以下内容不随 GitHub 提交：

| 本地文件或目录                          | 说明                                 |
| --------------------------------------- | ------------------------------------ |
| `gd21_digital_policies_2017_2022.csv`   | 运行爬虫脚本后生成的原始政策文本数据 |
| `gd21_digital_policies_2017_2022.xlsx`  | LLM 指标脚本默认读取的本地输入文件   |
| `gd21_policy_indicators_2017_2022.csv`  | 全量运行完成后生成的主结果 CSV       |
| `gd21_policy_indicators_2017_2022.xlsx` | 全量运行完成后生成的主结果 Excel     |

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

研究对象是广东 21 个地级市在 2017-2022 年间与数字经济相关的政策文本。项目目标是把政府网站上的非结构化政策文本转化为可复核、可统计、可进入后续实证分析的政策级结构化变量。

## 运行环境

建议使用 Python 3.10 或更高版本。

安装主要依赖：

```bash
pip install pandas openpyxl requests beautifulsoup4 openai tqdm
```

两个脚本都会在运行时检查部分依赖。为了减少运行中断，建议先手动安装上述包。

## 第一步：采集政策文本

运行：

```bash
python gd_digital_policy_crawler.py
```

脚本默认调用广东省政府政策文件检索接口：

```python
search_api_url = "https://search.gd.gov.cn/api/search/file"
```

默认采集口径：

| 项目       | 当前设置                                                     |
| ---------- | ------------------------------------------------------------ |
| 检索关键词 | `数字经济`                                                   |
| 时间范围   | 2017-2022 年                                                 |
| 空间范围   | 广东 21 个地级市                                             |
| 标题排除   | 标题中包含 `已废止` 的结果会被排除                           |
| 输出字段   | `keyword`, `city`, `title`, `date`, `year`, `url`, `content` |

运行完成后，本地会生成：

```text
gd21_digital_policies_2017_2022.csv
```

该 CSV 是本地运行产物，不上传到 GitHub。

## 爬虫设计说明

政府网站详情页可能出现 TLS 握手失败、证书链不完整、证书过期或 Python `requests` 与目标站点兼容性差等问题。采集脚本对详情页正文抓取采用三层兜底：

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

脚本还会根据标题和正文中的城市名称识别政策归属城市。若没有命中广东 21 个地级市，则标记为 `广东省级/未识别`，后续可人工复核。

## 第二步：准备 LLM 脚本输入

`gd_digital_policy_llm_indicators.py` 默认从项目根目录读取：

```text
gd21_digital_policies_2017_2022.xlsx
```

默认工作表名称为：

```text
gd21_digital_policies_2017_2022
```

如果第一步生成的是 CSV，可以先转换为脚本默认读取的 Excel 文件：

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

输入数据至少需要包含以下字段：

| 字段      | 说明                     |
| --------- | ------------------------ |
| `keyword` | 检索关键词               |
| `city`    | 城市归属                 |
| `title`   | 政策标题                 |
| `date`    | 发布日期                 |
| `year`    | 发布年份                 |
| `url`     | 政策网页链接             |
| `content` | 已抓取正文或网页文本片段 |

## 第三步：配置 LLM API

`gd_digital_policy_llm_indicators.py` 使用 OpenAI 兼容 SDK。默认读取 DeepSeek 环境变量：

```python
API_KEY = os.getenv("DEEPSEEK_API_KEY", "").strip()
BASE_URL = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com").strip()
MODEL = os.getenv("DEEPSEEK_MODEL", "deepseek-chat").strip()
```

运行前在终端设置：

```bash
export DEEPSEEK_API_KEY="your_api_key"
export DEEPSEEK_BASE_URL="https://api.deepseek.com"
export DEEPSEEK_MODEL="deepseek-chat"
```

不要把真实 API Key 写进代码、README、运行输出或 GitHub 仓库。

## 第四步：运行 LLM 指标构建

运行：

```bash
python gd_digital_policy_llm_indicators.py
```

脚本会完成：

- 字段名统一、日期解析和文本清洗；
- 生成 `policy_id`；
- 按城市、标题和年份去重；
- 标记疑似不完整正文；
- 按 URL 回溯抓取网页全文；
- 使用 LLM 多阶段识别数字经济产业政策；
- 抽取证据并基于证据编码；
- 构建政策级指标表；
- 输出质量检查信息。

## 批处理开关

脚本默认不会自动全量调用 API：

```python
RUN_FULL_BATCH = False
```

确认 API Key、样本口径和调用成本后，再改为：

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

| 参数                                     | 含义                                 |
| ---------------------------------------- | ------------------------------------ |
| `DROP_DUPLICATE_POLICIES`                | 是否按城市、标题和年份去重           |
| `KEEP_ONLY_21_CITIES`                    | 是否只保留广东 21 个地级市样本       |
| `FETCH_WEB_TEXT_BEFORE_API`              | 调用 API 前是否按 URL 回溯网页全文   |
| `FORCE_REFETCH_WEB_TEXT`                 | 是否忽略网页缓存并强制重抓           |
| `SKIP_INCOMPLETE_TEXT_BEFORE_API`        | 正文仍不完整时是否跳过 LLM           |
| `ENABLE_TITLE_COUNTERCHECK`              | 是否启用标题反向复核                 |
| `ENABLE_REVIEW_ROUND`                    | 是否启用最终复核修正                 |
| `RESET_RESULT_FILES_BEFORE_RUN`          | 全量重跑前是否清理旧结果文件         |
| `LOAD_EXISTING_RESULTS_WHEN_NOT_RUNNING` | 不调用 API 时是否读取已有 JSONL 结果 |

## 本地输出文件

运行 LLM 指标脚本后，常见本地输出包括：

| 文件                                                         | 说明                      |
| ------------------------------------------------------------ | ------------------------- |
| `outputs/chapter3_llm_policy_indicators/gd21_policy_llm_results.jsonl` | 逐条追加保存的 LLM 主结果 |
| `outputs/chapter3_llm_policy_indicators/gd21_policy_web_fulltext_cache.jsonl` | URL 网页全文缓存          |
| `outputs/chapter3_llm_policy_indicators/gd21_policy_web_fulltext_review.csv` | 网页补全复核表            |
| `outputs/chapter3_llm_policy_indicators/gd21_policy_indicators_partial_preview.csv` | 部分结果预览 CSV          |
| `outputs/chapter3_llm_policy_indicators/gd21_policy_indicators_partial_preview.xlsx` | 部分结果预览 Excel        |
| `gd21_policy_indicators_2017_2022.csv`                       | 全量完成后的主结果 CSV    |
| `gd21_policy_indicators_2017_2022.xlsx`                      | 全量完成后的主结果 Excel  |

实际是否生成某个文件，取决于是否运行全量批处理、是否已有可加载的 JSONL 结果，以及脚本当前参数设置。上述输出不上传到 GitHub。

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

LLM 指标脚本包含以下质量控制机制：

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

如果更换研究主题，建议同步修改 `gd_digital_policy_llm_indicators.py` 中的政策定义、证据抽取 prompt、结构化编码 prompt 和指标解释，避免爬虫口径与 LLM 编码口径不一致。

## 注意事项

- 政府网页结构和接口返回可能变化，爬虫结果需要以实际运行情况为准。
- LLM 调用会产生费用，全量运行前应先用小样本测试。
- 本仓库不包含真实 API Key。
- 本仓库不包含当前本地输出，因此 README 不声称具体样本数量或结果行数。
- 若需要完全复现实证结果，应在本地保留对应版本的数据文件、JSONL 中间结果和参数设置。
