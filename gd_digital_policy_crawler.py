import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import random
import shutil
import subprocess
from urllib.parse import urljoin, urlparse
import re


# =========================
# 1. 基本参数
# =========================

cities = [  # 广东 21 个地级市；城市识别时会在标题和正文中匹配这些名称
    '广州', '深圳', '珠海', '汕头', '佛山', '韶关', '河源', '梅州',
    '惠州', '汕尾', '东莞', '中山', '江门', '阳江', '湛江', '茂名',
    '肇庆', '清远', '潮州', '揭阳', '云浮'
]

keywords = [  # 政策检索关键词；换研究主题时优先修改这里
    '数字经济'
]
excluded_title_markers = ['已废止']  # 标题预筛选排除词；命中后不再抓取详情页

start_year = 2017  # 检索起始年份
end_year = 2022  # 检索结束年份

headers = {  # 政府检索接口请求头，模拟浏览器 AJAX 请求
    'User-Agent': (
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/120.0.0.0 Safari/537.36'
    ),
    'X-Requested-With': 'XMLHttpRequest',
    'Referer': 'https://search.gd.gov.cn/search/file/2',
    'Connection': 'close',
}

detail_headers = {  # 详情页请求头；与搜索接口分开，便于适配 HTML 页面
    'User-Agent': (
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/120.0.0.0 Safari/537.36'
    ),
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'zh-CN,zh;q=0.9',
    'Accept-Encoding': 'gzip, deflate',
    'Connection': 'close',
}

search_api_url = 'https://search.gd.gov.cn/api/search/file'  # 广东省政府政策文件检索接口
search_range = 'province'  # 检索范围：province=全省；site=仅广东省人民政府门户网站
search_max_retries = 4  # 搜索接口最大重试次数
search_base_delay = 1  # 搜索接口重试初始等待秒数
search_max_delay = 8  # 搜索接口重试最大等待秒数
detail_max_retries = 4  # 详情页 requests 抓取最大重试次数
detail_base_delay = 1  # 详情页重试初始等待秒数
detail_max_delay = 8  # 详情页重试最大等待秒数
search_page_delay_range = (2.0, 5.0)  # 翻页之间的随机暂停范围
detail_item_delay_range = (3.0, 7.0)  # 详情页之间的随机暂停范围
same_host_min_delay = 3.0  # 同一域名详情页请求的最小间隔
detail_failure_cooldown_threshold = 5  # 连续正文为空达到该次数后进入长暂停
detail_failure_cooldown_range = (60.0, 120.0)  # 连续失败后的长暂停范围

all_policies = []  # 采集到的政策记录列表
visited_urls = set()  # 已访问 URL，用于去重
last_detail_request_at = {}  # 记录每个域名最近一次详情页请求时间


# =========================
# 2. 工具函数
# =========================

def get_backoff_wait_seconds(attempt, base_delay, max_delay):
    """计算指数退避等待时间，并加入随机抖动。"""
    delay = min(base_delay * (2 ** attempt), max_delay)
    jitter = random.uniform(0, delay * 0.5)
    return delay + jitter


def sleep_random(delay_range):
    """在指定范围内随机暂停，避免请求节奏过于机械。"""
    time.sleep(random.uniform(*delay_range))


def wait_for_same_host(url):
    """同一域名详情页请求之间保持最小间隔。"""
    host = urlparse(url).netloc
    if not host:
        return

    now = time.time()
    last_requested_at = last_detail_request_at.get(host)
    if last_requested_at is not None:
        elapsed = now - last_requested_at
        if elapsed < same_host_min_delay:
            time.sleep(same_host_min_delay - elapsed + random.uniform(0, 1.5))

    last_detail_request_at[host] = time.time()


def clean_text(text):
    """清洗文本"""
    if not text:
        return ''
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


def normalize_title_for_match(text):
    """标题匹配前去掉空白，兼容“数 字 经 济”这类标题。"""
    if not text:
        return ''
    return re.sub(r'\s+', '', text)


def title_mentions_keyword(title, keyword):
    """只保留标题中出现当前关键词的结果。"""
    return normalize_title_for_match(keyword) in normalize_title_for_match(title)


def is_excluded_title(title):
    """过滤已废止等不需要的政策。"""
    return any(marker in title for marker in excluded_title_markers)


def extract_year(date_text):
    """从日期文本中提取年份"""
    if not date_text:
        return None

    match = re.search(r'(20\d{2})', date_text)
    if match:
        return int(match.group(1))
    return None


def judge_city(title, content):
    """
    根据标题和正文判断政策所属城市。
    如果文本中出现某个城市名，就认为属于该城市。
    如果没有匹配到，则返回“广东省级/未识别”。
    """
    text = f"{title} {content}"

    for city in cities:
        if city in text or f"{city}市" in text:
            return city

    return '广东省级/未识别'


def extract_detail_content(html):
    """从详情页 HTML 中提取正文。"""
    soup = BeautifulSoup(html, 'html.parser')

    # 广东政府网站正文区域可能使用的几种常见写法
    content_div = (
        soup.find('div', class_='content') or
        soup.find('div', class_='detail-content') or
        soup.find('div', class_='article-content') or
        soup.find('div', id='zoomcon') or
        soup.find('div', id='UCAP-CONTENT') or
        soup.find('div', class_='TRS_Editor')
    )

    if content_div:
        content = content_div.get_text('\n', strip=True)
    else:
        # 如果找不到正文区域，就退而求其次抓取整个页面文字
        content = soup.get_text('\n', strip=True)

    return clean_text(content)


def fetch_detail_html_with_requests(url):
    """每次请求都使用新的 Session，避免复用异常连接。"""
    wait_for_same_host(url)
    with requests.Session() as session:
        resp = session.get(url, headers=detail_headers, timeout=(10, 30))
        resp.raise_for_status()
        resp.encoding = resp.apparent_encoding
        return resp.text


def fetch_detail_html_with_curl(url, insecure=False):
    """使用系统 curl 兜底，绕过部分站点与 Python/OpenSSL 的 TLS 兼容问题。"""
    wait_for_same_host(url)
    curl_path = shutil.which('curl')
    if not curl_path:
        raise RuntimeError('系统未找到 curl 命令，无法使用 curl 兜底')

    cmd = [
        curl_path,
        '-L',
        '--silent',
        '--show-error',
        '--fail',
        '--http1.1',
        '--compressed',
        '--connect-timeout',
        '10',
        '--max-time',
        '25',
        '--retry',
        '1',
        '--retry-delay',
        '1',
        '-A',
        detail_headers['User-Agent'],
        '-H',
        f"Accept: {detail_headers['Accept']}",
        '-H',
        f"Accept-Language: {detail_headers['Accept-Language']}",
        url,
    ]
    if insecure:
        cmd.insert(1, '--insecure')

    result = subprocess.run(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        timeout=30
    )
    if result.returncode != 0:
        error = result.stderr.decode('utf-8', errors='ignore').strip()
        raise RuntimeError(error or f'curl 退出码：{result.returncode}')
    if not result.stdout:
        raise RuntimeError('curl 返回空内容')
    return result.stdout


def is_tls_compat_error(error):
    """判断是否是 requests/OpenSSL 与目标站 TLS 不兼容导致的错误。"""
    message = str(error)
    return (
        'BAD_ECPOINT' in message or
        'UNEXPECTED_EOF_WHILE_READING' in message or
        'SSLError' in message
    )


def is_certificate_verify_error(error):
    """判断是否是证书校验问题。"""
    message = str(error)
    return (
        'CERTIFICATE_VERIFY_FAILED' in message or
        'certificate verify failed' in message or
        'certificate has expired' in message or
        'SSL certificate problem' in message
    )


def try_curl_detail_content(url, insecure=False):
    try:
        mode = 'curl --insecure' if insecure else 'curl'
        print(f"详情页改用 {mode} 兜底抓取：{url}")
        html = fetch_detail_html_with_curl(url, insecure=insecure)
        return extract_detail_content(html)
    except Exception as e:
        print(f"详情页 {mode} 兜底失败：{url}，原因：{e}")
        return None


def get_detail_content(url):
    """抓取详情页正文"""
    last_error = None

    for attempt in range(detail_max_retries + 1):
        try:
            html = fetch_detail_html_with_requests(url)
            return extract_detail_content(html)

        except Exception as e:
            last_error = e

            if is_certificate_verify_error(e):
                content = try_curl_detail_content(url, insecure=True)
                if content is not None:
                    return content

                print(f"详情页证书异常且兜底失败，跳过正文：{url}，原因：{last_error}")
                return ''

            if is_tls_compat_error(e):
                content = try_curl_detail_content(url)
                if content is not None:
                    return content

                content = try_curl_detail_content(url, insecure=True)
                if content is not None:
                    return content

                print(f"详情页 TLS 异常且兜底失败，跳过正文：{url}，原因：{last_error}")
                return ''

            if attempt >= detail_max_retries:
                content = try_curl_detail_content(url)
                if content is not None:
                    return content

                print(f"详情页抓取失败：{url}，原因：{last_error}")
                return ''

            wait_seconds = get_backoff_wait_seconds(
                attempt, detail_base_delay, detail_max_delay
            )
            print(
                f"详情页抓取失败，刷新连接后准备重试：{url}，"
                f"第 {attempt + 1}/{detail_max_retries} 次重试，"
                f"等待 {wait_seconds:.1f} 秒，原因：{e}"
            )
            time.sleep(wait_seconds)


def remove_html_tags(text):
    """去掉搜索接口返回标题中的 <em> 高亮标签。"""
    if not text:
        return ''
    return clean_text(BeautifulSoup(text, 'html.parser').get_text(' ', strip=True))


def search_policy_page(keyword, page):
    """调用广东省人民政府站内检索“政策文件”接口。"""
    data = {
        'keywords': keyword,
        'sort': 'time',
        'site_id': '2',
        'range': search_range,
        'position': 'all',
        'page': page,
        'recommand': '0',
        'service_area': '1',
        'gdbsDivision': '440000',
    }

    for attempt in range(search_max_retries + 1):
        try:
            with requests.Session() as session:
                resp = session.post(
                    search_api_url,
                    headers=headers,
                    data=data,
                    timeout=(10, 30)
                )
            resp.raise_for_status()

            payload = resp.json()
            if payload.get('errcode') != 0:
                raise RuntimeError(payload.get('errmessage', '搜索接口返回异常'))

            break

        except Exception as e:
            if attempt >= search_max_retries:
                raise

            wait_seconds = get_backoff_wait_seconds(
                attempt, search_base_delay, search_max_delay
            )
            print(
                f"搜索页请求失败，准备重试：关键词={keyword}，页码={page}，"
                f"第 {attempt + 1}/{search_max_retries} 次重试，"
                f"等待 {wait_seconds:.1f} 秒，原因：{e}"
            )
            time.sleep(wait_seconds)

    results = []
    for item in payload.get('data', {}).get('list', []):
        title = remove_html_tags(item.get('title', ''))
        url = item.get('post_url') or item.get('url') or ''
        date = item.get('pub_time') or ''

        if not title or not url:
            continue

        results.append({
            'title': title,
            'url': urljoin('https://www.gd.gov.cn/', url),
            'date': date
        })

    return results


# =========================
# 3. 主程序：按关键词搜索
# =========================

for kw in keywords:
    print(f"\n========== 正在搜索关键词：{kw} ==========")

    reached_old_results = False
    consecutive_detail_failures = 0
    page = 1

    while True:
        print(f"正在请求搜索接口：{search_api_url}，页码={page}")

        try:
            list_items = search_policy_page(kw, page)

            if not list_items:
                print(f"关键词“{kw}”第 {page} 页没有找到结果，停止翻页")
                break

            for item in list_items:
                title = item['title']
                url = item['url']
                date = item['date']

                # 如果列表页能提取到日期，则先做年份筛选
                year = extract_year(date)
                if year is not None:
                    if year < start_year:
                        reached_old_results = True
                        continue
                    if year > end_year:
                        continue

                if is_excluded_title(title) or not title_mentions_keyword(title, kw):
                    continue

                if url in visited_urls:
                    continue
                visited_urls.add(url)

                print(f"发现政策：{title}")

                content = get_detail_content(url)
                if content:
                    consecutive_detail_failures = 0
                else:
                    consecutive_detail_failures += 1
                    if consecutive_detail_failures >= detail_failure_cooldown_threshold:
                        wait_seconds = random.uniform(*detail_failure_cooldown_range)
                        print(
                            f"连续 {consecutive_detail_failures} 条详情页正文为空，"
                            f"暂停 {wait_seconds:.1f} 秒后继续"
                        )
                        time.sleep(wait_seconds)
                        consecutive_detail_failures = 0

                # 如果列表页没日期，尝试从正文中提取日期
                if not date:
                    date_match = re.search(r'20\d{2}[-年./]\d{1,2}[-月./]\d{1,2}', content)
                    date = date_match.group(0) if date_match else ''

                year = extract_year(date)

                # 再次筛选 2017—2022
                if year is not None:
                    if year < start_year:
                        reached_old_results = True
                        continue
                    if year > end_year:
                        continue

                city = judge_city(title, content)

                all_policies.append({
                    'keyword': kw,
                    'city': city,
                    'title': title,
                    'date': date,
                    'year': year,
                    'url': url,
                    'content': content
                })

                print(f"已保存：{city} | {date} | {title}")

                sleep_random(detail_item_delay_range)

            sleep_random(search_page_delay_range)

            if reached_old_results:
                print(f"关键词“{kw}”已搜索到 {start_year - 1} 年及更早结果，停止翻页")
                break

            page += 1

        except Exception as e:
            print(f"搜索页多次重试后仍失败：关键词={kw}，页码={page}，原因：{e}")
            break


# =========================
# 4. 保存结果
# =========================

df = pd.DataFrame(all_policies)

# 去重
if not df.empty:
    df = df.drop_duplicates(subset=['title', 'url'])

# 调整列顺序；后续 LLM 指标构建脚本依赖这些基础字段。
columns = ['keyword', 'city', 'title', 'date', 'year', 'url', 'content']
df = df.reindex(columns=columns)

output_file = 'gd21_digital_policies_2017_2022.csv'  # 爬虫原始输出文件名

df.to_csv(output_file, index=False, encoding='utf-8-sig')

print("\n========== 爬取完成 ==========")
print(f"共爬取 {len(df)} 条政策文本")
print(f"文件已保存为：{output_file}")
