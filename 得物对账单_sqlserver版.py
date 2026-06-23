import os
import logging
import pyodbc  # 替换 pymysql 为 pyodbc
import hashlib
import time
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget,
    QVBoxLayout, QHBoxLayout, QPlainTextEdit, QPushButton, QLabel, QMessageBox)
from PyQt6.QtCore import pyqtSignal, Qt, QTimer
from PyQt6.QtGui import QFont
from urllib.parse import quote_plus
import requests
from datetime import datetime, timedelta
from tenacity import retry, stop_after_attempt, wait_exponential, RetryError
from typing import Dict, List, Tuple, Optional
import socket
import sys
import threading
import pandas as pd
import openpyxl
from openpyxl.workbook.views import BookView
from dotenv import load_dotenv

# 加载 .env 文件（敏感配置优先从环境变量读取）
load_dotenv()

# 修改基础路径配置为D盘绝对路径
BASE_DIR = 'D:\\dw_dzd'
LOG_DIR = os.path.join(BASE_DIR, 'dwd_bill_logs')
RESULT_DIR = os.path.join(BASE_DIR, 'dwd_bill_results')
DOWNLOAD_DIR = os.path.join(BASE_DIR, 'dwd_bill_downloads')
EXTRACT_DIR = os.path.join(BASE_DIR, 'dzd_tiqu')

# 数据库连接参数（全部从环境变量读取，不设生产环境默认值）
DB_CONFIG = {
    'server': os.environ.get('DB_SERVER'),
    'database': os.environ.get('DB_DATABASE'),
    'username': os.environ.get('DB_USERNAME'),
    'password': os.environ.get('DB_PASSWORD'),
    'port': int(os.environ.get('DB_PORT', '1433')),
    'timeout': 300  # 命令超时 300 秒
}

# 所有数据库配置项必须有值（port 除外，有默认值）
_missing_db = [k for k, v in DB_CONFIG.items() if not v and k != 'port']
if _missing_db:
    msg = f"数据库配置缺失：{', '.join(_missing_db)}，请在 .env 文件中设置"
    print(f"错误：{msg}")
    sys.exit(1)

# API配置
API_URL = "https://openapi.dewu.com/dop/api/v1/bill/period_list"
GENERATE_API_URL = "https://openapi.dewu.com/dop/api/v1/bill/generate"
DOWNLOAD_API_URL = "https://openapi.dewu.com/dop/api/v1/bill/export"
OSS_DOWNLOAD_EXPIRY = 3600

# 空值集合（Excel 中的多种空值表示）
EMPTY_VALUES = {'', 'NaN', 'nan', 'NAN', 'None', 'none', 'NONE'}

def is_empty(value) -> bool:
    """判断 Excel 单元格值是否为空"""
    return pd.isna(value) or str(value).strip() in EMPTY_VALUES


def detect_headers(df: pd.DataFrame) -> tuple:
    """自动检测 Excel 表头结构

    Returns:
        (note_rows, header_rows, first_data_row)
        - note_rows: 说明行索引列表
        - header_rows: 列名行索引列表（需要拼接为扁平表头的行）
        - first_data_row: 第一条数据行的索引
    """
    note_rows = []
    header_rows = []
    max_header_checks = min(len(df), 12)

    for i in range(max_header_checks):
        row = df.iloc[i]
        non_null = sum(1 for v in row if pd.notna(v))
        total_cols = len(row)
        first_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''

        # 说明行：首列以"说明"开头
        if first_val.startswith('说明'):
            note_rows.append(i)
            continue

        # 空白行：全空
        if non_null == 0:
            continue

        # 表头行：所有列都有值，不含浮点数（数据行的典型特征）
        # 排除已识别的说明行后，连续的全非空行都是表头
        has_float = any(isinstance(v, float) for v in row if pd.notna(v))
        if non_null >= total_cols * 0.8 and not has_float:
            header_rows.append(i)
        else:
            # 遇到第一条数据行，停止检测
            break

    first_data = header_rows[-1] + 1 if header_rows else (note_rows[-1] + 1 if note_rows else 1)

    return note_rows, header_rows, first_data


def merge_header_names(df: pd.DataFrame, header_rows: list) -> list:
    """将多行表头拼接为扁平列名

    对于 N 行表头，按列拼接每行的字符串值。
    示例：3行表头 ["订单基础信息","订单号","订单号"] → "订单基础信息订单号订单号"
    """
    if not header_rows:
        return []
    if len(header_rows) == 1:
        return [str(v).strip() for v in df.iloc[header_rows[0]].tolist()]

    # 多行拼接：每列逐行合并
    merged = []
    for col_idx in range(len(df.columns)):
        parts = []
        for row_idx in header_rows:
            v = df.iloc[row_idx, col_idx]
            parts.append(str(v).strip() if pd.notna(v) else '')
        merged.append(''.join(parts))
    return merged


class AppCredential:
    """应用凭证数据类"""

    def __init__(self, cred_id: str, app_key: str, app_secret: str):
        self.cred_id = cred_id
        self.app_key = app_key
        self.app_secret = app_secret

class QtLogHandler(logging.Handler):
    """将日志记录转发到 PyQt6 控件（通过信号线程安全）"""

    def __init__(self, signal):
        super().__init__()
        self.signal = signal
        self.formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

    def emit(self, record):
        msg = self.formatter.format(record)
        self.signal.emit(msg)

def setup_logging(text_handler) -> str:
    """初始化日志系统并绑定到文本控件"""
    log_file = os.path.join(LOG_DIR, f"dwd_bill_etl_{datetime.now().strftime('%Y%m%d%H%M%S')}.log")

    os.makedirs(os.path.dirname(log_file), exist_ok=True)

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    # 移除所有现有的处理器
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)

    # 添加文件处理器
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger.addHandler(file_handler)

    # 添加文本控件处理器
    logger.addHandler(text_handler)

    logging.info("日志系统初始化完成")  # 新增日志记录
    return log_file

class DBConnection:
    """数据库连接管理器（修改为 SQL Server）"""

    def __enter__(self):
        # 使用兼容性更强的驱动名称 + 增加超时设置
        conn_str = (
            f"DRIVER={{SQL Server}};"
            f"SERVER={DB_CONFIG['server']},{DB_CONFIG['port']};"
            f"DATABASE={DB_CONFIG['database']};"
            f"UID={DB_CONFIG['username']};"
            f"PWD={DB_CONFIG['password']};"
            f"Connection Timeout=60;"
            f"Login Timeout=60;"
        )
        try:
            self.conn = pyodbc.connect(conn_str)
            self.conn.timeout = DB_CONFIG.get('timeout', 300)
            return self.conn.cursor()
        except pyodbc.Error as e:
            logging.error(f"数据库连接失败：{str(e)}")
            safe_conn_str = conn_str.replace(DB_CONFIG['password'], '****')
            logging.error(f"连接字符串：{safe_conn_str}")
            raise

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.conn:
            if exc_type is None:
                self.conn.commit()
            else:
                self.conn.rollback()
            self.conn.close()

def test_db_connection() -> bool:
    """测试数据库连接"""
    try:
        logging.info("开始测试数据库连接...")
        with DBConnection() as cursor:
            cursor.execute("SELECT @@VERSION")
            version = cursor.fetchone()[0]
            logging.info(f"数据库连接成功！版本信息：{version[:100]}...")
            return True
    except pyodbc.Error as e:
        logging.error(f"数据库连接失败：{str(e)}")
        return False

def fetch_app_credentials() -> List[AppCredential]:
    """从数据库获取所有应用凭证"""
    try:
        with DBConnection() as cursor:
            # 修改SQL语句占位符为?
            cursor.execute(
                "SELECT ID, App_Key, App_Secret FROM dewu_app_credentials"
            )
            credentials = [
                AppCredential(row[0], row[1], row[2])
                for row in cursor.fetchall()
            ]
            logging.info(f"成功获取 {len(credentials)} 条应用凭证")
            return credentials
    except pyodbc.Error as e:  # 修改异常类型
        logging.error(f"获取应用凭证失败: {str(e)}")
        return []

def generate_sign(params: dict, app_secret: str) -> str:
    """生成请求签名"""
    filtered_params = {k: v for k, v in params.items() if v not in [None, ""]}
    sorted_params = sorted(filtered_params.items())
    sign_str = '&'.join([f"{k}={quote_plus(str(v))}" for k, v in sorted_params])
    return hashlib.md5((sign_str + app_secret).encode()).hexdigest().upper()

@retry(stop=stop_after_attempt(3), wait=wait_exponential(min=2, max=10))
def fetch_api_data(params: dict, credential: AppCredential) -> dict:
    """获取账单列表数据"""
    base_params = {
        "app_key": credential.app_key,
        "timestamp": str(int(time.time() * 1000)),
        "sign_method": "md5"
    }
    all_params = {**base_params, **params}
    all_params["sign"] = generate_sign(all_params, credential.app_secret)

    # 日志脱敏：隐藏 app_key 中间部分
    safe_params = {**all_params}
    if len(safe_params.get('app_key', '')) > 8:
        safe_params['app_key'] = safe_params['app_key'][:4] + '****' + safe_params['app_key'][-4:]
    logging.info(f"API 请求参数: {safe_params}")

    # 多页获取：循环拉取所有页面的账单列表
    all_items = []
    page = params.get("page_no", 1)
    page_size = params.get("page_size", 30)
    total_count = 0
    max_pages = 100

    while True:
        current_params = {**params}
        current_params["page_no"] = page
        all_params = {**base_params, **current_params}
        all_params["sign"] = generate_sign(all_params, credential.app_secret)

        try:
            response = requests.get(API_URL, params=all_params, timeout=30)
            response.raise_for_status()
            result = response.json()
        except requests.exceptions.RequestException as e:
            logging.error(f"API请求失败 [{credential.cred_id}]: {str(e)}")
            raise
        except OSError as e:
            logging.error(f"API网络连接错误 [{credential.cred_id}]: [{type(e).__name__}] {str(e)}", exc_info=True)
            raise

        data = result.get("data", {})
        page_items = data.get("list", [])
        all_items.extend(page_items)

        if total_count == 0:
            total_count = data.get("totalCount", data.get("total_count", data.get("total", 0)))

        # 判断是否还有更多页
        if not page_items or len(all_items) >= total_count:
            break
        page += 1
        if page > max_pages:
            logging.warning(f"分页超限(>{max_pages})，强制停止")
            break

    result["data"]["list"] = all_items
    return result

def parse_date(date_str: str) -> datetime:
    """智能日期解析"""
    date_formats = [
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
        "%Y/%m/%d %H:%M:%S",
        "%Y%m%d"
    ]

    for fmt in date_formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    raise ValueError(f"无法解析的日期格式: {date_str}")

def get_default_dates() -> dict:
    """获取默认日期（近30天）"""
    today = datetime.now()
    start_date = today - timedelta(days=30)
    return {
        "bill_start_date": start_date.strftime("%Y-%m-%d"),
        "bill_end_date": today.strftime("%Y-%m-%d"),
        "bill_no": None,
        "page_no": 1,
        "page_size": 30
    }

def process_api_data(data: dict, shop_name: str) -> List[tuple]:
    """处理API数据"""
    processed = []
    for item in data.get("list", []):
        try:
            record = (
                item["bill_no"],
                parse_date(item["bill_start_time"]),
                parse_date(item["bill_end_time"]),
                float(item.get("settle_amount", 0)) / 100,
                float(item.get("order_receive_amount", 0)) / 100,
                float(item.get("platform_service_fee", 0)) / 100,
                float(item.get("seller_subsidies_amount", 0)) / 100,
                float(item.get("refund_amount", 0)) / 100,
                item["status"],
                parse_date(item["update_time"]),
                shop_name
            )
            processed.append(record)
        except (KeyError, ValueError) as e:
            logging.warning(f"跳过无效记录: {str(e)}")
    return processed

def check_existing_bill_nos(cursor, shop_name: str) -> set:
    """检查已存在的 bill_no"""
    # 修改占位符为?
    query = """
    SELECT bill_no FROM dw_dwd_bill_records WHERE name = ?
    """
    cursor.execute(query, (shop_name,))
    existing_bill_nos = {row[0] for row in cursor.fetchall()}
    return existing_bill_nos

def save_records(records: List[tuple], existing_bill_nos: set) -> Tuple[List[Tuple[int, str]], List[str]]:
    """保存记录到数据库"""
    insert_sql = """
    INSERT INTO dw_dwd_bill_records (
        bill_no, bill_start_time, bill_end_time, settle_amount,
        order_receive_amount, platform_service_fee,
        seller_subsidies_amount, refund_amount, status, update_time, name
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """

    try:
        with DBConnection() as cursor:
            # 先检查当前最大ID（用于判断后续是否有新增）
            cursor.execute(
                "SELECT ISNULL(MAX(id), 0) FROM dw_dwd_bill_records"
            )
            max_id_before = cursor.fetchone()[0]

            # 逐条 INSERT 以便通过 SCOPE_IDENTITY() 获取自增 ID
            inserted_ids = []
            bill_nos = []
            new_records = [record for record in records if record[0] not in existing_bill_nos]
            
            for record in new_records:
                cursor.execute(insert_sql, record)
                cursor.execute("SELECT SCOPE_IDENTITY()")
                new_id = cursor.fetchone()[0]
                inserted_ids.append((new_id, record[0]))
                bill_nos.append(record[0])

            logging.info(f"成功插入/更新 {len(inserted_ids)} 条记录")
            return inserted_ids, bill_nos
    except pyodbc.Error as e:
        logging.error(f"数据库操作失败: {str(e)}")
        return [], []

class BillProcessor:
    """处理账单生成和下载"""

    def __init__(self, bill_nos: List[str], credential: AppCredential, progress_callback=None):
        self.bill_nos = bill_nos
        self.credential = credential
        self.results: Dict[str, dict] = {}
        self.progress_callback = progress_callback
        self.total_bills = len(bill_nos)
        self.completed_bills = 0

    @retry(stop=stop_after_attempt(3), wait=wait_exponential(min=2, max=10))
    def _generate_bill(self, bill_no: str) -> str:
        """调用生成API"""
        params = {
            "app_key": self.credential.app_key,
            "timestamp": str(int(time.time() * 1000)),
            "bill_no": bill_no
        }
        params["sign"] = generate_sign(params, self.credential.app_secret)

        try:
            response = requests.get(GENERATE_API_URL, params=params, timeout=30)
            response.raise_for_status()
            data = response.json()

            if data["code"] != 200:
                raise ValueError(f"API响应错误: {data.get('msg')}")
            return data["data"]
        except Exception as e:
            logging.error(f"生成账单失败 [{self.credential.cred_id}]: {str(e)}")
            raise

    @retry(stop=stop_after_attempt(3), wait=wait_exponential(min=2, max=10))
    def _get_download_url(self, file_key: str) -> str:
        """获取下载链接"""
        params = {
            "app_key": self.credential.app_key,
            "timestamp": str(int(time.time() * 1000)),
            "key": file_key,
            "expiry": OSS_DOWNLOAD_EXPIRY
        }
        params["sign"] = generate_sign(params, self.credential.app_secret)

        try:
            response = requests.get(DOWNLOAD_API_URL, params=params, timeout=30)
            response.raise_for_status()
            data = response.json()

            if data["code"] != 200:
                raise ValueError(f"API响应错误: {data.get('msg')}")
            return data["data"]
        except Exception as e:
            logging.error(f"获取下载链接失败 [{self.credential.cred_id}]: {str(e)}")
            raise

    def process_all(self):
        """处理所有账单"""
        for bill_no in self.bill_nos:
            try:
                file_key = self._generate_bill(bill_no)
                download_url = self._get_download_url(file_key)
                self.results[bill_no] = {'success': True, 'url': download_url}
                logging.info(f"成功获取下载链接 [{self.credential.cred_id}]: {download_url}")

            except Exception as e:
                self.results[bill_no] = {'success': False, 'error': str(e)}
                logging.error(f"处理账单失败 [{self.credential.cred_id}]: {str(e)}")

            self.completed_bills += 1
            if self.progress_callback:
                self.progress_callback(self.completed_bills, self.total_bills)

def download_files(download_results: Dict[str, dict], shop_name: str, progress_callback=None):
    """下载文件到指定目录"""
    shop_dir = os.path.join(DOWNLOAD_DIR, shop_name)
    os.makedirs(shop_dir, exist_ok=True)

    total_files = len(download_results)
    completed_files = 0

    for bill_no, result in download_results.items():
        if not result.get('success', False):
            logging.warning(f"跳过失败记录 [{shop_name}][{bill_no}]: {result.get('error', '未知错误')}")
            completed_files += 1
            if progress_callback:
                progress_callback(completed_files, total_files)
            continue

        try:
            url = result['url']
            logging.info(f"开始下载 [{shop_name}][{bill_no}]")

            response = requests.get(url, stream=True, timeout=60)
            response.raise_for_status()

            filename = f"{bill_no}.xlsx"
            filepath = os.path.join(shop_dir, filename)
            tmppath = filepath + '.tmp'

            with open(tmppath, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)

            os.replace(tmppath, filepath)
            logging.info(f"下载完成 [{shop_name}][{bill_no}] 到 {filepath}")
            completed_files += 1

        except Exception as e:
            logging.error(f"下载失败 [{shop_name}][{bill_no}]: {str(e)}")
            completed_files += 1

        if progress_callback:
            progress_callback(completed_files, total_files)

    logging.info(f"所有文件下载完成！文件保存路径: {shop_dir}")

def generate_result_file(inserted: List[Tuple[int, str]],
                         download_results: Dict[str, dict],
                         skipped: Dict[str, str],
                         shop_name: str,
                         error: Exception = None) -> str:
    """生成结果文件"""
    os.makedirs(RESULT_DIR, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    result_file = os.path.join(RESULT_DIR, f"result_{shop_name}_{timestamp}.txt")

    content = [
        f"店铺名称: {shop_name}",
        f"执行时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"状态: {'成功' if not error else '失败'}",
        f"数据库记录数: {len(inserted)}",
        f"成功下载数: {sum(1 for v in download_results.values() if v.get('success', False))}",
        f"跳过的账单数: {len(skipped)}",
        "\n数据库记录详情:"
    ]

    if inserted:
        content.extend([f"ID: {rid} \t Bill No: {bno}" for rid, bno in inserted])
    else:
        content.append("无新增数据库记录")

    content.append("\n下载结果详情:")
    for bno, result in download_results.items():
        content.append(f"Bill No: {bno}")
        content.append(f"状态: {'成功' if result.get('success', False) else result.get('error', '未知')}")
        content.append(f"链接: {result.get('url', '')}")
        content.append("-" * 50)

    content.append("\n跳过的账单详情:")
    for bno, reason in skipped.items():
        content.append(f"Bill No: {bno}")
        content.append(f"原因: {reason}")
        content.append("-" * 50)

    if error:
        content.append(f"\n错误详情:\n{str(error)}")

    try:
        with open(result_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(content))
        return result_file
    except Exception as e:
        logging.error(f"结果文件保存失败: {str(e)}")
        return ""

def diagnose_network(host: str) -> bool:
    """诊断网络连接"""
    try:
        logging.info(f"尝试解析域名: {host}")
        resolved_ip = socket.gethostbyname(host)
        logging.info(f"成功解析域名 {host} 到 IP 地址 {resolved_ip}")
        return True
    except socket.gaierror as e:
        logging.error(f"域名解析失败: {host}, 错误: {str(e)}")
        return False

def process_import(root, update_log, text_handler=None):
    """运行账单入库流程"""
    try:
        log_file = setup_logging(text_handler)
        logging.info("=== 账单入库流程启动 ===")
        update_log("账单入库流程启动...")

        import warnings
        warnings.filterwarnings('ignore', category=UserWarning, module=r'openpyxl\.styles\.stylesheet')

        for shop_folder in os.listdir(EXTRACT_DIR):
            shop_path = os.path.join(EXTRACT_DIR, shop_folder)
            if os.path.isdir(shop_path):
                for file in os.listdir(shop_path):
                    if file.endswith('.xlsx'):
                        file_path = os.path.join(shop_path, file)
                        logging.info(f"开始处理文件: {file_path}")
                        update_log(f"正在处理: {file_path}")

                        try:
                            filename = os.path.basename(file)
                            bill_no = filename.split('_')[0]
                            shop_name = shop_folder

                            with DBConnection() as cursor:
                                if check_if_imported_new(cursor, bill_no):
                                    logging.info(f"账单 {bill_no} 已导入新表，跳过")
                                    update_log(f"账单 {bill_no} 已导入新表，跳过")
                                    continue

                            imported = False
                            with pd.ExcelFile(file_path) as xls:
                                sheet_names = set(xls.sheet_names)

                            # 从账单总览 sheet 提取 bill_period
                            bill_period = extract_bill_period_from_file(file_path)

                            if '销售订单' in sheet_names:
                                # 带重试的销售订单导入
                                retry_count = 0
                                max_retries = 3
                                while retry_count < max_retries:
                                    try:
                                        import_sales_orders_from_file(file_path, shop_name, bill_no, bill_period)
                                        imported = True
                                        break
                                    except pyodbc.Error as e:
                                        if "08S01" in str(e) and retry_count < max_retries - 1:
                                            retry_count += 1
                                            logging.warning(f"检测到连接断开 ({retry_count}/{max_retries})，尝试重新连接...")
                                            update_log(f"检测到连接断开，5 秒后重试 ({retry_count}/{max_retries})...")
                                            time.sleep(5)
                                        else:
                                            raise

                            if '退货退款订单' in sheet_names:
                                # 带重试的退货订单导入
                                retry_count = 0
                                max_retries = 3
                                while retry_count < max_retries:
                                    try:
                                        import_refund_orders_from_file(file_path, shop_name, bill_no, bill_period)
                                        imported = True
                                        break
                                    except pyodbc.Error as e:
                                        if "08S01" in str(e) and retry_count < max_retries - 1:
                                            retry_count += 1
                                            logging.warning(f"检测到连接断开 ({retry_count}/{max_retries})，尝试重新连接...")
                                            update_log(f"检测到连接断开，5 秒后重试 ({retry_count}/{max_retries})...")
                                            time.sleep(5)
                                        else:
                                            raise

                            if '账单总览' in sheet_names:
                                # 带重试的账单总览导入
                                retry_count = 0
                                max_retries = 3
                                while retry_count < max_retries:
                                    try:
                                        import_bill_overview_from_file(file_path, shop_name)
                                        imported = True
                                        break
                                    except pyodbc.Error as e:
                                        if "08S01" in str(e) and retry_count < max_retries - 1:
                                            retry_count += 1
                                            logging.warning(f"检测到连接断开 ({retry_count}/{max_retries})，尝试重新连接...")
                                            update_log(f"检测到连接断开，5 秒后重试 ({retry_count}/{max_retries})...")
                                            time.sleep(5)
                                        else:
                                            raise

                            if '本期结算其他项费用' in sheet_names:
                                retry_count = 0
                                max_retries = 3
                                while retry_count < max_retries:
                                    try:
                                        import_other_fee_from_file(file_path, shop_name, bill_no, bill_period)
                                        imported = True
                                        break
                                    except pyodbc.Error as e:
                                        if "08S01" in str(e) and retry_count < max_retries - 1:
                                            retry_count += 1
                                            logging.warning(f"检测到连接断开 ({retry_count}/{max_retries})，尝试重新连接...")
                                            update_log(f"检测到连接断开，5 秒后重试 ({retry_count}/{max_retries})...")
                                            time.sleep(5)
                                        else:
                                            raise

                            if '扣减其他费用明细' in sheet_names:
                                retry_count = 0
                                max_retries = 3
                                while retry_count < max_retries:
                                    try:
                                        import_deduction_detail_from_file(file_path, shop_name, bill_no, bill_period)
                                        imported = True
                                        break
                                    except pyodbc.Error as e:
                                        if "08S01" in str(e) and retry_count < max_retries - 1:
                                            retry_count += 1
                                            logging.warning(f"检测到连接断开 ({retry_count}/{max_retries})，尝试重新连接...")
                                            update_log(f"检测到连接断开，5 秒后重试 ({retry_count}/{max_retries})...")
                                            time.sleep(5)
                                        else:
                                            raise

                            if '本期货损买进订单' in sheet_names:
                                retry_count = 0
                                max_retries = 3
                                while retry_count < max_retries:
                                    try:
                                        import_cargo_damage_from_file(file_path, shop_name, bill_no, bill_period)
                                        imported = True
                                        break
                                    except pyodbc.Error as e:
                                        if "08S01" in str(e) and retry_count < max_retries - 1:
                                            retry_count += 1
                                            logging.warning(f"检测到连接断开 ({retry_count}/{max_retries})，尝试重新连接...")
                                            update_log(f"检测到连接断开，5 秒后重试 ({retry_count}/{max_retries})...")
                                            time.sleep(5)
                                        else:
                                            raise

                            if imported:
                                try:
                                    with DBConnection() as cursor:
                                        record_import_new(cursor, bill_no, shop_name)
                                        logging.info(f"记录账单 {bill_no} 到新表")
                                        update_log(f"记录账单 {bill_no} 到新表")
                                except Exception as e:
                                    logging.error(f"记录导入状态失败 {bill_no}: {e}")
                                    update_log(f"⚠ 数据已入库但状态记录失败: {bill_no}，如补数请先清理 dw_dwd_bill_records_copy1")

                            logging.info(f"文件处理完成: {file_path}")
                            update_log(f"文件处理完成: {file_path}")

                        except Exception as e:
                            logging.error(f"处理失败 {file_path}: {str(e)}")
                            update_log(f"失败: {file_path} - {str(e)}")

        logging.info("=== 账单入库流程结束 ===")
        update_log("账单入库流程结束")
    except Exception as e:
        logging.error(f"未预期错误: {str(e)}", exc_info=True)
        update_log(f"处理失败: {str(e)}")

def import_sales_orders_from_file(file_path, shop_name, bill_no='', bill_period=''):
    """从文件导入销售订单"""
    data = pd.read_excel(file_path, sheet_name='销售订单')
    data = data.fillna('').replace(['NaN', 'nan', 'NAN', 'None', 'none', 'NONE'], '')

    with DBConnection() as cursor:
        import_sales_orders(cursor, data, shop_name, bill_no, bill_period)

def import_refund_orders_from_file(file_path, shop_name, bill_no='', bill_period=''):
    """从文件导入退货退款订单"""
    data = pd.read_excel(file_path, sheet_name='退货退款订单')
    data = data.fillna('').replace(['NaN', 'nan', 'NAN', 'None', 'none', 'NONE'], '')

    with DBConnection() as cursor:
        import_refund_orders(cursor, data, shop_name, bill_no, bill_period)

def process_import_with_logging(root, update_log, text_handler=None):
    """运行账单入库流程并更新日志"""
    try:
        logging.info("=== 账单入库流程启动 ===")
        update_log("账单入库流程启动...")

        process_import(root, update_log, text_handler)

    except Exception as e:
        logging.error(f"未预期错误: {str(e)}", exc_info=True)
        update_log(f"处理失败: {str(e)}")
    finally:
        update_log("账单入库流程结束")

def run_processing(root, update_log):
    exit_code = 0
    results = []

    try:
        logging.info("=== 账单处理流程启动 ===")
        update_log("账单处理流程启动...")
        credentials = fetch_app_credentials()
        if not credentials:
            raise ValueError("没有可用的应用凭证")

        input_params = get_default_dates()
        logging.info(f"默认输入参数: {input_params}")
        update_log(f"默认输入参数: {input_params}")

        for credential in credentials:
            logging.info(f"=== 开始处理店铺: {credential.cred_id} ===")
            update_log(f"开始处理店铺: {credential.cred_id}")
            inserted_records = []
            download_results = {}
            skipped_records = {}
            execution_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            if not diagnose_network("openapi.dewu.com"):
                logging.error(f"网络诊断失败，跳过处理店铺: {credential.cred_id}")
                update_log(f"网络诊断失败，跳过处理店铺: {credential.cred_id}")
                continue

            try:
                api_response = fetch_api_data(input_params, credential)
                if api_response.get("code") != 200:
                    raise ValueError(f"API错误: {api_response.get('message')}")

                processed_data = process_api_data(api_response.get("data", {}), credential.cred_id)
                bill_nos = []
                if not processed_data:
                    logging.warning(f"没有可处理的有效数据 [{credential.cred_id}]")
                    update_log(f"没有可处理的有效数据 [{credential.cred_id}]")
                else:
                    with DBConnection() as cursor:
                        existing_bill_nos = check_existing_bill_nos(cursor, credential.cred_id)
                        new_records = [record for record in processed_data if record[0] not in existing_bill_nos]
                        skipped_bill_nos = [record[0] for record in processed_data if record[0] in existing_bill_nos]

                        inserted_records, bill_nos = save_records(new_records, existing_bill_nos)
                        if not inserted_records:
                            logging.warning(f"未成功插入任何记录 [{credential.cred_id}]")
                            update_log(f"未成功插入任何记录 [{credential.cred_id}]")
                        else:
                            logging.info(f"成功插入/更新 {len(inserted_records)} 条记录")
                            update_log(f"成功插入/更新 {len(inserted_records)} 条记录")
                        for bill_no in skipped_bill_nos:
                            skipped_records[bill_no] = "账单已存在"

                if bill_nos:
                    processor = BillProcessor(bill_nos, credential,
                                              progress_callback=lambda x, y: update_log(f"进度: {x}/{y}"))
                    processor.process_all()
                    download_results = processor.results

                    update_log("等待 120 秒后开始下载...")
                    time.sleep(120)

                    download_files(download_results, credential.cred_id,
                                   progress_callback=lambda x, y: update_log(f"下载进度: {x}/{y}"))

                    results.append({
                        "execution_time": execution_time,
                        "shop_name": credential.cred_id,
                        "status": "成功",
                        "insert_count": len(inserted_records),
                        "download_success_count": sum(1 for v in download_results.values() if v.get("success", False)),
                        "skipped_count": len(skipped_records),
                        "inserted_records": inserted_records,
                        "download_results": download_results,
                        "skipped_records": skipped_records
                    })
                else:
                    results.append({
                        "execution_time": execution_time,
                        "shop_name": credential.cred_id,
                        "status": "成功",
                        "insert_count": len(inserted_records),
                        "download_success_count": 0,
                        "skipped_count": len(skipped_records),
                        "inserted_records": inserted_records,
                        "download_results": {},
                        "skipped_records": skipped_records
                    })

            except Exception as e:
                logging.error(f"处理失败 [{credential.cred_id}]: {str(e)}")
                update_log(f"处理失败 [{credential.cred_id}]: {str(e)}")
                results.append({
                    "execution_time": execution_time,
                    "shop_name": credential.cred_id,
                    "status": "失败",
                    "insert_count": 0,
                    "download_success_count": 0,
                    "skipped_count": 0,
                    "inserted_records": [],
                    "download_results": {},
                    "skipped_records": {}
                })

        logging.info(f"=== 流程结束 [退出码: {exit_code}] ===")
        update_log("账单处理流程结束")

    except Exception as e:
        exit_code = 1
        logging.error(f"未预期错误: {str(e)}", exc_info=True)
        update_log(f"未预期错误: {str(e)}")

def import_bills(root, update_log):
    """导入并处理本地账单文件"""
    try:
        logging.info("=== 本地账单导入流程启动 ===")
        update_log("本地账单导入流程启动...")

        SHEETS_TO_KEEP = ['账单总览', '销售订单', '退货退款订单',
                           '本期结算其他项费用', '扣减其他费用明细', '本期货损买进订单']

        all_files = []
        for shop_folder in os.listdir(DOWNLOAD_DIR):
            shop_path = os.path.join(DOWNLOAD_DIR, shop_folder)
            if os.path.isdir(shop_path):
                for file in os.listdir(shop_path):
                    if file.endswith('.xlsx') and not file.endswith('_tiqu.xlsx'):
                        src_path = os.path.join(shop_path, file)
                        dest_dir = os.path.join(EXTRACT_DIR, shop_folder)
                        dest_path = os.path.join(dest_dir, file.replace('.xlsx', '_tiqu.xlsx'))
                        if not os.path.exists(dest_path) and os.path.exists(src_path):
                            all_files.append((src_path, dest_dir, dest_path))

        if not all_files:
            logging.info("没有需要处理的新文件")
            update_log("没有需要处理的新文件")

            return

        total_files = len(all_files)
        for i, (src_path, dest_dir, dest_path) in enumerate(all_files):
            try:
                os.makedirs(dest_dir, exist_ok=True)
                logging.info(f"开始处理文件: {src_path}")
                update_log(f"正在处理: {src_path}")

                with pd.ExcelWriter(dest_path, engine='openpyxl') as writer:
                    with pd.ExcelFile(src_path) as xls:
                        sheets_to_process = {
                            sheet_name: pd.read_excel(xls, sheet_name=sheet_name, header=None)
                            for sheet_name in xls.sheet_names if sheet_name in SHEETS_TO_KEEP
                        }

                        if not sheets_to_process:
                            raise ValueError(f"文件中无有效工作表 {SHEETS_TO_KEEP}")

                        for sheet_name, data in sheets_to_process.items():
                            # 自动检测表头结构，不再依赖硬编码行号
                            note_rows, header_rows, first_data = detect_headers(data)

                            if len(header_rows) >= 2:
                                # 多行表头：拼接为扁平列名
                                flat_headers = merge_header_names(data, header_rows)
                                header_df = pd.DataFrame([flat_headers])
                                data_body = data.iloc[first_data:].reset_index(drop=True)
                                data = pd.concat([header_df, data_body]).reset_index(drop=True)
                            elif len(header_rows) == 1:
                                # 单行表头：跳过说明行后的第1行就是列名
                                data = data.iloc[first_data - 1:].reset_index(drop=True)
                            else:
                                logging.warning(f"未能检测到表头结构: {sheet_name}")
                                continue

                            data.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

                    workbook = writer.book
                    if workbook.worksheets:
                        workbook.worksheets[0].sheet_state = 'visible'
                        workbook.active = 0

                shop_name = os.path.basename(os.path.dirname(src_path))

                logging.info(f"文件处理完成: {src_path} -> {dest_path}")
                update_log(f"文件处理完成: {src_path} -> {dest_path}")
                update_log(f"进度: {i + 1}/{total_files}")

            except Exception as e:
                error_msg = f"处理失败 {src_path}: {str(e)}"
                logging.error(error_msg, exc_info=True)
                update_log(f"失败: {src_path} - {str(e)}")
                continue

        logging.info(f"=== 流程结束，共处理 {len(all_files)} 个文件 ===")
        update_log(f"=== 流程结束，共处理 {len(all_files)} 个文件 ===")


    except Exception as e:
        logging.error(f"未预期错误: {str(e)}", exc_info=True)
        update_log(f"未预期错误: {str(e)}")



def check_if_imported_new(cursor, bill_no: str) -> bool:
    """检查账单是否已经导入到新表"""
    query = "SELECT COUNT(*) FROM dw_dwd_bill_records_copy1 WHERE bill_no = ?"
    cursor.execute(query, (bill_no,))
    count = cursor.fetchone()[0]
    return count > 0


def extract_bill_period_from_file(file_path: str) -> str:
    """从 _tiqu.xlsx 的账单总览 sheet 提取账单起止时间"""
    try:
        overview = pd.read_excel(file_path, sheet_name='账单总览', header=None)
        note_rows, header_rows, first_data = detect_headers(overview)

        if not header_rows or len(overview) <= first_data:
            return ''

        # 用自动检测的扁平列名定位列
        flat_names = merge_header_names(overview, header_rows)
        data_row = overview.iloc[first_data]

        # 优先精确匹配
        target_col = '账单起止时间账单起止时间'
        if target_col in flat_names:
            col_idx = flat_names.index(target_col)
        else:
            # 模糊匹配：找到包含"账单起止时间"的列
            col_idx = -1
            for i, name in enumerate(flat_names):
                if '账单起止时间' in str(name):
                    col_idx = i
                    break

        if col_idx >= 0:
            val = data_row.iloc[col_idx]
            if pd.notna(val):
                return str(val).strip()

    except Exception as e:
        logging.warning(f"提取账单起止时间失败: {e}")
    return ''


def import_sales_orders(cursor, data: pd.DataFrame, shop_name: str, bill_no: str = '', bill_period: str = ''):
    """导入销售订单数据"""
    data = data.fillna('').replace(['NaN', 'nan', 'NAN', 'None', 'none', 'NONE'], '')

    field_mapping = {
        # 订单基础信息
        "订单基础信息订单号订单号": "order_id",
        "订单基础信息订单类型订单类型": "order_type",
        "订单基础信息商品名称商品名称": "product_name",
        "订单基础信息商品货号商品货号": "product_code",
        "订单基础信息数量数量": "quantity",
        "订单基础信息规格规格": "specification",
        "订单基础信息预约单号预约单号": "reservation_number",
        "订单基础信息商品金额商品金额": "product_amount",
        "订单基础信息联合营销费联合营销费": "joint_marketing_fee",
        "订单基础信息限时折扣/定金预售报名价限时折扣/定金预售报名价": "limited_discount_advance_price",
        "订单基础信息商品交易金额商品交易金额": "product_transaction_amount",
        "订单基础信息出价时间出价时间": "ask_time",
        "订单基础信息订单创建时间订单创建时间": "order_creation_time",
        "订单基础信息支付时间支付时间": "payment_time",
        "订单基础信息发货时间发货时间": "delivery_time",

        # 平台服务费信息（一口价）
        "平台服务费信息（一口价）是否参加活动是否参加活动": "is_participating_in_activity_ykj",
        "平台服务费信息（一口价）活动费率活动费率": "activity_fee_rate_ykj",
        "平台服务费信息（一口价）费率活动ID费率活动ID": "activity_fee_id_ykj",
        "平台服务费信息（一口价）适用费率适用费率": "applicable_rate_ykj",
        "平台服务费信息（一口价）费率下限费率下限": "fee_rate_lower_limit_ykj",
        "平台服务费信息（一口价）费率上限费率上限": "fee_rate_upper_limit_ykj",
        "平台服务费信息（一口价）优惠①:费率折扣优惠①:费率折扣": "discount_1_fee_rate_discount_ykj",
        "平台服务费信息（一口价）服务分费率折扣服务分费率折扣": "service_fee_rate_discount_ykj",
        "平台服务费信息（一口价）费率折扣优惠额费率折扣优惠额": "fee_rate_discount_amount_ykj",
        "平台服务费信息（一口价）任务达成折扣任务达成折扣": "task_achievement_discount_ykj",
        "平台服务费信息（一口价）任务达成折扣对应优惠额任务达成折扣对应优惠额": "task_achievement_discount_amount_ykj",
        "平台服务费信息（一口价）其中-服务费返利折扣金额其中-服务费返利折扣金额": "service_fee_rebate_discount_amount_ykj",
        "平台服务费信息（一口价）优惠②.技术服务费券优惠②.技术服务费券": "technical_service_fee_coupon_ykj",
        "平台服务费信息（一口价）优惠③服务费返利减免金额优惠③服务费返利减免金额": "service_fee_rebate_reduction_amount_ykj",
        "平台服务费信息（一口价）合计平台服务费金额(已扣减优惠①②③)合计平台服务费金额(已扣减优惠①②③)": "total_platform_service_fee_ykj",
        "平台服务费信息（一口价）商家返利商家返利": "merchant_rebate_ykj",
        "平台服务费信息（一口价）最终平台基础服务费最终平台基础服务费": "final_fundamental_fee_ykj",
        "平台服务费信息（一口价）其中:基础服务费金额其中:基础服务费金额": "fundamental_service_fee_ykj",
        "平台服务费信息（一口价）其中:履约服务费金额其中:履约服务费金额": "performance_service_fee_ykj",

        # 平台服务费信息技术服务费信息
        "平台服务费信息技术服务费活动信息是否参加活动": "is_participating_in_activity",
        "平台服务费信息技术服务费活动信息活动费率": "activity_fee_rate",
        "平台服务费信息技术服务费活动信息费率活动ID": "activity_fee_id",
        "平台服务费信息技术服务费信息技术服务费费率": "technical_service_fee_rate",
        "平台服务费信息技术服务费信息费率下限": "fee_rate_lower_limit",
        "平台服务费信息技术服务费信息费率上限": "fee_rate_upper_limit",
        "平台服务费信息技术服务费信息优惠①:费率折扣": "discount_1_fee_rate_discount",
        "平台服务费信息技术服务费信息服务分费率折扣": "service_fee_rate_discount",
        "平台服务费信息技术服务费信息费率折扣优惠额": "fee_rate_discount_amount",
        "平台服务费信息技术服务费信息任务达成折扣": "task_achievement_discount",
        "平台服务费信息技术服务费信息任务达成折扣对应优惠额": "task_achievement_discount_amount",
        "平台服务费信息技术服务费信息其中-服务费返利折扣金额": "service_fee_rebate_discount_amount",
        "平台服务费信息技术服务费信息优惠②.技术服务费券": "technical_service_fee_coupon",
        "平台服务费信息技术服务费信息优惠③服务费返利减免金额": "service_fee_rebate_reduction_amount",
        "平台服务费信息技术服务费信息技术服务费(已扣减优惠①②③)": "technical_service_fee_after_discounts",
        "平台服务费信息技术服务费信息商家返利": "merchant_rebate",
        "平台服务费信息技术服务费信息合计技术服务费": "total_technical_service_fee",

        # 平台服务费信息操作服务费信息
        "平台服务费信息操作服务费信息操作服务费": "operation_service_fee",
        "平台服务费信息操作服务费信息包含防尘袋包装费": "includes_dust_bag_packaging_fee",
        "平台服务费信息操作服务费信息包含礼盒费": "includes_gift_box_fee",
        "平台服务费信息操作服务费信息包含礼袋费": "includes_gift_bag_fee",
        "平台服务费信息操作类费用查验费": "inspection_fee",
        "平台服务费信息操作类费用鉴别费": "authentication_fee",
        "平台服务费信息操作类费用包装服务费": "packaging_service_fee",
        "平台服务费信息操作类费用转账手续费": "transfer_fee",
        "平台服务费信息操作类费用品牌服务费": "brand_service_fee",
        "平台服务费信息操作类费用客服托管服务费": "customer_hosting_service_fee",

        # 其他信息
        "平台服务费信息售后无忧服务费售后无忧服务费": "after_sales_service_fee",
        "平台服务费信息卖家退运服务费卖家退运服务费": "seller_return_service_fee",
        "平台服务费信息分销服务费分销服务费金额": "distribution_service_fee",
        "平台服务费信息分销服务费分销规则类型": "distribution_rule_type",
        "平台服务费信息分销服务费分销规则ID": "distribution_rule_id",
        "平台服务费信息合计平台服务费合计平台服务费": "total_platform_service_fee",

        # 结算信息
        "结算信息平台预付款收回金额平台预付款收回金额": "platform_advance_payment_recovery_amount",
        "结算信息以旧换新补贴金额以旧换新补贴金额": "trade_in_subsidy_amount",
        "结算信息卖家补贴金额卖家承担包邮金额": "seller_bears_free_shipping_amount",
        "结算信息卖家补贴金额消费者邮费补贴金额": "consumer_shipping_subsidy_amount",
        "结算信息卖家补贴金额卖家承担优惠券金额": "seller_bears_coupon_amount",
        "结算信息卖家补贴金额卖家承担折扣活动金额": "seller_bears_discount_activity_amount",
        "结算信息卖家补贴金额分期免息卖家承担金额": "installment_free_interest_seller_bears_amount",
        "结算信息售中降价(退款)售中降价(退款)": "defect_price_reduction_tk",
        "结算信息售中降价(退津贴)售中降价(退津贴)": "defect_price_reduction_tjt",
        "结算信息调整金额调整金额": "adjustment_amount",
        "结算信息应结金额应结金额": "payable_amount",
        "结算信息结算状态结算状态": "settlement_status",
        "结算信息结算渠道结算渠道": "settlement_channel",

        # 新增字段（2026-06 得物格式升级）
        "订单基础信息业务时间业务时间": "business_time",
        "订单基础信息交易成功时间交易成功时间": "transaction_success_time",
        "结算信息结算规则类型结算规则类型": "settlement_rule_type",
        "备注信息备注内容备注内容": "remark_content",
    }

    columns = ", ".join(field_mapping.values()) + ", ZH, bill_no, bill_period"
    placeholders = ", ".join(["?"] * (len(field_mapping) + 3))
    insert_sql = f"INSERT INTO dw_dzd_xs ({columns}) VALUES ({placeholders})"

    # 发货时间降级：检测业务时间列是否存在
    business_time_col = '订单基础信息业务时间业务时间'
    has_business_time = business_time_col in data.columns
    delivery_pos = list(field_mapping.values()).index('delivery_time') if 'delivery_time' in field_mapping.values() else -1
    fallback_count = 0
    if has_business_time and delivery_pos >= 0:
        logging.info("检测到业务时间列，发货时间为空时将自动降级使用业务时间")

    records = []
    for _, row in data.iterrows():
        record = []
        for excel_header, db_field in field_mapping.items():
            value = row.get(excel_header, '')
            if is_empty(value):
                record.append(None)
            else:
                record.append(value)

        # 发货时间降级：如果为空且有业务时间列，用业务时间填充
        if delivery_pos >= 0 and record[delivery_pos] in (None, '') and has_business_time:
            bt_val = row.get(business_time_col, '')
            if pd.notna(bt_val) and str(bt_val).strip() not in ('', 'NaN', 'nan', 'NAN', 'None', 'none', 'NONE'):
                record[delivery_pos] = bt_val
                fallback_count += 1

        record.append(shop_name)
        record.append(bill_no)
        record.append(bill_period)
        records.append(tuple(record))

    total_records = len(records)
    if fallback_count > 0:
        logging.info(f"发货时间降级: {fallback_count}/{total_records} 行使用了业务时间")

    # 分批插入，每 500 条提交一次，避免事务过长导致连接超时
    BATCH_SIZE = 500

    for i in range(0, total_records, BATCH_SIZE):
        batch = records[i:i+BATCH_SIZE]
        cursor.executemany(insert_sql, batch)

        # 每批提交后立即 commit，避免事务过长
        cursor.connection.commit()

        # 记录进度
        processed = min(i + BATCH_SIZE, total_records)
        if processed % 1000 == 0 or processed == total_records:
            logging.info(f"已插入 {processed}/{total_records} 条销售订单记录")

def import_refund_orders(cursor, data: pd.DataFrame, shop_name: str, bill_no: str = '', bill_period: str = ''):
    """导入退货退款订单数据"""
    data = data.fillna('').replace(['NaN', 'nan', 'NAN', 'None', 'none', 'NONE'], '')

    field_mapping = {
        "订单基础信息订单号订单号": "order_id",
        "订单基础信息退货订单号退货订单号": "return_order_id",
        "订单基础信息退货创建时间退货创建时间": "return_create_time",
        "订单基础信息退货订单账单起止时间退货订单账单起止时间": "return_order_bill_period",
        "订单基础信息订单类型订单类型": "order_type",
        "订单基础信息商品名称商品名称": "product_name",
        "订单基础信息商品货号商品货号": "product_code",
        "订单基础信息数量数量": "quantity",
        "订单基础信息规格规格": "specification",
        "订单基础信息预约单号预约单号": "reservation_number",
        "订单基础信息商品金额商品金额": "product_amount",
        "订单基础信息联合营销费联合营销费": "joint_marketing_fee",
        "订单基础信息限时折扣/定金预售报名价限时折扣/定金预售报名价": "limited_discount_advance_price",
        "订单基础信息商品交易金额商品交易金额": "product_transaction_amount",
        "订单基础信息出价时间出价时间": "ask_time",
        "订单基础信息订单创建时间订单创建时间": "order_create_time",
        "订单基础信息支付时间支付时间": "payment_time",
        "订单基础信息发货时间发货时间": "delivery_time",
        "平台服务费信息（一口价）是否参加活动是否参加活动": "is_participating_in_activity_ykj",
        "平台服务费信息（一口价）活动费率活动费率": "activity_fee_rate_ykj",
        "平台服务费信息（一口价）费率活动ID费率活动ID": "activity_fee_id_ykj",
        "平台服务费信息（一口价）适用费率适用费率": "applicable_rate_ykj",
        "平台服务费信息（一口价）费率下限费率下限": "fee_rate_lower_limit_ykj",
        "平台服务费信息（一口价）费率上限费率上限": "fee_rate_upper_limit_ykj",
        "平台服务费信息（一口价）优惠①:费率折扣优惠①:费率折扣": "discount_1_fee_rate_discount_ykj",
        "平台服务费信息（一口价）服务分费率折扣服务分费率折扣": "service_fee_rate_discount_ykj",
        "平台服务费信息（一口价）费率折扣优惠额费率折扣优惠额": "fee_rate_discount_amount_ykj",
        "平台服务费信息（一口价）任务达成折扣任务达成折扣": "task_achievement_discount_ykj",
        "平台服务费信息（一口价）任务达成折扣对应优惠额任务达成折扣对应优惠额": "task_achievement_discount_amount_ykj",
        "平台服务费信息（一口价）其中-服务费返利折扣金额其中-服务费返利折扣金额": "service_fee_rebate_discount_amount_ykj",
        "平台服务费信息（一口价）优惠②.技术服务费券优惠②.技术服务费券": "technical_service_fee_coupon_ykj",
        "平台服务费信息（一口价）优惠③服务费返利减免金额优惠③服务费返利减免金额": "service_fee_rebate_reduction_amount_ykj",
        "平台服务费信息（一口价）合计平台服务费金额(已扣减优惠①②③)合计平台服务费金额(已扣减优惠①②③)": "total_platform_service_fee_ykj",
        "平台服务费信息（一口价）商家返利商家返利": "merchant_rebate_ykj",
        "平台服务费信息（一口价）最终平台基础服务费最终平台基础服务费": "final_fundamental_fee_ykj",
        "平台服务费信息（一口价）其中:基础服务费金额其中:基础服务费金额": "fundamental_service_fee_ykj",
        "平台服务费信息（一口价）其中:履约服务费金额其中:履约服务费金额": "performance_service_fee_ykj",
        "平台服务费信息技术服务费活动信息是否参加活动": "is_participating_in_activity",
        "平台服务费信息技术服务费活动信息活动费率": "activity_fee_rate",
        "平台服务费信息技术服务费活动信息费率活动ID": "activity_fee_id",
        "平台服务费信息技术服务费信息技术服务费费率": "tech_service_fee_rate",
        "平台服务费信息技术服务费信息费率下限": "fee_rate_lower_limit",
        "平台服务费信息技术服务费信息费率上限": "fee_rate_upper_limit",
        "平台服务费信息技术服务费信息优惠①:费率折扣": "discount_1_fee_rate_discount",
        "平台服务费信息技术服务费信息服务分费率折扣": "service_fee_rate_discount",
        "平台服务费信息技术服务费信息费率折扣优惠额": "fee_rate_discount_amount",
        "平台服务费信息技术服务费信息任务达成折扣": "task_achievement_discount",
        "平台服务费信息技术服务费信息任务达成折扣对应优惠额": "task_achievement_discount_amount",
        "平台服务费信息技术服务费信息其中-服务费返利折扣金额": "service_fee_rebate_discount_amount",
        "平台服务费信息技术服务费信息优惠②.技术服务费券": "tech_service_fee_coupon",
        "平台服务费信息技术服务费信息优惠③服务费返利减免金额": "service_fee_rebate_reduction_amount",
        "平台服务费信息技术服务费信息技术服务费(已扣减优惠①②③)": "tech_service_fee_after_discounts",
        "平台服务费信息技术服务费信息商家返利": "merchant_rebate",
        "平台服务费信息技术服务费信息合计技术服务费": "total_tech_service_fee",
        "平台服务费信息操作服务费信息操作服务费": "operation_service_fee",
        "平台服务费信息操作服务费信息包含防尘袋包装费": "includes_dust_bag_packaging_fee",
        "平台服务费信息操作服务费信息包含礼盒费": "includes_gift_box_fee",
        "平台服务费信息操作服务费信息包含礼袋费": "includes_gift_bag_fee",
        "平台服务费信息操作类费用查验费": "inspection_fee",
        "平台服务费信息操作类费用鉴别费": "authentication_fee",
        "平台服务费信息操作类费用包装服务费": "packaging_service_fee",
        "平台服务费信息操作类费用转账手续费": "transfer_fee",
        "平台服务费信息操作类费用客服托管服务费": "customer_hosting_service_fee",
        "平台服务费信息其他费用品牌服务费": "brand_service_fee",
        "平台服务费信息合计平台服务费合计平台服务费": "total_platform_service_fee",
        "平台服务费信息售后无忧服务费售后无忧服务费": "after_sales_service_fee",
        "结算信息平台预付款收回金额平台预付款收回金额": "platform_advance_payment_recovery_amount",
        "结算信息以旧换新补贴金额以旧换新补贴金额": "trade_in_subsidy_amount",
        "结算信息卖家补贴金额卖家承担包邮金额": "seller_bears_free_shipping_amount",
        "结算信息卖家补贴金额消费者邮费补贴金额": "consumer_shipping_subsidy_amount",
        "结算信息卖家补贴金额卖家承担优惠券金额": "seller_bears_coupon_amount",
        "结算信息卖家补贴金额卖家承担折扣活动金额": "seller_bears_discount_activity_amount",
        "结算信息卖家补贴金额分期免息卖家承担金额": "installment_free_interest_seller_bears_amount",
        "结算信息售中降价(退款)售中降价(退款)": "defect_price_reduction_tk",
        "结算信息售中降价(退津贴)售中降价(退津贴)": "defect_price_reduction_tjt",
        "结算信息调整金额调整金额": "adjustment_amount",
        "结算信息应结金额应结金额": "payable_amount",
        "结算信息结算状态结算状态": "settlement_status",
        "退货信息收款账期收款账期": "payment_period",

        # 新增字段（2026-06 得物格式升级）
        "备注信息备注内容备注内容": "remark_content",
    }

    columns = ", ".join(field_mapping.values()) + ", ZH, bill_no, bill_period"
    placeholders = ", ".join(["?"] * (len(field_mapping) + 3))
    insert_sql = f"INSERT INTO dw_dzd_thtk ({columns}) VALUES ({placeholders})"

    # 发货时间降级：检测业务时间列是否存在
    business_time_col = '订单基础信息业务时间业务时间'
    has_business_time = business_time_col in data.columns
    delivery_pos = list(field_mapping.values()).index('delivery_time') if 'delivery_time' in field_mapping.values() else -1
    fallback_count = 0
    if has_business_time and delivery_pos >= 0:
        logging.info("检测到业务时间列，发货时间为空时将自动降级使用业务时间")

    records = []
    for _, row in data.iterrows():
        record = []
        for excel_header, db_field in field_mapping.items():
            value = row.get(excel_header, '')
            if is_empty(value):
                record.append(None)
            else:
                record.append(value)

        # 发货时间降级：如果为空且有业务时间列，用业务时间填充
        if delivery_pos >= 0 and record[delivery_pos] in (None, '') and has_business_time:
            bt_val = row.get(business_time_col, '')
            if pd.notna(bt_val) and str(bt_val).strip() not in ('', 'NaN', 'nan', 'NAN', 'None', 'none', 'NONE'):
                record[delivery_pos] = bt_val
                fallback_count += 1

        record.append(shop_name)
        record.append(bill_no)
        record.append(bill_period)
        records.append(tuple(record))

    total_records = len(records)
    if fallback_count > 0:
        logging.info(f"发货时间降级: {fallback_count}/{total_records} 行使用了业务时间")

    # 分批插入，每 500 条提交一次，避免事务过长导致连接超时
    BATCH_SIZE = 500

    for i in range(0, total_records, BATCH_SIZE):
        batch = records[i:i+BATCH_SIZE]
        cursor.executemany(insert_sql, batch)

        # 每批提交后立即 commit，避免事务过长
        cursor.connection.commit()

        # 记录进度
        processed = min(i + BATCH_SIZE, total_records)
        if processed % 1000 == 0 or processed == total_records:
            logging.info(f"已插入 {processed}/{total_records} 条退货退款订单记录")


# ============================================================
# 账单总览 sheet 字段映射 → dw_dzd_bill_overview
# Excel 表头为两行拼接格式（与销售订单一致）：{行2}{行3}
# 例如 "账单编号"+"账单编号" = "账单编号账单编号"
# ============================================================
BILL_OVERVIEW_FIELD_MAPPING = {
    '账单编号账单编号': 'bill_no',
    '公司名称公司名称': 'company_name',
    '结算周期结算周期': 'settlement_cycle',
    '账单起止时间账单起止时间': 'bill_period',
    '对账单更新时间对账单更新时间': 'update_time_desc',
    '本期商品交易金额本期商品交易金额': 'product_transaction_amount',
    '本期交易类平台服务费金额本期交易类平台服务费金额': 'platform_service_fee',
    '分销服务费分销服务费': 'distribution_service_fee',
    '平台预付款收回金额平台预付款收回金额': 'advance_payment_recovery',
    '卖家补贴金额卖家补贴金额': 'seller_subsidy_amount',
    '调整项调整项': 'adjustment_item',
    '卖家退运服务费卖家退运服务费': 'seller_return_shipping_fee',
    '售后无忧售后无忧': 'after_sales_service',
    '以旧换新补贴金额以旧换新补贴金额': 'trade_in_subsidy_amount',
    '本期结算其他项费用其他非交易类应收商品金额': 'other_non_transaction_receivable',
    '本期结算其他项费用其他非交易类平台费用': 'other_non_transaction_platform_fee',
    '本期结算其他项费用平台预付款收回金额': 'other_advance_payment_recovery',
    '本期结算其他项费用其他非交易类应结金额': 'other_non_transaction_settlement',
    '本期结算其他项费用扣减其他费用': 'other_deductions',
    '售中降价(退款)售中降价(退款)': 'price_reduction_refund',
    '售中降价(退津贴)售中降价(退津贴)': 'price_reduction_subsidy',
    '应结总金额应结总金额': 'total_payable_amount',
    '结算状态结算状态': 'settlement_status',
}


def import_bill_overview(cursor, data: pd.DataFrame, shop_name: str):
    """将账单总览数据插入 dw_dzd_bill_overview"""
    data = data.fillna('').replace(['NaN', 'nan', 'NAN', 'None', 'none', 'NONE'], '')
    field_mapping = BILL_OVERVIEW_FIELD_MAPPING

    # 从扁平表头中提取 data row
    # data 结构: Row0=扁平表头, Row1=数据行
    data_row = data.iloc[1:2].copy()
    data_row.columns = data.iloc[0]

    # 拼装记录
    record = []
    for excel_header, db_field in field_mapping.items():
        value = data_row.iloc[0].get(excel_header, '')
        if is_empty(value):
            record.append(None)
        else:
            record.append(value)
    record.append(shop_name)

    # 检查是否已存在（按 bill_no 去重）
    bill_no = record[0]  # bill_no 是映射的第一个字段
    if bill_no:
        cursor.execute("SELECT COUNT(1) FROM dw_dzd_bill_overview WHERE bill_no = ?", (bill_no,))
        if cursor.fetchone()[0] > 0:
            logging.info(f"账单总览 {bill_no} 已存在，跳过")
            return
    else:
        logging.warning("账单总览 bill_no 为空，跳过去重检查，直接插入")

    columns = ", ".join(field_mapping.values()) + ", ZH"
    placeholders = ", ".join(["?"] * (len(field_mapping) + 1))
    insert_sql = f"INSERT INTO dw_dzd_bill_overview ({columns}) VALUES ({placeholders})"

    cursor.execute(insert_sql, tuple(record))
    logging.info(f"账单总览 {bill_no} 入库成功")


def import_bill_overview_from_file(file_path: str, shop_name: str):
    """从 _tiqu.xlsx 的账单总览 sheet 读取并入库"""
    data = pd.read_excel(file_path, sheet_name='账单总览', header=None)
    data = data.fillna('').replace(['NaN', 'nan', 'NAN', 'None', 'none', 'NONE'], '')

    with DBConnection() as cursor:
        import_bill_overview(cursor, data, shop_name)


# ============================================================
# 本期结算其他项费用 field_mapping → dw_dzd_other_fee
# 2行表头（说明行+列名行），使用简单列名
# ============================================================
OTHER_FEE_FIELD_MAPPING = {
    '费用类型': 'fee_type',
    '订单号': 'order_id',
    '原订单号': 'original_order_id',
    '订单类型': 'order_type',
    '发生时间': 'occur_time',
    '商品金额': 'product_amount',
    '平台预付款收回金额': 'advance_payment_recovery',
    '操作服务费': 'operation_service_fee',
    '认证直发服务费': 'certification_direct_fee',
    '转账手续费': 'transfer_fee',
    '技术服务费': 'technical_service_fee',
    '客服托管服务费': 'customer_service_fee',
    '卖家补贴买家(商品)': 'seller_subsidy_buyer_product',
    '售中降价(退款)': 'price_reduction_refund',
    '售中降价(退津贴)': 'price_reduction_subsidy',
    '出口推广服务费': 'export_promotion_fee',
    '卖家补贴买家(分期手续费)': 'seller_subsidy_installment_fee',
    '其他赔付项': 'other_compensation',
    '结算金额': 'settlement_amount',
    '结算汇率': 'settlement_rate',
}


def import_other_fee(cursor, data: pd.DataFrame, shop_name: str, bill_no: str = '', bill_period: str = ''):
    """导入本期结算其他项费用"""
    data = data.fillna('').replace(['NaN', 'nan', 'NAN', 'None', 'none', 'NONE'], '')
    field_mapping = OTHER_FEE_FIELD_MAPPING

    columns = ", ".join(field_mapping.values()) + ", ZH, bill_no, bill_period"
    placeholders = ", ".join(["?"] * (len(field_mapping) + 3))
    insert_sql = f"INSERT INTO dw_dzd_other_fee ({columns}) VALUES ({placeholders})"

    records = []
    for _, row in data.iterrows():
        record = []
        for excel_header, db_field in field_mapping.items():
            value = row.get(excel_header, '')
            if is_empty(value):
                record.append(None)
            else:
                record.append(value)

        record.append(shop_name)
        record.append(bill_no)
        record.append(bill_period)
        records.append(tuple(record))

    total_records = len(records)
    if total_records == 0:
        return

    BATCH_SIZE = 500
    for i in range(0, total_records, BATCH_SIZE):
        batch = records[i:i + BATCH_SIZE]
        cursor.executemany(insert_sql, batch)
        cursor.connection.commit()

        processed = min(i + BATCH_SIZE, total_records)
        if processed % 1000 == 0 or processed == total_records:
            logging.info(f"已插入 {processed}/{total_records} 条其他项费用记录")


def import_other_fee_from_file(file_path: str, shop_name: str, bill_no: str = '', bill_period: str = ''):
    """从 _tiqu.xlsx 的 本期结算其他项费用 sheet 读取并入库"""
    data = pd.read_excel(file_path, sheet_name='本期结算其他项费用')
    data = data.fillna('').replace(['NaN', 'nan', 'NAN', 'None', 'none', 'NONE'], '')

    with DBConnection() as cursor:
        import_other_fee(cursor, data, shop_name, bill_no, bill_period)


# ============================================================
# 扣减其他费用明细 field_mapping → dw_dzd_deduction_detail
# ============================================================
DEDUCTION_DETAIL_FIELD_MAPPING = {
    '费用类型': 'fee_type',
    '偿还总金额': 'total_repayment',
    '明细费用项': 'detail_item',
    '偿还金额': 'repayment_amount',
    '币种': 'currency',
}


def import_deduction_detail(cursor, data: pd.DataFrame, shop_name: str, bill_no: str = '', bill_period: str = ''):
    """导入扣减其他费用明细"""
    data = data.fillna('').replace(['NaN', 'nan', 'NAN', 'None', 'none', 'NONE'], '')
    field_mapping = DEDUCTION_DETAIL_FIELD_MAPPING

    columns = ", ".join(field_mapping.values()) + ", ZH, bill_no, bill_period"
    placeholders = ", ".join(["?"] * (len(field_mapping) + 3))
    insert_sql = f"INSERT INTO dw_dzd_deduction_detail ({columns}) VALUES ({placeholders})"

    records = []
    for _, row in data.iterrows():
        record = []
        for excel_header, db_field in field_mapping.items():
            value = row.get(excel_header, '')
            if is_empty(value):
                record.append(None)
            else:
                record.append(value)

        record.append(shop_name)
        record.append(bill_no)
        record.append(bill_period)
        records.append(tuple(record))

    total_records = len(records)
    if total_records == 0:
        return

    BATCH_SIZE = 500
    for i in range(0, total_records, BATCH_SIZE):
        batch = records[i:i + BATCH_SIZE]
        cursor.executemany(insert_sql, batch)
        cursor.connection.commit()

        processed = min(i + BATCH_SIZE, total_records)
        if processed % 1000 == 0 or processed == total_records:
            logging.info(f"已插入 {processed}/{total_records} 条扣减明细记录")


def import_deduction_detail_from_file(file_path: str, shop_name: str, bill_no: str = '', bill_period: str = ''):
    """从 _tiqu.xlsx 的 扣减其他费用明细 sheet 读取并入库"""
    data = pd.read_excel(file_path, sheet_name='扣减其他费用明细')
    data = data.fillna('').replace(['NaN', 'nan', 'NAN', 'None', 'none', 'NONE'], '')

    with DBConnection() as cursor:
        import_deduction_detail(cursor, data, shop_name, bill_no, bill_period)


# ============================================================
# 本期货损买进订单 field_mapping → dw_dzd_cargo_damage
# ============================================================
CARGO_DAMAGE_FIELD_MAPPING = {
    '账单编号': 'bill_no_from_sheet',
    '买进订单号': 'purchase_order_id',
    '原订单号': 'original_order_id',
    '订单类型': 'order_type',
    '发生时间': 'occur_time',
    '商品金额（元）': 'product_amount',
    '平台预付款收回金额（元）': 'advance_payment_recovery',
    '操作服务费（元）': 'operation_service_fee',
    '认证直发服务费（元）': 'certification_direct_fee',
    '技术服务费（元）': 'technical_service_fee',
    '平台基础服务费（元）': 'platform_base_service_fee',
    '其中:基础服务费金额（元）': 'base_service_fee_amount',
    '其中:履约服务费金额（元）': 'performance_service_fee_amount',
    '转账手续费（元）': 'transfer_fee',
    '售后无忧服务费（元）': 'after_sales_service_fee',
    '消费者邮费补贴（元）': 'consumer_shipping_subsidy',
    '客服托管服务费（元）': 'customer_service_fee',
    '包装服务费（元）': 'packaging_service_fee',
    '结算金额（元）': 'settlement_amount',
}


def import_cargo_damage(cursor, data: pd.DataFrame, shop_name: str, bill_no: str = '', bill_period: str = ''):
    """导入本期货损买进订单"""
    data = data.fillna('').replace(['NaN', 'nan', 'NAN', 'None', 'none', 'NONE'], '')
    field_mapping = CARGO_DAMAGE_FIELD_MAPPING

    columns = ", ".join(field_mapping.values()) + ", ZH, bill_no, bill_period"
    placeholders = ", ".join(["?"] * (len(field_mapping) + 3))
    insert_sql = f"INSERT INTO dw_dzd_cargo_damage ({columns}) VALUES ({placeholders})"

    records = []
    for _, row in data.iterrows():
        record = []
        for excel_header, db_field in field_mapping.items():
            value = row.get(excel_header, '')
            if is_empty(value):
                record.append(None)
            else:
                record.append(value)

        record.append(shop_name)
        record.append(bill_no)
        record.append(bill_period)
        records.append(tuple(record))

    total_records = len(records)
    if total_records == 0:
        return

    BATCH_SIZE = 500
    for i in range(0, total_records, BATCH_SIZE):
        batch = records[i:i + BATCH_SIZE]
        cursor.executemany(insert_sql, batch)
        cursor.connection.commit()

        processed = min(i + BATCH_SIZE, total_records)
        if processed % 1000 == 0 or processed == total_records:
            logging.info(f"已插入 {processed}/{total_records} 条货损买进记录")


def import_cargo_damage_from_file(file_path: str, shop_name: str, bill_no: str = '', bill_period: str = ''):
    """从 _tiqu.xlsx 的 本期货损买进订单 sheet 读取并入库"""
    data = pd.read_excel(file_path, sheet_name='本期货损买进订单')
    data = data.fillna('').replace(['NaN', 'nan', 'NAN', 'None', 'none', 'NONE'], '')

    with DBConnection() as cursor:
        import_cargo_damage(cursor, data, shop_name, bill_no, bill_period)



def record_import_new(cursor, bill_no: str, shop_name: str):
    """记录账单到新表"""
    insert_sql = """
    INSERT INTO dw_dwd_bill_records_copy1 (bill_no, name)
    VALUES (?, ?)
    """
    cursor.execute(insert_sql, (bill_no, shop_name))

def test_db_connection_gui(update_log):
    """GUI 测试数据库连接"""
    try:
        update_log("正在测试数据库连接...")
        if test_db_connection():
            update_log("数据库连接成功！")
            QMessageBox.information(None, "成功", "数据库连接测试成功！")
        else:
            update_log("数据库连接失败！")
            QMessageBox.critical(None, "失败", "数据库连接测试失败！")
    except Exception as e:
        update_log(f"测试异常: {str(e)}")
        QMessageBox.critical(None, "错误", f"测试异常: {str(e)}")

def import_bills_with_logging(root, update_log):
    """运行账单导入流程并更新日志"""
    try:
        logging.info("=== 本地账单导入流程启动 ===")
        update_log("本地账单导入流程启动...")

        import_bills(root, update_log)

    except Exception as e:
        logging.error(f"未预期错误: {str(e)}", exc_info=True)
        update_log(f"处理失败: {str(e)}")
    finally:
        update_log("本地账单导入流程结束")

def run_processing_with_logging(root, update_log):
    """运行账单处理流程并更新日志"""
    try:
        logging.info("=== 账单处理流程启动 ===")
        update_log("账单处理流程启动...")

        run_processing(root, update_log)

    except Exception as e:
        logging.error(f"未预期错误: {str(e)}", exc_info=True)
        update_log(f"处理失败: {str(e)}")
    finally:
        update_log("账单处理流程结束")


class MainWindow(QMainWindow):
    """得物对账单控制系统 - 主窗口（深色玻璃拟态风格）"""
    log_signal = pyqtSignal(str)
    countdown_signal = pyqtSignal(int)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("得物对账单控制系统")
        self.setMinimumSize(800, 500)

        os.makedirs(LOG_DIR, exist_ok=True)
        os.makedirs(RESULT_DIR, exist_ok=True)
        os.makedirs(DOWNLOAD_DIR, exist_ok=True)
        os.makedirs(EXTRACT_DIR, exist_ok=True)

        # 颜色方案
        self.C_BG = "#303438"
        self.C_FG = "#eef5fb"
        self.C_FG_SEC = "#a3adbb"
        self.C_ACCENT = "#b7ead4"
        self.C_CARD = "#1c1c24"
        self.C_BORDER = "#42474d"
        self.C_SUCCESS = "#4ec77f"
        self.C_DANGER = "#f06a7b"

        # 信号连接（线程安全）
        self.log_signal.connect(self._append_log)
        self.countdown_signal.connect(lambda v: self.countdown_label.setText(str(v)))

        # 日志处理器（传递给业务函数）
        self.log_handler = QtLogHandler(self.log_signal)

        self._setup_ui()
        self._setup_auto_run()

        # 启动倒计时
        QTimer.singleShot(0, self._start_countdown)

    def _setup_ui(self):
        """构建界面"""
        central = QWidget()
        central.setStyleSheet(f"background-color: {self.C_BG};")
        self.setCentralWidget(central)

        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(14, 14, 14, 14)
        root_layout.setSpacing(8)

        # ── 日志卡片 ──
        log_card = QWidget()
        log_card.setObjectName("logCard")
        log_card.setStyleSheet(f"""
            QWidget#logCard {{
                background-color: {self.C_CARD};
                border: 1px solid {self.C_BORDER};
                border-radius: 8px;
            }}
        """)

        card_layout = QVBoxLayout(log_card)
        card_layout.setContentsMargins(0, 0, 0, 0)

        self.log_text = QPlainTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 11))
        self.log_text.setStyleSheet(f"""
            QPlainTextEdit {{
                background-color: {self.C_CARD};
                color: {self.C_FG};
                border: none;
                padding: 12px 10px;
                selection-background-color: {self.C_ACCENT}40;
            }}
            QScrollBar:vertical {{
                width: 6px;
                background: transparent;
            }}
            QScrollBar::handle:vertical {{
                background: #808a9b45;
                border-radius: 3px;
                min-height: 30px;
            }}
            QScrollBar::handle:vertical:hover {{
                background: #808a9b80;
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0px;
            }}
        """)
        card_layout.addWidget(self.log_text)
        root_layout.addWidget(log_card, 1)

        # ── 按钮行 ──
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(4)

        btn_style = f"""
            QPushButton {{
                background-color: {self.C_CARD};
                color: {self.C_FG};
                font-family: Consolas;
                font-size: 12px;
                padding: 8px 14px;
                border: 1px solid {self.C_BORDER};
                border-radius: 6px;
            }}
            QPushButton:hover {{
                background-color: {self.C_ACCENT};
                color: #1c1c24;
                border-color: {self.C_ACCENT};
            }}
            QPushButton:pressed {{
                background-color: #8dd4b0;
            }}
        """

        buttons = [
            ("下载账单", lambda: self._run_thread(
                lambda: run_processing_with_logging(self, self._update_log))),
            ("账单处理", lambda: self._run_thread(
                lambda: import_bills_with_logging(self, self._update_log))),
            ("账单入库", lambda: self._run_thread(
                lambda: process_import_with_logging(self, self._update_log, self.log_handler))),
            ("测试连接", lambda: self._run_thread(
                lambda: test_db_connection_gui(self._update_log))),
        ]

        for text, callback in buttons:
            btn = QPushButton(text)
            btn.setStyleSheet(btn_style)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.clicked.connect(callback)
            btn_layout.addWidget(btn)

        root_layout.addLayout(btn_layout)

        # ── 状态行 ──
        status_layout = QHBoxLayout()
        status_layout.setSpacing(4)

        self.countdown_label = QLabel("10")
        self.countdown_label.setFont(QFont("Consolas", 28))
        self.countdown_label.setStyleSheet(f"color: {self.C_ACCENT}; background: transparent;")
        status_layout.addWidget(self.countdown_label)

        self.status_label = QLabel("\u25cf \u7a7a\u95f2")
        self.status_label.setFont(QFont("Consolas", 10))
        self.status_label.setStyleSheet(f"color: {self.C_FG_SEC}; background: transparent;")
        status_layout.addWidget(self.status_label)

        status_layout.addStretch(1)

        pause_btn = QPushButton("\u6682\u505c")
        pause_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.C_CARD};
                color: {self.C_DANGER};
                font-family: Consolas;
                font-size: 12px;
                padding: 8px 14px;
                border: 1px solid {self.C_BORDER};
                border-radius: 6px;
            }}
            QPushButton:hover {{
                background-color: {self.C_DANGER};
                color: white;
            }}
        """)
        pause_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        pause_btn.clicked.connect(lambda: setattr(self.auto_run, 'paused', True))
        status_layout.addWidget(pause_btn)

        root_layout.addLayout(status_layout)

    def _setup_auto_run(self):
        """自动运行控制"""
        class AutoRun:
            def __init__(self, window):
                import threading
                self._window = window
                self._lock = threading.Lock()
                self._paused = False
                self._running = False

            @property
            def paused(self):
                with self._lock:
                    return self._paused

            @paused.setter
            def paused(self, value):
                with self._lock:
                    self._paused = value
                if value:
                    self._window.status_label.setText("\u25cf \u5df2\u6682\u505c")
                    self._window.status_label.setStyleSheet(f"color: {self._window.C_DANGER}; background: transparent;")
                else:
                    self._window.status_label.setText("\u25cf \u8fd0\u884c\u4e2d")
                    self._window.status_label.setStyleSheet(f"color: {self._window.C_SUCCESS}; background: transparent;")

            @property
            def running(self):
                with self._lock:
                    return self._running

            @running.setter
            def running(self, value):
                with self._lock:
                    self._running = value
                if not value and not self._paused:
                    self._window.status_label.setText("\u25cf \u7a7a\u95f2")
                    self._window.status_label.setStyleSheet(f"color: {self._window.C_FG_SEC}; background: transparent;")

        self.auto_run = AutoRun(self)

    def _append_log(self, message):
        """主线程中更新日志（由信号触发）"""
        self.log_text.appendPlainText(message)

    def _update_log(self, message):
        """从工作线程调用的日志更新（线程安全）"""
        self.log_signal.emit(message)

    def _run_thread(self, func):
        """在后台线程中运行任务"""
        thread = threading.Thread(target=func, daemon=True)
        thread.start()

    def _start_countdown(self):
        """启动倒计时和自动运行"""
        def run():
            if self.auto_run.paused or self.auto_run.running:
                return
            self.auto_run.running = True
            countdown = 10
            self.countdown_signal.emit(countdown)
            for i in range(countdown, 0, -1):
                if self.auto_run.paused:
                    self.auto_run.running = False
                    return
                self.countdown_signal.emit(i)
                time.sleep(1)
            self.countdown_signal.emit(0)
            self.auto_run.paused = False
            self._run_auto_sequence()
        threading.Thread(target=run, daemon=True).start()

    def _sleep_cancellable(self, seconds):
        for _ in range(seconds):
            if self.auto_run.paused:
                return False
            time.sleep(1)
        return True

    def _run_auto_sequence(self):
        while not self.auto_run.paused:
            try:
                logging.info("=== \u5f00\u59cb\u81ea\u52a8\u8fd0\u884c\u5e8f\u5217 ===")
                self._update_log("=== \u5f00\u59cb\u81ea\u52a8\u8fd0\u884c\u5e8f\u5217 ===")
                run_processing_with_logging(self, self._update_log)
                if self.auto_run.paused: break
                self._update_log("\u4e0b\u8f7d\u8d26\u5355\u6d41\u7a0b\u5b8c\u6210\uff0c\u7b49\u5f8515\u79d2...")
                if not self._sleep_cancellable(15): break
                import_bills_with_logging(self, self._update_log)
                if self.auto_run.paused: break
                self._update_log("\u8d26\u5355\u5904\u7406\u6d41\u7a0b\u5b8c\u6210\uff0c\u7b49\u5f8515\u79d2...")
                if not self._sleep_cancellable(15): break
                process_import_with_logging(self, self._update_log, self.log_handler)
                if self.auto_run.paused: break
                self._update_log("\u8d26\u5355\u5165\u5e93\u6d41\u7a0b\u5b8c\u6210\uff0c\u7b49\u5f8521600\u79d2\u540e\u91cd\u65b0\u5f00\u59cb\u5faa\u73af...")
                if not self._sleep_cancellable(21600): break
            except Exception as e:
                logging.error(f"\u81ea\u52a8\u8fd0\u884c\u5e8f\u5217\u5f02\u5e38: {str(e)}", exc_info=True)
                self._update_log(f"\u81ea\u52a8\u8fd0\u884c\u5e8f\u5217\u5f02\u5e38: {str(e)}"
                                 f"\uff0c60\u79d2\u540e\u91cd\u8bd5...")
                if not self._sleep_cancellable(60): break
        self.auto_run.running = False
        logging.info("\u81ea\u52a8\u8fd0\u884c\u5df2\u505c\u6b62")

    def closeEvent(self, event):
        """窗口关闭事件"""
        reply = QMessageBox.question(self, "\u9000\u51fa",
                                     "\u786e\u5b9a\u8981\u9000\u51fa\u7a0b\u5e8f\u5417\uff1f\n\u6b63\u5728\u8fd0\u884c\u7684\u4efb\u52a1\u5c06\u88ab\u4e2d\u65ad\u3002",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.auto_run.paused = True
            self.auto_run.running = False
            event.accept()
        else:
            event.ignore()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())