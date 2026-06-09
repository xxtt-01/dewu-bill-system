import os
import logging
import pyodbc  # 替换 pymysql 为 pyodbc
import hashlib
import time
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
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

# 修改基础路径配置为D盘绝对路径
BASE_DIR = 'D:\\dw_dzd'
LOG_DIR = os.path.join(BASE_DIR, 'dwd_bill_logs')
RESULT_DIR = os.path.join(BASE_DIR, 'dwd_bill_results')
DOWNLOAD_DIR = os.path.join(BASE_DIR, 'dwd_bill_downloads')
EXTRACT_DIR = os.path.join(BASE_DIR, 'dzd_xintiqu')

# 修改为 SQL Server 连接参数（密码优先从环境变量 DB_PASSWORD 读取）
DB_CONFIG = {
    'server': '123.57.247.228',
    'database': 'YKYC',
    'username': 'sa',
    'password': os.environ.get('DB_PASSWORD', 'yike@2025'),
    'port': 1433,
    'timeout': 300  # 命令超时 300 秒
}

# API配置
API_URL = "https://openapi.dewu.com/dop/api/v1/bill/period_list"
GENERATE_API_URL = "https://openapi.dewu.com/dop/api/v1/bill/generate"
DOWNLOAD_API_URL = "https://openapi.dewu.com/dop/api/v1/bill/export"
OSS_DOWNLOAD_EXPIRY = 3600

class AppCredential:
    """应用凭证数据类"""

    def __init__(self, cred_id: str, app_key: str, app_secret: str):
        self.cred_id = cred_id
        self.app_key = app_key
        self.app_secret = app_secret

class TextHandler(logging.Handler):
    """将日志记录转发到Tkinter文本控件"""

    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
        self.formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

    def emit(self, record):
        """异步更新GUI"""
        msg = self.formatter.format(record)
        self.text_widget.after(0, self._append_message, msg)

    def _append_message(self, msg):
        self.text_widget.config(state=tk.NORMAL)
        self.text_widget.insert(tk.END, msg + "\n")
        self.text_widget.yview(tk.END)
        self.text_widget.config(state=tk.DISABLED)

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

    logging.info(f"API 请求参数: {all_params}")

    try:
        response = requests.get(API_URL, params=all_params, timeout=30)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"API请求失败 [{credential.cred_id}]: {str(e)}")
        raise
    except OSError as e:
        logging.error(f"API网络连接错误 [{credential.cred_id}]: [{type(e).__name__}] {str(e)}", exc_info=True)
        raise

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
    # 移除ON DUPLICATE KEY UPDATE部分（SQL Server不支持）

    inserted_ids = []
    bill_nos = []
    new_records = [record for record in records if record[0] not in existing_bill_nos]

    try:
        with DBConnection() as cursor:
            # 修改查询语句
            cursor.execute(
                "SELECT ISNULL(MAX(id), 0) FROM dw_dwd_bill_records"
            )
            max_id_before = cursor.fetchone()[0]

            # 批量插入需要逐条处理（SQL Server不支持executemany的INSERT）
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
        self.results: Dict[str, str] = {}
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
                self.results[bill_no] = download_url
                logging.info(f"成功获取下载链接 [{self.credential.cred_id}]: {download_url}")

            except Exception as e:
                self.results[bill_no] = f"错误: {str(e)}"
                logging.error(f"处理账单失败 [{self.credential.cred_id}]: {str(e)}")

            self.completed_bills += 1
            if self.progress_callback:
                self.progress_callback(self.completed_bills, self.total_bills)

def download_files(download_results: Dict[str, str], shop_name: str, progress_callback=None):
    """下载文件到指定目录"""
    shop_dir = os.path.join(DOWNLOAD_DIR, shop_name)
    os.makedirs(shop_dir, exist_ok=True)

    total_files = len(download_results)
    completed_files = 0

    for bill_no, url in download_results.items():
        if url.startswith("错误"):
            logging.error(f"跳过下载 [{shop_name}][{bill_no}]: {url}")
            completed_files += 1
            if progress_callback:
                progress_callback(completed_files, total_files)
            continue

        try:
            logging.info(f"开始下载 [{shop_name}][{bill_no}]")

            response = requests.get(url, stream=True, timeout=30)
            response.raise_for_status()

            filename = f"{bill_no}.xlsx"
            filepath = os.path.join(shop_dir, filename)

            with open(filepath, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)

            logging.info(f"✅ 下载完成 [{shop_name}][{bill_no}] 到 {filepath}")
            completed_files += 1

        except Exception as e:
            logging.error(f"❌ 下载失败 [{shop_name}][{bill_no}]: {str(e)}")
            completed_files += 1

        if progress_callback:
            progress_callback(completed_files, total_files)

    logging.info(f"🎉 所有文件下载完成！文件保存路径: {shop_dir}")

def generate_result_file(inserted: List[Tuple[int, str]],
                         download_results: Dict[str, str],
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
        f"成功下载数: {sum(1 for v in download_results.values() if not v.startswith('错误'))}",
        f"跳过的账单数: {len(skipped)}",
        "\n数据库记录详情:"
    ]

    if inserted:
        content.extend([f"ID: {rid} \t Bill No: {bno}" for rid, bno in inserted])
    else:
        content.append("无新增数据库记录")

    content.append("\n下载结果详情:")
    for bno, url in download_results.items():
        content.append(f"Bill No: {bno}")
        content.append(f"状态: {'成功' if not url.startswith('错误') else '失败'}")
        content.append(f"链接: {url}")
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
        warnings.filterwarnings('ignore', category=UserWarning, module=r'openpyxl.*')

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

                            xls = pd.ExcelFile(file_path)
                            imported = False

                            # 从账单总览 sheet 提取 bill_period
                            bill_period = extract_bill_period_from_file(file_path)

                            if '销售订单' in xls.sheet_names:
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

                            if '退货退款订单' in xls.sheet_names:
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

                            if '账单总览' in xls.sheet_names:
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

                            if imported:
                                with DBConnection() as cursor:
                                    record_import_new(cursor, bill_no, shop_name)
                                    logging.info(f"记录账单 {bill_no} 到新表")
                                    update_log(f"记录账单 {bill_no} 到新表")

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

def main_gui():
    """创建初始GUI界面"""
    root = tk.Tk()
    root.title("账单处理控制系统")
    
    # 确保D盘目录存在（使用绝对路径）
    os.makedirs(LOG_DIR, exist_ok=True)
    os.makedirs(RESULT_DIR, exist_ok=True)
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    os.makedirs(EXTRACT_DIR, exist_ok=True)

    frame = ttk.Frame(root, padding="20")
    frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    log_text = tk.Text(frame, wrap=tk.WORD, height=20, width=80)
    log_text.grid(row=0, column=0, columnspan=2, padx=10, pady=10)
    log_text.config(state=tk.DISABLED)

    text_handler = TextHandler(log_text)
    log_file = setup_logging(text_handler)

    def update_log(message):
        """更新日志到 Text 控件"""
        log_text.config(state=tk.NORMAL)
        log_text.insert(tk.END, message + "\n")
        log_text.yview(tk.END)
        log_text.config(state=tk.DISABLED)

    def run_in_thread(func):
        """在单独线程中运行耗时任务"""

        def wrapper():
            try:
                func()
            except Exception as e:
                logging.error(f"线程任务异常: {str(e)}", exc_info=True)

        thread = threading.Thread(target=wrapper)
        thread.daemon = True
        thread.start()

    start_button = ttk.Button(
        frame,
        text="下载账单",
        command=lambda: run_in_thread(lambda: run_processing_with_logging(root, update_log))
    )
    start_button.grid(row=1, column=0, padx=10, pady=10)

    import_button = ttk.Button(
        frame,
        text="账单处理",
        command=lambda: run_in_thread(lambda: import_bills_with_logging(root, update_log))
    )
    import_button.grid(row=1, column=1, padx=10, pady=10)

    import_db_button = ttk.Button(
        frame,
        text="账单入库",
        command=lambda: run_in_thread(lambda: process_import_with_logging(root, update_log, text_handler))
    )
    import_db_button.grid(row=2, column=0, padx=10, pady=10)

    # 新增测试数据库连接按钮
    test_db_button = ttk.Button(
        frame,
        text="测试数据库连接",
        command=lambda: run_in_thread(lambda: test_db_connection_gui(update_log))
    )
    test_db_button.grid(row=2, column=1, padx=10, pady=10)

    # 新增倒计时功能
    countdown_var = tk.StringVar(value="10")
    countdown_label = ttk.Label(frame, textvariable=countdown_var, font=("Arial", 24))
    countdown_label.grid(row=3, column=0, padx=10, pady=10)

    pause_button = ttk.Button(frame, text="暂停", command=lambda: setattr(auto_run, 'paused', True))
    pause_button.grid(row=3, column=1, padx=10, pady=10)

    # 自动运行控制
    class AutoRun:
        def __init__(self):
            self.paused = False
            self.running = False
    auto_run = AutoRun()

    def countdown_and_auto_run():
        if auto_run.paused or auto_run.running:
            return

        auto_run.running = True
        countdown = 10
        countdown_var.set(str(countdown))

        for i in range(countdown, 0, -1):
            if auto_run.paused:
                auto_run.running = False
                return
            countdown_var.set(str(i))
            time.sleep(1)

        countdown_var.set("0")
        auto_run.paused = False
        run_auto_sequence()

    def sleep_cancellable(seconds):
        """可中断的等待，每 1 秒检查一次暂停状态"""
        for _ in range(seconds):
            if auto_run.paused:
                return False
            time.sleep(1)
        return True

    def run_auto_sequence():
        while not auto_run.paused:
            try:
                logging.info("=== 开始自动运行序列 ===")
                update_log("=== 开始自动运行序列 ===")

                # 运行 "下载账单" 按钮的逻辑
                run_processing_with_logging(root, update_log)
                if auto_run.paused:
                    break
                update_log("下载账单流程完成，等待15秒...")
                if not sleep_cancellable(15):
                    break

                # 运行 "账单处理" 按钮的逻辑
                import_bills_with_logging(root, update_log)
                if auto_run.paused:
                    break
                update_log("账单处理流程完成，等待15秒...")
                if not sleep_cancellable(15):
                    break

                # 运行 "账单入库" 按钮的逻辑
                process_import_with_logging(root, update_log, text_handler)
                if auto_run.paused:
                    break
                update_log("账单入库流程完成，等待21600秒后重新开始循环...")
                if not sleep_cancellable(21600):
                    break

            except Exception as e:
                logging.error(f"自动运行序列异常: {str(e)}", exc_info=True)
                update_log(f"自动运行序列异常: {str(e)}，60秒后重试...")
                if not sleep_cancellable(60):
                    break

        auto_run.running = False
        logging.info("自动运行已停止")

    # 启动倒计时
    threading.Thread(target=countdown_and_auto_run, daemon=True).start()

    # 窗口关闭时优雅退出
    def on_closing():
        if messagebox.askokcancel("退出", "确定要退出程序吗？\n正在运行的任务将被中断。"):
            auto_run.paused = True
            auto_run.running = False
            root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()

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
                        "download_success_count": sum(1 for v in download_results.values() if not v.startswith("错误")),
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

        SHEETS_TO_KEEP = ['账单总览', '销售订单', '退货退款订单']

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
                            data = data.iloc[1:]
                            if sheet_name == '账单总览':
                                # 只保留前3行: [表头行, 子表头行, 数据行]，去掉结算渠道
                                data = data.iloc[:3].reset_index(drop=True)
                                # 合并两行表头为扁平表头（与销售订单的拼接逻辑一致）
                                flat_header = data.iloc[:2].fillna('').astype(str).agg(''.join, axis=0)
                                data = pd.concat([
                                    data.iloc[:2],
                                    pd.DataFrame([flat_header]),
                                    data.iloc[2:]
                                ]).reset_index(drop=True)
                                data = data.iloc[2:]
                            else:
                                if len(data) >= 3:
                                    summary_row = data.iloc[:3].fillna('').astype(str).agg(''.join, axis=0)
                                    data = pd.concat([
                                        data.iloc[:3],
                                        pd.DataFrame([summary_row]),
                                        data.iloc[3:]
                                    ]).reset_index(drop=True)

                                data = data.iloc[3:]

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
                raise RuntimeError(f"文件处理失败: {src_path}") from e

        logging.info(f"=== 流程结束，共处理 {len(all_files)} 个文件 ===")
        update_log(f"=== 流程结束，共处理 {len(all_files)} 个文件 ===")


    except Exception as e:
        logging.error(f"未预期错误: {str(e)}", exc_info=True)
        update_log(f"未预期错误: {str(e)}")


def check_if_imported(cursor, bill_no: str) -> bool:
    """检查账单是否已经导入"""
    # 修改占位符从%s为?
    query = "SELECT COUNT(*) FROM dw_dzd_bill_import_records WHERE bill_no = ?"
    cursor.execute(query, (bill_no,))
    count = cursor.fetchone()[0]
    return count > 0

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
        if len(overview) < 2:
            return ''
        # 用列名取值代替硬编码索引，避免得物调整列顺序
        overview.columns = overview.iloc[0]
        data = overview.iloc[1:]
        val = data.iloc[0].get('账单起止时间账单起止时间', '')
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
        "结算信息结算渠道结算渠道": "settlement_channel"
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
            if pd.isna(value) or str(value).strip() in ('', 'NaN', 'nan', 'NAN', 'None', 'none', 'NONE'):
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
        "订单基础信息退货完成时间退货完成时间": "return_create_time",
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
        "退货信息收款账期收款账期": "payment_period"
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
            if pd.isna(value) or str(value).strip() in ('', 'NaN', 'nan', 'NAN', 'None', 'none', 'NONE'):
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
        if pd.isna(value) or str(value).strip() in ('', 'NaN', 'nan', 'NAN', 'None', 'none', 'NONE'):
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


def record_import(cursor, bill_no: str, shop_name: str):
    """记录账单导入信息"""
    insert_sql = """
    INSERT INTO dw_dzd_bill_import_records (
        bill_no, shop_name, import_time
    ) VALUES (?, ?, ?)
    """
    import_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute(insert_sql, (bill_no, shop_name, import_time))

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
            update_log("✅ 数据库连接成功！")
            messagebox.showinfo("成功", "数据库连接测试成功！")
        else:
            update_log("❌ 数据库连接失败！")
            messagebox.showerror("失败", "数据库连接测试失败，请检查配置！")
    except Exception as e:
        update_log(f"❌ 测试异常：{str(e)}")
        messagebox.showerror("异常", f"测试过程中发生错误：{str(e)}")

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

if __name__ == "__main__":
    main_gui()