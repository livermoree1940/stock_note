import os

# 禁用全局代理：让程序直接连接网络，不走 Clash 
os.environ['HTTP_PROXY'] = ''
os.environ['HTTPS_PROTOCOL'] = ''
os.environ['ALL_PROXY'] = ''
os.environ['no_proxy'] = '*'  # 强制所有请求直连
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import xml.etree.ElementTree as ET
import pandas as pd
import adata
import threading
import time
import datetime
from datetime import timedelta
import pyautogui
import pygetwindow as gw
import json
import os
import numpy as np
from concurrent.futures import ThreadPoolExecutor, as_completed

# 高DPI支持设置
import ctypes
ctypes.windll.shcore.SetProcessDpiAwareness(1)

# 尝试导入Ashare，用于获取历史行情
try:
    from Ashare import *
    USE_AKSHARE = True
    print("已成功导入Ashare")
except ImportError:
    USE_AKSHARE = False
    print("Ashare未找到，将使用adata获取历史数据")

# 配置参数
CONFIG = {
    "ui": {
        "font": ("微软雅黑", 20),
        "colors": {
            "bg": "#000000",
            "text": "#FFFFFF",
            "rise": "#FF0000",
            "fall": "#008000",
            "purple": "#800080",
            "pinned": "#FFD700",
            # 20个不同的颜色分组
            "group1": "#FFFFA8",
            "group2": "#9FFFFF",
            "group3": "#FFAAFF",
            "group4": "#B6FFB6",
            "group5": "#FFEDCB",
            "group6": "#FFCCCC",
            "group7": "#CCCCFF",
            "group8": "#FFD700",
            "group9": "#98FB98",
            "group10": "#87CEFA",
            "group11": "#FFB6C1",
            "group12": "#F0E68C",
            "group13": "#E6E6FA",
            "group14": "#ADD8E6",
            "group15": "#F5DEB3",
            "group16": "#FFA07A",
            "group17": "#20B2AA",
            "group18": "#DDA0DD",
            "group19": "#90EE90",
            "group20": "#FFC0CB",
        },
    },
    "refresh_interval": 20,
    "data_file": "custom_data.json",
    "max_workers": 10,  # 减少线程池大小，降低内存占用 
    # 结果测试  还是10线程最快   再高会竞争
    "batch_size": 100,  # 减小批量处理大小，降低内存占用
    "volume_ratio_days": 10,  # 计算近十日量比
}


class BlockAnalyzer:

    def __init__(self, xml_path, specific_block=None):
        self.xml_path = xml_path
        self.specific_block = specific_block
        self.blocks = self._parse_xml()
        self.prev_data = None
        self.five_min_data = []
        self.custom_data = self._load_custom_data()
        self.ma5_cache = {}
        self.volume_ratio_cache = {}
        self.history_data_cache = {}  # 历史数据缓存
        self.thread_pool = ThreadPoolExecutor(max_workers=CONFIG["max_workers"])
        self.last_ma5_fetch_date = None
        self.last_volume_ratio_fetch_date = None
        self.last_history_data_fetch_date = None

    def _parse_xml(self):
        """解析XML文件获取板块数据"""
        tree = ET.parse(self.xml_path)
        root = tree.getroot()

        block_dict = {}
        for block in root.findall("Block"):
            block_name = block.get("name")
            codes = []
            for security in block.findall("security"):
                market = security.get("market")
                code = security.get("code")
                if market in ["USHA", "USZA"]:
                    codes.append(code)
            block_dict[block_name] = codes
        return block_dict

    def _load_custom_data(self):
        """加载自定义数据"""
        if os.path.exists(CONFIG["data_file"]):
            try:
                with open(CONFIG["data_file"], "r", encoding="utf-8") as f:
                    return json.load(f)
            except:
                return {}
        return {}

    def save_custom_data(self):
        """保存自定义数据"""
        with open(CONFIG["data_file"], "w", encoding="utf-8") as f:
            json.dump(self.custom_data, f, ensure_ascii=False, indent=2)

    def _get_stock_history_data(self, stock_code):
        """获取股票历史数据，支持缓存"""
        today = datetime.datetime.now().date()
        
        # 检查缓存中是否有今天的数据
        if stock_code in self.history_data_cache:
            cached_time, cached_data = self.history_data_cache[stock_code]
            if cached_time.date() == today:
                return cached_data
        
        # 调整股票代码格式
        symbol = stock_code
        if len(stock_code) == 6:
            if stock_code.startswith('6'):
                symbol = f'sh{stock_code}'
            else:
                symbol = f'sz{stock_code}'
        
        try:
            # 使用Ashare的get_price函数获取最近20天的日线数据
            df = get_price(symbol, frequency='1d', count=20)
            
            if df is not None and not df.empty:
                # 缓存数据
                self.history_data_cache[stock_code] = (datetime.datetime.now(), df)
                self.last_history_data_fetch_date = today
                return df
        except Exception as e:
            print(f"获取{stock_code}历史数据失败: {str(e)}")
        
        return None

    def get_ma5_data_batch(self, stock_codes):
        """批量获取5日均线数据，优化内存使用，支持akshare和adata两种方式"""
        today = datetime.datetime.now().date()

        # 如果今天已经获取过数据，直接使用缓存
        if self.last_ma5_fetch_date == today:
            results = {}
            for code in stock_codes:
                if code in self.ma5_cache:
                    cached_time, ma5_value = self.ma5_cache[code]
                    if cached_time.date() == today:
                        results[code] = ma5_value
            return results

        # 否则批量获取数据，减少线程池使用
        results = {}
        
        # 只处理有实时数据的股票
        valid_codes = [code for code in stock_codes if code not in results]
        
        # 分小批次处理，减少内存占用
        batch_size = 10  # 每次处理10个股票
        for i in range(0, len(valid_codes), batch_size):
            batch = valid_codes[i:i+batch_size]
            futures = {}
            
            for code in batch:
                # 如果缓存中有今天的数据，直接使用
                if code in self.ma5_cache:
                    cached_time, ma5_value = self.ma5_cache[code]
                    if cached_time.date() == today:
                        results[code] = ma5_value
                        continue
                
                # 否则提交任务到线程池
                future = self.thread_pool.submit(self._get_single_ma5_data, code)
                futures[future] = code
            
            # 等待当前批次完成
            for future in as_completed(futures):
                code = futures[future]
                try:
                    ma5_value = future.result()
                    if ma5_value is not None:
                        results[code] = ma5_value
                        self.ma5_cache[code] = (datetime.datetime.now(), ma5_value)
                except Exception as e:
                    # 减少打印，只在必要时输出
                    # print(f"获取{code}的5日均线数据失败: {str(e)}")
                    pass

        self.last_ma5_fetch_date = today
        return results

    def _get_single_ma5_data(self, stock_code):
        """获取单个股票的5日均线数据，使用共享的历史数据"""
        try:
            # 首先尝试使用共享的历史数据
            df = self._get_stock_history_data(stock_code)
            
            if df is not None and not df.empty:
                print(f"使用共享历史数据计算{stock_code}的MA5值...")
                
                # 确保收盘价是数值类型
                df['close'] = pd.to_numeric(df['close'], errors='coerce')
                df = df.dropna(subset=['close'])
                
                if len(df) >= 5:
                    # 计算最近5天的平均值
                    recent_5_days = df['close'].tail(5)
                    ma5_value = recent_5_days.mean()
                    result = round(ma5_value, 2)
                    print(f"{stock_code}的MA5值: {result}")
                    return result
                else:
                    print(f"{stock_code}数据不足5条，无法计算MA5")
            
            # 如果共享数据获取失败，尝试使用adata获取历史数据
            try:
                print(f"尝试使用adata获取{stock_code}的历史数据...")
                # 获取最近20个交易日的数据
                end_date = datetime.datetime.now().strftime("%Y-%m-%d")
                start_date = (datetime.datetime.now() - timedelta(days=20)).strftime("%Y-%m-%d")
                
                # 调整股票代码格式，A股需要加上市场前缀
                formatted_code = stock_code
                if len(stock_code) == 6:
                    # 沪市股票以6开头
                    if stock_code.startswith('6'):
                        formatted_code = f'sh{stock_code}'
                    # 深市股票以0、3、2开头
                    elif stock_code.startswith(('0', '3', '2')):
                        formatted_code = f'sz{stock_code}'

                df = adata.stock.market.get_market(
                    formatted_code,
                    start_date=start_date,
                    end_date=end_date,
                    k_type=1,
                    adjust_type=1,
                )

                if df is not None and not df.empty:
                    print(f"adata获取{stock_code}数据成功，共{len(df)}条")
                    # 确保收盘价是数值类型
                    df["close"] = pd.to_numeric(df["close"], errors="coerce")
                    df = df.dropna(subset=["close"])

                    if len(df) >= 5:
                        # 计算5日均线
                        df["ma5"] = df["close"].rolling(window=5).mean()
                        ma5_value = df["ma5"].iloc[-1]
                        result = round(ma5_value, 2)
                        print(f"{stock_code}的MA5值: {result}")
                        return result
                    else:
                        print(f"{stock_code}数据不足5条，无法计算MA5")
                else:
                    print(f"adata获取{stock_code}数据为空")
            except Exception as e:
                print(f"adata获取{stock_code}的5日均线数据失败: {str(e)}")
                
        except Exception as e:
            print(f"获取{stock_code}的5日均线数据失败: {str(e)}")

        return None

    def get_volume_ratio_data_batch(self, stock_codes):
        """批量获取近十日单日最大量比数据"""
        today = datetime.datetime.now().date()

        # 如果今天已经获取过数据，直接使用缓存
        if self.last_volume_ratio_fetch_date == today:
            results = {}
            for code in stock_codes:
                if code in self.volume_ratio_cache:
                    cached_time, volume_ratio_value = self.volume_ratio_cache[code]
                    if cached_time.date() == today:
                        results[code] = volume_ratio_value
            return results

        # 否则批量获取数据
        results = {}
        
        # 只处理有实时数据的股票
        valid_codes = [code for code in stock_codes if code not in results]
        
        # 分小批次处理，减少内存占用
        batch_size = 10  # 每次处理10个股票
        for i in range(0, len(valid_codes), batch_size):
            batch = valid_codes[i:i+batch_size]
            futures = {}
            
            for code in batch:
                # 如果缓存中有今天的数据，直接使用
                if code in self.volume_ratio_cache:
                    cached_time, volume_ratio_value = self.volume_ratio_cache[code]
                    if cached_time.date() == today:
                        results[code] = volume_ratio_value
                        continue
                
                # 否则提交任务到线程池
                future = self.thread_pool.submit(self._get_single_volume_ratio_data, code)
                futures[future] = code
            
            # 等待当前批次完成
            for future in as_completed(futures):
                code = futures[future]
                try:
                    volume_ratio_value = future.result()
                    if volume_ratio_value is not None:
                        results[code] = volume_ratio_value
                        self.volume_ratio_cache[code] = (datetime.datetime.now(), volume_ratio_value)
                except Exception as e:
                    # 减少打印，只在必要时输出
                    # print(f"获取{code}的量比数据失败: {str(e)}")
                    pass

        self.last_volume_ratio_fetch_date = today
        return results

    def _get_single_volume_ratio_data(self, stock_code):
        """获取单个股票的近十日单日最大量比数据，使用共享的历史数据"""
        try:
            # 首先尝试使用共享的历史数据
            df = self._get_stock_history_data(stock_code)
            
            if df is not None and not df.empty:
                print(f"使用共享历史数据计算{stock_code}的量比值...")
                
                # 检查数据列名
                volume_col = None
                if 'volume' in df.columns:
                    volume_col = 'volume'
                elif 'amount' in df.columns:
                    volume_col = 'amount'
                else:
                    print(f"历史数据中没有找到成交量或成交额列")
                    return None
                
                # 确保数据是数值类型
                df[volume_col] = pd.to_numeric(df[volume_col], errors='coerce')
                df = df.dropna(subset=[volume_col])
                
                if len(df) >= 2:
                    # 计算量比：当日成交量/前一日成交量
                    df['volume_ratio'] = df[volume_col] / df[volume_col].shift(1)
                    # 移除NaN值
                    df = df.dropna(subset=['volume_ratio'])
                    
                    if len(df) >= 1:
                        # 获取近10天的最大量比
                        recent_data = df.tail(CONFIG["volume_ratio_days"])
                        max_volume_ratio = recent_data['volume_ratio'].max()
                        result = round(max_volume_ratio, 2)
                        print(f"{stock_code}的近十日最大量比: {result}")
                        return result
                    else:
                        print(f"{stock_code}量比计算数据不足")
                else:
                    print(f"{stock_code}数据不足2条，无法计算量比")
            
            # 如果共享数据获取失败，尝试使用adata获取历史数据
            try:
                print(f"尝试使用adata获取{stock_code}的量比数据...")
                # 获取最近20个交易日的数据
                end_date = datetime.datetime.now().strftime("%Y-%m-%d")
                start_date = (datetime.datetime.now() - timedelta(days=20)).strftime("%Y-%m-%d")
                
                # 调整股票代码格式，A股需要加上市场前缀
                formatted_code = stock_code
                if len(stock_code) == 6:
                    # 沪市股票以6开头
                    if stock_code.startswith('6'):
                        formatted_code = f'sh{stock_code}'
                    # 深市股票以0、3、2开头
                    elif stock_code.startswith(('0', '3', '2')):
                        formatted_code = f'sz{stock_code}'

                df = adata.stock.market.get_market(
                    formatted_code,
                    start_date=start_date,
                    end_date=end_date,
                    k_type=1,
                    adjust_type=1,
                )

                if df is not None and not df.empty:
                    print(f"adata获取{stock_code}量比数据成功，共{len(df)}条")
                    
                    # 检查数据列名
                    volume_col = None
                    if 'volume' in df.columns:
                        volume_col = 'volume'
                    elif 'amount' in df.columns:
                        volume_col = 'amount'
                    else:
                        print(f"adata数据中没有找到成交量或成交额列")
                        return None
                    
                    # 确保数据是数值类型
                    df[volume_col] = pd.to_numeric(df[volume_col], errors="coerce")
                    df = df.dropna(subset=[volume_col])

                    if len(df) >= 2:
                        # 计算量比：当日成交量/前一日成交量
                        df["volume_ratio"] = df[volume_col] / df[volume_col].shift(1)
                        df = df.dropna(subset=["volume_ratio"])
                        
                        if len(df) >= 1:
                            # 获取近10天的最大量比
                            recent_data = df.tail(CONFIG["volume_ratio_days"])
                            max_volume_ratio = recent_data["volume_ratio"].max()
                            result = round(max_volume_ratio, 2)
                            print(f"{stock_code}的近十日最大量比: {result}")
                            return result
                        else:
                            print(f"{stock_code}量比计算数据不足")
                    else:
                        print(f"{stock_code}数据不足2条，无法计算量比")
                else:
                    print(f"adata获取{stock_code}量比数据为空")
            except Exception as e:
                print(f"adata获取{stock_code}的量比数据失败: {str(e)}")
                
        except Exception as e:
            print(f"获取{stock_code}的量比数据失败: {str(e)}")

        return None

    def calculate_ma5_distance(self, current_price, ma5_price):
        """计算当前价格与5日线的距离百分比"""
        if ma5_price is None or ma5_price == 0 or current_price is None:
            return None

        try:
            current_price = float(current_price)
            ma5_price = float(ma5_price)

            distance = ((current_price - ma5_price) / ma5_price) * 100
            return round(distance, 2)
        except (ValueError, TypeError):
            return None

    def calculate_amplitude_10d(self, stock_code, current_price):
        """计算近10日振幅：(近10日最高价 - 近10日最低价) / 当前股价"""
        try:
            # 获取历史数据
            df = self._get_stock_history_data(stock_code)
            
            if df is None or df.empty:
                return None
            
            # 确保数据是数值类型
            df['high'] = pd.to_numeric(df['high'], errors='coerce')
            df['low'] = pd.to_numeric(df['low'], errors='coerce')
            df = df.dropna(subset=['high', 'low'])
            
            if len(df) < 10:
                return None
            
            # 获取近10天的数据
            recent_10d = df.tail(10)
            
            # 计算近10日最高价和最低价
            high_10d = recent_10d['high'].max()
            low_10d = recent_10d['low'].min()
            
            # 计算振幅
            if current_price and current_price > 0:
                amplitude = ((high_10d - low_10d) / current_price) * 100
                return round(amplitude, 2)
            
            return None
        except Exception as e:
            print(f"计算{stock_code}近10日振幅失败: {str(e)}")
            return None

    def analyze(self):
        """分析板块实时数据"""
        return self._analyze_specific_block()

    def _analyze_specific_block(self):
        """分析特定板块的个股"""
        # 获取指定板块的股票代码
        codes = self.blocks.get(self.specific_block, [])
        if not codes:
            return []

        # 分批次获取数据
        batch_size = CONFIG["batch_size"]
        total_batches = (len(codes) + batch_size - 1) // batch_size
        all_data = []

        for i in range(total_batches):
            start = i * batch_size
            end = min((i + 1) * batch_size, len(codes))
            batch_codes = codes[start:end]

            try:
                df = adata.stock.market.list_market_current(code_list=batch_codes)
                if not df.empty:
                    df["change_pct"] = pd.to_numeric(df["change_pct"], errors="coerce")
                    df["price"] = pd.to_numeric(df["price"], errors="coerce")
                    all_data.append(df.dropna(subset=["change_pct", "price"]))
                else:
                    print(f"第{i+1}批数据为空")
            except Exception as e:
                print(f"第{i+1}批数据获取失败: {str(e)}")

        if not all_data:
            return []

        df_all = pd.concat(all_data)

        # 批量获取MA5数据
        ma5_data = self.get_ma5_data_batch(codes)
        
        # 批量获取量比数据
        volume_ratio_data = self.get_volume_ratio_data_batch(codes)

        # 准备个股数据
        stock_data = []

        # 计算个股涨速
        for _, row in df_all.iterrows():
            stock_code = row["stock_code"]
            stock_name = row.get("short_name", "未知")
            current_change = row["change_pct"]
            current_price = row["price"]

            # 计算 1 分钟涨速
            speed_change_1min = 0
            if self.prev_data is not None:
                prev_row = self.prev_data[self.prev_data["stock_code"] == stock_code]
                if not prev_row.empty:
                    prev_change = prev_row.iloc[0]["change_pct"]
                    try:
                        current_change = float(current_change)
                        prev_change = float(prev_change)
                        speed_change_1min = current_change - prev_change
                    except (ValueError, TypeError): 
                        speed_change_1min = 0

            # 计算 5 分钟涨速
            speed_change_5min = 0
            if len(self.five_min_data) > 0:
                if len(self.five_min_data) < 5:
                    prev_five_df = self.five_min_data[0]
                else:
                    prev_five_df = self.five_min_data[-5]

                prev_five_row = prev_five_df[prev_five_df["stock_code"] == stock_code]
                if not prev_five_row.empty:
                    prev_five_change = prev_five_row.iloc[0]["change_pct"]
                    try:
                        current_change = float(current_change)
                        prev_five_change = float(prev_five_change)
                        speed_change_5min = current_change - prev_five_change
                    except (ValueError, TypeError):
                        speed_change_5min = 0

            # 获取5日均线数据
            ma5_price = ma5_data.get(stock_code)

            # 计算与5日线的距离
            ma5_distance = self.calculate_ma5_distance(current_price, ma5_price)
            
            # 获取量比数据
            max_volume_ratio = volume_ratio_data.get(stock_code)
            
            # 计算近10日振幅
            amplitude_10d = self.calculate_amplitude_10d(stock_code, current_price)

            # 获取自定义数据
            custom_text = ""
            pinned = False
            if stock_code in self.custom_data:
                custom_data = self.custom_data[stock_code]
                custom_text = custom_data.get("text", "")
                pinned = custom_data.get("pinned", False)

            stock_data.append(
                {
                    "name": stock_code,
                    "stock_name": stock_name,
                    "custom_text": custom_text,
                    "real_time_return": round(current_change, 2),
                    "speed_change_1min": round(speed_change_1min, 2),
                    "speed_change_5min": round(speed_change_5min, 2),
                    "ma5_distance": ma5_distance,
                    "max_volume_ratio": max_volume_ratio,
                    "amplitude_10d": amplitude_10d,
                    "pinned": pinned,
                }
            )

        # 更新上一分钟的数据
        self.prev_data = df_all
        # 更新 5 分钟数据缓存
        self.five_min_data.append(df_all)
        if len(self.five_min_data) > 5:
            self.five_min_data.pop(0)

        # 按照实时涨幅从高到低排序，但置顶的股票排在前面
        stock_data.sort(
            key=lambda x: (not x["pinned"], x["real_time_return"]), reverse=True
        )

        return stock_data


class StockMonitor(tk.Tk):
    def __init__(self, xml_path, specific_block=None):
        super().__init__()
        self.analyzer = BlockAnalyzer(xml_path, specific_block)
        self.paused = False
        self.topmost = False
        self.specific_block = specific_block
        self.text_colors = {}
        self.current_sort_column = None
        self.current_sort_reverse = False
        self._init_ui()
        self._initial_refresh()

    def _init_ui(self):
        """初始化用户界面"""
        title = "智能选股监控 v1.0"
        if self.specific_block:
            title += f" - {self.specific_block}板块"
        self.title(title)
        self.geometry("1500x800")

        # 配置样式
        style = ttk.Style()
        style.configure(
            "Treeview",
            font=CONFIG["ui"]["font"],
            background=CONFIG["ui"]["colors"]["bg"],
            fieldbackground=CONFIG["ui"]["colors"]["bg"],
            foreground=CONFIG["ui"]["colors"]["text"],
            rowheight=60,  # 进一步增大行高，确保20号字体能完整显示
        )
        # 配置表头样式
        style.configure(
            "Treeview.Heading",
            font=CONFIG["ui"]["font"],
        )

        # 创建时间标签
        self.time_label = ttk.Label(self, font=CONFIG["ui"]["font"])

        # 创建控件
        columns = (
            "股票名称",
            "实时涨幅",
            "1 分钟涨速",
            "5 分钟涨速",
            "距5日线",
            "近10日振幅",
            "量比>2",
            "自定义文本",
        )

        self.tree = ttk.Treeview(self, columns=columns, show="headings")
        self.scroll = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.btn_refresh = ttk.Button(self, text="立即刷新", command=self.safe_refresh)
        self.btn_pause = ttk.Button(self, text="暂停刷新", command=self.toggle_pause)
        self.btn_topmost = ttk.Button(self, text="置顶", command=self.toggle_topmost)

        # 创建两个状态标签
        self.status_time = ttk.Label(self, anchor="e", font=CONFIG["ui"]["font"])
        self.status_cost = ttk.Label(self, anchor="e", font=CONFIG["ui"]["font"])

        # 布局管理
        self.grid_rowconfigure(0, minsize=150)  # 增大第一行高度，确保所有控件有足够空间
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=0)
        self.grid_columnconfigure(2, weight=0)

        # 按钮布局，增大间距以适配高分辨率
        self.btn_refresh.grid(row=0, column=0, sticky="w", padx=10, pady=10)
        self.btn_pause.grid(row=0, column=0, sticky="w", padx=200, pady=10)
        self.btn_topmost.grid(row=0, column=0, sticky="w", padx=400, pady=10)

        # 左上角的时间标签
        self.time_label.grid(row=0, column=1, sticky="nw", padx=20, pady=10)

        # 右上角的状态标签，分两行显示
        self.status_time.grid(row=0, column=2, sticky="ne", padx=20, pady=(10, 0))  # 最后更新时间
        self.status_cost.grid(row=0, column=2, sticky="ne", padx=20, pady=(60, 0))  # 耗时，垂直偏移

        # 表格和滚动条
        self.tree.grid(row=1, column=0, columnspan=3, sticky="nsew")
        self.scroll.grid(row=1, column=3, sticky="ns")

        # 配置列
        col_widths = [
            ("股票名称", 400),
            ("实时涨幅", 200),
            ("1 分钟涨速", 200),
            ("5 分钟涨速", 200),
            ("距5日线", 200),
            ("近10日振幅", 200),
            ("量比>2", 200),
            ("自定义文本", 400),
        ]

        for col, width in col_widths:
            self.tree.heading(
                col,
                text=col,
                command=lambda c=col: self.sort_by_column(
                    c,
                    (
                        not self.current_sort_reverse
                        if self.current_sort_column == c
                        else False
                    ),
                ),
            )
            self.tree.column(col, width=width, anchor="center")

        # 配置颜色标签
        self.tree.tag_configure("rise", foreground=CONFIG["ui"]["colors"]["rise"])
        self.tree.tag_configure("fall", foreground=CONFIG["ui"]["colors"]["fall"])
        self.tree.tag_configure("purple", foreground=CONFIG["ui"]["colors"]["purple"])
        self.tree.tag_configure("pinned", background=CONFIG["ui"]["colors"]["pinned"])
        
        # 为量比>2添加特殊颜色标签
        self.tree.tag_configure("volume_ratio_high", foreground="#FF4500")  # 橙红色

        # 为20种不同的自定义文本值创建颜色标签
        colors = [
            CONFIG["ui"]["colors"]["group1"],
            CONFIG["ui"]["colors"]["group2"],
            CONFIG["ui"]["colors"]["group3"],
            CONFIG["ui"]["colors"]["group4"],
            CONFIG["ui"]["colors"]["group5"],
            CONFIG["ui"]["colors"]["group6"],
            CONFIG["ui"]["colors"]["group7"],
            CONFIG["ui"]["colors"]["group8"],
            CONFIG["ui"]["colors"]["group9"],
            CONFIG["ui"]["colors"]["group10"],
            CONFIG["ui"]["colors"]["group11"],
            CONFIG["ui"]["colors"]["group12"],
            CONFIG["ui"]["colors"]["group13"],
            CONFIG["ui"]["colors"]["group14"],
            CONFIG["ui"]["colors"]["group15"],
            CONFIG["ui"]["colors"]["group16"],
            CONFIG["ui"]["colors"]["group17"],
            CONFIG["ui"]["colors"]["group18"],
            CONFIG["ui"]["colors"]["group19"],
            CONFIG["ui"]["colors"]["group20"],
        ]

        for i, color in enumerate(colors, 1):
            self.tree.tag_configure(f"text_color_{i}", background=color)

        # 绑定双击事件
        self.tree.bind("<Double-1>", self.on_expand)

        # 绑定右键点击事件
        self.tree.bind("<Button-3>", self.on_right_click)

    def toggle_topmost(self):
        """切换置顶状态"""
        self.topmost = not self.topmost
        self.attributes("-topmost", self.topmost)
        self.btn_topmost.config(text="取消置顶" if self.topmost else "置顶")

    def sort_by_column(self, column, reverse=False):
        """根据列进行排序"""
        # 保存当前排序状态
        self.current_sort_column = column
        self.current_sort_reverse = reverse

        data = [(self.tree.set(k, column), k) for k in self.tree.get_children("")]

        # 尝试将数据转换为浮点数，以便正确排序数值列
        try:
            # 特别处理百分比数据（如红盘比例），去掉百分号并转换为浮点数
            if "%" in data[0][0]:  # 检查是否为百分比格式
                data = [(float(val.strip("%")), k) for val, k in data]
            else:
                data = [(float(val), k) for val, k in data]
        except (ValueError, IndexError):
            pass  # 如果无法转换为数值，则保持原样排序

        # 排序
        data.sort(reverse=reverse)

        # 更新 Treeview 显示
        for index, (_, k) in enumerate(data):
            self.tree.move(k, "", index)

        # 更新列标题的排序指示器
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)

        # 为当前排序列添加排序指示器
        sort_indicator = " ↓" if reverse else " ↑"
        self.tree.heading(column, text=column + sort_indicator)

    def on_right_click(self, event):
        """处理右键点击事件"""
        item = self.tree.identify_row(event.y)
        if item:
            # 获取股票代码
            stock_name = self.tree.item(item, "values")[0]
            stock_code = self._get_stock_code_by_name(stock_name)

            if stock_code:
                # 获取当前置顶状态
                pinned = False
                if stock_code in self.analyzer.custom_data:
                    custom_data = self.analyzer.custom_data[stock_code]
                    pinned = custom_data.get("pinned", False)

                # 创建右键菜单
                menu = tk.Menu(self, tearoff=0)
                menu.add_command(
                    label="编辑自定义文本",
                    command=lambda: self.edit_custom_text(stock_code, item),
                )

                # 根据置顶状态添加不同的菜单项
                if pinned:
                    menu.add_command(
                        label="取消置顶",
                        command=lambda: self.toggle_stock_pin(stock_code, item, False),
                    )
                else:
                    menu.add_command(
                        label="添加置顶",
                        command=lambda: self.toggle_stock_pin(stock_code, item, True),
                    )

                menu.tk_popup(event.x_root, event.y_root)

    def toggle_stock_pin(self, stock_code, item, pin_status):
        """切换个股置顶状态"""
        if stock_code not in self.analyzer.custom_data:
            self.analyzer.custom_data[stock_code] = {}

        self.analyzer.custom_data[stock_code]["pinned"] = pin_status
        self.analyzer.save_custom_data()

        # 更新界面显示
        self.safe_refresh()

    def _get_stock_code_by_name(self, stock_name):
        """通过股票名称获取股票代码"""
        for stock in self.analyzer.stock_data:
            if stock.get("stock_name", "") == stock_name:
                return stock["name"]
        return None

    def edit_custom_text(self, stock_code, item):
        """编辑自定义文本"""
        current_text = self.tree.item(item, "values")[7]  # 获取当前自定义文本
        new_text = simpledialog.askstring(
            "编辑自定义文本",
            "请输入自定义文本:",
            parent=self,
            initialvalue=current_text,
        )

        if new_text is not None:
            # 更新自定义数据
            if stock_code not in self.analyzer.custom_data:
                self.analyzer.custom_data[stock_code] = {}

            self.analyzer.custom_data[stock_code]["text"] = new_text
            self.analyzer.save_custom_data()

            # 更新界面显示
            values = list(self.tree.item(item, "values"))
            values[7] = new_text  # 更新自定义文本列
            self.tree.item(item, values=values)

            # 更新颜色
            self._update_item_color(item, new_text)

    def _get_color_for_text(self, text):
        """根据文本内容获取颜色标签"""
        if not text:
            return None

        # 如果文本已经分配了颜色，返回该颜色
        if text in self.text_colors:
            return self.text_colors[text]

        # 如果没有，则分配一个新颜色
        colors = [
            "text_color_1",
            "text_color_2",
            "text_color_3",
            "text_color_4",
            "text_color_5",
            "text_color_6",
            "text_color_7",
            "text_color_8",
            "text_color_9",
            "text_color_10",
            "text_color_11",
            "text_color_12",
            "text_color_13",
            "text_color_14",
            "text_color_15",
            "text_color_16",
            "text_color_17",
            "text_color_18",
            "text_color_19",
            "text_color_20",
        ]

        # 计算文本的哈希值来确定颜色
        color_index = hash(text) % len(colors)
        color_tag = colors[color_index]
        self.text_colors[text] = color_tag

        return color_tag

    def _update_item_color(self, item, text):
        """更新项的颜色"""
        # 清除所有文本颜色标签
        for i in range(1, 21):  # 清除20种可能的颜色标签
            tag = f"text_color_{i}"
            if tag in self.tree.item(item, "tags"):
                tags = list(self.tree.item(item, "tags"))
                tags.remove(tag)
                self.tree.item(item, tags=tags)

        # 添加新的文本颜色标签
        if text:
            color_tag = self._get_color_for_text(text)
            if color_tag:
                tags = list(self.tree.item(item, "tags")) + [color_tag]
                self.tree.item(item, tags=tags)

    def safe_refresh(self):
        """线程安全刷新"""
        if hasattr(self, "_is_refreshing"):
            return
        self._is_refreshing = True
        threading.Thread(target=self._refresh_data, daemon=True).start()

    def _refresh_data(self):
        """后台数据刷新任务"""
        try:
            start_time = time.time()
            data = self.analyzer.analyze()
            self._update_display(data)
            # 更新状态标签 - 分为两行
            current_time =  time.strftime("%H:%M:%S")
            cost_time = time.time() - start_time
            self.status_time.config(text=f"最后更新: {current_time}")
            self.status_cost.config(text=f"耗时: {cost_time:.1f}秒")
        except Exception as e:
            self.status_time.config(text="更新失败")
            self.status_cost.config(text=f"{str(e)}")
        finally:
            if hasattr(self, "_is_refreshing"):
                del self._is_refreshing

    def _update_display(self, data):
        """更新界面显示"""
        # 保存当前数据用于后续查找
        self.analyzer.stock_data = data

        # 保存当前选中的项目
        selected_items = self.tree.selection()

        # 先清空现有数据
        self.tree.delete(*self.tree.get_children())

        # 插入新数据，不进行排序
        for item in data:
            change = item["real_time_return"]
            tags = ["rise"] if change >= 0 else ["fall"]

            # 获取自定义文本
            custom_text = item.get("custom_text", "")

            # 获取5日线距离
            ma5_distance = item.get("ma5_distance", None)
            ma5_text = f"{ma5_distance}%" if ma5_distance is not None else "N/A"

            # 获取量比数据
            max_volume_ratio = item.get("max_volume_ratio", None)
            volume_ratio_text = f"{max_volume_ratio}" if max_volume_ratio is not None else "N/A"
            
            # 获取近10日振幅数据
            amplitude_10d = item.get("amplitude_10d", None)
            amplitude_text = f"{amplitude_10d}%" if amplitude_10d is not None else "N/A"
            
            # 如果量比>2，添加特殊颜色标签
            if max_volume_ratio is not None and max_volume_ratio > 2:
                tags.append("volume_ratio_high")

            # 如果5日线距离在-3%到3%之间，添加紫色标签
            if ma5_distance is not None and -3 <= ma5_distance <= 3:
                tags.append("purple")

            # 添加文本颜色标签
            if custom_text:
                color_tag = self._get_color_for_text(custom_text)
                if color_tag:
                    tags.append(color_tag)

            # 如果股票被置顶，添加置顶标签
            if item.get("pinned", False):
                tags.append("pinned")

            # 显示个股数据
            self.tree.insert(
                "",
                "end",
                values=(
                    item.get("stock_name", "未知"),  # 股票名称
                    f"{change}%" if change is not None else "N/A",  # 实时涨幅
                    f"{item['speed_change_1min']}%",  # 1 分钟涨速
                    f"{item['speed_change_5min']}%",  # 5 分钟涨速
                    ma5_text,  # 距5日线
                    amplitude_text,  # 近10日振幅
                    volume_ratio_text,  # 量比>2
                    custom_text,  # 自定义文本
                ),
                tags=tuple(tags),
            )

        # 如果有当前排序设置，应用排序
        if self.current_sort_column:
            self.sort_by_column(self.current_sort_column, self.current_sort_reverse)

        # 恢复选中的项目
        for item in selected_items:
            try:
                self.tree.selection_add(item)
            except:
                pass  # 如果项目不存在，忽略错误

    def _initial_refresh(self):
        """初始化刷新"""
        self._update_time()
        self.safe_refresh()
        self.after(CONFIG["refresh_interval"] * 1000, self._auto_refresh)

    def _update_time(self):
        current_time = datetime.datetime.now().strftime("%m/%d %A %H:%M")
        self.time_label.config(text=current_time)
        self.after(60000, self._update_time)

    def toggle_pause(self):
        self.paused = not self.paused
        self.btn_pause.config(text="恢复刷新" if self.paused else "暂停刷新")
        if not self.paused:
            self.safe_refresh()
            self._auto_refresh()

    def _auto_refresh(self):
        """自动定时刷新"""
        if self.winfo_exists() and not self.paused:
            self.safe_refresh()
            self.after(CONFIG["refresh_interval"] * 1000, self._auto_refresh)

    def on_expand(self, event):
        """双击展开板块详情"""
        item = self.tree.focus()
        if self.tree.parent(item):
            return

        # 新增的pyautogui函数调用
        self.pyautogui(item)

    def pyautogui(self, item):
        """通过pyautogui操作同花顺窗口"""
        # 双击个股时打开该股票
        # 由于不再显示股票代码，需要通过股票名称获取股票代码
        stock_name = self.tree.item(item, "values")[0]
        # 在分析器中查找股票代码
        stock_code = None
        for stock in self.analyzer.stock_data:
            if stock.get("stock_name", "") == stock_name:
                stock_code = stock["name"]
                break

        if not stock_code:
            messagebox.showwarning("警告", f"未找到股票 {stock_name} 的代码")
            return

        try:
            # 激活同花顺窗口
            ths_windows = gw.getWindowsWithTitle("同花顺远航版")
            if ths_windows:
                ths = ths_windows[0]
                if ths.isMinimized:
                    ths.restore()
                ths.activate()

                # 等待窗口激活
                time.sleep(0.1)
                pyautogui.write(stock_code, interval=0.03)  # 输入股票代码
                time.sleep(0.1)  # 等待输入完成
                pyautogui.press("enter")  # 回车确认
            else:
                messagebox.showwarning("警告", "请先打开同花顺远航版")
        except Exception as e:
            messagebox.showerror("错误", f"自动化操作失败: {str(e)}")


if __name__ == "__main__":
    XML_PATH = r"D:\同花顺远航版\bin\users\mx_713570454\blockstockV3.xml"

    # 在这里指定要分析的板块，例如：specific_block = "传媒"
    # specific_block = "股性"  # 修改这里为想要的板块名称
    # specific_block = "跟踪"  # 修改这里为想要的板块名称

    # specific_block = "沪深300"  # 修改这里为想要的板块名称
    specific_block = "股性500强"  # 修改这里为想要的板块名称

    app = StockMonitor(XML_PATH, specific_block)
    app.mainloop()
