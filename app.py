from flask import Flask, render_template, request, jsonify
import json
import os
import pyautogui
import pygetwindow as gw
import time
from datetime import datetime
import shutil
import pandas as pd

# 尝试导入akshare
akshare_available = False
try:
    import akshare as ak
    akshare_available = True
    print("akshare导入成功")
except ImportError:
    print("akshare未安装，将使用手动输入的行业信息")

app = Flask(__name__)

# 配置
DEFAULT_USER = "LYY"  # 默认用户名

# 用户数据目录结构
USER_DATA_DIR = "user"
USER_FOLDER = os.path.join(USER_DATA_DIR, DEFAULT_USER)
CONFIG = {
    "data_file": os.path.join(USER_FOLDER, "stock_notes.json"),
    "folders_file": os.path.join(USER_FOLDER, "folders.json"),
    "expanded_folders_file": os.path.join(USER_FOLDER, "expanded_folders.json"),
    "calendar_file": os.path.join(USER_FOLDER, "calendar.json"),
    "images_folder": os.path.join(USER_FOLDER, "stock_images")
}

# 确保目录存在
if not os.path.exists(USER_DATA_DIR):
    os.makedirs(USER_DATA_DIR)
if not os.path.exists(USER_FOLDER):
    os.makedirs(USER_FOLDER)
if not os.path.exists(CONFIG["images_folder"]):
    os.makedirs(CONFIG["images_folder"])

# 加载笔记数据
def load_notes():
    if os.path.exists(CONFIG["data_file"]):
        try:
            with open(CONFIG["data_file"], "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return {}
    return {}

# 保存笔记数据
def save_notes(notes):
    with open(CONFIG["data_file"], "w", encoding="utf-8") as f:
        json.dump(notes, f, ensure_ascii=False, indent=2)

# 加载文件夹结构
def load_folders():
    if os.path.exists(CONFIG["folders_file"]):
        try:
            with open(CONFIG["folders_file"], "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return {"root": {"name": "根目录", "items": []}}
    return {"root": {"name": "根目录", "items": []}}

# 保存文件夹结构
def save_folders(folders):
    try:
        print(f'保存文件夹结构到: {CONFIG["folders_file"]}')
        print(f'文件夹结构内容: {folders}')
        
        # 确保目录存在
        os.makedirs(os.path.dirname(CONFIG["folders_file"]), exist_ok=True)
        
        with open(CONFIG["folders_file"], "w", encoding="utf-8") as f:
            json.dump(folders, f, ensure_ascii=False, indent=2)
        
        print('文件夹结构保存成功')
    except Exception as e:
        print(f'保存文件夹结构失败: {str(e)}')
        raise

# 加载展开的文件夹状态
def load_expanded_folders():
    if os.path.exists(CONFIG["expanded_folders_file"]):
        try:
            with open(CONFIG["expanded_folders_file"], "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return ["root"]
    return ["root"]

# 保存展开的文件夹状态
def save_expanded_folders(expanded_folders):
    with open(CONFIG["expanded_folders_file"], "w", encoding="utf-8") as f:
        json.dump(expanded_folders, f, ensure_ascii=False, indent=2)

# 加载日历数据
def load_calendar():
    if os.path.exists(CONFIG["calendar_file"]):
        try:
            with open(CONFIG["calendar_file"], "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return {}
    return {}

# 保存日历数据
def save_calendar(calendar_data):
    with open(CONFIG["calendar_file"], "w", encoding="utf-8") as f:
        json.dump(calendar_data, f, ensure_ascii=False, indent=2)

# 加载本地行业信息
def load_industry_data():
    """加载本地行业信息文件"""
    industry_file = "sw_three_industries_2026.json"
    if os.path.exists(industry_file):
        try:
            with open(industry_file, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"加载行业信息文件失败: {e}")
            return {}
    else:
        print(f"行业信息文件 {industry_file} 不存在")
        return {}

# ==================== 核心计算函数 ====================

def calculate_ma5_distance(current_price, ma5_price):
    """
    计算当前价格与5日线的距离百分比
    公式: ((当前价格 - MA5价格) / MA5价格) * 100
    
    Args:
        current_price: 当前价格
        ma5_price: 5日均线价格
    
    Returns:
        float: 距离百分比（保留2位小数），计算失败返回None
    """
    if ma5_price is None or ma5_price == 0 or current_price is None:
        return None

    try:
        current_price = float(current_price)
        ma5_price = float(ma5_price)
        
        distance = ((current_price - ma5_price) / ma5_price) * 100
        return round(distance, 2)
    except (ValueError, TypeError):
        return None


def calculate_ma5_from_history(df_history):
    """
    从历史数据计算5日均线值
    
    Args:
        df_history: 包含'close'列的DataFrame（最近20天日线数据）
    
    Returns:
        float: MA5值（保留2位小数），数据不足返回None
    """
    if df_history is None or df_history.empty:
        return None
    
    # 确保收盘价是数值类型
    df_history['close'] = pd.to_numeric(df_history['close'], errors='coerce')
    df_history = df_history.dropna(subset=['close'])
    
    if len(df_history) >= 5:
        # 计算最近5天的平均值
        recent_5_days = df_history['close'].tail(5)
        ma5_value = recent_5_days.mean()
        return round(ma5_value, 2)
    else:
        return None


def get_stock_history_data(stock_code, days=20):
    """
    获取股票历史日线数据（用于计算MA5）
    
    Args:
        stock_code: 股票代码（如 '000001' 或 'sh000001'）
        days: 获取天数，默认20天
    
    Returns:
        DataFrame: 包含历史数据，失败返回None
    """
    try:
        # 导入Ashare模块
        import Ashare
        
        # 调整股票代码格式
        symbol = stock_code
        if len(stock_code) == 6:
            if stock_code.startswith('6'):
                symbol = f'sh{stock_code}'
            else:
                symbol = f'sz{stock_code}'
        
        # 获取最近N天的日线数据
        df = Ashare.get_price(symbol, frequency='1d', count=days)
        return df
        
    except Exception as e:
        print(f"获取{stock_code}历史数据失败: {str(e)}")
        return None

# 获取股票行业信息
def get_stock_industry(code):
    """从本地JSON文件获取股票的行业信息"""
    try:
        # 直接使用6位股票代码
        stock_code = code[-6:] if len(code) > 6 else code
        print(f"尝试获取股票 {stock_code} 的行业信息...")
        
        # 加载行业数据
        industry_data = load_industry_data()
        
        # 遍历所有行业
        for industry_name, industry_info in industry_data.items():
            stocks = industry_info.get('stocks', [])
            for stock in stocks:
                # 获取股票代码并提取6位数字部分
                stock_code_full = stock.get('代码', '')
                if stock_code_full:
                    # 提取6位数字代码
                    stock_code_in_file = stock_code_full.split('.')[0]
                    if stock_code_in_file == stock_code:
                        print(f"找到股票 {stock_code} 的行业: {industry_name}")
                        return industry_name
        
        print(f"未在本地行业信息中找到股票 {stock_code}")
                
    except Exception as e:
        print(f"获取行业信息失败: {e}")
    return None

# 打开同花顺
def open_in_ths(code):
    try:
        # 尝试多种可能的窗口标题
        possible_titles = ["同花顺远航版", "同花顺", "同花顺金融服务网", "同花顺-行情"]
        ths = None
        
        print("正在查找同花顺窗口...")
        # 遍历所有窗口标题，打印出来以便调试
        all_windows = gw.getAllTitles()
        print(f"当前所有窗口标题: {[title for title in all_windows if '同花顺' in title]}")
        
        for title in possible_titles:
            ths_windows = gw.getWindowsWithTitle(title)
            print(f"查找 '{title}': 找到 {len(ths_windows)} 个窗口")
            if ths_windows:
                ths = ths_windows[0]
                print(f"找到同花顺窗口: {ths.title}")
                break
        
        if ths:
            print(f"窗口信息: 标题={ths.title}, 位置=({ths.left}, {ths.top}), 大小=({ths.width}, {ths.height})")
            
            try:
                if ths.isMinimized:
                    print("窗口已最小化，正在还原...")
                    ths.restore()
                    time.sleep(0.3)
                
                print("正在激活窗口...")
                ths.activate()
                time.sleep(0.3)
                
                print("正在最大化窗口...")
                ths.maximize()
                time.sleep(0.3)
                
                # 确保窗口在前台
                print("正在将窗口置于前台...")
                ths.activate()
                time.sleep(0.5)
                
                # 使用不同的方法确保窗口激活
                try:
                    # 尝试使用SetForegroundWindow API
                    import win32gui
                    import win32con
                    hwnd = ths._hWnd
                    win32gui.SetForegroundWindow(hwnd)
                    print("使用Win32 API激活窗口成功")
                except Exception as e:
                    print(f"使用Win32 API激活窗口失败: {e}")
                
                # 等待窗口激活
                time.sleep(0.5)
                
                # 尝试点击窗口内部
                try:
                    click_x = ths.left + 100
                    click_y = ths.top + 100
                    print(f"正在点击窗口位置: ({click_x}, {click_y})")
                    pyautogui.click(click_x, click_y)
                    time.sleep(0.3)
                except Exception as e:
                    print(f"点击窗口失败: {e}")
                
                # 尝试输入股票代码
                print(f"正在输入股票代码: {code}")
                
                # 先尝试清空输入
                try:
                    pyautogui.hotkey('ctrl', 'a')
                    pyautogui.press('backspace')
                    time.sleep(0.2)
                except Exception as e:
                    print(f"清空输入失败: {e}")
                
                # 输入股票代码
                pyautogui.write(code, interval=0.05)
                time.sleep(0.3)
                
                # 按下回车键
                print("正在按下回车键...")
                pyautogui.press("enter")
                time.sleep(0.2)
                
                print("操作完成，同花顺已打开")
                return True
            except Exception as e:
                print(f"操作窗口失败: {e}")
                # 即使遇到异常，也尝试继续执行基本操作
                try:
                    print("尝试直接输入股票代码...")
                    pyautogui.write(code, interval=0.05)
                    time.sleep(0.3)
                    pyautogui.press("enter")
                    print("直接输入成功")
                    return True
                except Exception as e2:
                    print(f"直接输入也失败: {e2}")
                    return False
        else:
            print("未找到同花顺窗口")
            return False
    except Exception as e:
        print(f"打开同花顺失败: {str(e)}")
        # 即使发生异常，也尝试直接输入
        try:
            print("尝试全局输入股票代码...")
            pyautogui.write(code, interval=0.05)
            time.sleep(0.3)
            pyautogui.press("enter")
            print("全局输入成功")
            return True
        except Exception as e2:
            print(f"全局输入也失败: {e2}")
            return False

@app.route('/')
def index():
    from flask import send_file
    return send_file('股票笔记本.html')

@app.route('/api/notes', methods=['GET'])
def get_notes():
    notes = load_notes()
    
    # 每次获取笔记时更新所有股票的涨幅和偏离五日线幅度
    updated_notes = {}
    for code, note in notes.items():
        # 获取历史数据
        df_history = get_stock_history_data(code)
        if df_history is not None and not df_history.empty:
            # 使用历史数据中的最新收盘价
            current_price = df_history['close'].iloc[-1]
            note['close'] = current_price
            
            # 计算MA5
            ma5_price = calculate_ma5_from_history(df_history)
            if ma5_price:
                # 计算距五日线幅度
                ma5_distance = calculate_ma5_distance(current_price, ma5_price)
                note['ma5'] = ma5_price
                note['ma5_distance'] = ma5_distance
            
            # 计算涨幅（如果有昨日收盘价）
            if len(df_history) >= 2:
                yesterday_close = df_history['close'].iloc[-2]
                if yesterday_close > 0:
                    change_percent = ((float(current_price) - yesterday_close) / yesterday_close) * 100
                    note['change_percent'] = round(change_percent, 2)
                    note['yesterdayPrice'] = yesterday_close
        updated_notes[code] = note
    
    # 保存更新后的笔记
    save_notes(updated_notes)
    return jsonify(updated_notes)

@app.route('/api/notes', methods=['POST'])
def save_note():
    data = request.json
    notes = load_notes()
    code = data.get('code')
    if code:
        # 检查是否需要自动获取行业信息
        if not data.get('industry') or data.get('industry').strip() == '':
            print(f"自动获取股票 {code} 的行业信息...")
            industry = get_stock_industry(code)
            if industry:
                print(f"获取到行业信息: {industry}")
                data['industry'] = industry
            else:
                print("无法获取行业信息，使用空值")
        
        # 获取历史数据计算收盘价、涨幅和距五日线幅度
        df_history = get_stock_history_data(code)
        if df_history is not None and not df_history.empty:
            # 使用历史数据中的最新收盘价
            current_price = df_history['close'].iloc[-1]
            data['close'] = current_price
            
            # 计算MA5
            ma5_price = calculate_ma5_from_history(df_history)
            if ma5_price:
                # 计算距五日线幅度
                ma5_distance = calculate_ma5_distance(current_price, ma5_price)
                data['ma5'] = ma5_price
                data['ma5_distance'] = ma5_distance
            
            # 计算涨幅（如果有昨日收盘价）
            if len(df_history) >= 2:
                yesterday_close = df_history['close'].iloc[-2]
                if yesterday_close > 0:
                    change_percent = ((float(current_price) - yesterday_close) / yesterday_close) * 100
                    data['change_percent'] = round(change_percent, 2)
                    data['yesterdayPrice'] = yesterday_close
        
        notes[code] = data
        save_notes(notes)
        return jsonify({"status": "success"})
    return jsonify({"status": "error", "message": "股票代码不能为空"}), 400

@app.route('/api/notes/<code>', methods=['DELETE'])
def delete_note(code):
    notes = load_notes()
    if code in notes:
        del notes[code]
        save_notes(notes)
        return jsonify({"status": "success"})
    return jsonify({"status": "error", "message": "笔记不存在"}), 404

@app.route('/api/open_ths', methods=['POST'])
def open_ths():
    data = request.json
    code = data.get('code')
    if code:
        success = open_in_ths(code)
        if success:
            return jsonify({"status": "success"})
        else:
            return jsonify({"status": "error", "message": "请先打开同花顺远航版"}), 400
    return jsonify({"status": "error", "message": "股票代码不能为空"}), 400

# 文件夹结构相关API
@app.route('/api/folders', methods=['GET'])
def get_folders():
    folders = load_folders()
    expanded_folders = load_expanded_folders()
    return jsonify({"folders": folders, "expanded_folders": expanded_folders})

@app.route('/api/folders', methods=['POST'])
def save_folders_api():
    try:
        data = request.json
        print('收到保存文件夹结构的请求')
        print('请求数据:', data)
        
        # 核心：确保即使前端传空，也有默认结构
        folders = data.get('folders', {"root": {"name": "根目录", "items": []}})
        expanded_folders = data.get('expanded_folders', ["root"])
        
        print('保存的文件夹结构:', folders)
        print('保存的展开状态:', expanded_folders)
        
        # 强制创建用户目录
        os.makedirs(os.path.dirname(CONFIG["folders_file"]), exist_ok=True)
        os.makedirs(os.path.dirname(CONFIG["expanded_folders_file"]), exist_ok=True)
        
        # 直接保存文件，不使用封装函数以减少潜在问题
        with open(CONFIG["folders_file"], "w", encoding="utf-8") as f:
            json.dump(folders, f, ensure_ascii=False, indent=2)
        
        with open(CONFIG["expanded_folders_file"], "w", encoding="utf-8") as f:
            json.dump(expanded_folders, f, ensure_ascii=False, indent=2)
        
        print('文件夹结构保存成功')
        return jsonify({"status": "success"})
    except Exception as e:
        print(f'保存文件夹结构失败: {str(e)}')
        return jsonify({"status": "error", "message": str(e)}), 500

# 日历数据相关API
@app.route('/api/calendar', methods=['GET'])
def get_calendar():
    calendar_data = load_calendar()
    return jsonify(calendar_data)

@app.route('/api/calendar', methods=['POST'])
def save_calendar_api():
    calendar_data = request.json
    save_calendar(calendar_data)
    return jsonify({"status": "success"})

# 获取K线数据
@app.route('/api/kline', methods=['GET'])
def get_kline():
    try:
        # 导入Ashare模块
        import Ashare
        
        code = request.args.get('code')
        frequency = request.args.get('frequency', '1d')  # 默认日线
        count = int(request.args.get('count', 30))  # 默认30条数据
        
        if not code:
            return jsonify({"status": "error", "message": "股票代码不能为空"}), 400
        
        # 转换股票代码格式
        stock_code = code
        if code.startswith('60'):
            stock_code = f"{code}.XSHG"
        elif code.startswith(('00', '30')):
            stock_code = f"{code}.XSHE"
        
        # 获取K线数据
        df = Ashare.get_price(stock_code, frequency=frequency, count=count)
        
        # 转换为JSON格式
        if not df.empty:
            # 处理数据格式
            kline_data = []
            for idx, row in df.iterrows():
                kline_data.append({
                    'date': idx.strftime('%Y-%m-%d'),
                    'open': float(row['open']),
                    'close': float(row['close']),
                    'high': float(row['high']),
                    'low': float(row['low']),
                    'volume': float(row['volume'])
                })
            return jsonify({"status": "success", "data": kline_data})
        else:
            return jsonify({"status": "error", "message": "获取K线数据失败"}), 400
            
    except Exception as e:
        print(f"获取K线数据失败: {e}")
        return jsonify({"status": "error", "message": f"获取K线数据失败: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(debug=True, host='localhost', port=5000)
