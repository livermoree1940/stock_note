# -*- coding:utf-8 -*-
import json, requests, datetime
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ================= Ashare 核心函数 (数据获取) =================

def get_price_day_tx(code, end_date='', count=10, frequency='1d'): 
    unit='week' if frequency in '1w' else 'month' if frequency in '1M' else 'day'
    if end_date: end_date=end_date.strftime('%Y-%m-%d') if isinstance(end_date,datetime.date) else end_date.split(' ')[0]
    end_date='' if end_date==datetime.datetime.now().strftime('%Y-%m-%d') else end_date 
    URL=f'http://web.ifzq.gtimg.cn/appstock/app/fqkline/get?param={code},{unit},,{end_date},{count},qfq' 
    st= json.loads(requests.get(URL).content); ms='qfq'+unit; stk=st['data'][code] 
    buf=stk[ms] if ms in stk else stk[unit]
    if len(buf) > 0:
        buf = [item[:6] for item in buf if len(item)>=6]
    if len(buf) == 0: return pd.DataFrame()
    df=pd.DataFrame(buf,columns=['time','open','close','high','low','volume'])
    df[['open','close','high','low','volume']] = df[['open','close','high','low','volume']].astype('float')
    df.time=pd.to_datetime(df.time); df.set_index(['time'], inplace=True); df.index.name=''
    return df

def get_price_sina(code, end_date='', count=10, frequency='60m'): 
    frequency=frequency.replace('1d','240m').replace('1w','1200m').replace('1M','7200m'); mcount=count
    ts=int(frequency[:-1]) if frequency[:-1].isdigit() else 1
    if (end_date!='') & (frequency in ['240m','1200m','7200m']): 
        end_date=pd.to_datetime(end_date)
        unit=4 if frequency=='1200m' else 29 if frequency=='7200m' else 1
        count=count+(datetime.datetime.now()-end_date).days//unit
    URL=f'http://money.finance.sina.com.cn/quotes_service/api/json_v2.php/CN_MarketData.getKLineData?symbol={code}&scale={ts}&ma=5&datalen={count}' 
    res = requests.get(URL).content
    if not res: return pd.DataFrame()
    dstr= json.loads(res); 
    df= pd.DataFrame(dstr,columns=['day','open','high','low','close','volume'])
    df[['open','high','low','close','volume']] = df[['open','high','low','close','volume']].astype(float)
    df.day=pd.to_datetime(df.day); df.set_index(['day'], inplace=True); df.index.name=''
    if (end_date!='') & (frequency in ['240m','1200m','7200m']): return df[df.index<=end_date][-mcount:]
    return df

def get_price(code, end_date='', count=10, frequency='1d'):
    """唯一的对外暴露接口"""
    xcode= code.replace('.XSHG','').replace('.XSHE','')
    xcode='sh'+xcode if ('XSHG' in code) else 'sz'+xcode if ('XSHE' in code) else code 
    if frequency in ['1d','1w','1M']:
        try: return get_price_sina(xcode, end_date=end_date, count=count, frequency=frequency)
        except: return get_price_day_tx(xcode, end_date=end_date, count=count, frequency=frequency)
    return pd.DataFrame()

# ================= Plotly 绘图函数 =================

def draw_kline_with_ashare(code, count=100):
    df = get_price(code, frequency='1d', count=count)
    if df is None or df.empty:
        print(f"无法获取数据: {code}")
        return None

    # 计算均线
    df['MA5'] = df['close'].rolling(5).mean()
    df['MA10'] = df['close'].rolling(10).mean()
    df['MA20'] = df['close'].rolling(20).mean()

    # 创建画布
    fig = make_subplots(rows=2, cols=1, shared_xaxes=True, 
                        vertical_spacing=0.05, row_width=[0.3, 0.7])

    # K线图
    fig.add_trace(go.Candlestick(
        x=df.index, open=df['open'], high=df['high'], low=df['low'], close=df['close'],
        name='K线', increasing_line_color='#f64e60', decreasing_line_color='#0bb783'
    ), row=1, col=1)

    # 均线
    fig.add_trace(go.Scatter(x=df.index, y=df['MA5'], name='MA5', line=dict(color='orange', width=1)), row=1, col=1)
    fig.add_trace(go.Scatter(x=df.index, y=df['MA20'], name='MA20', line=dict(color='blue', width=1.5)), row=1, col=1)

    # 成交量
    colors = ['#f64e60' if c >= o else '#0bb783' for o, c in zip(df['open'], df['close'])]
    fig.add_trace(go.Bar(x=df.index, y=df['volume'], name='成交量', marker_color=colors), row=2, col=1)

    fig.update_layout(title=f'股票: {code}', xaxis_rangeslider_visible=False, height=600, template='plotly_white')
    return fig

# ================= 运行入口 =================

if __name__ == '__main__':
    target_code = 'sh600519' # 茅台
    fig = draw_kline_with_ashare(target_code, count=100)
    if fig:
        fig.show()