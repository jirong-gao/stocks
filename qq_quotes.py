#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
from urllib import request, error
import re #正则表达式
import csv
import time

# DEBUG状态定义
# 非DEBUG状态下的输出信息少很多
DEBUG = False

# 存放获得的证券行情数据文件的路径，及文件名
# 该文件将会作为Excel文件的外部数据源，如果改动路径和文件名，需要在Excel文件中更新数据连接
# 该文件是类似csv格式的文件,只不过以波浪号"~"作为分隔符
QUOTES_FILE_PATH = r'd:\tmp'
QUOTES_FILE_NAME = 'quotes.dat'

# 存放所关注证券代码的文件，文件中包含了相应的查询代码
# 文件中的所有证券的行情信息将会被更新
# 查询代码的格式为：Market + Symbol
# Market是证券交易场所（大小写敏感）：sh -- 上海，sz -- 深圳，hk -- 香港，s_jj -- 基金
# Symbol：证券代码
WATCHING_STOCKS_FILE = 'watching_stocks.dat'

# 定义腾讯查询API一次调用允许的最大查询代码数，太多会出现返回不完全的现象；
# 经探索，超过30个代码一般会出错；这里我将一次查询限制为最大20个
MAX_QUERY_CODES = 20

# 每次调用腾讯行情数据API之前需要延时一下，以避免调用过于频繁被服务器端拒绝
# 延时0.5秒
API_CALLING_INTERVAL = 0.5

# 如果在调用腾讯API时发生错误时，可以重试的次数
API_CALLING_RETRIES = 3

def check_env():
    '''
    检查所需的环境，如果某些环境配置不存在，进行初始化工作
    比如存储行情数据文件的路径
    
    返回True，表示环境准备好，可以继续执行
    返回False，表示环境有问题，必须中断执行
    '''
    
    # 如果存放证券行情数据文件的目录不存在,则创建
    if not os.path.exists(QUOTES_FILE_PATH):
        os.mkdir(QUOTES_FILE_PATH)
    
    # 检查存放证券代码的文件是否存在
    file_path = os.path.join(os.getcwd(), WATCHING_STOCKS_FILE)
    if not os.path.isfile(file_path):
        # 如不存在，返回False
        print('环境检查出错：')
        print('    ', file_path, '文件不存在，请配置好该文件再重新运行！')
        return False
    
    return True


def load_watching_stocks():
    '''
    从WATCHING_STOCKS_FILE文件中读取查询代码
    文件中的空行，或者查询代码仅包含空格的，都会被跳过
    
    返回一个list，包含所有要查询的证券查询代码
    '''
    # 组装文件的全路径
    file_name = os.path.join(os.getcwd(), WATCHING_STOCKS_FILE)
    
    watching_codes = []
    
    # 使用with，Python会自动地调用close()，不用再写了
    with open(file_name, 'r', newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile, delimiter=',')
        
        for row in reader:
            if row:
                # 不是空行，才处理
                if row[0].strip():
                    # 如果只包含空格,则跳过
                    # 不作非常严格的检查,以便将来出现新的查询代码时不用修改代码
                    watching_codes.append(row[0])
                
    if not watching_codes:
        print('文件', file_name, '中没有内容，请添加需要关注的证券查询代码!\n')
        
    return watching_codes

def refresh_quotes():
    '''
    完成获取行情数据的调用
    并且将分解好的行情数据存入证券行情数据文件中，以备Excel文件同步
    '''
    
    # 从WATCHING_STOCKS_FILE文件中读取查询代码
    watching_codes_all = load_watching_stocks()
    
    if not watching_codes_all:
        return
    
    # 分组，每组不超过20个代码
    grouped_watching_codes = []
    for i in range(0, len(watching_codes_all), MAX_QUERY_CODES):
        grouped_watching_codes.append(watching_codes_all[i:i+MAX_QUERY_CODES])
    
    print('共有', len(watching_codes_all), '个证券查询代码，分为', len(grouped_watching_codes), '组进行查询！\n')
       
    print('开始从腾讯获取行情数据...\n')
    
    # 这是存放分解之后的行情数据，最终被存储到行情数据文件中
    # 每一个元素都是一个list，代表一个证券的行情数据，对应文件中的一行
    quote_list = [] 
    
    # 第一行是头
    quote_list.append(['查询代码', '证券代码', '市场', '名称', '价格', '涨跌', r'涨跌幅(%)', 'PE', 'PB', '总市值'])
    
    # 第二行是当前获取行情数据的时间信息,这样便于在Excel文件中查看数据是否更新
    localtime = time.localtime(time.time())
    ftime = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    
    # 查询代码为tmquotes，用于在Excel文件中使用VLOOKUP定位
    quote_list.append(['tmquotes','', '', '行情数据时间', ftime])
    
    
    #调用腾讯的股市行情API获取行情数据
    # quote_strs = get_quotes_qq(['sh000001','sz399300','sh600519','hk00857','sz000858','s_jj160706'])
    for watching_codes in grouped_watching_codes:  
        quote_strs = get_quotes_qq(watching_codes)
    
        # 对每一条行情数据进行分解，获取我们关心的信息
        for quote_str in quote_strs:
            # 跳过空字符串，这可能是腾讯返回的字符串中最后一个分号引起的
            if not quote_str:
                debug_print('分解行情数据：跳过一条空字符串！\n')
                continue
                
            debug_print('开始分解行情数据：')
            debug_print(quote_str, '\n')
            
            quote = parse_full_qq_stock_quote(quote_str)
                    
            if not (quote is None):
                quote_list.append(quote)
                
                debug_print('本条行情数据分解成功：')
                debug_print(quote, '\n')
    
    print('所有行情数据下载、分解已完成，正在保存行情数据...\n')
    
    # 将分解出来的行情数据存储到文件中
    store_quotes(quote_list)
    
    print('行情数据保存成功，可以从Excel中进行同步操作了！\n')

def get_quotes_qq(codes):
    '''
    调用腾讯的股市行情API，返回指定证券的行情数据
    腾讯API调用示例：http://qt.gtimg.cn/q=sz000858,sh600010,hk00857,s_jj160706
    
    传入的证券代码为一个list，证券代码中的英文字母大小写敏感
    由于腾讯API的限制，每次调用超过30个证券时，就有可能出错
    因此这里要求传入的查询代码不要要超过20个，如果超过也只取前面的20个进行查询
    
    腾讯API返回的行情数据是以分号";"为分隔符的一个字符串
    成功获取后，将其分解成为一个list返回，list中每一个元素（字符串）代表一个证券的行情数据
    
    注意：腾讯返回的字符串末尾有一个分号，所以返回的list中有一个空的字符串
    '''
    
    # 证券代码通过逗号“，”拼接起来作为腾讯API的参数传入，最多取前20个
    url = 'http://qt.gtimg.cn/q=' + ','.join(codes[0:MAX_QUERY_CODES])
    
    # 服务器发生错误时，重试三次    
    retry = 0
    api_call_success = False
    
    while (not api_call_success) and (retry < API_CALLING_RETRIES):    
        try:
            #每次调用前先延时一下
            time.sleep(API_CALLING_INTERVAL)
            retry = retry + 1
            
            # urlopen函数返回一个http.client.HTTPResponse对象
            with request.urlopen(url) as resp:
                # 腾迅的返回内容是以GBK编码的
                # 可以对返回内容的编码进行智能检测，但这样会要求安装第三方库chardet
                # 为了使用的方便，降低环境准备的复杂度，这里直接使用硬代码
                data = resp.read().decode('GBK')
                
                data = data.replace('\n', '').replace('r', '') # 去掉里面的回车换行符
                
                '''
                最后一个分号会产生一个空字符串
                不在这里处理最后一个分号，外部代码应该负责进行校验
                
                if data.endswith(';'):
                    # 如果字符串末尾有一个分号，这是腾讯API返回的
                    # 先在这里去掉
                    data = data[:-1] # 去掉字符串末尾的一个分号
                '''
                
                if DEBUG:
                    debug_print('Status:', resp.status, resp.reason)
                
                    for k, v in resp.getheaders():
                        debug_print('%s: %s' % (k, v))

                    debug_print('行情数据:\n', data, '\n')
                    
        except error.HTTPError as err:
            print('调用腾讯API发生异常：\n')
            print(err.reason, err.code, err.headers, sep='\n')
            print('第', retry, '次重试...\n')
        except error.URLError as err:
            print('调用腾讯API发生URL异常，请检查，不再重试！\n')
            print(err.reason)
            retry = API_CALLING_RETRIES # 终止执行
        else: # 无异常发生，即调用成功
            api_call_success = True
            
    
    if api_call_success:
        return data.split(';')
    else:
        return []
        


def parse_full_qq_stock_quote(quote_str):
    '''
    分解腾讯股票接口数据，返回一个list，包含我们需要的各个字段
    支持分解港股
    传入的数据格式如下例：
    v_sz000858="51~五 粮 液~000858~24.80~25.11~25.13~207959~90194~117764~24.80~1438~24.79~451~24.78~819~24.77~115~24.76~293~24.81~628~24.82~539~24.83~557~24.84~191~24.85~126~15:00:18/24.80/2269/S/5626376/19587|14:57:00/24.81/9/B/22329/19437|14:57:00/24.81/73/B/181076/19435|14:56:57/24.81/23/B/57063/19433|14:56:54/24.80/49/S/121557/19428|14:56:51/24.81/14/B/34734/19425~20130308150351~-0.31~-1.23~25.25~24.79~24.81/205690/514010790~207959~51964~0.55~9.05~~25.25~24.79~1.83~941.28~941.40~3.25~27.62~22.60~"
    v_hk00700="100~腾讯控股~00700~294.400~293.800~294.000~9286639.0~0~0~294.400~0~0~0~0~0~0~0~0~0~294.400~0~0~0~0~0~0~0~0~0~9286639.0~2018/10/10 10:34:10~0.600~0.20~296.800~293.600~294.400~9286639.0~2738811478.770~0~33.09~~0~0~1.09~28031.731~28031.731~TENCENT~0.30~475.720~293.600~2.29~-7.48~0~0~0~0~0~28.81~9.24~0.10~100~"
    
    如果传入的字符串不是一个合法的行情数据字符串，返回None
    '''
    
    # 使用正则表达式提取行情字符串的信息，正确的字符串应该包含两段内容
    # 查询代码位于 v_ 和随后的 =" 之间，然后是到最后的 " 之间，是行情数据
    # 下面字符串前面的r前缀是python提供的一种写法，说明后面的字符串都是普通字符，不用再用"\"进行转义了
    pattern = re.compile(r'v_(.+)="(.+)"') # 两个括号定义了两个组，用来提取想要的信息
    
    result = pattern.match(quote_str)
    
    #如果没有匹配上，match返回None
    if result is None:
        print('行情数据字符串格式不正确：')
        print('    ', quote_str, '\n')
        
        return None # 返回None
    
    strs = result.groups()
    if len(strs) != 2:
        print('行情数据字符串包含的信息不足：')
        print('    ', quote_str, '\n')
        
        return None
    
    # 正确取出所需字段，第一个是查询代码，第二个是行情数据
    
    # 根据查询代码判断，该行情数据的内容是基金还是股票
    if strs[0].startswith('s_jj'):
        # 查询代码以s_jj开头，表示是基金
        return parse_qq_fund_quote(strs[0], strs[1])
    else:
        # 其它的认为是股票
        return parse_qq_stock_quote(strs[0], strs[1])
    

def parse_qq_stock_quote(query_code, quote_str):
    '''
    分解股票行情数据
    query_code: 查询代码
    quote_str: 不带查询代码的行情数据字符串
    '''
    rtn_strs = []
    rtn_strs.append(query_code) # 1 查询代码 - QueryCode
    
    # 分隔行情数据字符串，波浪号"~"为分隔符
    fields = quote_str.split('~') 
    
    rtn_strs.append(fields[2])      # 2 股票代码 - StockSymbol
    rtn_strs.append(query_code[:2]) # 3 股票市场 - sz 或者 sh, 港股为hk，就是查询代码的前两位
    rtn_strs.append(fields[1])      # 4 股票名称 - StockName
    rtn_strs.append(fields[3])      # 5 当前价格 - StockPrice
    rtn_strs.append(fields[31])     # 6 涨跌 - Change
    rtn_strs.append(fields[32])     # 7 涨跌幅 - ChangePercentage
    
    # 港股与A股不同的地方
    if rtn_strs[2] == 'hk':
        # 如果是港股
        rtn_strs.append(fields[57]) # 8 市盈率 - PE
        rtn_strs.append(fields[58]) # 9 市净率 - PB
        rtn_strs.append(fields[45]) # 总市值 - TMV，单位是亿港币
    else:
        # 如果是A股，包括sh和sz
        rtn_strs.append(fields[39]) # 8 市盈率 - PE
        rtn_strs.append(fields[46]) # 9 市净率 - PB
        rtn_strs.append(fields[45]) # 总市值 - TMV，单位是亿人民币

    return rtn_strs

def parse_qq_fund_quote(query_code, quote_str):
    '''
    分解基金行情数据
    query_code: 查询代码
    quote_str: 不带查询代码的行情数据字符串
    '''
    rtn_strs = []
    rtn_strs.append(query_code) # 1 查询代码 - QueryCode
    
    # 分隔行情数据字符串，波浪号"~"为分隔符
    fields = quote_str.split('~') 
    
    rtn_strs.append(fields[0])      # 2 基金代码 - Symbol
    rtn_strs.append(query_code[:4]) # 3 市场 - s_jj，查询代码前4位
    rtn_strs.append(fields[1])      # 4 基金名称 - Name
    rtn_strs.append(fields[3])      # 5 净值 - Price
    rtn_strs.append(fields[4])      # 6 累计净值
    rtn_strs.append(fields[2])      # 7 净值日期
    
    return rtn_strs

def store_quotes(quote_list):
    '''
    将分解出来的行情数据存储到文件中
    '''
    if not quote_list:
        #如果传入的list是空的，直接返回
        return
    
    # 组装行情数据文件的全路径
    file_name = os.path.join(QUOTES_FILE_PATH, QUOTES_FILE_NAME)
    
    # 使用with，Python会自动地调用close()，不用再写了
    # 使用系统缺省的编码方式，see locale.getpreferredencoding()
    with open(file_name, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile, delimiter='~')
        writer.writerows(quote_list)
        
    

# 在DEBUG状态下，才打印信息
def debug_print(*kwargs):
    if DEBUG:
        print(*kwargs)
        

def main():
    if not check_env():
        print('环境检查失败，运行中止！')
        return
         
    refresh_quotes()

if __name__ == '__main__':
    main()
    
