# coding:utf-8
import os
import datetime
import pandas as pd
import openpyxl
import numpy as np
from decimal import Decimal

# 二维数组转一维数组
def toSingleList(data):
  return [i[0] for i in data]

# 两个list拼成字典
def getDict(list1, list2):
  return dict(map(lambda x,y:[x,y],list1,list2))

# 创建文件夹
def mkdir(path):
    # 去除首位空格
    path = path.strip()
    # 去除尾部 \ 符号
    path = path.rstrip("\\")
    # 判断路径是否存在
    isExists = os.path.exists(path)

    # 判断结果
    if not isExists:
      # 如果不存在则创建目录
      # 创建目录操作函数
      os.makedirs(path)
      print(path+' 创建成功')
      return True
    else:
      # 如果目录存在则不创建，并提示目录已存在
      print(path+' 目录已存在')
      return False

def my_sum(lst):
  sum_res = 0
  for item in lst:
    if isinstance(item,(float,int)):
      sum_res = round(Decimal(sum_res),2) + round(Decimal(item),2)
      # sum_res = sum_res.round(2) + item.round(2)
  return round(sum_res,2)



start_time = datetime.datetime.now()
# 创建存储目录
mkdir('./result');

pd.set_option('display.precision', 2)
# 读取数据文件
# df = pd.read_excel('./test.xlsx')
df = pd.read_excel('./source.xlsx')

# 构建透视表
df_filtered = df.loc[df['所属K框'].str.contains('K3'),:]
# df_filtered['累量消耗'] = df_filtered['累量消耗'].apply(lambda x: Decimal(str(x)))
df_filtered['累量消耗'] = df_filtered['累量消耗'].round(2)
# df_filtered['累量消耗'] = df_filtered['累量消耗'].astype('float16')
df_filtered['日期'] = df_filtered['日期'].apply(lambda x:x.strftime('%Y年%m月%d日'))
pivot_table = pd.pivot_table(df_filtered, values='累量消耗', index=['一级事业部'], columns=['日期'], aggfunc='sum', margins=1, margins_name='总计')

sortList = ['家居事业部', '金融事业部', '北区事业部', '集团渠道', '在线效果事业部', '南区事业部', '总计']
cat_type = pd.api.types.CategoricalDtype(categories=sortList, ordered=True)
pivot_table.index = pivot_table.index.astype(cat_type)

# 按照索引排序
pivot_table = pivot_table.sort_index()
# 倒数第二列的值
lastOriginData = pivot_table.iloc[:, [-2]]
# 最后一列的值
totalOriginData = pivot_table.iloc[:, [-1]]

# 行名
index = pivot_table.index.tolist()
# 行名
columns = pivot_table.columns.tolist();
# 转成一维数组
before2DayData = toSingleList(pivot_table.iloc[:, [-3]].values);
lastDayData = toSingleList(lastOriginData.values);
totalData = toSingleList(totalOriginData.values);
print('k框数据===>')
print(columns[-3], before2DayData)
print(columns[-2], lastDayData)
print(columns[-1], totalData)

print('客户数据===>')
# 构建客户的透视表
customer_table = pd.pivot_table(df_filtered, values='累量消耗', index=['集团简称'], columns=['日期'], aggfunc=np.sum, margins=1, margins_name='总计')
customer_index = customer_table.index.tolist();
customer_columns = customer_table.columns.tolist();
before2CustomerDic = getDict(customer_index, toSingleList(customer_table.iloc[:, [-3]].values))
lastCustomerDic = getDict(customer_index, toSingleList(customer_table.iloc[:, [-2]].values))
totalCustomerDic = getDict(customer_index, toSingleList(customer_table.iloc[:, [-1]].values))
print(lastCustomerDic)
print(totalCustomerDic)

updateDate = datetime.datetime.strptime(columns[-2], '%Y年%m月%d日');
# 将透视表保存到文件
pivotTableName = './result/' + updateDate.strftime('%Y%m%d') + '__' + 'pivotTable.xlsx'
p_writer = pd.ExcelWriter(pivotTableName)
pivot_table.to_excel(p_writer, sheet_name='Sheet1')
customer_table.to_excel(p_writer, sheet_name='Sheet2')
# p_writer.save()

# 消耗统计写入日报
# 调试数据
# columns = ['2023年09月15日', '2023年09月16日', '总计']
# lastDayData = [616039.27, 318476.63, 293637.12, 387793.85000000003, 787179.45, 233916.26, 2637042.58]
# totalData = [40854438.59, 30487650.33, 25065261.43, 24326936.53, 60277616.25, 11786265.57, 192798168.7]
# lastCustomerDic = {'TATA木门': 10617.539999999999, '东方教育': 2873.1, '东易日盛': 6222.99, '东高': 33644.9, '东鹏': 25312.399999999998, '中信证券': float('nan'), '中信银行': 13732.820000000002, '中职通': 15144.28, '中诺口腔': float('nan'), '丰田': float('nan'), '九方云': 3793.42, '亚特兰蒂斯': float('nan'), '伴鱼英语': 3672.63, '全友家居': 197389.12, '冈本': 10036.060000000001, '千聊': float('nan'), '博洛尼': 45102.88, '去哪儿': 50734.439999999995, '可啦啦': 2789.22, '同程旅游': 2013.8, '和讯': 98015.29, '国诚投资': 21688.489999999998, '圣诺游艇': float('nan'), '小叶子': 1218.71, '尚层': 2204.3700000000003, '尚层装饰': 39069.72, '居然之家': 5539.12, '巨丰投资': 211590.81, '平安银行': 19401.19, '广发证券': 209.56, '广州秒可': float('nan'), '建信住房服务': float('nan'), '微淼财商': 8115.709999999999, '快财': 520399.71, '慕思寝具': 66580.51, '拉尔森': float('nan'), '招商银行': 16516.56, '指南针': 21297.530000000002, '新诗懿': float('nan'), '无忧贷款': float('nan'), '星火保': 77201.15, '春满欢苗': 8.43, '曲美家居': float('nan'), '有书': 111258.27, '林氏木业': 25099.11, '欧派家居': 57992.67, '民生银行': float('nan'), '汤臣倍健': float('nan'), '泰康财险': 10499.92, '浦发银行': 3214.5, '深爱居': 15565.35, '爱邦保险': 5428.64, '猿辅导': 154.41, '瑞达洲际': float('nan'), '生活家': 36107.39, '画啦啦': 11565.51, '百信银行': float('nan'), '百安居': 20394.85, '百度金融': float('nan'), '盘子女人坊': 34417.22, '瞳学贸易': float('nan'), '福仁康大药房': 43599.66, '福特电马': float('nan'), '立邦': 3802.12, '立邦涂料': float('nan'), '索菲亚家居': 25984.26, '红杉树': 2454.5299999999997, '红松': 21513.760000000002, '维京游轮': float('nan'), '网易有道': 246511.0, '聪明核桃': 76247.64, '自如': 2.33, '艺旗网络': 77308.13, '芊丝诺': 3391.37, '芝华仕家居': float('nan'), '莲姿娜': 33408.350000000006, '蓝城健康': float('nan'), '贝壳': 27.63, '赛益世': float('nan'), '超鸟': float('nan'), '跟谁学': float('nan'), '邓禄普': float('nan'), '金城银行': 40575.68, '金掌柜贷款': float('nan'), '金牌家居': 25525.48, '顶点财经': 41696.52, '顾家家居': 18874.73, '高途': 15.53, '高顿教育': 36016.52, '黑牛保': 59542.73, '齐家网': 16710.31, '总计': 2637042.58}
# totalCustomerDic = {'TATA木门': 693357.81, '东方教育': 47716.15, '东易日盛': 395542.26, '东高': 5948132.17, '东鹏': 1244113.0, '中信证券': 40966.2, '中信银行': 692159.63, '中职通': 1311526.45, '中诺口腔': 127452.51, '丰田': 6886.59, '九方云': 120598.33, '亚特兰蒂斯': 64463.869999999995, '伴鱼英语': 3683.47, '全友家居': 15024007.18, '冈本': 160455.94, '千聊': 48.61, '博洛尼': 2066132.49, '去哪儿': 4012818.53, '可啦啦': 202370.01, '同程旅游': 6713.96, '和讯': 5361476.01, '国诚投资': 302371.94, '圣诺游艇': 5000.01, '小叶子': 400530.75, '尚层': 65976.5, '尚层装饰': 3255328.03, '居然之家': 5841314.26, '巨丰投资': 8745514.28, '平安银行': 487998.21, '广发证券': 27094.86, '广州秒可': 0.11, '建信住房服务': 346033.86, '微淼财商': 1097712.8699999999, '快财': 44770817.44, '慕思寝具': 534438.21, '拉尔森': 92.98, '招商银行': 1382793.17, '指南针': 765818.1, '新诗懿': 2944.79, '无忧贷款': 523797.55, '星火保': 11558814.22, '春满欢苗': 1000.64, '曲美家居': 11760.619999999999, '有书': 7487830.16, '林氏木业': 2452267.0, '欧派家居': 4763250.59, '民生银行': 12112.0, '汤臣倍健': 4621.04, '泰康财险': 31734.22, '浦发银行': 573665.07, '深爱居': 190249.57, '爱邦保险': 238322.06, '猿辅导': 154.41, '瑞达洲际': 1250.1899999999998, '生活家': 1688582.31, '画啦啦': 578533.44, '百信银行': 66043.91, '百安居': 705520.27, '百度金融': 78861.11, '盘子女人坊': 1602542.63, '瞳学贸易': 121925.07, '福仁康大药房': 1550701.19, '福特电马': 197489.17, '立邦': 295523.2, '立邦涂料': 67773.64, '索菲亚家居': 200155.56, '红杉树': 623410.29, '红松': 3159320.25, '维京游轮': 77442.21, '网易有道': 13438917.44, '聪明核桃': 3730065.46, '自如': 2848129.01, '艺旗网络': 4060282.4899999998, '芊丝诺': 46803.33, '芝华仕家居': 1226489.31, '莲姿娜': 190060.93, '蓝城健康': 17053.51, '贝壳': 3857187.0100000002, '赛益世': 1288.14, '超鸟': 1362193.3699999999, '跟谁学': 114.0, '邓禄普': 20000.0, '金城银行': 168679.84, '金掌柜贷款': 32955.96, '金牌家居': 224887.84, '顶点财经': 2748442.49, '顾家家居': 4911737.24, '高途': 15.53, '高顿教育': 3327468.57, '黑牛保': 5210131.34, '齐家网': 948210.76, '总计': 192798168.7}
temFile = './K3日报09-07.xlsx' #模板文件
savePath = './result/k3日报'+updateDate.strftime('%Y%m%d')+'.xlsx'

# 读取模板文件
workbook = openpyxl.load_workbook(temFile)
# 选择工作表
worksheet = workbook.active
# 修改数据
worksheet['B2'] = updateDate.strftime('%Y/%m/%d')
# worksheet['G4'] = worksheet['F4'].value #前日总消耗
worksheet['G4'] = before2DayData[-1] #前日总消耗
worksheet['F4'] = lastDayData[-1] #昨日总消耗
worksheet['C4'] = totalData[-1] #总消耗

# 昨日消耗
worksheet['D9'] = lastDayData[0]
worksheet['D10'] = lastDayData[1]
worksheet['D11'] = lastDayData[2]
worksheet['D12'] = lastDayData[3]
worksheet['D13'] = lastDayData[4]
worksheet['D14'] = lastDayData[5]
# 累计消耗
worksheet['C9'] = totalData[0]
worksheet['C10'] = totalData[1]
worksheet['C11'] = totalData[2]
worksheet['C12'] = totalData[3]
worksheet['C13'] = totalData[4]
worksheet['C14'] = totalData[5]

def notNan(val):
  return val == val
# 销售数据
b_column_data = [(cell.value, str(cell.row)) for cell in worksheet['B']]
for it in b_column_data[17:]:
  key = it[0]
  if it[0] == None or it[0] == '客户名称' or it[0] == '-':
    continue;
  else:
    # 前日数据copy
    # worksheet['E' + it[1]] = worksheet['D' + it[1]].value
    worksheet['C' + it[1]] = totalCustomerDic.get(key) if notNan(totalCustomerDic.get(key)) else '-'
    worksheet['D' + it[1]] = lastCustomerDic.get(key) if notNan(totalCustomerDic.get(key)) else '-'
    worksheet['E' + it[1]] = before2CustomerDic.get(key) if notNan(before2CustomerDic.get(key)) else '-'


# 保存新文件
workbook.save(savePath)
end_time = datetime.datetime.now()
print("耗时: {}秒".format(end_time - start_time))
