# --*-- coding:utf-8 --*--

# Author: Taoyong
# Time: 2019/1/7 12:16

import pandas as pd
from pyecharts import Line
from datetime import date, timedelta
from docx import Document
from docx.shared import Inches

# 毕节、六枝曲线图
def plot_pic1(arr1, data1, d):
    line = Line(title='毕节、六枝逐3小时气温预报', title_pos='center', title_top='5%', height=600, width=1000)
    line.add('毕节', arr1, data1.loc['毕节'], is_label_show=True, is_smooth=True, line_width=2, mark_point=['max', 'min'],
             mark_point_symbolsize=25, mark_point_symbol='circle', line_color="red")
    line.add('六枝', arr1, data1.loc['六枝'], is_label_show=True, is_smooth=True, line_width=2, mark_point=['max', 'min'],
             mark_point_symbolsize=25, mark_point_symbol='circle', line_color="blue")
    line.add('4℃', arr1, data1.loc['特征值1'], line_type='dashed', line_color="orange")
    line.add('7℃', arr1, data1.loc['特征值2'], xaxis_name='时间', yaxis_name='温度', is_xaxislabel_align=True,
             is_splitline_show=False, line_type='dashed')
    line.add('11℃', arr1, data1.loc['特征值3'], xaxis_name='时间', yaxis_name='温度', is_xaxislabel_align=True,
             is_splitline_show=False, line_type='dashed')
    line.render(path=d+'毕节、六枝'+'.gif')
# 安顺、贵阳、都匀曲线图
def plot_pic2(arr2, data2, d):
    line = Line(title='安顺、贵阳、都匀逐3小时气温预报', title_pos='center', title_top='5%', height=600, width=1000)
    line.add('安顺', arr2, data2.loc['安顺'], is_label_show=True, is_smooth=True, line_width=2, mark_point=['max', 'min'],
             mark_point_symbolsize=25, mark_point_symbol='circle', line_color="red")
    line.add('贵阳', arr2, data2.loc['贵阳'], is_label_show=True, is_smooth=True, line_width=2, mark_point=['max', 'min'],
             mark_point_symbolsize=25, mark_point_symbol='circle', line_color="blue")
    line.add('都匀', arr2, data2.loc['都匀'], is_label_show=True, is_smooth=True, line_width=2, mark_point=['max', 'min'],
             mark_point_symbolsize=25, mark_point_symbol='circle', line_color="green")
    line.add('4℃', arr2, data2.loc['特征值1'], line_type='dashed', line_color="orange")
    line.add('7℃', arr2, data2.loc['特征值2'], xaxis_name='时间', yaxis_name='温度', is_xaxislabel_align=True,
             is_splitline_show=False, line_type='dashed')
    line.add('11℃', arr2, data2.loc['特征值3'], xaxis_name='时间', yaxis_name='温度', is_xaxislabel_align=True,
             is_splitline_show=False, line_type='dashed')
    line.render(path=d+'安顺、贵阳、都匀'+'.gif')
# 遵义、播州区曲线图
def plot_pic3(arr3, data3, d):
    line = Line(title='遵义、播州区逐3小时气温预报', title_pos='center', title_top='5%', height=600, width=1000)
    line.add('遵义', arr3, data3.loc['遵义'], is_label_show=True, is_smooth=True, line_width=2, mark_point=['max', 'min'],
             mark_point_symbolsize=25, mark_point_symbol='circle', line_color="red")
    line.add('播州区', arr3, data3.loc['播州区'], is_label_show=True, is_smooth=True, line_width=2, mark_point=['max', 'min'],
             mark_point_symbolsize=25, mark_point_symbol='circle', line_color="blue")
    line.add('4℃', arr3, data3.loc['特征值1'], line_type='dashed', line_color="orange")
    line.add('7℃', arr3, data3.loc['特征值2'], xaxis_name='时间', yaxis_name='温度', is_xaxislabel_align=True,
             is_splitline_show=False, line_type='dashed')
    line.add('11℃', arr3, data3.loc['特征值3'], xaxis_name='时间', yaxis_name='温度', is_xaxislabel_align=True,
             is_splitline_show=False, line_type='dashed')
    line.render(path=d+'遵义、播州区'+'.gif')
# 打开word模板，插入生成的图片，最后保存为最新的文档
def insert_word(t2):
    document = Document('贵州燃气公司专题服务模板.docx')
    # para = document.paragraphs
    # for p in para:
    #     print(p.text)
    # for p in range(len(document.paragraphs)):
    #     if p>3:
    #         document.paragraphs[p].clear()
    document.add_picture(t2 + '毕节、六枝.gif', width=Inches(7.25))
    document.add_picture(t2 + '安顺、贵阳、都匀.gif', width=Inches(7.25))
    document.add_picture(t2 + '遵义、播州区.gif', width=Inches(7.25))
    document.save('贵州燃气公司专题服务'+t2+'.docx')

def main():
    # 获取昨天、今天、明天日期
    # zt = (date.today() + timedelta(days=-1)).strftime("%Y%m%d")
    jt = date.today().strftime("%Y%m%d")
    tomorrow = (date.today() + timedelta(days=1)).strftime("%d")
    postnatal = (date.today() + timedelta(days=2)).strftime("%d")
    arr = ['08时', '11时', '14时', '17时', '20时', '23时', '02时', '05时', '08时']
    # 处理日期隔天
    for i in range(0, len(arr)):
        if i < 6:
            arr[i] = tomorrow + '日' + arr[i]
        else:
            arr[i] = postnatal + '日' + arr[i]
    # 读取数据
    data = pd.read_table('data_3h.txt', header=None, skiprows=1, encoding='utf-8', index_col=0)
    # print(data)

    plot_pic1(arr, data, jt)
    plot_pic2(arr, data, jt)
    plot_pic3(arr, data, jt)
    insert_word(jt)


if __name__ == "__main__":
    main()
