# 要添加一个新单元，输入 '# %%'
# 要添加一个新的标记单元，输入 '# %% [markdown]'
# %% [markdown]
# # 1.读取数据源（警报统计EXCEL）

# %%
from datetime import datetime, date, timedelta
import datetime
import pandas as pd  
import time
from pyecharts.globals import CurrentConfig


CurrentConfig.ONLINE_HOST = "./templates/static/js/"

# %% [markdown]
# ### 1.1获取相关数据

# %%
def ReadData(nub):
    file = eval((date.today() + timedelta(days = -1)).strftime("%Y%m"))
    filepath = '//192.168.218.243/PIMS_report/SE report/AlarmData/Month/CSV_'
    file = filepath+str(file)+'.csv'
    data = pd.read_csv(file)
    csg = list(data.iloc[:,nub])
    tm = list(data.iloc[:,0])
    tm = list(map(str, tm))
    plantname = data.iloc[:,nub].name
    return tm,csg,plantname

# %% [markdown]
# # 2.绘制看板图

# %%
from pyecharts import options as opts
from pyecharts.charts import Bar, Grid, Line, Liquid, Page, Pie,Timeline
from pyecharts.commons.utils import JsCode
from pyecharts.components import Table
import random
from pyecharts.globals import ThemeType
from pyecharts.components import Image
from pyecharts.options import ComponentTitleOpts

# %% [markdown]
# ### 2.1封装条形图

# %%
def bar_datazoom_slider(x,y) -> Bar:
    c = (
        
        Bar(init_opts=opts.InitOpts(bg_color='rgba(255,250,205,0.2)',
                                    width = '390px',
                                     height = '250px',
                                    #  js_host="./js/",
                                     theme=ThemeType.CHALK))
        .add_xaxis(x)
        .add_yaxis("全厂总报警数量", y)
        .set_global_opts(
            title_opts=opts.TitleOpts(title="全厂总报警数量趋势图"),
            datazoom_opts=[opts.DataZoomOpts()],
            legend_opts=opts.LegendOpts(type_="scroll", pos_left="right", orient="vertical")
        )
    )
    return c

# %% [markdown]
# ### 2.2封装折线图

# %%
def line_markpoint(plant,xdata,ydata) -> Line:
    title = plant + "警报数量趋势图"
#     themelist = ["ThemeType.LIGHT",
#              'ThemeType.DARK',
#              'ThemeType.CHALK','ThemeType.ESSOS','ThemeType.INFOGRAPHIC',
#              'ThemeType.MACARONS','ThemeType.PURPLE_PASSION','ThemeType.ROMA',
#             'ThemeType.ROMANTIC','ThemeType.VINTAGE',
#             'ThemeType.WALDEN','ThemeType.WESTEROS','ThemeType.WONDERLAND']
#     a1 = themelist[random.randint(0,len(themelist))]
#     print(ydata)
    c = (
        Line(init_opts=opts.InitOpts(bg_color='rgba(255,250,205,0.2)',
                                    width = '390px',
                                     height = '250px',
                                    #  js_host="./js/",
                                     theme =ThemeType.DARK
                                    #theme=ThemeType.MACARONS括号内可改变主题
                                    )) 
        
        .add_xaxis(xdata)
        .add_yaxis(
            plant,
            ydata,
            markpoint_opts=opts.MarkPointOpts(data=[opts.MarkLineItem(type_="min", name="最小值"),
                opts.MarkLineItem(type_="max", name="最大值"),
                ]),
        )
        .set_global_opts(title_opts=opts.TitleOpts(title=title),
#                          datazoom_opts=opts.DataZoomOpts(),#可调节横坐标
                        legend_opts=opts.LegendOpts(type_="scroll", pos_left="right", orient="vertical"))
        .set_series_opts(
        label_opts=opts.LabelOpts(is_show=False),
        markline_opts=opts.MarkLineOpts(
            data=[
                opts.MarkLineItem(type_="average", name="平均值"),
            ]
        ),
    )

    )
    return c

# %% [markdown]
# ### 2.3封装表格

# %%
def table_base() -> Table:
    table = Table()
    rl = GetRange()
    list1 = list(ReadDataCsv(rl[0]).iloc[:,0])
    list2 = list(ReadDataCsv(rl[0]).iloc[:,7])
    list1, list2 = zip(*sorted(zip(list2, list1),reverse = True))
    headers = ["排名", "部门", "数量"]
    rows = []
    for i in range(3):
        rows.append([i+1,list2[i],list1[i]])
    table.add(headers, rows).set_global_opts(
        title_opts=opts.ComponentTitleOpts(title="前一日警报数量排行")
    )
    return table

# %% [markdown]
# ### 2.4封装布局

# %%
def page_simple_layout():
#    page = Page()   默认布局
    page = Page(layout=Page.SimplePageLayout)    # 简单布局
#     page = Page(layout=Page.DraggablePageLayout)   #可移动布局
    # 将上面定义好的图添加到 page
    for i in range(1,18):
        TmData,AlarmData,plantname = ReadData(i)
#         TmData,AlarmDataAO,plantname = ReadData(file,stname,3)
        page.add(
            
            line_markpoint(plantname,TmData,AlarmData),
          
        )
#     x,y,z= ReadData(18)
    page.add(
        MoveLine()
    )
#     page.add(
#         table_base()
#     )
    page.render("./templates/index1.0.html")

# %% [markdown]
# ### 2.5封装图片

# %%
def Image_base() -> Image:
    image = Image()

    img_src = (
            "ccp.png"
    )
    image.add(
        src=img_src,
        style_opts={"width": "80px", "height": "60px", "style": "margin-top: 0px"},
    )
    return image

# %% [markdown]
# ### 2.6封装动态图

# %%
def get_year_overlap_chart(year: int,total_data) -> Bar:
    x = eval((date.today() + timedelta(days = -1)).strftime("%Y%m%d"))
    y = eval((date.today() + timedelta(days = -1)).strftime("%Y%m")+"00")
    rl = GetRange(x-y)
    name_list = MakeData(0)[rl[0]]
    bar = (
        Bar(init_opts=opts.InitOpts(bg_color='rgba(255,255,255,0.2)',
                                    width = '1200px',
                                     height = '800px',
                                    #  js_host="./js/",
                                     theme=ThemeType.LIGHT))
        .add_xaxis(xaxis_data=name_list)
        .add_yaxis(
            series_name="Alarm",
            y_axis=total_data["dataalr"][year],
#             is_selected=False,
            label_opts=opts.LabelOpts(is_show=True),
        )
        .add_yaxis(
            series_name="HH",
            y_axis=total_data["datahh"][year],
#             is_selected=False,
            label_opts=opts.LabelOpts(is_show=False),
        )
        .add_yaxis(
            series_name="IOP",
            y_axis=total_data["dataiop"][year],
#             is_selected=False,
            label_opts=opts.LabelOpts(is_show=False),
        )
        .add_yaxis(
            series_name="IOP-",
            y_axis=total_data["dataiop2"][year],
            label_opts=opts.LabelOpts(is_show=False),
        )
        .add_yaxis(
            series_name="LL",
            y_axis=total_data["datall"][year],
            label_opts=opts.LabelOpts(is_show=False),
        )
        .add_yaxis(
            series_name="TRIP",
            y_axis=total_data["datatrip"][year],
            label_opts=opts.LabelOpts(is_show=False),
        )
        .set_global_opts(
            title_opts=opts.TitleOpts(
                title="{}各部门每日警报情况".format(year), subtitle="数据来自PIMS"
            ),
            tooltip_opts=opts.TooltipOpts(
                is_show=True, trigger="axis", axis_pointer_type="shadow"
            ),
        )
    )
    pie = (
        Pie()
        .add(
            series_name="警报占比",
            data_pair=[
                ["ALM", total_data["dataalr"]["{}sum".format(year)]],
                ["HH", total_data["datahh"]["{}sum".format(year)]],
                ["LL", total_data["datall"]["{}sum".format(year)]],
                ["IOP", total_data["dataiop"]["{}sum".format(year)]],
                ["IOP-", total_data["dataiop2"]["{}sum".format(year)]],
                ["TRIP", total_data["datatrip"]["{}sum".format(year)]]
            ],
            center=["90%", "20%"],
            radius="28%",
        )
        .set_series_opts(tooltip_opts=opts.TooltipOpts(is_show=True, trigger="item"))
    )
    return bar.overlap(pie)


# %%
def GetRange(n):
    tmlist=[]
    for i in range(n,0,-1):
        tmlist.append(eval((date.today() + timedelta(days = -i)).strftime("%Y%m%d")))
    return tmlist  


# %%
def ReadDataCsv(file):
    filepath = '//192.168.218.243/PIMS_report/SE report/AlarmData/Day/CSV_'
    file = filepath+str(file)+'.csv'
    data = pd.read_csv(file)
    return data

# %% [markdown]
# 数据构造

# %%
def MakeData(col):
    x = eval((date.today() + timedelta(days = -1)).strftime("%Y%m%d"))
    y = eval((date.today() + timedelta(days = -1)).strftime("%Y%m")+"00")
    rl = GetRange(x-y)
    data_alr = {}
    for i in range(x-y):
        data_alr.update({rl[i]: list(ReadDataCsv(rl[i]).iloc[:,col])})
    return data_alr


# %%
def format_data(data: dict) -> dict:
    x = eval((date.today() + timedelta(days = -1)).strftime("%Y%m%d"))
    y = eval((date.today() + timedelta(days = -1)).strftime("%Y%m")+"00")
    rangelist = GetRange(x-y)
    name_list = MakeData(0)[rangelist[0]]
    for year in rangelist:
        max_data, sum_data = 0, 0
        temp = data[year]
        max_data = max(temp)
        for i in range(len(temp)):
            sum_data += temp[i]
            data[year][i] = {"name": name_list[i], "value": temp[i]}
        data[str(year) + "max"] = int(max_data / 100) * 100
        data[str(year) + "sum"] = sum_data
    return data


# %%
def MoveLine():
    total_data = {}
    x = eval((date.today() + timedelta(days = -1)).strftime("%Y%m%d"))
    y = eval((date.today() + timedelta(days = -1)).strftime("%Y%m")+"00")
    rl = GetRange(x-y)
    name_list = MakeData(0)[rl[0]] #x轴
    #y轴
    data_alr = MakeData(1)
    data_hh = MakeData(2)
    data_iop = MakeData(3)
    data_iop2 = MakeData(4)
    data_ll = MakeData(5)
    data_trip = MakeData(6)
    total_data["dataalr"] = format_data(data=data_alr)
    # hh
    total_data["datahh"] = format_data(data=data_hh)
    # iop
    total_data["dataiop"] = format_data(data=data_iop)
    # iop2
    total_data["dataiop2"] = format_data(data=data_iop2)
    # ll
    total_data["datall"] = format_data(data=data_ll)
    # trip
    total_data["datatrip"] = format_data(data=data_trip)

    # 生成时间轴的图
    timeline = Timeline(init_opts=opts.InitOpts(width="1200px", height="600px",theme=ThemeType.DARK))
    Rangelist = GetRange(x-y)
    for y in Rangelist:
        timeline.add(get_year_overlap_chart(y,total_data), time_point=str(y))
    timeline.add_schema(is_auto_play=True, play_interval=1000)
    return timeline

# %% [markdown]
# # 3.主函数

# %%
if __name__ == "__main__":
    page_simple_layout()

# %% [markdown]
# # 4.HTML布置

# %%
from bs4 import BeautifulSoup


# %%
with open("./templates/index1.0.html", "r+", encoding='utf-8') as html:
    html_bf = BeautifulSoup(html, 'lxml')
    title=html_bf.find('title')
    title.string.replace_with("CCJS Alarm Kanban")
    divs = html_bf.select('.chart-container')
    meta=html_bf.find('meta')
    meta['name']="viewport"
    meta['content']="width=device-width, initial-scale=1"
    for i in range(17):
        wid = 390
        hei = 250
        left =(i % 4)*400
        top = (i//4)*300+800
        divs[i]["style"] = "width:%dpx;height:%dpx;position:absolute;top:%dpx;left:%dpx;border-style:solid;border-color:black;border-width:0px;" % (wid,hei,top,left)
#     divs[1]["style"] = "width:%dpx;height:%dpx;position:absolute;top:%dpx;left:%dpx;border-style:solid;border-color:black;border-width:0px;" % (390,250,5,400)
#     divs[2]["style"] = "width:390px;height:250px;position:absolute;top:5px;left:800px;border-style:solid;border-color:black;border-width:0px;"
    divs[17]["style"] = "width:1200px;height:600px;position:absolute;top:100px;left:200px;border-style:solid;border-color:black;border-width:0px;"
#     divs[18]["style"] = "width:400px;height:300px;position:absolute;top:50px;left:0px;border-style:solid;border-color:white;border-width:0px;"
    
    body = html_bf.find("body")
    body["style"] = "background-color:black;"
    div_title="<div align=\"center\" style=\"width:100%;\">\n<span style=\"font-size:150%;font face=\'微软雅黑\';color:#FFFF00\"><h3>CCJS 各制程警报管理数据可视化看板</h3></div>"  
    #修改页面背景色、追加标题
    body.insert(0,BeautifulSoup(div_title,"lxml").div)
    html_new = str(html_bf)
    html.seek(0, 0)
    html.truncate()
    html.write(html_new)
    html.close()


