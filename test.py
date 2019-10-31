# from xlrd import open_workbook
# from xlutils.copy import copy
#
# rb = open_workbook("C:\\Users\\Administrator\\Desktop\\需要调整表格\\需要调整表格\\1.地表.xls", formatting_info=True, on_demand=True)
# wb = copy(rb)
# table=wb.get_sheet(0)
# table.write(1,0,'changed!')
# wb.save('output.xls')


from pyecharts.charts import Line
from pyecharts.render import make_snapshot
from pyecharts import options as opts

# 使用 snapshot-selenium 渲染图片
from snapshot_selenium import snapshot

bar = (
    Line()
        .add_xaxis(['2019-04-20', '2019-04-21', '2019-04-22', '2019-04-23', '2019-04-24','2019-04-20', '2019-04-21', '2019-04-22', '2019-04-23', '2019-04-24'])
        .add_yaxis("JCD6-2", [1.2,3.4,2.5,0.5,6.5,1.2,3.4,2.5,0.5,6.5])
        .set_global_opts(
            title_opts=opts.TitleOpts(title="Line-数值 X 轴")
        )

    )

make_snapshot(snapshot, bar.render(), "line.png")