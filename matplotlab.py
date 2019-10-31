# -*- coding: UTF-8 -*-
import numpy as np
import matplotlib as mpl
import matplotlib.pyplot as plt
from pylab import mpl

mpl.rcParams['font.sans-serif'] = ['FangSong'] # 指定默认字体
mpl.rcParams['axes.unicode_minus'] = False # 解决保存图像是负号'-'显示为方块的问题


x_axix=['2019-04-20', '2019-04-21', '2019-04-22', '2019-04-23', '2019-04-24','2019-04-25', '2019-04-26', '2019-04-27', '2019-04-28', '2019-04-29']
#开始画图
# sub_axix = filter(lambda x:x%200 == 0, x_axix)
train_acys=[1.2,3.4,2.5,0.5,6.5,1.2,3.4,2.5,0.5,6.5]
train_acys1=[1.5,3.4,2.7,1.5,9.5,3.2,1.4,6.5,2.5,7.5]
plt.title('Result Analysis')
plt.plot(x_axix, train_acys, color='green', label='电线杆')
plt.plot(x_axix, train_acys1, color='red', label='井盖')

# plt.plot(sub_axix, test_acys, color='red', label='testing accuracy')
# plt.plot(x_axix, train_pn_dis,  color='skyblue', label='PN distance')
# plt.plot(x_axix, thresholds, color='blue', label='threshold')
plt.legend() # 显示图例

plt.xlabel('iteration times')
xticks=list(range(0,len(x_axix),2)) # 这里设置的是x轴点的位置
xlabels=[x_axix[x] for x in xticks] #这里设置X轴上的点对应那个totalseed中的值
xticks.append(9)
# xlabels.append(x_axix[len(x_axix)-1])

plt.xticks(xticks)
plt.ylabel('rate')
plt.savefig('./test2.jpg')
# plt.show()