# geomagic igs解析
用于geomagic导出特征的igs\iges解析.
做种植体测量的时候想导出特征数据,但没有能直接读的格式.
网上搜了一圈也没找到现成的轮子,于是自己整了一个.
写得很丑陋,但能用.


食用方法:

0.python及openpyxml库

1.将temp.xlsx igs.py 及待解igs析文件放在同一文件夹

2.运行igs.py

3.结果保存在temp2.xlsx


注意:

1.igs文件在导出[点]或[线]或[点+线]时与导出[点+线+面]的结构不同.
本py仅适配了[点+线+面]的情况,请保证导出的igs包含[点线面]三种特征,或修改re规则

2.导出的igs文件中[点]和[线]的顺序似乎和geo程序中相反,请注意
