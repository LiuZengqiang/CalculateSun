import numpy as np
import math
import xlrd
import xlwt
import xlutils
from xlutils.copy import copy
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D

fig = plt.figure(1)
ax = fig.gca( projection='3d')

filename="sun.xls"
data = xlrd.open_workbook(filename)
table = data.sheets()[0] # 打开第一张表
nrows = table.nrows # 获取表的行数
ncols = table.ncols

station=table.col_values(0)[1:]
year=table.col_values(1)[1:]
month=table.col_values(2)[1:]
day=table.col_values(3)[1:]
hour=table.col_values(4)[1:]
min=table.col_values(5)[1:]
sec=table.col_values(6)[1:]
lon=table.col_values(7)[1:]
lat =table.col_values(8)[1:]
TimeZone =table.col_values(9)[1:]

wb = xlutils.copy.copy(data)
ws = wb.get_sheet(0)

ws.write(0, 11, 'Day of Year')
ws.write(0, 12, 'Local Time') #实际是gtdt
ws.write(0, 13, 'Sun Angle')#(sitar)
ws.write(0, 14, 'Declination Angle')
ws.write(0, 15, 'Equation of Time')
style = xlwt.easyxf('pattern: pattern solid, fore_color yellow;')
ws.write(0, 16, 'ZenithAngle(deg)',style)
ws.write(0, 17, 'HeightAngle(deg)',style)
ws.write(0, 18, 'AzimuthAngle(deg)',style)


altitude= np.array([])
azimuth = np.array([])
x = np.array([])
y = np.array([])
z = np.array([])

for n in range(1,nrows):
    m=n-1
#年积日的计算
    #儒略日 Julian day(由通用时转换到儒略日)
    JD0 = int(365.25*(year[m]-1))+int(30.6001*(1+13))+1+hour[m]/24+1720981.5

    if month[m]<=2:
        JD2 = int(365.25*(year[m]-1))+int(30.6001*(month[m]+13))+day[m]+hour[m]/24+1720981.5
    else:
        JD2 = int(365.25*year[m])+int(30.6001*(month[m]+1))+day[m]+hour[m]/24+1720981.5

    #年积日 Day of year
    DOY = JD2-JD0+1

#N0   sitar=θ
    N0 = 79.6764 + 0.2422*(year[m]-1985) - int((year[m]-1985)/4.0)
    sitar = 2*math.pi*(DOY-N0)/365.2422
    ED1 = 0.3723 + 23.2567*math.sin(sitar) + 0.1149*math.sin(2*sitar) - 0.1712*math.sin(3*sitar)- 0.758*math.cos(sitar) + 0.3656*math.cos(2*sitar) + 0.0201*math.cos(3*sitar)
    ED = ED1*math.pi/180           #ED本身有符号

    if lon[m] >= 0:
        if TimeZone == -13:
            dLon = lon[m] - (math.floor((lon[m]*10-75)/150)+1)*15.0
        else:
            dLon = lon[m] - TimeZone[m]*15.0   #地球上某一点与其所在时区中心的经度差
    else:
        if TimeZone[m] == -13:
            dLon =  (math.floor((lon[m]*10-75)/150)+1)*15.0- lon[m]
        else:
            dLon =  TimeZone[m]*15.0- lon[m]
    #时差
    Et = 0.0028 - 1.9857*math.sin(sitar) + 9.9059*math.sin(2*sitar) - 7.0924*math.cos(sitar)- 0.6882*math.cos(2*sitar)
    gtdt1 = hour[m] + min[m]/60.0 + sec[m]/3600.0 + dLon/15        #地方时
    gtdt = gtdt1 + Et/60.0
    dTimeAngle1 = 15.0*(gtdt-12)
    dTimeAngle = dTimeAngle1*math.pi/180
    latitudeArc = lat[m]*math.pi/180

# 高度角计算公式
    HeightAngleArc = math.asin(math.sin(latitudeArc)*math.sin(ED)+math.cos(latitudeArc)*math.cos(ED)*math.cos(dTimeAngle))
# 方位角计算公式
    CosAzimuthAngle = (math.sin(HeightAngleArc)*math.sin(latitudeArc)-math.sin(ED))/math.cos(HeightAngleArc)/math.cos(latitudeArc)
    AzimuthAngleArc = math.acos(CosAzimuthAngle)
    HeightAngle = HeightAngleArc*180/math.pi
    ZenithAngle = 90-HeightAngle
    AzimuthAngle1 = AzimuthAngleArc *180/math.pi

    if dTimeAngle < 0:
        AzimuthAngle = 180 - AzimuthAngle1
    else:
        AzimuthAngle = 180 + AzimuthAngle1
    altitude=np.append(altitude,HeightAngle)
    azimuth = np.append(azimuth,AzimuthAngle)
    aa = np.cos(HeightAngle / 180 * np.pi) * np.cos(AzimuthAngle / 180 * np.pi)
    cc = np.sin(HeightAngle / 180 * np.pi)
    bb = np.cos(HeightAngle/ 180 * np.pi) * np.sin(AzimuthAngle / 180 * np.pi)
    ll = aa * aa + bb * bb + cc * cc
    ll = np.sqrt(ll)
    aa = aa / ll
    bb = bb / ll
    cc = cc / ll
    x=np.append(x,aa)
    y=np.append(y,bb)
    z=np.append(z,cc)
    print('站位：'+station[m]+' 时间：%d:%d 高度角(deg)：%f 方位角(deg)：%f ' % (hour[m],min[m],HeightAngle,AzimuthAngle))

ax.plot(x,y,z)
#ax.legend()
plt.xlabel("x")
plt.ylabel("y")
plt.show()