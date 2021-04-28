# %%import
import datetime
import os
import time

import pandas as pd
import win32com.client

# %%
# 创建数据存储路径
target_excel_dir = r'D:/PythonProject/STK'
if not os.path.exists(target_excel_dir):
    os.makedirs(target_excel_dir)

# %%run stk
# Get reference to running STK instance
app = win32com.client.Dispatch('STK11.Application')
app.Visible = True
# Get our IAgStkObjectRoot interface
root = app.Personality2
# creat a new scenario
# IAgStkObjectRoot root: STK Object Model Root
root.NewScenario('Beidou3G2')  # 场景名称不能有空格等非法字符
scenario = root.CurrentScenario
root.UnitPreferences.SetCurrentUnit('DateFormat', 'UTCG')
# 设置场景时间
root.CurrentScenario.StartTime = '8 Jun 2020 16:00:00.00'
root.CurrentScenario.StopTime = '10 Jun 2020 04:00:00.00'
# %%定义函数创建卫星并导出数据
# path_tle:D:/xxx/xxx.txt存储目标的tle文件夹


def createSatellite(path_tle, starttime, stoptime):
    # 创建卫星
    # IAgStkObjectRoot root: STK Object Model Root
    satellite = root.CurrentScenario.Children.New(
        18, 'Beidou3G2')  # eSatellite
    # IAgSatellite satellite: Satellite object
    satellite.SetPropagatorType(4)  # ePropagatorSGP4
    propagator = satellite.Propagator
    propagator.UseScenarioAnalysisTime
    # 用tle根数生成卫星，注意：
    # 1、该函数有两个输入参数皆为字符串，'NORADID','根数文件路径'
    # 2、根数文件里同一目标不能有相同历元根数
    # propagator.CommonTasks.AddSegsFromFile('ID','path.txt'),txt中要有该卫星的根数
    propagator.CommonTasks.AddSegsFromFile(
        '45344', path_tle)
    propagator.AutoUpdateEnabled = True
    propagator.Propagate()

    # 通过Dataprovider接口计算卫星参数
    # 推荐使用这些方式连接接口
    # There are 4 Methods to get DP From a Path depending on the kind of DP:
    #   GetDataPrvTimeVarFromPath
    #   GetDataPrvIntervalFromPath
    #   GetDataPrvInfoFromPath
    #   GetDataPrvFixeDataromPath
    # Orbit元素：半长轴等
    satOrbitDP = satellite.DataProviders.GetDataPrvTimeVarFromPath(
        'Classical Elements/J2000')
    results_orbit = satOrbitDP.Exec(
        starttime, stoptime, 300)
    # (t1,t2,时间步长)t1,t2都必须是符合STK格式的时间字符串：'1 Jan 2020 04:00:00'
    # LLA State 经纬度轨迹
    satLLADP = satellite.DataProviders.GetDataPrvTimeVarFromPath(
        'LLA State/Fixed')
    results_LLA = satLLADP.Exec(
        starttime, stoptime, 300)
    # 组：Classical Elements，类：J2000，计算元素：Time x y z等等
    # 使用Exec，会计算类下所有元素的值，更多详情请参考STK帮助文档
    # GET DATA
    # 从DataSets中获取数据有两种方式：1、序号；2、元素名称。
    # satax = results_orbit.DataSets.GetDataSetByName('Semi-major Axis').GetValues()
    # [0]:'Time' [1]:'Semi-major Axis'
    sat_time = results_orbit.DataSets[0].GetValues()
    sat_sma = results_orbit.DataSets[1].GetValues()
    sat_inc = results_orbit.DataSets[3].GetValues()
    sat_lat = results_LLA.DataSets[1].GetValues()
    sat_lon = results_LLA.DataSets[2].GetValues()
    # create new dataframes to store these data
    Data = pd.DataFrame(
        columns=('Time (UTC)',
                 'Semi-major Axis (km)', 'Inclination (deg)',
                 'latitude (deg)', 'longitude (deg)'))
    Data_a = pd.DataFrame(columns=('Time', 'Semi-major Axis (km)'))
    # 给Data,Data_a赋值
    for j in range(0, len(sat_time)):
        t = sat_time[j].split('.', 1)[0]
        # 转换输出时间为时间戳方便筛选核计算数据
        # STK输出的时间格式样式：1 Jan 2020 04:00:00
        t_stamp = time.strptime(t, '%d %b %Y %H:%M:%S')
        t_stamp = time.mktime(t_stamp) / 3600
        # 转换输出时间为格式时间便于显示以及读取
        t = datetime.datetime.strptime(t, '%d %b %Y %H:%M:%S')
        t = datetime.datetime.strftime(t, '%Y-%m-%d %H:%M:%S')
        sma = sat_sma[j]
        inc = sat_inc[j]
        lat = sat_lat[j]
        lon = sat_lon[j]
        Data_a = Data_a.append(pd.DataFrame(
            {'Time': [t_stamp], 'Semi-major Axis (km)': [sma]}),
            ignore_index=True)
        print(Data_a.head())
        Data = Data.append(pd.DataFrame(
            {'Time (UTC)': [t],
             'Semi-major Axis (km)': [sma], 'Inclination (deg)': [inc],
             'latitude (deg)': [lat], 'longitude (deg)': [lon]}),
            ignore_index=True)
        print(Data.head())
    # 存储Data
    Data.to_excel(target_excel_dir + '/' + 'test.xlsx',
                  'sheet1', float_format='%.3f', index=False)
    Data_a.to_excel(target_excel_dir + '/' + 'test.xlsx',
                    'sheet1', float_format='%.3f', index=False)


# %%执行函数
path_tle = r'D:/PythonProject/STK/beidou3.txt'
t1 = scenario.StartTime
t2 = scenario.StopTime
# 你也可以自定义计算卫星参数的时间区间，但需要在场景时间范围内。
t3 = '9 Jun 2020 04:00:00.00'
t4 = '9 Jun 2020 16:00:00.00'
# 可以用root多次设置场景时间
createSatellite(path_tle,t1,t2)