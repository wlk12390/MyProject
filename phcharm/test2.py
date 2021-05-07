import re
import numpy
from numpy import *
import numpy as np
import time
from tqdm import tqdm

startTime = time.time()
from comtypes.gen import STKObjects, STKUtil, AgStkGatorLib
from comtypes.client import CreateObject, GetActiveObject, GetEvents, CoGetObject, ShowEvents
from ctypes import *
import comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0
from comtypes import GUID
from comtypes import helpstring
from comtypes import COMMETHOD
from comtypes import dispid
from ctypes.wintypes import VARIANT_BOOL
from ctypes import HRESULT
from comtypes import BSTR
from comtypes.automation import VARIANT
from comtypes.automation import _midlSAFEARRAY
from comtypes import CoClass
from comtypes import IUnknown
import comtypes.gen._00DD7BD4_53D5_4870_996B_8ADB8AF904FA_0_1_0
import comtypes.gen._8B49F426_4BF0_49F7_A59B_93961D83CB5D_0_1_0
from comtypes.automation import IDispatch
import comtypes.gen._42D2781B_8A06_4DB2_9969_72D6ABF01A72_0_1_0
from comtypes import DISPMETHOD, DISPPROPERTY, helpstring

"""
SET TO TRUE TO USE ENGINE, FALSE TO USE GUI
"""
useStkEngine = False
Read_Scenario = True

############################################################################
# Scenario Setup
############################################################################

if useStkEngine:
    # Launch STK Engine
    print("Launching STK Engine...")
    stkxApp = CreateObject("STKX11.Application")

    # Disable graphics. The NoGraphics property must be set to true before the root object is created.
    stkxApp.NoGraphics = True

    # Create root object
    stkRoot = CreateObject('AgStkObjects11.AgStkObjectRoot')

else:
    # Launch GUI
    print("Launching STK...")
    if not Read_Scenario:
        uiApp = CreateObject("STK11.Application")
    else:
        uiApp = GetActiveObject("STK11.Application")
    uiApp.Visible = True
    uiApp.UserControl = True

    # Get root object
    stkRoot = uiApp.Personality2

# Set date format
stkRoot.UnitPreferences.SetCurrentUnit("DateFormat", "UTCG")
# Create new scenario
print("Creating scenario...")
if not Read_Scenario:
    # stkRoot.NewScenario('Kuiper')
    stkRoot.NewScenario('StarLink')
scenario = stkRoot.CurrentScenario
scenario2 = scenario.QueryInterface(STKObjects.IAgScenario)
#scenario2.StartTime = '24 Sep 2020 16:00:00.00'
#scenario2.StopTime = '25 Sep 2020 16:00:00.00'

totalTime = time.time() - startTime
splitTime = time.time()
print("--- Scenario creation: {a:4.3f} sec\t\tTotal time: {b:4.3f} sec ---".format(a=totalTime, b=totalTime))


# 创建卫星星系
def Creat_satellite(numOrbitPlanes, numSatsPerPlane, hight, Inclination=53, name=''):
    # Create constellation object
    constellation = scenario.Children.New(STKObjects.eConstellation, name)
    constellation2 = constellation.QueryInterface(STKObjects.IAgConstellation)

    # Insert the constellation of Satellites
    for orbitPlaneNum in range(numOrbitPlanes):  # RAAN in degrees

        for satNum in range(numSatsPerPlane):  # trueAnomaly in degrees
            # Insert satellite
            satellite = scenario.Children.New(STKObjects.eSatellite, f"{name}{orbitPlaneNum}_{satNum}")
            satellite2 = satellite.QueryInterface(STKObjects.IAgSatellite)

            # Select Propagator
            satellite2.SetPropagatorType(STKObjects.ePropagatorTwoBody)

            # Set initial state
            twoBodyPropagator = satellite2.Propagator.QueryInterface(STKObjects.IAgVePropagatorTwoBody)
            keplarian = twoBodyPropagator.InitialState.Representation.ConvertTo(
                STKUtil.eOrbitStateClassical).QueryInterface(STKObjects.IAgOrbitStateClassical)

            keplarian.SizeShapeType = STKObjects.eSizeShapeSemimajorAxis
            keplarian.SizeShape.QueryInterface(
                STKObjects.IAgClassicalSizeShapeSemimajorAxis).SemiMajorAxis = hight + 6371  # km
            keplarian.SizeShape.QueryInterface(STKObjects.IAgClassicalSizeShapeSemimajorAxis).Eccentricity = 0

            keplarian.Orientation.Inclination = int(Inclination)  # degrees
            keplarian.Orientation.ArgOfPerigee = 0  # degrees
            keplarian.Orientation.AscNodeType = STKObjects.eAscNodeRAAN
            RAAN = 360 / numOrbitPlanes * orbitPlaneNum
            keplarian.Orientation.AscNode.QueryInterface(STKObjects.IAgOrientationAscNodeRAAN).Value = RAAN  # degrees

            keplarian.LocationType = STKObjects.eLocationTrueAnomaly
            trueAnomaly = 360 / numSatsPerPlane * satNum
            keplarian.Location.QueryInterface(STKObjects.IAgClassicalLocationTrueAnomaly).Value = trueAnomaly

            # Propagate
            satellite2.Propagator.QueryInterface(STKObjects.IAgVePropagatorTwoBody).InitialState.Representation.Assign(
                keplarian)
            satellite2.Propagator.QueryInterface(STKObjects.IAgVePropagatorTwoBody).Propagate()
            # Add to constellation object
            constellation2.Objects.AddObject(satellite)


# 为每个卫星加上发射机和接收机
def Add_transmitter_receiver(sat_list):
    for each in sat_list:
        Instance_name = each.InstanceName
        #  new transmitter and receiver
        transmitter = each.Children.New(STKObjects.eTransmitter, "Transmitter_" + Instance_name)
        reciver = each.Children.New(STKObjects.eReceiver, "Reciver_" + Instance_name)
        # sensor = each.Children.New(STKObjects.eSensor, 'Sensor_' + Instance_name)


# 设置发射机参数
def Set_Transmitter_Parameter(transmitter, frequency=12, EIRP=20, DataRate=14):
    transmitter2 = transmitter.QueryInterface(STKObjects.IAgTransmitter)  # 建立发射机的映射，以便对其进行设置
    transmitter2.SetModel('Simple Transmitter Model')
    txModel = transmitter2.Model
    txModel = txModel.QueryInterface(STKObjects.IAgTransmitterModelSimple)
    txModel.Frequency = frequency  # GHz range:10.7-12.7GHz
    txModel.EIRP = EIRP  # dBW
    txModel.DataRate = DataRate  # Mb/sec


# 设置接收机参数
def Set_Receiver_Parameter(receiver, GT=20, frequency=12):
    receiver2 = receiver.QueryInterface(STKObjects.IAgReceiver)  # 建立发射机的映射，以便对其进行设置
    receiver2.SetModel('Simple Receiver Model')
    recModel = receiver2.Model
    recModel = recModel.QueryInterface(STKObjects.IAgReceiverModelSimple)
    recModel.AutoTrackFrequency = False
    recModel.Frequency = frequency  # GHz range:10.7-12.7GHz
    recModel.GOverT = GT  # dB/K
    return receiver2


# 获得接收机示例，并设置其参数
def Get_sat_receiver(sat, GT=20, frequency=12):
    receiver = sat.Children.GetElements(STKObjects.eReceiver)[0]  # 找到该卫星的接收机
    receiver2 = Set_Receiver_Parameter(receiver=receiver, GT=GT, frequency=frequency)
    return receiver2


def Compute_access(access):
    access.ComputeAccess()
    accessDP = access.DataProviders.Item('Access Data')
    accessDP2 = accessDP.QueryInterface(STKObjects.IAgDataPrvInterval)
    Elements = ["Access Number", "Start Time", "Stop Time",
                "Duration"]  # , 'EIRP', 'Prop Loss', 'Rcvr Gain', "Xmtr Gain", "Eb/No", "BER"
    # scenario2.StartTime = "24 Sep 2020 16:00:00.00"
    # scenario2.StopTime = "25 Sep 2020 16:00:00.00"

    results = accessDP2.Exec(scenario2.StartTime, scenario2.StopTime)  # , Elements
    STimes = results.DataSets.GetDataSetByName('Start Time').GetValues()  # 时间
    ETimes = results.DataSets.GetDataSetByName('Stop Time').GetValues()  # 时间
    # EbNo = results.DataSets.GetDataSetByName('Eb/No').GetValues()  # 码元能量
    #  BER = results.DataSets.GetDataSetByName('BER').GetValues()  # 误码率
    Link_Name = results.DataSets.GetDataSetByName('Access Number').GetValues()
    # Link_sec = results.DataSets.GetDataSetByName("Duration").GetValues()
    # Prop_Loss = results.DataSets.GetDataSetByName('Prop Loss').GetValues()
    # Xmtr_Gain = results.DataSets.GetDataSetByName('Xmtr Gain').GetValues()
    # EIRP = results.DataSets.GetDataSetByName('EIRP').GetValues()
    # Range = results.DataSets.GetDataSetByName('Range').GetValues()
    return STimes, ETimes, Link_Name  # , BER, EbNo, Prop_Loss, Xmtr_Gain, EIRP,,Link_sec


def Creating_All_Access():
    # 首先清空场景中所有的链接
    print('Clearing All Access')
    stkRoot.ExecuteCommand('RemoveAllAccess /')
    T = scenario2.StartTime
    Temp = re.sub("\D", "", T)  # 时间戳
    Temp = int(Temp) - 100000
    tick = 0
    link_t = np.zeros((12, 12))  # 链接矩阵,行从KPA0_0开始排列，列从KPB0_0开始排列
    a1 = np.zeros((12,12,100))
    b1 = np.zeros((12,12,100))
    durationnum = np.zeros((12, 12))
    #将一天分为1440段，步长一分钟
    for t in range(0,1440):
        Temp = int(Temp) + 100000
        tick = tick + 1
        if tick>=60:
            tick = 0
            Temp = int(Temp) + 400000
        if int(Temp)==282021240000000:
            Temp=292021000000000
        number1 = -1  # 卫星对象标识
    # 计算某个卫星与其他卫星的链路质量，并生成报告
        if t==0:
            for each_sat in tqdm(sat_list):
                    number1 = number1 + 1
                    number2 = -1  # 卫星对象标识
                    for k in range(0, 2):
                       for j in range(0,3):
                         satellite = each_sat.QueryInterface(STKObjects.IAgSatellite)
                         now_sat_name = each_sat.InstanceName
                         now_plane_num = k
                         now_sat_num = j
                         number2 = 1 + number2
                         now_sat_transmitter = each_sat.Children.GetElements(STKObjects.eTransmitter)[0]  # 找到该卫星的发射机
                         Set_Transmitter_Parameter(now_sat_transmitter)  # , EIRP=20)
                         s = 'KPA' + str(now_plane_num) + '_' + str(now_sat_num)
                         #print(now_sat_name , s) #链接对象
                         access = now_sat_transmitter.GetAccessToObject(
                                Get_sat_receiver(sat_dic['KPA' + str(now_plane_num) + '_' + str(now_sat_num)]))

                         access.ComputeAccess()

                         accessIntervals = access.ComputedAccessIntervalTimes

                         accessDataProvider = access.DataProviders.Item('Access Data')

                         accessDataProvider2 = accessDataProvider.QueryInterface(STKObjects.IAgDataPrvInterval)

                         dataProviderElements = ['Start Time', 'Stop Time']
                         x = number1
                         y = number2
                         durationnum[x][y] = accessIntervals.Count
                         for i in range(0, accessIntervals.Count):
                             timesi = accessIntervals.GetInterval(i)
                             #print(timesi)
                             sxyi = list(timesi)
                             a1[x][y][i] = re.sub("\D", "", sxyi[0])  # 正则表达式
                             b1[x][y][i] = re.sub("\D", "", sxyi[1])
                             #return a1xyi,b1xyi,durationnumxy

                    for k in range(0, 2):
                        for j in range(0, 3):
                            now_sat_name = each_sat.InstanceName
                            satellite = each_sat.QueryInterface(STKObjects.IAgSatellite)
                            now_plane_num = k
                            now_sat_num = j
                            number2 = 1 + number2
                            satellite = each_sat.QueryInterface(STKObjects.IAgSatellite)

                            now_sat_transmitter = each_sat.Children.GetElements(STKObjects.eTransmitter)[0]  # 找到该卫星的发射机
                            Set_Transmitter_Parameter(now_sat_transmitter)  # , EIRP=20)
                            s1 = 'KPB' + str(now_plane_num) + '_' + str(now_sat_num)
                            #print(now_sat_name , s1)
                            access = now_sat_transmitter.GetAccessToObject(
                                Get_sat_receiver(sat_dic['KPB' + str(now_plane_num) + '_' + str(now_sat_num)]))

                            now_sat_num = now_sat_num + 1
                            access.ComputeAccess()
                            accessIntervals = access.ComputedAccessIntervalTimes

                            accessDataProvider = access.DataProviders.Item('Access Data')

                            accessDataProvider2 = accessDataProvider.QueryInterface(STKObjects.IAgDataPrvInterval)

                            dataProviderElements = ['Start Time', 'Stop Time']
                            x = number1
                            y = number2
                            durationnum[x][y] = accessIntervals.Count
                            for i in range(0, accessIntervals.Count):
                                timesi = accessIntervals.GetInterval(i)
                                #print(timesi)
                                sxyi = list(timesi)
                                a1[x][y][i] = re.sub("\D", "", sxyi[0])  # 正则表达式
                                b1[x][y][i] = re.sub("\D", "", sxyi[1])
            print('第', t, '个链接矩阵')
            print(link_t)
        else:
            for x in range(0, 12):
                for y in range(0, 12):
                    for i in range(0, int(durationnum[x][y])):
                        if Temp >= int(a1[x][y][i]) and Temp <= int(b1[x][y][i]):
                            link_t[x, y] = 1
                            break
            print('第', t, '个链接矩阵')
            print(link_t)
    '''
        access_backward = now_sat_transmitter.GetAccessToObject(
            Get_sat_receiver(sat_dic['Sat' + str(now_plane_num) + '_' + str((now_sat_num + 1) % 5)]))
        # 与前面的卫星的接收机相连
        access_forward = now_sat_transmitter.GetAccessToObject(
            Get_sat_receiver(sat_dic['Sat' + str(now_plane_num) + '_' + str((now_sat_num - 1) % 5)]))
        # 与左面的卫星的接收机相连
        access_left = now_sat_transmitter.GetAccessToObject(
            Get_sat_receiver(sat_dic['Sat' + str((now_plane_num - 1) % 12) + '_' + str(now_sat_num)]))
        # 与右面的卫星的接收机相连
        access_right = now_sat_transmitter.GetAccessToObject(
            Get_sat_receiver(sat_dic['Sat' + str((now_plane_num + 1) % 12) + '_' + str(now_sat_num)]))
        #B_Times, B_Link_Name = Compute_access(access_backward)
        #F_Times, F_Link_Name = Compute_access(access_forward)


        L_STimes,L_ETimes,L_Link_Names = Compute_access(access_left)#L_Link_Name,,L_Link_sec
      #  R_Times, R_Link_Name = Compute_access(access_right)
       # print('{0}\r', R_Times, R_Link_Name)
        print('{0}\r', L_STimes, L_ETimes,L_Link_Names)# R_Times, R_Link_Name, B_Times, B_Link_Name, F_Times, F_Link_Name, L_Times, L_Link_Names,L_Link_Name,L_Link_sec
    # stkRoot.ExecuteCommand('RemoveAllAccess /')
    '''


# 修改卫星及其轨道的颜色
def Change_Sat_color(sat_list):
    print('Changing Color of Satellite')
    for each_sat in tqdm(sat_list):
        now_sat_name = each_sat.InstanceName
        now_plane_num = int(now_sat_name.split('_')[0][3:])
        now_sat_num = int(now_sat_name.split('_')[1])
        satellite = each_sat.QueryInterface(STKObjects.IAgSatellite)
        graphics = satellite.Graphics
        graphics.SetAttributesType(1)  # eAttributesBasic
        attributes = graphics.Attributes
        attributes_color = attributes.QueryInterface(STKObjects.IAgVeGfxAttributesBasic)
        attributes_color.Inherit = False
        # 16436871 浅蓝色
        # 2330219 墨绿色
        # 42495 橙色
        # 9234160 米黄色
        # 65535 黄色
        # 255 红色
        # 16776960 青色
        color_sheet = [16436871, 2330219, 42495, 9234160, 65535, 255, 16776960]
        if now_sat_name[2] == 'A':
            color = 255
        elif now_sat_name[2] == 'B':
            color = 42495
        elif now_sat_name[2] == 'C':
            color = 16436871
        attributes_color.Color = color
        # 找出轨道对应的属性接口
        orbit = attributes.QueryInterface(STKObjects.IAgVeGfxAttributesOrbit)
        orbit.IsOrbitVisible = False  # 将轨道设置为不可见


# 如果不是读取当前场景，即首次创建场景
if not Read_Scenario:
    Creat_satellite(numOrbitPlanes=2, numSatsPerPlane=3, hight=750, Inclination=53, name='KPB')  # Starlink
    Creat_satellite(numOrbitPlanes=2, numSatsPerPlane=3, hight=450, Inclination=53, name='KPA')
    # Kuiper
    # Creat_satellite(numOrbitPlanes=34, numSatsPerPlane=34, hight=630, Inclination=51.9, nanow_sat_name = each_sat.InstanceNameme='KPA')  # Phase A
    # Creat_satellite(numOrbitPlanes=32, numSatsPerPlane=32, hight=610, Inclination=42, name='KPB')  # Phase B
    # Creat_satellite(numOrbitPlanes=28, numSatsPerPlane=28, hight=590, Inclination=33, name='KPC')  # Phase C
    sat_list = stkRoot.CurrentScenario.Children.GetElements(STKObjects.eSatellite)
    sat_dic = {}
    print('Creating Satellite Dictionary')
    for sat in tqdm(sat_list):
        sat_dic[sat.InstanceName] = sat
    Plane_num = []
    for i in range(0, 2):
        Plane_num.append(i)
    Sat_num = []
    for i in range(0, 1):
        Sat_num.append(i)
    Add_transmitter_receiver(sat_list)
    Creating_All_Access()
else:
# 创建卫星的字典，方便根据名字对卫星进行查找
    sat_list = stkRoot.CurrentScenario.Children.GetElements(STKObjects.eSatellite)
    sat_dic = {}
    print('Creating Satellite Dictionary')
    for sat in tqdm(sat_list):
       sat_dic[sat.InstanceName] = sat
    Plane_num = []
    for i in range(0, 2):
       Plane_num.append(i)
    Sat_num = []
    for i in range(0, 1):
       Sat_num.append(i)
    Creating_All_Access()
    access_list = stkRoot.CurrentScenario.Children.GetElements(STKObjects.eAccess)
    for access in tqdm(access_list):
      Times, Link_Name = Compute_access(access=access)
      print(Times, Link_Name)





