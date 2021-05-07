
from comtypes.client import CreateObject
from comtypes.gen import STKObjects

import time, datetime

uiApplication = CreateObject("STK11.Application")
uiApplication.Visible=True
uiApplication.UserControl=True
root = uiApplication.Personality2

# 1. Create a new scenario.

root.NewScenario("Python_Starter")

scenario = root.CurrentScenario

# 2. Set the analytical time period.

scenario2 = scenario.QueryInterface(STKObjects.IAgScenario)
# root.CurrentScenario.StartTime = '8 Jun 2020 16:00:00.00'
# root.CurrentScenario.StopTime = '9 Jun 2020 04:00:00.00'

# 3. Reset the animation time.

root.Rewind()

#TASK 3

#1. Add a target object to the scenario.

target = scenario.Children.New(STKObjects.eTarget,"GroundTarget");

target2 = target.QueryInterface(STKObjects.IAgTarget)

#2. Move the Target object to a desired location.

target2.Position.AssignGeodetic(50,-100,0)

#3. Add a Satellite object to the scenario.


satellite = scenario.Children.New(STKObjects.eSatellite, "LeoSat")

#4. Propagate the Satellite object's orbit.

root.ExecuteCommand('SetState */Satellite/LeoSat Classical TwoBody "' + scenario2.StartTime + '" "'+ scenario2.StopTime +'" 60 ICRF "'+ scenario2.StartTime + '" 7200000.0 0.0 90 0.0 0.0 0.0')


access = satellite.GetAccessToObject(target)

access.ComputeAccess()

accessIntervals = access.ComputedAccessIntervalTimes

accessDataProvider = access.DataProviders.Item('Access Data')

accessDataProvider2 = accessDataProvider.QueryInterface(STKObjects.IAgDataPrvInterval)

dataProviderElements = ['Start Time', 'Stop Time']

for i in range(0,accessIntervals.Count):

 times = accessIntervals.GetInterval(i)
 print(times)

print(accessIntervals.GetInterval[0][1])
print(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
print(datetime.datetime.now()+datetime.timedelta(minutes=1)).strftime("%Y-%m-%d %H:%M:%S")







