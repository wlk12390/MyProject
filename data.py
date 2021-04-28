# Get reference to running STK instance using win32com
from win32com.client import GetActiveObject
uiApplication = GetActiveObject('STK11.Application')
uiApplication.Visible = True

# Get our IAgStkObjectRoot interface
root = uiApplication.Personality2


satellite = root.CurrentScenario

# IAgStkObjectRoot root: STK Object Model root
# IAgSatellite satellite: Satellite object
# IAgFacility facility: Facility object

# Change DateFormat dimension to epoch seconds to make the data easier to handle in
# Python
root.UnitPreferences.Item('DateFormat').SetCurrentUnit('EpSec')
# Get the current scenario
scenario = root.CurrentScenario
# Set up the access object
access = satellite.GetAccessToObject()
access.ComputeAccess()
# Get the Access AER Data Provider
accessDP = access.DataProviders.Item('Access Data').Exec(scenario.StartTime, scenario.StopTime)

accessStartTimes = accessDP.DataSets.GetDataSetByName('Start Time').GetValues
accessStopTimes = accessDP.DataSets.GetDataSetByName('Stop Time').GetValues
