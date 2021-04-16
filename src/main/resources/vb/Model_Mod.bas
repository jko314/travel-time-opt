Attribute VB_Name = "Model_Mod"
Option Explicit


Private Const c_strColHeadings_Airports As String = "Code;Cat;City;State;Name;Latitude;Longitude;Op Start;Op Hrs;Time Zone;"
Private Const c_strColHeadings_ServiceAreas As String = "Name;City;State;Latitude;Longitude;"
Private Const c_strColHeadings_AirportServiceAreas As String = "Airport Code;Service Area;Priority;IsBase;CM Travel Time;"
Private Const c_strColHeadings_EquipmentModels As String = "Equipment Type;Manufacturer;Model;"
Private Const c_strColHeadings_AirportEquipment As String = "Airport;Model;Count;SerialNum;"
Private Const c_strColHeadings_Equipment As String = "Equipment ID;Make-Model;Airport;"
Private Const c_strColHeadings_EquipmentPM As String = "Make.Model;Name;Periodicity;Allowed Slack;Labor-Initial;Labor-Wait;Labor-Final;Consumables-Cost;Technician-Count;"
Private Const c_strColHeadings_EquipmentCM As String = "Make.Model;Name;Frequency;CM-Time;CM-StndDev;CM-Min;CM-Max;Parts-Cost;Parts-Time;Parts-StndDev;Parts-Min;Parts-Max;Consumables-Cost;Tech-Count;"
Private Const c_strColHeadings_PMStatus As String = "Airport;Equipment ID;Equipment Type;Periodicity;Month;Day;ScheduleIndex;LastPM;"
Private Const c_strColHeadings_DailyPMTimes As String = "Date;PM Time;"
Private Const c_strColHeadings_PMSchedule As String = "Airport;Equipment ID;Equipment Type;PM-Dates;"

Private Const c_dblEarthRadius_Equitorial As Double = 3963.191
Private Const c_dblEarthRadius_Polar As Double = 3949.9028
Private Const c_dblPiOver180 As Double = 1.74532925199433E-02



Public Enum FILTERTYPE_ENUM
  FILTER_ModelNum = 1
End Enum

Public Type FILTER_TYPE
  enmFilterTypes As FILTERTYPE_ENUM
  strModelNumList As String
End Type





Public Type ITEMINDEXLIST_TYPE
  lngItemCount As Long
  lngItemIndexes() As Long
End Type



Public Enum MODELSTEPTYPE_ENUM
  MDLSTEP_initialize = 1
  MDLSTEP_randomize = 2
  MDLSTEP_loadAirports = 32
  MDLSTEP_loadServiceAreas = 40
  MDLSTEP_loadAirportServiceAreas = 41
  MDLSTEP_assignAirportServiceAreas = 42
  MDLSTEP_loadEquipmentModels = 48
  MDLSTEP_loadEquipment = 49
  MDLSTEP_loadPMRequirements = 50
  MDLSTEP_loadCMRequirements = 51
  MDLSTEP_loadDepotRequirements = 52
  MDLSTEP_loadPMStatus = 56
  MDLSTEP_createPMStatus = 57
End Enum
  
  
Public Type MODELSTEP_TYPE
  enmStepType As MODELSTEPTYPE_ENUM
  lngParameterCount As Long
  strParameters() As String
  strStatus As String
End Type
  



' Enumerations

Public Enum MODELDATATYPE_ENUM
  MODDATA_Airports = 1
  MODDATA_EquipmentTypes = 2
  MODDATA_Equipment = 4
End Enum


Public Enum LOCATIONTYPE_ENUM
  LOCTYPE_Airport = 1
  LOCTYPE_ServiceCenter = 2
End Enum

Public Enum ACTIVITYTYPE_ENUM
  ' Activity category
  ACTTYPE_PM = 16777216 * 1
  ACTTYPE_CM = 16777216 * 2
  ACTTYPE_Depot = 16777216 * 3
  ' Repair activities
  ACTTYPE_Diagnosis = 1
  ACTTYPE_PartsRequest = 2
  ACTTYPE_PartsFulfillment = 3
  ACTTYPE_PartsLocalLogistics = 4
  ACTTYPE_Repair = 5
  ACTTYPE_Test = 6
  ACTTYPE_RequestTechSupport = 7
  ACTTYPE_ProvideTechSupport = 8
  ACTTYPE_Signoff = 15
  ' Travel activities
  ACTTYPE_DriveOwnCar = 48
  ACTTYPE_DriveCompanyCar = 49
  ACTTYPE_Fly = 50
  ACTTYPE_Taxi = 51
  ACTTYPE_EnterAirport = 52
End Enum
  
Public Enum PMPERIODICITY_ENUM
  PMPER_NoPM = 0
  PMPER_Daily = 1
  PMPER_Weekly = 2
  PMPER_Biweekly = 3
  PMPER_Monthly = 4
  PMPER_Quarterly = 5
  PMPER_SemiAnnually = 6
  PMPER_Annually = 7
End Enum
Public Const c_lngPMPeriodicity_MaxValue As Long = 7
Public Const c_strPMPeriodicityList As String = "0;NoPM;1;Daily;2;Weekly;3;Biweekly;4;Monthly;5;Quarterly;6;Semi-Annual;7;Annual;"


Public Type TRAVELCOSTS_TYPE
  strID As String
  dblLodging As Double
  dblPerDiem As Double
  dblRentalCar As Double
  dblFacilityParking As Double
End Type



Public Type MAINTACTIVITY_TYPE
  lngEquipmentIndex As Long
  lngPMIndex As Long
  dteScheduledStart As Date
  dteActualStartDate As Date
  dteActualEndDate As Date
  ' Need details info
End Type
  
Public Type PMSCHEDULE_TYPE
  enmPeriodicity As PMPERIODICITY_ENUM
  lngMonth As Long                         ' Specifies the month within a cycle in which PM is performed.
                                           ' If weekly or monthly, this is 0
                                           ' If quarterly, this is 0 to 2, indicating the month within the quarter
                                           ' If semiannually, this is 0 to 5,  indicating the month within the semiannual cycle
                                           ' If annually, this is 0 to 11
  lngDay As Long                           ' Specifies on which day PM is performed.
                                           ' If weekly, it is the day of the week (1-7)
                                           ' Otherwise, it is the day of the month (1-31)
  lngPMScheduleIndex As Long               ' Index into lngPMSchedule array of the EquipmentModel structure
  dteLastPMCompleted As Date
  lngPMItemCount As Long
  lngPMItemIndexes() As Long               ' Indexes into m_udtPMItems
End Type

Public Type PMACTIVITY_TYPE
  lngPMIndex As Long
  dteStartTime As Date
  dteEndTime As Date
End Type

Public Type CMACTIVITY_TYPE
  lngCMIndex As Long
  dteFailureTime As Date
  dteCallTime As Date
  dteDispatchTime As Date
  dteArrivalTime As Date
  dteDiagnosisEndTime As Date
  dtePartsRequestTime As Date
  dtePartsFulfillmentTime As Date
  dtePartsLocalLogisticsTime As Date
  dteRepairTime As Date
  dteTestTime As Date
  dteSignoffTime As Date
End Type



'''''''''' Probability and statistics data structures

Public Type DISTRIBUTION_TYPE
  dblAvg As Double
  dblStndDev As Double
  dblMin As Double
  dblMax As Double
End Type

'''''''''' Airport and airport equipment data structures

Public Type EQUIPMENTTYPE_TYPE
  lngID As Long
  strName As String
End Type

Public Type EQUIPMENTITEM_TYPE
  lngID As Long
  lngEquipmentModelIndex As Long
  ' Airport Reference
  lngAirportIndex As Long
  lngAirportEquipmentModelIndex As Long
  lngAirportEquipmentIndex As Long
  ' PM Requirements
  lngPMRequirementItemCount As Long
  lngPMRequirementsIndexes() As Long        ' Refers to the PM Requirements in Equipment_Models
  ' PM Scheduling
  
  
  udtPMSchedule As PMSCHEDULE_TYPE
  dteNextPMDue As Date
  lngNextPMRequirementIndex As Long
  ' PM Activity Tracking
  lngPMCount As Long
  lngPMIndexes() As Long
  ' CM Activities
  lngCMActivityCount As Long
  udtCMActivities() As CMACTIVITY_TYPE
End Type

Public Type EQUIPMENT_TYPE
  lngEquipmentCount As Long
  udtEquipment() As EQUIPMENTITEM_TYPE
End Type


Public Type AIRPORTEQUIPMENT_TYPE
  lngEquipmentModelIndex As Long
  lngCMRequirementIndex As Long
  lngCount As Long
  lngEquipmentIndexes() As Long
End Type

Public Type AIRPORTSERVICEAREA_TYPE
  lngServiceAreaIndex As Long
  lngPriority As Long
  booIsBase As Boolean
  dblCMTravelTime As Double
End Type


Public Type AIRPORT_TYPE
  ' Airport Info
  lngID As Long
  strCode As String
  strCat As String
  strCity As String
  strState As String
  strName As String
  dblLatitude As Double
  dblLongitude As Double
  lngOperatingStartHour As Long
  lngOperatingHours As Long
  lngTimeZoneAdjustment As Long
  ' Service Areas servicing airport
  lngBaseServiceAreaIndex As Long
  lngServiceAreaCount As Long
  udtServiceAreaIndexes() As AIRPORTSERVICEAREA_TYPE
  dblCMTravelTime As Double
  ' Airport equipment
  lngEquipmentCount As Long
  udtEquipment() As AIRPORTEQUIPMENT_TYPE
  ' Airport PM
  enmPMPeriodicity As PMPERIODICITY_ENUM                  ' Most frequent PM at airport
  dblPMTime(1 To c_lngPMPeriodicity_MaxValue) As Double   ' Total PM time in each category
End Type

Public Type AIRPORTS_TYPE
  lngAirportCount As Long
  udtAirport() As AIRPORT_TYPE
  colAirports As Collection
End Type



'''''''''' Service area data structures

Public Type SERVICEAREA_TYPE
  ' Service Area Info
  lngID As Long
  strName As String
  strCity As String
  strState As String
  dblLatitude As Double
  dblLongitude As Double
  ' Airports in Service Area
  lngAirportCount As Long
  lngAirportIndexes() As Long
  ' Maintenance activities
  lngScheduledMaintCount As Long
  lngScheduledMaintIndexes() As Long
  lngCompletedMaintCount As Long
  lngCompletedMaintIndexes() As Long
  lngInteruptedMaintCount As Long
  lngInteruptedMaintIndexes() As Long
End Type

Public Type SERVICEAREAS_TYPE
  lngServiceAreaCount As Long
  udtServiceAreas() As SERVICEAREA_TYPE
  colServiceAreas As Collection
End Type


'''''''''' PM data structures

Public Type PMREQUIREMENTS_TYPE
  strName As String
  lngAllowedSlack As Long
  lngEventsPerYear As Long                      ' Events per year, corrected for less frequent overlaps
  dblLabor_Initial As Double                    ' Per-event labor
  dblLabor_Wait As Double
  dblLabor_Final As Double
  dblConsumables As Double
  lngTechnicianCount As Long
End Type



Public Type PMITEM_TYPE
  lngEquipmentIndex As Long
  lngPMRequirementIndex As Long
  lngTripIndex As Long
  dteScheduledStart As Date
  dteStartTime As Date
  dteEndTime As Date
End Type







Public Type TRAVEL_TYPE
  enmFromType As LOCATIONTYPE_ENUM
  lngFromIndex As Long
  enmToType As LOCATIONTYPE_ENUM
  lngToIndex As Long
  dteScheduleStart As Date
  dteScheduleEnd As Date
  dteActualStart As Date
  dteActualEnd As Date
End Type

  
Public Type CMREQUIREMENT_TYPE
  strModelNum As String
  strName As String
  dblFrequency As Double
  udtCMTime As DISTRIBUTION_TYPE
  dblPartsCost As Double
  udtPartsTime As DISTRIBUTION_TYPE
  dblConsumablesCost As Double
  lngTechnicianCount As Long
End Type

Public Type DEPOTREQUIREMENT_TYPE
  strName As String
  dblFrequency As Double
  udtDiagnosisTime As DISTRIBUTION_TYPE
  udtReinstallTime As DISTRIBUTION_TYPE
End Type

Public Type EQUIPMENTMODEL_TYPE
  lngID As Long
  lngEquipmentTypeIndex As Long
  strManufacturer As String
  strModel As String
  ' PM Requirements
  enmPeriodicity As PMPERIODICITY_ENUM
  udtPM(0 To c_lngPMPeriodicity_MaxValue) As PMREQUIREMENTS_TYPE
  lngPMSchedule() As Long                ' Lists in sequential order the PM requirements for a year.
                                         ' Each entry is an index into udtPM
  dblPMTime(0 To c_lngPMPeriodicity_MaxValue) As Double   ' Total PM time in each category
  ' CM Requirements
'  lngCMCount As Long
'  udtCM() As CMREQUIREMENT_TYPE
  '
End Type

Public Type EQUIPMENTMODELS_TYPE
  lngEquipmentModelCount As Long
  udtEquipmentModels() As EQUIPMENTMODEL_TYPE
  colEquipmentModels As Collection
End Type


Public Type CMREQUIREMENTS_TYPE
  lngCMRequirementCount As Long
  udtCMRequirements() As CMREQUIREMENT_TYPE
  colCMRequirement As Collection
End Type
  




Public Type METRICITEM_TYPE
  lngAirportIndex As Long
  lngEquipmentModelIndex As Long
  lngEquipmentTypeIndex As Long
  lngEquipmentCount As Long
  dblOperatingTime As Double
  lngPMEvents As Long
  dblPMTime As Double
  lngCMEvents As Long
  dblCMTime As Double
End Type

Public Type METRICS_TYPE
  lngEquipmentModelCount As Long
  udtEquipmentModelMetrics() As METRICITEM_TYPE
  lngEquipmentTypeCount As Long
  udtEquipmentTypeMetrics() As METRICITEM_TYPE
End Type


Public Enum MTRIP_STATUS_ENUM
  TRIPSTAT_Scheduled = 1
  TRIPSTAT_Active = 2
  TRIPSTAT_Completed = 3
End Enum

Public Enum MTRIP_ITEMTYPE_ENUM
  TRIPITEM_Travel = 1
  TRIPITEM_PM = 2
  TRIPITEM_CM = 3
  TRIPITEM_Other = 4
End Enum

Public Type MTRIP_ITEM_TYPE
  enmItemType As MTRIP_ITEMTYPE_ENUM
  lngItemIndex As Long
  dteStartTime As Date
  dteEndTime As Date
End Type


Public Type MAINTENANCETRIP_TYPE
  lngID As Long
  enmTripStatus As MTRIP_STATUS_ENUM
  dteScheduledStart As Date
  dblScheduledLength As Double
  lngItemCount As Long
  lngItemCurrent As Long
  lngItems() As MTRIP_ITEM_TYPE
End Type
  
  
Public Type EQUIPMENTTYPES_TYPE
  lngEquipmentTypeCount As Long
  udtEquipmentTypes() As EQUIPMENTTYPE_TYPE
  colEquipmentTypes As Collection
End Type
  
Private Type ITEMLIST_TYPE
  lngCount As Long
  lngItemIndexes() As Long
End Type
  
  
Public Type AIRPORTSMODEL_TYPE
  dteModelStartDate As Date
  dteModelEndDate As Date
  udtAirports As AIRPORTS_TYPE
  udtServiceAreas As SERVICEAREAS_TYPE
  udtEquipmentTypes As EQUIPMENTTYPES_TYPE
  udtEquipmentModels As EQUIPMENTMODELS_TYPE
  udtCMRequirements As CMREQUIREMENTS_TYPE
  udtEquipment As EQUIPMENT_TYPE
  dblAirportDistances() As Double
End Type
  
  
 
Private m_objStatusRange As Excel.Range
Private m_objErrorRange As Excel.Range

Private m_lngModelStepCount As Long
Private m_udtModelSteps() As MODELSTEP_TYPE

Private m_dteModelStartDate As Date
Private m_dteModelEndDate As Date
  
'Private m_lngAirportCount As Long
'Private m_udtAirports() As AIRPORTS_TYPE
'Private m_colAirports As Collection

'Private m_lngServiceAreaCount As Long
'Private m_udtServiceAreas() As SERVICEAREA_TYPE
'Private m_colServiceAreas As Collection

'Private m_lngEquipmentTypeCount As Long
'Private m_udtEquipmentTypes() As EQUIPMENTTYPE_TYPE
'Private m_colEquipmentTypes As Collection

'Private m_lngEquipmentModelCount As Long
'Private m_udtEquipmentModels() As EQUIPMENTMODEL_TYPE
'Private m_colEquipmentModels As Collection

'Private m_lngEquipmentCount As Long
'Private m_udtEquipment() As EQUIPMENT_TYPE

Private m_lngTravelCostCount As Long
Private m_udtTravelCosts() As TRAVELCOSTS_TYPE
Private m_colTravelCosts As Collection

Private m_lngPMItemCount As Long
Private m_udtPMItems() As PMITEM_TYPE

'Private m_dblAirportDistances() As Double






'''''''''' Methods for running the model
''''''
''
  
  

''''''''''' Import / Export Methods
''''''
''
  
''''''''''' Model Management

Public Sub initializeModel( _
    ByVal dteStartDate As Date, ByVal dteEndDate As Date, _
    Optional ByVal objStatusRange As Excel.Range = Nothing, _
    Optional ByVal objErrorRange As Excel.Range = Nothing)

  m_dteModelStartDate = dteStartDate
  m_dteModelEndDate = dteEndDate
  Set m_objStatusRange = objStatusRange
  Set m_objErrorRange = objErrorRange
  
  m_lngModelStepCount = 1
  ReDim m_udtModelSteps(0 To 63)
  With m_udtModelSteps(0)
    .enmStepType = MDLSTEP_initialize
    .lngParameterCount = 2
    ReDim .strParameters(0 To 1)
    .strParameters(0) = dteStartDate
    .strParameters(1) = dteEndDate
    .strStatus = "Done"
  End With

End Sub

  
Public Function randomizeModel(strStepParameter As String) As Boolean

  Dim lngStepNumber As Long

  lngStepNumber = recordModelStep(MDLSTEP_randomize, strStepParameter)

  If m_udtModelSteps(lngStepNumber).lngParameterCount = 0 Then
    Randomize
  ElseIf m_udtModelSteps(lngStepNumber).lngParameterCount = 1 Then
    If IsNumeric(m_udtModelSteps(lngStepNumber).strParameters(0)) Then
      Randomize -1
      Randomize CLng(m_udtModelSteps(lngStepNumber).strParameters(0))
    Else
      m_udtModelSteps(lngStepNumber).strStatus = "Error: Invalid parameter"
      Exit Function
    End If
  Else
    m_udtModelSteps(lngStepNumber).strStatus = "Error: Invalid parameter"
    Exit Function
  End If

  m_udtModelSteps(lngStepNumber).strStatus = "Done"
  randomizeModel = True

End Function
  
  
  
''''''''''' Airports
  
Public Function loadAirports( _
    ByRef udtAirports As AIRPORTS_TYPE, _
    objWorksheet As Excel.Worksheet, strStepParameter As String) As Boolean

  Dim varValues As Variant
  Dim lngRowNumber As Long, lngStepNumber As Long
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Loading airports ..."
  End If
  
  lngStepNumber = recordModelStep(MDLSTEP_loadAirports, strStepParameter)

  If Not verifyColumnHeadings(objWorksheet, c_strColHeadings_Airports, 8) Then
    recordError m_udtModelSteps(lngStepNumber), "Incorrect headings in Airports worksheet"
    Exit Function
  End If
  
  lngRowNumber = objWorksheet.Range("A10:A10").CurrentRegion.Rows.Count
  varValues = objWorksheet.Range("A10:J" & (9 + lngRowNumber)).Value
  
  udtAirports.lngAirportCount = UBound(varValues, 1)
  ReDim udtAirports.udtAirport(0 To udtAirports.lngAirportCount - 1)
  Set udtAirports.colAirports = New Collection
  For lngRowNumber = 1 To udtAirports.lngAirportCount
    With udtAirports.udtAirport(lngRowNumber - 1)
      .lngID = lngRowNumber
      .strCode = varValues(lngRowNumber, 1)
      .strCat = varValues(lngRowNumber, 2)
      .strCity = varValues(lngRowNumber, 3)
      .strState = varValues(lngRowNumber, 4)
      .strName = varValues(lngRowNumber, 5)
      .dblLatitude = varValues(lngRowNumber, 6)
      .dblLongitude = varValues(lngRowNumber, 7)
      .lngOperatingStartHour = varValues(lngRowNumber, 8)
      .lngOperatingHours = varValues(lngRowNumber, 9)
      If varValues(lngRowNumber, 10) <> "" Then
        .lngTimeZoneAdjustment = varValues(lngRowNumber, 10)
      End If
      udtAirports.colAirports.Add lngRowNumber - 1, "A:" & .strCode
    End With
  Next
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Loading airports ... done"
  End If
  m_udtModelSteps(lngStepNumber).strStatus = "Done"
  
  loadAirports = True

End Function


Public Function exportAirports( _
    udtAirports As AIRPORTS_TYPE, _
    objWorkbook As Excel.Workbook, objAfterWorksheet As Excel.Worksheet, _
    strStepParameter As String) As Excel.Worksheet

  Dim objWorksheet As Excel.Worksheet
  Dim varValues As Variant
  Dim lngRowNumber As Long, lngAirportIndex As Long
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Exporting airports ..."
  End If
  
  On Error Resume Next
  Set objWorksheet = objWorkbook.Worksheets.Item("Airports")
  If objWorksheet Is Nothing Then
    If objAfterWorksheet Is Nothing Then
      Set objWorksheet = objWorkbook.Worksheets.Add()
    Else
      Set objWorksheet = objWorkbook.Worksheets.Add(, , objAfterWorksheet)
    End If
    objWorksheet.Name = "Airports"
    exportColumnHeadings objWorksheet, c_strColHeadings_Airports
  End If
  On Error GoTo 0
  
  ' Clear current contents (if any)
  lngRowNumber = objWorksheet.Range("A3:A3").CurrentRegion.Rows.Count
  objWorksheet.Range("A3:J" & (2 + lngRowNumber)).Clear
  
  ' Export airports
  varValues = objWorksheet.Range("A3:J" & (2 + udtAirports.lngAirportCount)).Value
  For lngAirportIndex = 0 To udtAirports.lngAirportCount - 1
    lngRowNumber = lngAirportIndex + 1
    With udtAirports.udtAirport(lngAirportIndex)
      varValues(lngRowNumber, 1) = .strCode
      varValues(lngRowNumber, 2) = .strCat
      varValues(lngRowNumber, 3) = .strCity
      varValues(lngRowNumber, 4) = .strState
      varValues(lngRowNumber, 5) = .strName
      varValues(lngRowNumber, 6) = .dblLatitude
      varValues(lngRowNumber, 7) = .dblLongitude
      varValues(lngRowNumber, 8) = .lngOperatingStartHour
      varValues(lngRowNumber, 9) = .lngOperatingHours
      varValues(lngRowNumber, 10) = .lngTimeZoneAdjustment
    End With
  Next
  objWorksheet.Range("A3:J" & (2 + udtAirports.lngAirportCount)).Value = varValues
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Exporting airports ... done"
  End If
  
  Set exportAirports = objWorksheet
  Set objWorksheet = Nothing

End Function


''''''''''' Service Areas

Public Function loadServiceAreas( _
    ByRef udtServiceAreas As SERVICEAREAS_TYPE, _
    objWorkbook As Excel.Workbook, strStepParameter As String) As Boolean

  Dim objWorksheet As Excel.Worksheet
  Dim varValues As Variant
  Dim lngRowNumber As Long
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Loading service areas ..."
  End If
  
  Set objWorksheet = objWorkbook.Worksheets.Item("ServiceAreas")
  
  If Not verifyColumnHeadings(objWorksheet, c_strColHeadings_ServiceAreas) Then
    If Not (m_objErrorRange Is Nothing) Then
      m_objErrorRange.Value = "Invalid column headings in ServiceAreas worksheet"
      Exit Function
    End If
  End If
  
  lngRowNumber = objWorksheet.Range("A3:A3").CurrentRegion.Rows.Count
  varValues = objWorksheet.Range("A3:E" & (2 + lngRowNumber)).Value
  udtServiceAreas.lngServiceAreaCount = UBound(varValues, 1)
  ReDim m_udtServiceAreas(0 To udtServiceAreas.lngServiceAreaCount - 1)
  Set udtServiceAreas.colServiceAreas = New Collection
  For lngRowNumber = 1 To udtServiceAreas.lngServiceAreaCount
    With udtServiceAreas.udtServiceAreas(lngRowNumber - 1)
      .lngID = lngRowNumber
      .strName = varValues(lngRowNumber, 1)
      .strCity = varValues(lngRowNumber, 2)
      .strState = varValues(lngRowNumber, 3)
      .dblLatitude = varValues(lngRowNumber, 4)
      .dblLongitude = varValues(lngRowNumber, 5)
      udtServiceAreas.colServiceAreas.Add lngRowNumber - 1, "SA:" & .strName
    End With
  Next
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Loading service areas ... done"
  End If
  
  loadServiceAreas = True
  
End Function


Public Function exportServiceAreas( _
    udtServiceAreas As SERVICEAREAS_TYPE, _
    ByVal objWorkbook As Excel.Workbook, ByVal objAfterWorksheet As Excel.Worksheet, _
    strStepParameter As String) As Excel.Worksheet

  Dim objWorksheet As Excel.Worksheet
  Dim varValues As Variant
  Dim lngRowNumber As Long, lngServiceAreaIndex As Long
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Exporting service areas ..."
  End If
  
  On Error Resume Next
  Set objWorksheet = objWorkbook.Worksheets.Item("ServiceAreas")
  If objWorksheet Is Nothing Then
    If objAfterWorksheet Is Nothing Then
      Set objWorksheet = objWorkbook.Worksheets.Add()
    Else
      Set objWorksheet = objWorkbook.Worksheets.Add(, objAfterWorksheet)
    End If
    objWorksheet.Name = "ServiceAreas"
    exportColumnHeadings objWorksheet, c_strColHeadings_ServiceAreas
  End If
  
  ' Clear existing content
  lngRowNumber = objWorksheet.Range("A3:A3").CurrentRegion.Rows.Count
  objWorksheet.Range("A3:E" & (2 + lngRowNumber)).Clear
  
  varValues = objWorksheet.Range("A3:E" & (2 + udtServiceAreas.lngServiceAreaCount)).Value
  For lngServiceAreaIndex = 0 To udtServiceAreas.lngServiceAreaCount - 1
    lngRowNumber = lngServiceAreaIndex + 1
    With udtServiceAreas.udtServiceAreas(lngRowNumber - 1)
      varValues(lngRowNumber, 1) = .strName
      varValues(lngRowNumber, 2) = .strCity
      varValues(lngRowNumber, 3) = .strState
      varValues(lngRowNumber, 4) = .dblLatitude
      varValues(lngRowNumber, 5) = .dblLongitude
    End With
  Next
  objWorksheet.Range("A3:E" & (2 + udtServiceAreas.lngServiceAreaCount)).Value = varValues
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Exporting service areas ... done"
  End If
  
  Set exportServiceAreas = objWorksheet
  Set objWorksheet = Nothing
  
End Function


''''''''''' Airport Service Areas

Public Function loadAirportServiceAreas( _
    ByRef udtAirports As AIRPORTS_TYPE, ByRef udtServiceAreas As SERVICEAREAS_TYPE, _
    objWorkbook As Excel.Workbook, strStepParameter As String) As Boolean

  Dim objWorksheet As Excel.Worksheet
  Dim varValues As Variant
  Dim lngRowNumber As Long, lngAirportIndex As Long, lngServiceAreaIndex As Long, _
      strAirportCode As String, strServiceAreaName As String
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Loading airport service areas ..."
  End If
  
  Set objWorksheet = objWorkbook.Worksheets.Item("Airport_ServiceAreas")
  If Not verifyColumnHeadings(objWorksheet, c_strColHeadings_AirportServiceAreas) Then
    If Not (m_objErrorRange Is Nothing) Then
      m_objErrorRange.Value = "Invalid column headings in Airport_ServiceAreas"
    End If
    Exit Function
  End If
  
  For lngAirportIndex = 0 To udtAirports.lngAirportCount - 1
    With udtAirports.udtAirport(lngAirportIndex)
      .lngServiceAreaCount = 0
      .lngBaseServiceAreaIndex = -1
      ReDim .udtServiceAreaIndexes(0 To 7)
    End With
  Next
  For lngServiceAreaIndex = 0 To udtServiceAreas.lngServiceAreaCount - 1
    With udtServiceAreas.udtServiceAreas(lngServiceAreaIndex)
      .lngAirportCount = 0
      ReDim .lngAirportIndexes(0 To 31)
    End With
  Next
  
  lngRowNumber = objWorksheet.Range("A3:A3").CurrentRegion.Rows.Count
  varValues = objWorksheet.Range("A3:E" & (2 + lngRowNumber)).Value
  For lngRowNumber = 1 To lngRowNumber
    strAirportCode = varValues(lngRowNumber, 1)
    lngAirportIndex = udtAirports.colAirports.Item("A:" & strAirportCode)
    strServiceAreaName = varValues(lngRowNumber, 2)
    lngServiceAreaIndex = udtServiceAreas.colServiceAreas.Item("SA:" & strServiceAreaName)
    
    With udtAirports.udtAirport(lngAirportIndex)
      If UBound(.udtServiceAreaIndexes) < .lngServiceAreaCount Then
        ReDim Preserve .udtServiceAreaIndexes(0 To 7 + .lngServiceAreaCount)
      End If
      With .udtServiceAreaIndexes(.lngServiceAreaCount)
        .lngServiceAreaIndex = lngServiceAreaIndex
        .lngPriority = varValues(lngRowNumber, 3)
        .booIsBase = varValues(lngRowNumber, 4)
        .dblCMTravelTime = varValues(lngRowNumber, 5)
      End With
      If .udtServiceAreaIndexes(.lngServiceAreaCount).booIsBase Then
        If .lngBaseServiceAreaIndex <> -1 Then Err.Raise 5
        .lngBaseServiceAreaIndex = lngServiceAreaIndex
      End If
      If .lngServiceAreaCount = 0 Then
        .dblCMTravelTime = varValues(lngRowNumber, 5) / 24
      ElseIf varValues(lngRowNumber, 5) / 24 < .dblCMTravelTime Then
        .dblCMTravelTime = varValues(lngRowNumber, 5) / 24
      End If
      .lngServiceAreaCount = .lngServiceAreaCount + 1
    End With
    
    With udtServiceAreas.udtServiceAreas(lngServiceAreaIndex)
      If UBound(.lngAirportIndexes) < .lngAirportCount Then
        ReDim Preserve .lngAirportIndexes(0 To 31 + .lngAirportCount)
      End If
      .lngAirportIndexes(.lngAirportCount) = lngAirportIndex
      .lngAirportCount = .lngAirportCount + 1
    End With

  Next
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Loading airport service areas ... done"
  End If
  
  loadAirportServiceAreas = True

End Function


Public Function exportAirportServiceAreas( _
    udtAirports As AIRPORTS_TYPE, udtServiceAreas As SERVICEAREAS_TYPE, _
    ByVal objWorkbook As Excel.Workbook, ByVal objAfterWorksheet As Excel.Worksheet, _
    strStepParameter As String) As Excel.Worksheet

  Dim objWorksheet As Excel.Worksheet
  Dim varValues As Variant, lngRowCount As Long
  Dim lngRowNumber As Long, lngAirportIndex As Long, lngAirportServiceAreaIndex As Long, _
      lngServiceAreaIndex As Long, strAirportCode As String, strServiceAreaName As String
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Exporting airport service areas ..."
  End If
  
  On Error Resume Next
  Set objWorksheet = objWorkbook.Worksheets.Item("Airport_ServiceAreas")
  If objWorksheet Is Nothing Then
    If objAfterWorksheet Is Nothing Then
      Set objWorksheet = objWorkbook.Worksheets.Add()
    Else
      Set objWorksheet = objWorkbook.Worksheets.Add(, objAfterWorksheet)
    End If
    objWorksheet.Name = "Airport_ServiceAreas"
    exportColumnHeadings objWorksheet, c_strColHeadings_AirportServiceAreas
  End If
  
  lngRowNumber = objWorksheet.Range("A3:A3").CurrentRegion.Rows.Count
  objWorksheet.Range("A3:E" & (2 + lngRowNumber)).Clear
  
  For lngAirportIndex = 0 To udtAirports.lngAirportCount - 1
    lngRowCount = lngRowCount + udtAirports.udtAirport(lngAirportIndex).lngServiceAreaCount
  Next
  If lngRowCount = 0 Then GoTo Err_None
  
  varValues = objWorksheet.Range("A3:E" & (2 + lngRowCount)).Value
  lngRowNumber = 1
  For lngAirportIndex = 0 To udtAirports.lngAirportCount - 1
    With udtAirports.udtAirport(lngAirportIndex)
      strAirportCode = .strCode
      For lngAirportServiceAreaIndex = 0 To .lngServiceAreaCount - 1
        With .udtServiceAreaIndexes(lngAirportServiceAreaIndex)
          varValues(lngRowNumber, 1) = strAirportCode
          varValues(lngRowNumber, 2) = udtServiceAreas.udtServiceAreas(.lngServiceAreaIndex).strName
          varValues(lngRowNumber, 3) = .lngPriority
          varValues(lngRowNumber, 4) = IIf(.booIsBase, "TRUE", "FALSE")
          varValues(lngRowNumber, 5) = .dblCMTravelTime
        End With
        lngRowNumber = lngRowNumber + 1
      Next
    End With
  Next
  objWorksheet.Range("A3:E" & (2 + lngRowCount)).Value = varValues
  
Err_None:
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Exporting airport service areas ... done"
  End If
  
  Set exportAirportServiceAreas = objWorksheet
  Set objWorksheet = Nothing

End Function


''''''''''' Equipment Models

Public Function loadEquipmentModels( _
    udtEquipmentModels As EQUIPMENTMODELS_TYPE, udtEquipmentTypes As EQUIPMENTTYPES_TYPE, _
    objWorksheet As Excel.Worksheet, strStepParameter As String) As Boolean

  Dim varValues As Variant
  Dim lngRowNumber As Long, lngEquipmentTypeIndex As Long, strEquipmentTypeName As String
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Loading equipment models ..."
  End If
    
  If Not objWorksheet.Range("A1").Value = "Worksheet Type:" Then Exit Function
  If Not objWorksheet.Range("B1").Value = "Equipment Models" Then Exit Function
  If Not verifyColumnHeadings(objWorksheet, c_strColHeadings_EquipmentModels, 8) Then
    If Not (m_objErrorRange Is Nothing) Then
      m_objErrorRange.Value = "Invalid column headings in EquipmentModels"
    End If
    Exit Function
  End If
  
  lngRowNumber = 10
  lngRowNumber = objWorksheet.Range("A10:A10").CurrentRegion.Rows.Count
  varValues = objWorksheet.Range("A10:C" & (9 + lngRowNumber)).Value
  udtEquipmentTypes.lngEquipmentTypeCount = 0
  ReDim udtEquipmentTypes.udtEquipmentTypes(0 To 15)
  Set udtEquipmentTypes.colEquipmentTypes = New Collection
  
  udtEquipmentModels.lngEquipmentModelCount = UBound(varValues, 1)
  ReDim udtEquipmentModels.udtEquipmentModels(0 To udtEquipmentModels.lngEquipmentModelCount - 1)
  Set udtEquipmentModels.colEquipmentModels = New Collection
  
  For lngRowNumber = 1 To udtEquipmentModels.lngEquipmentModelCount
  
    strEquipmentTypeName = varValues(lngRowNumber, 1)
    On Error Resume Next
    lngEquipmentTypeIndex = -1
    lngEquipmentTypeIndex = udtEquipmentTypes.colEquipmentTypes.Item("ET:" & strEquipmentTypeName)
    On Error GoTo 0
    If lngEquipmentTypeIndex = -1 Then
      lngEquipmentTypeIndex = udtEquipmentTypes.lngEquipmentTypeCount
      If UBound(udtEquipmentTypes.udtEquipmentTypes) < udtEquipmentTypes.lngEquipmentTypeCount Then
        ReDim Preserve udtEquipmentTypes.udtEquipmentTypes(0 To 7 + udtEquipmentTypes.lngEquipmentTypeCount)
      End If
      With udtEquipmentTypes.udtEquipmentTypes(lngEquipmentTypeIndex)
        .lngID = lngEquipmentTypeIndex + 1
        .strName = strEquipmentTypeName
      End With
      udtEquipmentTypes.lngEquipmentTypeCount = udtEquipmentTypes.lngEquipmentTypeCount + 1
      udtEquipmentTypes.colEquipmentTypes.Add lngEquipmentTypeIndex, "ET:" & strEquipmentTypeName
    End If
  
    With udtEquipmentModels.udtEquipmentModels(lngRowNumber - 1)
      .lngID = lngRowNumber
      .lngEquipmentTypeIndex = lngEquipmentTypeIndex
      .strManufacturer = varValues(lngRowNumber, 2)
      .strModel = varValues(lngRowNumber, 3)
      udtEquipmentModels.colEquipmentModels.Add lngRowNumber - 1, "EM:" & .strModel
    End With
  Next
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Loading equipment models ... done"
  End If
  
  loadEquipmentModels = True
  
End Function


Public Function exportEquipmentModels( _
    udtEquipmentModels As EQUIPMENTMODELS_TYPE, udtEquipmentTypes As EQUIPMENTTYPES_TYPE, _
    ByVal objWorkbook As Excel.Workbook, ByVal objAfterWorksheet As Excel.Worksheet, _
    strStepParameter As String) As Excel.Worksheet

  Dim objWorksheet As Excel.Worksheet
  Dim varValues As Variant
  Dim lngRowNumber As Long, lngEquipmentTypeIndex As Long, strEquipmentTypeName As String
  Dim lngEquipmentModelIndex As Long
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Exporting equipment models ..."
  End If
  
  On Error Resume Next
  Set objWorksheet = objWorkbook.Worksheets.Item("EquipmentModels")
  If objWorksheet Is Nothing Then
    If objAfterWorksheet Is Nothing Then
      Set objWorksheet = objWorkbook.Worksheets.Add()
    Else
      Set objWorksheet = objWorkbook.Worksheets.Add(, objAfterWorksheet)
    End If
    objWorksheet.Name = "EquipmentModels"
    exportColumnHeadings objWorksheet, c_strColHeadings_EquipmentModels
  End If
  
  lngRowNumber = objWorksheet.Range("A3:A3").CurrentRegion.Rows.Count
  objWorksheet.Range("A3:C" & (2 + lngRowNumber)).Clear
  
  varValues = objWorksheet.Range("A3:C" & (2 + udtEquipmentModels.lngEquipmentModelCount)).Value
  For lngEquipmentModelIndex = 0 To udtEquipmentModels.lngEquipmentModelCount - 1
    lngRowNumber = lngEquipmentModelIndex + 1
    With udtEquipmentModels.udtEquipmentModels(lngEquipmentModelIndex)
      varValues(lngRowNumber, 1) = udtEquipmentTypes.udtEquipmentTypes(.lngEquipmentTypeIndex).strName
      varValues(lngRowNumber, 2) = .strManufacturer
      varValues(lngRowNumber, 3) = .strModel
    End With
  Next
  objWorksheet.Range("A3:C" & (2 + udtEquipmentModels.lngEquipmentModelCount)).Value = varValues
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Exporting equipment models ... done"
  End If
  
  Set exportEquipmentModels = objWorksheet
  Set objWorksheet = Nothing
  
End Function


''''''''''' Airport Equipment

Public Function loadEquipment( _
    ByRef udtAirports As AIRPORTS_TYPE, udtEquipmentModels As EQUIPMENTMODELS_TYPE, _
    ByRef udtEquipment As EQUIPMENT_TYPE, _
    objWorksheet As Excel.Worksheet, strStepParameter As String) As Boolean

  If strStepParameter = "ByCount" Then
    loadEquipment = loadEquipment_ByCount(udtAirports, udtEquipmentModels, udtEquipment, objWorksheet)
  ElseIf strStepParameter = "ByItem" Then
    loadEquipment = loadEquipment_ByItem(udtAirports, udtEquipmentModels, udtEquipment, objWorksheet)
  Else
    If Not (m_objErrorRange Is Nothing) Then m_objErrorRange.Value = "Invalid parameter in Load equipment"
  End If

End Function


Public Function exportEquipment( _
    udtAirports As AIRPORTS_TYPE, udtEquipmentModels As EQUIPMENTMODELS_TYPE, _
    udtEquipment As EQUIPMENT_TYPE, _
    ByVal objWorkbook As Excel.Workbook, ByVal objAfterWorksheet As Excel.Worksheet, _
    strStepParameter As String) As Excel.Worksheet

  If strStepParameter = "ByCount" Then
    Set exportEquipment = exportEquipment_ByCount(udtAirports, udtEquipmentModels, udtEquipment, objWorkbook, objAfterWorksheet)
  ElseIf strStepParameter = "ByItem" Then
    Set exportEquipment = exportEquipment_ByItem(udtAirports, udtEquipmentModels, udtEquipment, objWorkbook, objAfterWorksheet)
  Else
    If Not (m_objErrorRange Is Nothing) Then m_objErrorRange.Value = "Invalid parameter in Export equipment"
  End If
  
End Function


Private Function loadEquipment_ByCount( _
    udtAirports As AIRPORTS_TYPE, udtEquipmentModels As EQUIPMENTMODELS_TYPE, udtEquipment As EQUIPMENT_TYPE, _
    objWorksheet As Excel.Worksheet) As Boolean

  Dim varValues As Variant
  Dim lngRowNumber As Long, lngAirportIndex As Long, strAirportCode As String, _
      lngAirportEquipmentModelIndex As Long, strMakeModel As String, lngEquipmentModelIndex As Long
  Dim lngIndex As Long
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Loading airport equipment ..."
  End If
  
  If Not verifyColumnHeadings(objWorksheet, c_strColHeadings_AirportEquipment, 8) Then
    If Not (m_objErrorRange Is Nothing) Then
      m_objErrorRange.Value = "Invalid column headings in Airport_Equipment worksheet"
      Set objWorksheet = Nothing
      Exit Function
    End If
  End If
  
  For lngAirportIndex = 0 To udtAirports.lngAirportCount - 1
    With udtAirports.udtAirport(lngAirportIndex)
      .lngEquipmentCount = 0
      ReDim .udtEquipment(0 To 31)
    End With
  Next
  
  udtEquipment.lngEquipmentCount = 0
  lngRowNumber = objWorksheet.Range("A10:A10").CurrentRegion.Rows.Count
  varValues = objWorksheet.Range("A10:C" & (9 + lngRowNumber)).Value
  For lngRowNumber = 1 To UBound(varValues, 1)
  
    If varValues(lngRowNumber, 3) = 0 Then GoTo Label_NoEquipment
    strAirportCode = varValues(lngRowNumber, 1)
    strMakeModel = varValues(lngRowNumber, 2)
    
    On Error Resume Next
    lngAirportIndex = -1
    lngAirportIndex = udtAirports.colAirports.Item("A:" & strAirportCode)
    On Error GoTo 0
    If lngAirportIndex = -1 Then
      If Not (m_objErrorRange Is Nothing) Then m_objErrorRange.Value = "Invalid airport code (" * strAirportCode & ") listed"
      GoTo Label_NoEquipment
    End If
    
    On Error Resume Next
    lngEquipmentModelIndex = -1
    lngEquipmentModelIndex = udtEquipmentModels.colEquipmentModels.Item("EM:" & strMakeModel)
    On Error GoTo 0
    If lngEquipmentModelIndex = -1 Then
      If Not (m_objErrorRange Is Nothing) Then
        m_objErrorRange.Value = "Invalid equipment model (" & strMakeModel & ") listed"
      End If
      GoTo Label_NoEquipment
    End If
      
    With udtAirports.udtAirport(lngAirportIndex)
    
      For lngIndex = 0 To .lngEquipmentCount - 1
        If .udtEquipment(lngIndex).lngEquipmentModelIndex = lngEquipmentModelIndex Then
          If Not (m_objErrorRange Is Nothing) Then
            m_objErrorRange.Value = "Duplicate equipment model listed for airport " & strAirportCode
          End If
          GoTo Label_Abort
        End If
      Next
    
      If UBound(.udtEquipment) < .lngEquipmentCount Then
        ReDim Preserve .udtEquipment(0 To 15 + .lngEquipmentCount)
      End If
      If 0 < varValues(lngRowNumber, 3) Then
        With .udtEquipment(.lngEquipmentCount)
          .lngEquipmentModelIndex = lngEquipmentModelIndex
          .lngCount = varValues(lngRowNumber, 3)
          udtEquipment.lngEquipmentCount = udtEquipment.lngEquipmentCount + .lngCount
        End With
        .lngEquipmentCount = .lngEquipmentCount + 1
      End If
    End With
    
Label_NoEquipment:
  Next
  
  ReDim udtEquipment.udtEquipment(0 To udtEquipment.lngEquipmentCount - 1)
  udtEquipment.lngEquipmentCount = 0
  For lngAirportIndex = 0 To udtAirports.lngAirportCount - 1
    For lngAirportEquipmentModelIndex = 0 To udtAirports.udtAirport(lngAirportIndex).lngEquipmentCount - 1
      With udtAirports.udtAirport(lngAirportIndex).udtEquipment(lngAirportEquipmentModelIndex)
    
        lngEquipmentModelIndex = .lngEquipmentModelIndex
        ReDim .lngEquipmentIndexes(0 To .lngCount - 1)
        For lngIndex = 0 To .lngCount - 1
          .lngEquipmentIndexes(lngIndex) = udtEquipment.lngEquipmentCount
          With udtEquipment.udtEquipment(udtEquipment.lngEquipmentCount)
            .lngID = udtEquipment.lngEquipmentCount + 1
            .lngEquipmentModelIndex = lngEquipmentModelIndex
            .lngAirportIndex = lngAirportIndex
            .lngAirportEquipmentModelIndex = lngAirportEquipmentModelIndex
            .lngAirportEquipmentIndex = lngIndex
          End With
          udtEquipment.lngEquipmentCount = udtEquipment.lngEquipmentCount + 1
        Next
        
      End With
    Next
  Next
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Loading airport equipment ... done"
  End If
  
  loadEquipment_ByCount = True
  
Label_Abort:
  
End Function


Private Function exportEquipment_ByCount( _
    udtAirports As AIRPORTS_TYPE, udtEquipmentModels As EQUIPMENTMODELS_TYPE, udtEquipment As EQUIPMENT_TYPE, _
    ByVal objWorkbook As Excel.Workbook, ByVal objAfterWorksheet As Excel.Worksheet) As Excel.Worksheet

  Dim objWorksheet As Excel.Worksheet
  Dim varValues As Variant
  Dim lngRowNumber As Long, lngAirportIndex As Long, strAirportCode As String, _
      lngAirportEquipmentModelIndex As Long, strMakeModel As String, lngEquipmentModelIndex As Long
  Dim lngIndex As Long, lngRowCount As Long
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Exporting airport equipment ..."
  End If
  
  On Error Resume Next
  Set objWorksheet = objWorkbook.Worksheets.Item("Airport_Equipment")
  If objWorksheet Is Nothing Then
    If objAfterWorksheet Is Nothing Then
      Set objWorksheet = objWorkbook.Worksheets.Add()
    Else
      Set objWorksheet = objWorkbook.Worksheets.Add(, objAfterWorksheet)
    End If
    objWorksheet.Name = "Airport_Equipment"
    exportColumnHeadings objWorksheet, c_strColHeadings_AirportEquipment
  End If
  On Error GoTo 0
  
  lngRowNumber = objWorksheet.Range("A3:A3").CurrentRegion.Rows.Count
  objWorksheet.Range("A3:C" & (2 + lngRowNumber)).Clear
  
  For lngAirportIndex = 0 To udtAirports.lngAirportCount - 1
    lngRowCount = lngRowCount + udtAirports.udtAirport(lngAirportIndex).lngEquipmentCount
  Next
  If lngRowCount = 0 Then GoTo Err_None
  
  lngRowNumber = 1
  varValues = objWorksheet.Range("A3:C" & (2 + lngRowCount)).Value
  For lngAirportIndex = 0 To udtAirports.lngAirportCount - 1
    With udtAirports.udtAirport(lngAirportIndex)
      For lngAirportEquipmentModelIndex = 0 To .lngEquipmentCount - 1
        varValues(lngRowNumber, 1) = .strCode
        With udtEquipmentModels.udtEquipmentModels(.udtEquipment(lngAirportEquipmentModelIndex).lngEquipmentModelIndex)
          varValues(lngRowNumber, 2) = .strManufacturer & "." & .strModel
        End With
        varValues(lngRowNumber, 3) = .udtEquipment(lngAirportEquipmentModelIndex).lngCount
        lngRowNumber = lngRowNumber + 1
      Next
    End With
  Next
  objWorksheet.Range("A3:C" & (2 + lngRowCount)).Value = varValues
  
Err_None:
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Exporting airport equipment ... done"
  End If
  
  Set exportEquipment_ByCount = objWorksheet
  Set objWorksheet = Nothing
  
End Function


Private Function loadEquipment_ByItem( _
    udtAirports As AIRPORTS_TYPE, udtEquipmentModels As EQUIPMENTMODELS_TYPE, udtEquipment As EQUIPMENT_TYPE, _
    objWorksheet As Excel.Worksheet) As Boolean

  Dim varValues As Variant
  Dim lngRowNumber As Long, lngAirportIndex As Long, strAirportCode As String, _
      lngAirportEquipmentModelIndex As Long, strMakeModel As String, lngEquipmentModelIndex As Long, _
      lngAirportEquipmentIndex As Long
  Dim lngEquipmentID As Long
  Dim lngIndex As Long
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Loading airport equipment ..."
  End If
  
  If Not verifyColumnHeadings(objWorksheet, c_strColHeadings_Equipment, 8) Then
    If Not (m_objErrorRange Is Nothing) Then
      m_objErrorRange.Value = "Invalid column headings in Equipment worksheet"
      Set objWorksheet = Nothing
      Exit Function
    End If
  End If
  
  For lngAirportIndex = 0 To udtAirports.lngAirportCount - 1
    With udtAirports.udtAirport(lngAirportIndex)
      .lngEquipmentCount = 0
      ReDim .udtEquipment(0 To 31)
    End With
  Next
  
  lngRowNumber = objWorksheet.Range("A10:A10").CurrentRegion.Rows.Count
  varValues = objWorksheet.Range("A10:C" & (9 + lngRowNumber)).Value
  udtEquipment.lngEquipmentCount = 0
  ReDim udtEquipment.udtEquipment(0 To lngRowNumber)
  For lngRowNumber = 1 To UBound(varValues, 1)
  
    lngEquipmentID = varValues(lngRowNumber, 1)
    If 0 < udtEquipment.lngEquipmentCount Then
      If lngEquipmentID <= udtEquipment.udtEquipment(udtEquipment.lngEquipmentCount - 1).lngID Then
        If Not (m_objErrorRange Is Nothing) Then
          m_objErrorRange.Value = "Invalid equipment ID (" & lngEquipmentID & ") - IDs must be in ascending order."
        End If
        GoTo Label_Abort
      End If
    End If
    If strAirportCode <> varValues(lngRowNumber, 3) Then
      strAirportCode = varValues(lngRowNumber, 3)
      lngAirportIndex = udtAirports.colAirports.Item("A:" & strAirportCode)
    End If
    If strMakeModel <> varValues(lngRowNumber, 2) Then
      strMakeModel = varValues(lngRowNumber, 2)
      On Error Resume Next
      lngEquipmentModelIndex = -1
      lngEquipmentModelIndex = udtEquipmentModels.colEquipmentModels.Item("EM:" & strMakeModel)
      On Error GoTo 0
      If lngEquipmentModelIndex = -1 Then
        If Not (m_objErrorRange Is Nothing) Then
          m_objErrorRange.Value = "Invalid equipment model (" & strMakeModel & ") listed"
        End If
        GoTo Label_Abort
      End If
    End If
      
    With udtAirports.udtAirport(lngAirportIndex)
      For lngAirportEquipmentModelIndex = 0 To .lngEquipmentCount - 1
        If .udtEquipment(lngAirportEquipmentModelIndex).lngEquipmentModelIndex = lngEquipmentModelIndex Then Exit For
      Next
      If lngAirportEquipmentModelIndex < .lngEquipmentCount Then
        With .udtEquipment(lngAirportEquipmentModelIndex)
          lngAirportEquipmentIndex = .lngCount
          If UBound(.lngEquipmentIndexes) < lngAirportEquipmentIndex Then
            ReDim Preserve .lngEquipmentIndexes(0 To 64 + lngAirportEquipmentIndex)
          End If
          .lngEquipmentIndexes(lngAirportEquipmentIndex) = udtEquipment.lngEquipmentCount
          .lngCount = .lngCount + 1
        End With
      Else
        If UBound(.udtEquipment) < .lngEquipmentCount Then
          ReDim Preserve .udtEquipment(0 To 16 + .lngEquipmentCount)
        End If
        With .udtEquipment(.lngEquipmentCount)
          .lngEquipmentModelIndex = lngEquipmentModelIndex
          ReDim .lngEquipmentIndexes(0 To 15)
          lngAirportEquipmentIndex = 0
          .lngEquipmentIndexes(0) = udtEquipment.lngEquipmentCount
          .lngCount = 1
        End With
        .lngEquipmentCount = .lngEquipmentCount + 1
      End If
    End With
      
    With udtEquipment.udtEquipment(udtEquipment.lngEquipmentCount)
      .lngID = lngEquipmentID
      .lngEquipmentModelIndex = lngEquipmentModelIndex
      .lngAirportIndex = lngAirportIndex
      .lngAirportEquipmentModelIndex = lngEquipmentModelIndex
      .lngAirportEquipmentIndex = lngAirportEquipmentIndex
    End With
  
  Next
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Loading equipment ... done"
  End If
  
  loadEquipment_ByItem = True
  
Label_Abort:
  Set objWorksheet = Nothing
  
End Function


Private Function exportEquipment_ByItem( _
    udtAirports As AIRPORTS_TYPE, udtEquipmentModels As EQUIPMENTMODELS_TYPE, _
    udtEquipment As EQUIPMENT_TYPE, _
    ByVal objWorkbook As Excel.Workbook, ByVal objAfterWorksheet As Excel.Worksheet) As Excel.Worksheet

  Dim objWorksheet As Excel.Worksheet
  Dim varValues As Variant
  Dim lngRowNumber As Long, lngAirportIndex As Long, strAirportCode As String, _
      lngAirportEquipmentModelIndex As Long, strMakeModel As String, lngEquipmentModelIndex As Long, _
      lngEquipmentIndex As Long
  Dim lngIndex As Long, lngRowCount As Long
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Exporting equipment ..."
  End If
  
  On Error Resume Next
  Set objWorksheet = objWorkbook.Worksheets.Item("Equipment")
  If objWorksheet Is Nothing Then
    If objAfterWorksheet Is Nothing Then
      Set objWorksheet = objWorkbook.Worksheets.Add()
    Else
      Set objWorksheet = objWorkbook.Worksheets.Add(, objAfterWorksheet)
    End If
    objWorksheet.Name = "Equipment"
    exportColumnHeadings objWorksheet, c_strColHeadings_Equipment
  End If
  On Error GoTo 0
  
  lngRowNumber = objWorksheet.Range("A3:A3").CurrentRegion.Rows.Count
  objWorksheet.Range("A3:C" & (2 + lngRowNumber)).Clear
  
  If udtEquipment.lngEquipmentCount = 0 Then GoTo Err_None
  
  lngRowNumber = 1
  varValues = objWorksheet.Range("A3:C" & (2 + udtEquipment.lngEquipmentCount)).Value
  For lngEquipmentIndex = 0 To udtEquipment.lngEquipmentCount - 1
    With udtEquipment.udtEquipment(lngEquipmentIndex)
      varValues(lngRowNumber, 1) = .lngID
      varValues(lngRowNumber, 2) = udtEquipmentModels.udtEquipmentModels(.lngEquipmentModelIndex).strManufacturer & "." _
          & udtEquipmentModels.udtEquipmentModels(.lngEquipmentModelIndex).strModel
      varValues(lngRowNumber, 3) = udtAirports.udtAirport(.lngAirportIndex).strCode
    End With
    lngRowNumber = lngRowNumber + 1
  Next
  objWorksheet.Range("A3:C" & (2 + udtEquipment.lngEquipmentCount)).Value = varValues
  
Err_None:
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Exporting equipment ... done"
  End If
  
  Set exportEquipment_ByItem = objWorksheet
  Set objWorksheet = Nothing
  
End Function



''''''''''' Equipment PM Requirements

' BUG: Remove ByType and BySchedule
Public Function loadEquipmentPM( _
    udtAirports As AIRPORTS_TYPE, udtEquipmentModels As EQUIPMENTMODELS_TYPE, _
    objWorksheet As Excel.Worksheet, strStepParameter As String) As Boolean
    
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Loading equipment PM requirements ..."
  End If
  
  Select Case strStepParameter
  
    Case "ByType"
      If Not loadEquipmentPM_ByType(udtAirports, udtEquipmentModels, objWorksheet) Then Exit Function
    
    Case "BySchedule"
      'If Not loadEquipmentPM_BySchedule(objWorkbook) Then Exit Function

    Case Else
      If Not (m_objErrorRange Is Nothing) Then m_objErrorRange.Value = "Invalid or missing parameter in load Equipment PM step"
      Exit Function
      
  End Select
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Loading equipment PM requirements ... done"
  End If
  
  loadEquipmentPM = True
    
End Function


Private Function loadEquipmentPM_ByType( _
    udtAirports As AIRPORTS_TYPE, udtEquipmentModels As EQUIPMENTMODELS_TYPE, _
    objWorksheet As Excel.Worksheet) As Boolean

  Dim varValues As Variant
  Dim lngRowNumber As Long, strMakeModel As String, lngEquipmentModelIndex As Long, lngPMIndex As Long, _
      enmPeriodicity As PMPERIODICITY_ENUM
  Dim lngAirportIndex As Long, lngAirportEquipmentModelIndex As Long
  Dim udtNullPMRequirement As PMREQUIREMENTS_TYPE
  Dim lngIndex As Long

  If Not objWorksheet.Range("A1").Value = "Worksheet Type:" Then Exit Function
  If Not objWorksheet.Range("B1").Value = "Equipment PM" Then Exit Function
  If Not verifyColumnHeadings(objWorksheet, c_strColHeadings_EquipmentPM, 8) Then
    If Not (m_objErrorRange Is Nothing) Then
      m_objErrorRange.Value = "Invalid column headings in EquipmentPM worksheet"
    End If
    Set objWorksheet = Nothing
    Exit Function
  End If
  
  ' Initialize structures to indicate empty PM
  For lngEquipmentModelIndex = 0 To udtEquipmentModels.lngEquipmentModelCount - 1
    With udtEquipmentModels.udtEquipmentModels(lngEquipmentModelIndex)
      .enmPeriodicity = PMPER_NoPM
      For lngPMIndex = 1 To c_lngPMPeriodicity_MaxValue
        .udtPM(lngPMIndex) = udtNullPMRequirement
        .dblPMTime(lngPMIndex) = 0#
      Next
    End With
  Next
  
  ' Load equipment PM data
  lngRowNumber = objWorksheet.Range("A10:A10").CurrentRegion.Rows.Count
  varValues = objWorksheet.Range("A10:I" & (9 + lngRowNumber)).Value
  For lngRowNumber = 1 To UBound(varValues, 1)

    strMakeModel = varValues(lngRowNumber, 1)
    lngEquipmentModelIndex = udtEquipmentModels.colEquipmentModels.Item("EM:" & strMakeModel)
    
    With udtEquipmentModels.udtEquipmentModels(lngEquipmentModelIndex)
      
      ' Sort entries in Periodicity order
      enmPeriodicity = decodeName(CStr(varValues(lngRowNumber, 3)), c_strPMPeriodicityList)
      With .udtPM(enmPeriodicity)
        If 0 < Len(.strName) Then
          If Not (m_objErrorRange Is Nothing) Then m_objErrorRange.Value = "Duplicate PM-type for " & strMakeModel
          GoTo Label_Abort
        End If
        .strName = varValues(lngRowNumber, 2)
        .lngAllowedSlack = varValues(lngRowNumber, 4)
        .dblLabor_Initial = varValues(lngRowNumber, 5)
        .dblLabor_Wait = varValues(lngRowNumber, 6)
        .dblLabor_Final = varValues(lngRowNumber, 7)
        .dblConsumables = varValues(lngRowNumber, 8)
        .lngTechnicianCount = varValues(lngRowNumber, 9)
        Select Case enmPeriodicity
          Case PMPER_Weekly
            .lngEventsPerYear = 52
          Case PMPER_Monthly
            .lngEventsPerYear = 12
          Case PMPER_Quarterly
            .lngEventsPerYear = 4
          Case PMPER_SemiAnnually
            .lngEventsPerYear = 2
          Case PMPER_Annually
            .lngEventsPerYear = 1
          Case PMPER_NoPM
            .lngEventsPerYear = 0
          Case Else
            Err.Raise 5
        End Select
      End With
      
      If .enmPeriodicity = PMPER_NoPM Then
        .enmPeriodicity = enmPeriodicity
      ElseIf enmPeriodicity < .enmPeriodicity Then
        .enmPeriodicity = enmPeriodicity
      End If

    End With
  Next
  
  ' Correct PM events per year for less frequent events
  For lngEquipmentModelIndex = 0 To udtEquipmentModels.lngEquipmentModelCount - 1
    With udtEquipmentModels.udtEquipmentModels(lngEquipmentModelIndex)
    
      For lngPMIndex = c_lngPMPeriodicity_MaxValue - 1 To 1 Step -1
        If 0 < .udtPM(lngPMIndex).lngEventsPerYear Then
          For lngIndex = lngPMIndex + 1 To c_lngPMPeriodicity_MaxValue
            .udtPM(lngPMIndex).lngEventsPerYear = .udtPM(lngPMIndex).lngEventsPerYear - .udtPM(lngIndex).lngEventsPerYear
          Next
          .dblPMTime(lngPMIndex) = .udtPM(lngPMIndex).lngEventsPerYear * (.udtPM(lngPMIndex).dblLabor_Initial _
              + .udtPM(lngPMIndex).dblLabor_Wait + .udtPM(lngPMIndex).dblLabor_Final)
        End If
      Next
      
    End With
  Next

  ' Create PM schedule
  For lngEquipmentModelIndex = 0 To udtEquipmentModels.lngEquipmentModelCount - 1
    With udtEquipmentModels.udtEquipmentModels(lngEquipmentModelIndex)

      Select Case .enmPeriodicity
      
        Case PMPER_NoPM
          ' Do nothing
      
        Case PMPER_Weekly
          ReDim .lngPMSchedule(0 To 51)
          For lngIndex = 0 To 51
            .lngPMSchedule(lngIndex) = PMPER_Weekly
          Next
          If 0 < .udtPM(PMPER_Monthly).lngEventsPerYear Then
            .lngPMSchedule(0) = PMPER_Monthly
            .lngPMSchedule(5) = PMPER_Monthly
            .lngPMSchedule(9) = PMPER_Monthly
            .lngPMSchedule(13) = PMPER_Monthly
            .lngPMSchedule(18) = PMPER_Monthly
            .lngPMSchedule(22) = PMPER_Monthly
            .lngPMSchedule(26) = PMPER_Monthly
            .lngPMSchedule(31) = PMPER_Monthly
            .lngPMSchedule(35) = PMPER_Monthly
            .lngPMSchedule(39) = PMPER_Monthly
            .lngPMSchedule(44) = PMPER_Monthly
            .lngPMSchedule(48) = PMPER_Monthly
          End If
          If 0 < .udtPM(PMPER_Quarterly).lngEventsPerYear Then
            .lngPMSchedule(0) = PMPER_Quarterly
            .lngPMSchedule(13) = PMPER_Quarterly
            .lngPMSchedule(26) = PMPER_Quarterly
            .lngPMSchedule(39) = PMPER_Quarterly
          End If
          If 0 < .udtPM(PMPER_SemiAnnually).lngEventsPerYear Then
            .lngPMSchedule(0) = PMPER_SemiAnnually
            .lngPMSchedule(26) = PMPER_SemiAnnually
          End If
          If 0 < .udtPM(PMPER_Annually).lngEventsPerYear Then
            .lngPMSchedule(0) = PMPER_Annually
          End If
                
        Case PMPER_Monthly
          ReDim .lngPMSchedule(0 To 11)
          For lngIndex = 0 To 11
            .lngPMSchedule(lngIndex) = PMPER_Monthly
          Next
          If 0 < .udtPM(PMPER_Quarterly).lngEventsPerYear Then
            .lngPMSchedule(0) = PMPER_Quarterly
            .lngPMSchedule(3) = PMPER_Quarterly
            .lngPMSchedule(6) = PMPER_Quarterly
            .lngPMSchedule(9) = PMPER_Quarterly
          End If
          If 0 < .udtPM(PMPER_SemiAnnually).lngEventsPerYear Then
            .lngPMSchedule(0) = PMPER_SemiAnnually
            .lngPMSchedule(6) = PMPER_SemiAnnually
          End If
          If 0 < .udtPM(PMPER_Annually).lngEventsPerYear Then
            .lngPMSchedule(0) = PMPER_Annually
          End If

        Case PMPER_Quarterly
          ReDim .lngPMSchedule(0 To 3)
          For lngIndex = 0 To 3
            .lngPMSchedule(lngIndex) = PMPER_Quarterly
          Next
          If 0 < .udtPM(PMPER_SemiAnnually).lngEventsPerYear Then
            .lngPMSchedule(0) = PMPER_SemiAnnually
            .lngPMSchedule(2) = PMPER_SemiAnnually
          End If
          If 0 < .udtPM(PMPER_Annually).lngEventsPerYear Then
            .lngPMSchedule(0) = PMPER_Annually
          End If
          
        Case PMPER_SemiAnnually
          ReDim .lngPMSchedule(0 To 1)
          For lngIndex = 0 To 1
            .lngPMSchedule(lngIndex) = PMPER_SemiAnnually
          Next
          If 0 < .udtPM(PMPER_Annually).lngEventsPerYear Then
            .lngPMSchedule(0) = PMPER_Annually
          End If
        
        Case PMPER_Annually
          ReDim .lngPMSchedule(0 To 0)
          .lngPMSchedule(0) = PMPER_Annually

        Case Else
          Err.Raise 5

      End Select
      
    End With
  Next
 
  
  ' Set airport PM periodicity
  For lngAirportIndex = 0 To udtAirports.lngAirportCount - 1
    With udtAirports.udtAirport(lngAirportIndex)
    
      .enmPMPeriodicity = c_lngPMPeriodicity_MaxValue + 1
      For lngPMIndex = 1 To c_lngPMPeriodicity_MaxValue
        .dblPMTime(lngPMIndex) = 0#
      Next
      
      For lngAirportEquipmentModelIndex = 0 To .lngEquipmentCount - 1
        lngEquipmentModelIndex = .udtEquipment(lngAirportEquipmentModelIndex).lngEquipmentModelIndex
        If udtEquipmentModels.udtEquipmentModels(lngEquipmentModelIndex).enmPeriodicity <> PMPER_NoPM Then
          If udtEquipmentModels.udtEquipmentModels(lngEquipmentModelIndex).enmPeriodicity < .enmPMPeriodicity Then
            .enmPMPeriodicity = udtEquipmentModels.udtEquipmentModels(lngEquipmentModelIndex).enmPeriodicity
          End If
          For lngPMIndex = 1 To c_lngPMPeriodicity_MaxValue
            .dblPMTime(lngPMIndex) = .dblPMTime(lngPMIndex) + .udtEquipment(lngAirportEquipmentModelIndex).lngCount _
                * udtEquipmentModels.udtEquipmentModels(lngEquipmentModelIndex).udtPM(lngPMIndex).lngEventsPerYear * _
                    (udtEquipmentModels.udtEquipmentModels(lngEquipmentModelIndex).udtPM(lngPMIndex).dblLabor_Initial _
                    + udtEquipmentModels.udtEquipmentModels(lngEquipmentModelIndex).udtPM(lngPMIndex).dblLabor_Wait _
                    + udtEquipmentModels.udtEquipmentModels(lngEquipmentModelIndex).udtPM(lngPMIndex).dblLabor_Final)
          Next
        End If
      Next
      If c_lngPMPeriodicity_MaxValue < .enmPMPeriodicity Then .enmPMPeriodicity = PMPER_NoPM
    End With
  Next
  
  loadEquipmentPM_ByType = True
  
Label_Abort:
  Set objWorksheet = Nothing
  
End Function


Public Function exportEquipmentPM( _
    udtEquipmentModels As EQUIPMENTMODELS_TYPE, _
    ByVal objWorkbook As Excel.Workbook, ByVal objAfterWorksheet As Excel.Worksheet, _
    strStepParameter As String) As Excel.Worksheet
  
  Select Case strStepParameter
  
    Case "ByType"
      Set exportEquipmentPM = exportEquipmentPM_ByType(udtEquipmentModels, objWorkbook, objAfterWorksheet)
    
    Case "BySchedule"
      'set exportEquipmentPM =  exportEquipmentPM_BySchedule(objWorkbook)

    Case Else
      If Not (m_objStatusRange Is Nothing) Then
        m_objStatusRange.Value = "Exporting equipment PM requirements ..."
      End If
      If Not (m_objErrorRange Is Nothing) Then m_objErrorRange.Value = "Invalid or missing parameter in export Equipment PM step"
      
  End Select

End Function


Public Function exportEquipmentPM_ByType( _
    udtEquipmentModels As EQUIPMENTMODELS_TYPE, _
    ByVal objWorkbook As Excel.Workbook, ByVal objAfterWorksheet As Excel.Worksheet) As Excel.Worksheet

  Dim objWorksheet As Excel.Worksheet
  Dim varValues As Variant, lngRowCount As Long, strMakeModel As String
  Dim lngRowNumber As Long, lngEquipmentModelIndex As Long, lngPMIndex As Long, _
      enmPeriodicity As PMPERIODICITY_ENUM
  Dim lngIndex As Long
 
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Exporting equipment PM requirements (by type) ..."
  End If
 
  On Error Resume Next
  Set objWorksheet = objWorkbook.Worksheets.Item("EquipmentPM")
  If objWorksheet Is Nothing Then
    If objAfterWorksheet Is Nothing Then
      Set objWorksheet = objWorkbook.Worksheets.Add()
    Else
      Set objWorksheet = objWorkbook.Worksheets.Add(, objAfterWorksheet)
    End If
    objWorksheet.Name = "EquipmentPM"
    exportColumnHeadings objWorksheet, c_strColHeadings_EquipmentPM
  End If

  lngRowNumber = objWorksheet.Range("A3:A3").CurrentRegion.Rows.Count
  objWorksheet.Range("A3:I" & (2 + lngRowNumber)).Clear

  For lngEquipmentModelIndex = 0 To udtEquipmentModels.lngEquipmentModelCount - 1
    With udtEquipmentModels.udtEquipmentModels(lngEquipmentModelIndex)
      For lngPMIndex = 1 To c_lngPMPeriodicity_MaxValue
        If 0 < .udtPM(lngPMIndex).lngEventsPerYear Then lngRowCount = lngRowCount + 1
      Next
    End With
  Next
  
  lngRowNumber = 1
  varValues = objWorksheet.Range("A3:I" & (2 + lngRowCount)).Value
  For lngEquipmentModelIndex = 0 To udtEquipmentModels.lngEquipmentModelCount - 1
    With udtEquipmentModels.udtEquipmentModels(lngEquipmentModelIndex)
      strMakeModel = .strManufacturer & "." & .strModel
      For lngPMIndex = 1 To c_lngPMPeriodicity_MaxValue
        If 0 < .udtPM(lngPMIndex).lngEventsPerYear Then
          With .udtPM(lngPMIndex)
            varValues(lngRowNumber, 1) = strMakeModel
            varValues(lngRowNumber, 2) = .strName
            varValues(lngRowNumber, 3) = encodeName(lngPMIndex, c_strPMPeriodicityList)
            varValues(lngRowNumber, 4) = .lngAllowedSlack
            varValues(lngRowNumber, 5) = .dblLabor_Initial
            varValues(lngRowNumber, 6) = .dblLabor_Wait
            varValues(lngRowNumber, 7) = .dblLabor_Final
            varValues(lngRowNumber, 8) = .dblConsumables
            varValues(lngRowNumber, 9) = .lngTechnicianCount
          End With
          lngRowNumber = lngRowNumber + 1
        End If
      Next
    End With
  Next
  objWorksheet.Range("A3:I" & (2 + lngRowCount)).Value = varValues
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Exporting equipment PM requirements (by type) ... done"
  End If
  
  Set exportEquipmentPM_ByType = objWorksheet
  
End Function


''''''''''' Equipment CM Expectations

Public Function loadCMRequirements( _
    udtCMRequirements As CMREQUIREMENTS_TYPE, _
    objWorksheet As Excel.Worksheet, strStepParameter As String) As Boolean

  Dim varValues As Variant
  Dim lngRowNumber As Long, strMakeModel As String, lngEquipmentModelIndex As Long
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Loading equipment CM requirements ..."
  End If
  
  If Not verifyColumnHeadings(objWorksheet, c_strColHeadings_EquipmentCM, 8) Then
    If Not (m_objErrorRange Is Nothing) Then
      m_objErrorRange.Value = "Invalid column headings in EquipmentCM worksheet"
   End If
   Exit Function
 End If
  
  lngRowNumber = objWorksheet.Range("A10:A10").CurrentRegion.Rows.Count
  varValues = objWorksheet.Range("A10:N" & (9 + lngRowNumber)).Value
  udtCMRequirements.lngCMRequirementCount = 0
  ReDim udtCMRequirements.udtCMRequirements(0 To lngRowNumber)
  Set udtCMRequirements.colCMRequirement = New Collection
  For lngRowNumber = 1 To UBound(varValues, 1)
    If varValues(lngRowNumber, 1) <> "" Then

      With udtCMRequirements.udtCMRequirements(udtCMRequirements.lngCMRequirementCount)
        .strModelNum = varValues(lngRowNumber, 1)
        .strName = varValues(lngRowNumber, 2)
        .dblFrequency = varValues(lngRowNumber, 3)
        .udtCMTime.dblAvg = varValues(lngRowNumber, 4)
        .udtCMTime.dblStndDev = varValues(lngRowNumber, 5)
        .udtCMTime.dblMin = varValues(lngRowNumber, 6)
        .udtCMTime.dblMax = varValues(lngRowNumber, 7)
        .dblPartsCost = varValues(lngRowNumber, 8)
        .udtPartsTime.dblAvg = varValues(lngRowNumber, 9)
        .udtPartsTime.dblStndDev = varValues(lngRowNumber, 10)
        .udtPartsTime.dblMin = varValues(lngRowNumber, 11)
        .udtPartsTime.dblMax = varValues(lngRowNumber, 12)
        .dblConsumablesCost = varValues(lngRowNumber, 13)
        .lngTechnicianCount = varValues(lngRowNumber, 14)
      End With
      udtCMRequirements.colCMRequirement.Add udtCMRequirements.lngCMRequirementCount, "I:" & CStr(varValues(lngRowNumber, 2))
      udtCMRequirements.lngCMRequirementCount = udtCMRequirements.lngCMRequirementCount + 1
    
    End If
  Next
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Loading equipment CM requirements ... done"
  End If
  
  loadCMRequirements = True
  
End Function


Public Function applyCMRequirements(udtCMRequirements As CMREQUIREMENTS_TYPE, udtAirports As AIRPORTS_TYPE, _
    udtEquipmentModels As EQUIPMENTMODELS_TYPE) As Boolean

  Dim lngAirportIndex As Long, lngAirportEquipmentIndex As Long, lngCMRequirementIndex As Long
  Dim strLocation As String, strCategory As String, strModelNum As String, booFail As Boolean
  
  booFail = False
  For lngAirportIndex = 0 To udtAirports.lngAirportCount - 1
    With udtAirports.udtAirport(lngAirportIndex)
      strLocation = .strCode
      strCategory = .strCat
    
      For lngAirportEquipmentIndex = 0 To .lngEquipmentCount - 1
        With .udtEquipment(lngAirportEquipmentIndex)
          strModelNum = udtEquipmentModels.udtEquipmentModels(.lngEquipmentModelIndex).strModel
          
          On Error Resume Next
          lngCMRequirementIndex = -1
          lngCMRequirementIndex = udtCMRequirements.colCMRequirement("I:" & strModelNum & ":" & strLocation)
          If lngCMRequirementIndex = -1 Then
            lngCMRequirementIndex = udtCMRequirements.colCMRequirement("I:" & strModelNum & ":" & strCategory)
          End If
          If lngCMRequirementIndex = -1 Then
            lngCMRequirementIndex = udtCMRequirements.colCMRequirement("I:" & strModelNum)
          End If
          On Error GoTo 0
          .lngCMRequirementIndex = lngCMRequirementIndex
          If lngCMRequirementIndex = -1 Then Err.Raise 5
          
        End With
      Next
      
    End With
  Next
          
  applyCMRequirements = Not booFail

End Function




Public Function exportEquipmentCM( _
    udtCMRequirements As CMREQUIREMENTS_TYPE, _
    ByVal objWorkbook As Excel.Workbook, ByVal objAfterWorksheet As Excel.Worksheet, _
    strStepParameter As String) As Excel.Worksheet

  Dim objWorksheet As Excel.Worksheet
  Dim varValues As Variant, lngRowCount As Long
  Dim lngRowNumber As Long, strMakeModel As String, lngEquipmentModelIndex As Long, lngIndex As Long
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Exporting equipment CM requirements ..."
  End If
  
  On Error Resume Next
  Set objWorksheet = objWorkbook.Worksheets.Item("EquipmentCM")
  If objWorksheet Is Nothing Then
    If objAfterWorksheet Is Nothing Then
      Set objWorksheet = objWorkbook.Worksheets.Add()
    Else
      Set objWorksheet = objWorkbook.Worksheets.Add(, objAfterWorksheet)
    End If
    objWorksheet.Name = "EquipmentCM"
    exportColumnHeadings objWorksheet, c_strColHeadings_EquipmentCM
  End If

  lngRowNumber = objWorksheet.Range("A3:A3").CurrentRegion.Rows.Count
  objWorksheet.Range("A3:N" & (2 + lngRowNumber)).Clear
  
  lngRowCount = udtCMRequirements.lngCMRequirementCount
  If lngRowCount = 0 Then GoTo Err_None
  
  lngRowNumber = 1
  varValues = objWorksheet.Range("A3:N" & (2 + lngRowCount)).Value
  For lngEquipmentModelIndex = 0 To udtCMRequirements.lngCMRequirementCount - 1
    With udtCMRequirements.udtCMRequirements(lngEquipmentModelIndex)
      varValues(lngRowNumber, 1) = .strModelNum
      varValues(lngRowNumber, 2) = .strName
      varValues(lngRowNumber, 3) = .dblFrequency
      varValues(lngRowNumber, 4) = .udtCMTime.dblAvg
      varValues(lngRowNumber, 5) = .udtCMTime.dblStndDev
      varValues(lngRowNumber, 6) = .udtCMTime.dblMin
      varValues(lngRowNumber, 7) = .udtCMTime.dblMax
      varValues(lngRowNumber, 8) = .dblPartsCost
      varValues(lngRowNumber, 9) = .udtPartsTime.dblAvg
      varValues(lngRowNumber, 10) = .udtPartsTime.dblStndDev
      varValues(lngRowNumber, 11) = .udtPartsTime.dblMin
      varValues(lngRowNumber, 12) = .udtPartsTime.dblMax
      varValues(lngRowNumber, 13) = .dblConsumablesCost
      varValues(lngRowNumber, 14) = .lngTechnicianCount
      lngRowNumber = lngRowNumber + 1
    End With
  Next
  objWorksheet.Range("A3:N" & (2 + lngRowCount)).Value = varValues
  
Err_None:
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Exporting equipment CM requirements ... done"
  End If
  
  Set exportEquipmentCM = objWorksheet
  Set objWorksheet = Nothing
  
End Function


'''''''''' Model steps
''''''
''

' WORK IN PROGRESS

Public Function loadPMStatus( _
    udtAirports As AIRPORTS_TYPE, udtEquipmentModels As EQUIPMENTMODELS_TYPE, _
    udtEquipment As EQUIPMENT_TYPE, _
    objWorkbook As Excel.Workbook, strStepParameter As String) As Boolean

  Dim lngStepNumber As Long

  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Loading PM status ..."
  End If
  
  lngStepNumber = recordModelStep(MDLSTEP_loadPMStatus, strStepParameter)
  If m_udtModelSteps(lngStepNumber).lngParameterCount <> 0 Then
    recordError m_udtModelSteps(lngStepNumber), "Invalid parameter"
    Exit Function
  End If
  
  ' BUG: Should move loadPMStatus_Import code into this function
  If loadPMStatus_Import(udtAirports, udtEquipmentModels, udtEquipment, objWorkbook) Then
    If Not (m_objStatusRange Is Nothing) Then
      m_objStatusRange.Value = "Loading PM status ... done"
    End If
    m_udtModelSteps(lngStepNumber).strStatus = "Done"
    loadPMStatus = True
  Else
    recordError m_udtModelSteps(lngStepNumber), "Error"
  End If

End Function


Public Function createPMStatus( _
    udtAirports As AIRPORTS_TYPE, udtEquipmentModels As EQUIPMENTMODELS_TYPE, _
    udtEquipment As EQUIPMENT_TYPE, _
    objWorkbook As Excel.Workbook, strStepParameter As String) As Boolean

  Dim lngStepNumber As Long
  Dim lngParameterCount As Long, strParameters() As String
  Dim lngAirportIndex As Long
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Creating PM status ..."
  End If
  lngStepNumber = recordModelStep(MDLSTEP_createPMStatus, strStepParameter)
  
  If m_udtModelSteps(lngStepNumber).lngParameterCount <> 2 Then
    recordError m_udtModelSteps(lngStepNumber), "Invalid parameters"
    Exit Function
  End If
  
  Select Case strParameters(1)
    Case "Random"
      For lngAirportIndex = 0 To udtAirports.lngAirportCount - 1
        createPMStatus_Random udtAirports.udtAirport(lngAirportIndex), udtEquipmentModels, udtEquipment, m_dteModelStartDate
      Next
    Case "Synchronize"
      For lngAirportIndex = 0 To udtAirports.lngAirportCount - 1
        createPMStatus_Synchronize udtAirports.udtAirport(lngAirportIndex), udtEquipmentModels, udtEquipment, m_dteModelStartDate, CDbl(strParameters(2))
      Next
    Case "Optimize"
      For lngAirportIndex = 0 To udtAirports.lngAirportCount - 1
        If udtAirports.udtAirport(lngAirportIndex).lngBaseServiceAreaIndex <> -1 Then
          createPMStatus_Random udtAirports.udtAirport(lngAirportIndex), udtEquipmentModels, udtEquipment, m_dteModelStartDate
        Else
          createPMStatus_Synchronize udtAirports.udtAirport(lngAirportIndex), udtEquipmentModels, udtEquipment, m_dteModelStartDate, CDbl(strParameters(2))
        End If
      Next
    Case Else
      If Not (m_objErrorRange Is Nothing) Then
        m_objErrorRange.Value = "Invalid parameter for load PM Status step"
        Exit Function
      End If
  End Select

  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Loading PM status ... done"
  End If
  m_udtModelSteps(lngStepNumber).strStatus = "Done"
  
  createPMStatus = True

End Function


' Sets the most recent PM date for the most frequent PM requirement to dteLastPMDate and sets the
' remaining PM requirements to a random previous date
'Private Sub setEquipmentPMStatus(udtEquipment As EQUIPMENT_TYPE, ByVal dteLastPMDate As Date)
'
'  Dim lngPMIndex As Long, lngEquipmentModelIndex As Long
'
'  lngEquipmentModelIndex = udtEquipment.lngEquipmentModelIndex
'  With udtEquipment.udtPMSchedule
'
'    For lngPMIndex = 1 To c_lngPMPeriodicity_MaxValue
'      .dteLastPMCompleted(lngPMIndex) = 0#
'    Next
'
'    Select Case m_udtEquipmentModels(udtEquipment.lngEquipmentModelIndex).enmPeriodicity
'
'      Case PMPER_NoPM
'        ' Do nothing
'
'      Case PMPER_Weekly
'        .dteLastPMCompleted(PMPER_Weekly) = dteLastPMDate
'        If 0 < m_udtEquipmentModels(lngEquipmentModelIndex).udtPM(PMPER_Monthly).lngEventsPerYear Then
'          .dteLastPMCompleted(PMPER_Monthly) = .dteLastPMCompleted(PMPER_Weekly) - 7 * Fix(4 * Rnd())
'        End If
'        If 0 < m_udtEquipmentModels(lngEquipmentModelIndex).udtPM(PMPER_Quarterly).lngEventsPerYear Then
'          .dteLastPMCompleted(PMPER_Quarterly) = .dteLastPMCompleted(PMPER_Weekly) - 7 * Fix(13 * Rnd())
'        End If
'        If 0 < m_udtEquipmentModels(lngEquipmentModelIndex).udtPM(PMPER_SemiAnnually).lngEventsPerYear Then
'          .dteLastPMCompleted(PMPER_SemiAnnually) = .dteLastPMCompleted(PMPER_Weekly) - 7 * Fix(26 * Rnd())
'        End If
'        If 0 < m_udtEquipmentModels(lngEquipmentModelIndex).udtPM(PMPER_Annually).lngEventsPerYear Then
'          .dteLastPMCompleted(PMPER_Annually) = .dteLastPMCompleted(PMPER_Weekly) - 7 * Fix(52 * Rnd())
'        End If
'
'      Case PMPER_Monthly
'        .dteLastPMCompleted(PMPER_Monthly) = dteLastPMDate
'        If 0 < m_udtEquipmentModels(lngEquipmentModelIndex).udtPM(PMPER_Quarterly).lngEventsPerYear Then
'          .dteLastPMCompleted(PMPER_Quarterly) = .dteLastPMCompleted(PMPER_Monthly) - 30 * Fix(3 * Rnd())
'        End If
'        If 0 < m_udtEquipmentModels(lngEquipmentModelIndex).udtPM(PMPER_SemiAnnually).lngEventsPerYear Then
'          .dteLastPMCompleted(PMPER_SemiAnnually) = .dteLastPMCompleted(PMPER_Monthly) - 30 * Fix(6 * Rnd())
'        End If
'        If 0 < m_udtEquipmentModels(lngEquipmentModelIndex).udtPM(PMPER_Annually).lngEventsPerYear Then
'          .dteLastPMCompleted(PMPER_Annually) = .dteLastPMCompleted(PMPER_Monthly) - 30 * Fix(12 * Rnd())
'        End If
'
'      Case PMPER_Quarterly
'        .dteLastPMCompleted(PMPER_Quarterly) = dteLastPMDate
'        If 0 < m_udtEquipmentModels(lngEquipmentModelIndex).udtPM(PMPER_SemiAnnually).lngEventsPerYear Then
'          .dteLastPMCompleted(PMPER_SemiAnnually) = .dteLastPMCompleted(PMPER_Quarterly) - 91 * Fix(2 * Rnd())
'        End If
'        If 0 < m_udtEquipmentModels(lngEquipmentModelIndex).udtPM(PMPER_Annually).lngEventsPerYear Then
'          .dteLastPMCompleted(PMPER_Annually) = .dteLastPMCompleted(PMPER_Quarterly) - 91 * Fix(4 * Rnd())
'        End If
'
'      Case PMPER_SemiAnnually
'        .dteLastPMCompleted(PMPER_SemiAnnually) = dteLastPMDate
'        If 0 < m_udtEquipmentModels(lngEquipmentModelIndex).udtPM(PMPER_Annually).lngEventsPerYear Then
'          .dteLastPMCompleted(PMPER_Annually) = .dteLastPMCompleted(PMPER_SemiAnnually) - 182 * Fix(2 * Rnd())
'        End If
'
'      Case PMPER_Annually
'        .dteLastPMCompleted(PMPER_Annually) = dteLastPMDate
'
'      Case Else
'        Err.Raise 5
'
'    End Select
'  End With
'
'End Sub


Private Sub createPMStatus_Random( _
    ByRef udtAirport As AIRPORT_TYPE, udtEquipmentModels As EQUIPMENTMODELS_TYPE, _
    ByRef udtEquipment As EQUIPMENT_TYPE, _
    ByVal dteStartDate As Date)
      
  Dim lngAirportEquipmentModelIndex As Long, lngEquipmentModelIndex As Long, _
      enmPMPeriodicity As PMPERIODICITY_ENUM, lngAirportEquipmentIndex As Long, _
      lngEquipmentIndex As Long
      
  ' OPTIMIZE
  For lngAirportEquipmentModelIndex = 0 To udtAirport.lngEquipmentCount - 1
    
    lngEquipmentModelIndex = udtAirport.udtEquipment(lngAirportEquipmentModelIndex).lngEquipmentModelIndex
    enmPMPeriodicity = udtEquipmentModels.udtEquipmentModels(lngEquipmentModelIndex).enmPeriodicity
      
    For lngAirportEquipmentIndex = 0 To udtAirport.udtEquipment(lngAirportEquipmentModelIndex).lngCount - 1
      lngEquipmentIndex = udtAirport.udtEquipment(lngAirportEquipmentModelIndex).lngEquipmentIndexes(lngAirportEquipmentIndex)
      
        Select Case enmPMPeriodicity
        
          Case PMPER_NoPM
            With udtEquipment.udtEquipment(lngEquipmentIndex).udtPMSchedule
              .enmPeriodicity = PMPER_NoPM
              .dteLastPMCompleted = 0#
              .lngMonth = 0
              .lngDay = 0
              .lngPMScheduleIndex = 0
            End With
        
          Case PMPER_Weekly
            With udtEquipment.udtEquipment(lngEquipmentIndex).udtPMSchedule
              .dteLastPMCompleted = dteStartDate - 1 - Fix(7 * Rnd())
              .enmPeriodicity = PMPER_Weekly
              .lngMonth = 0
              .lngDay = Weekday(.dteLastPMCompleted)
              .lngPMScheduleIndex = Fix(52 * Rnd())
            End With

          Case PMPER_Monthly
            With udtEquipment.udtEquipment(lngEquipmentIndex)
              .udtPMSchedule.enmPeriodicity = PMPER_Monthly
              .udtPMSchedule.lngMonth = 0
              .udtPMSchedule.lngDay = 1 + Fix(31 * Rnd())
              .udtPMSchedule.lngPMScheduleIndex = Fix(12 * Rnd())
              .udtPMSchedule.dteLastPMCompleted = computePMEventDate(dteStartDate, .udtPMSchedule, 0)
              If dteStartDate <= .udtPMSchedule.dteLastPMCompleted Then
                .udtPMSchedule.dteLastPMCompleted = moveDateMonth(.udtPMSchedule.dteLastPMCompleted, -1)
              End If
            End With
          
          Case PMPER_Quarterly
            With udtEquipment.udtEquipment(lngEquipmentIndex)
              .udtPMSchedule.enmPeriodicity = PMPER_Quarterly
              .udtPMSchedule.lngMonth = Fix(3 * Rnd())
              .udtPMSchedule.lngDay = 1 + Fix(31 * Rnd())
              .udtPMSchedule.lngPMScheduleIndex = Fix(4 * Rnd())
              .udtPMSchedule.dteLastPMCompleted = computePMEventDate(dteStartDate, .udtPMSchedule, 0)
              If dteStartDate <= .udtPMSchedule.dteLastPMCompleted Then
                .udtPMSchedule.dteLastPMCompleted = moveDateMonth(.udtPMSchedule.dteLastPMCompleted, -3)
              End If
            End With
          
          Case PMPER_SemiAnnually
            With udtEquipment.udtEquipment(lngEquipmentIndex)
              .udtPMSchedule.enmPeriodicity = PMPER_SemiAnnually
              .udtPMSchedule.lngMonth = Fix(6 * Rnd())
              .udtPMSchedule.lngDay = 1 + Fix(31 * Rnd())
              .udtPMSchedule.lngPMScheduleIndex = Fix(2 * Rnd())
              .udtPMSchedule.dteLastPMCompleted = computePMEventDate(dteStartDate, .udtPMSchedule, 0)
              If dteStartDate <= .udtPMSchedule.dteLastPMCompleted Then
                .udtPMSchedule.dteLastPMCompleted = moveDateMonth(.udtPMSchedule.dteLastPMCompleted, -6)
              End If
            End With
          
          Case PMPER_Annually
            With udtEquipment.udtEquipment(lngEquipmentIndex)
              .udtPMSchedule.enmPeriodicity = PMPER_SemiAnnually
              .udtPMSchedule.lngMonth = Fix(12 * Rnd())
              .udtPMSchedule.lngDay = 1 + Fix(31 * Rnd())
              .udtPMSchedule.lngPMScheduleIndex = 0
              .udtPMSchedule.dteLastPMCompleted = computePMEventDate(dteStartDate, .udtPMSchedule, 0)
              If dteStartDate <= .udtPMSchedule.dteLastPMCompleted Then
                .udtPMSchedule.dteLastPMCompleted = moveDateMonth(.udtPMSchedule.dteLastPMCompleted, -12)
              End If
            End With

          Case Else
            Err.Raise 5
            
        End Select

    Next

  Next

End Sub


' Synchronization occurs separately for weekly and month events
Private Sub createPMStatus_Synchronize( _
    ByRef udtAirport As AIRPORT_TYPE, udtEquipmentModels As EQUIPMENTMODELS_TYPE, _
    ByRef udtEquipment As EQUIPMENT_TYPE, _
    ByVal dteStartDate As Date, _
    ByVal dblMaxPMPerSync As Double)
      
  Dim lngAirportEquipmentModelIndex As Long, lngEquipmentModelIndex As Long, _
      enmPMPeriodicity As PMPERIODICITY_ENUM, lngAirportEquipmentIndex As Long, _
      lngEquipmentIndex As Long, lngPMCount As Long
  Dim booPMPeriodicityIsPresent(1 To c_lngPMPeriodicity_MaxValue)
  Dim udtPMSchedules() As PMSCHEDULE_TYPE, lngPMEventCount As Long, _
      lngLastPMEvent(1 To c_lngPMPeriodicity_MaxValue) As Long, _
      lngPMSubEventPeriod As Long
  Dim lngIndex As Long

  ' Determine PM frequencies in use and configure NoPM equipment
  For lngAirportEquipmentModelIndex = 0 To udtAirport.lngEquipmentCount - 1
    With udtAirport.udtEquipment(lngAirportEquipmentModelIndex)
      enmPMPeriodicity = udtEquipmentModels.udtEquipmentModels(.lngEquipmentModelIndex).enmPeriodicity
      If enmPMPeriodicity = PMPER_NoPM Then
        For lngAirportEquipmentIndex = 0 To .lngCount - 1
          With udtEquipment.udtEquipment(.lngEquipmentIndexes(lngAirportEquipmentIndex)).udtPMSchedule
            .enmPeriodicity = PMPER_NoPM
            .lngMonth = 0
            .lngDay = 0
            .lngPMScheduleIndex = 0
            .dteLastPMCompleted = 0#
          End With
        Next
      Else
        booPMPeriodicityIsPresent(udtEquipmentModels.udtEquipmentModels(.lngEquipmentModelIndex).enmPeriodicity) = True
      End If
    End With
  Next

  ' Synchronize weekly-based PM
  If booPMPeriodicityIsPresent(PMPER_Weekly) Then
  
    ' Set up the PM schedule
    lngPMEventCount = 1 + Fix(udtAirport.dblPMTime(PMPER_Weekly) / dblMaxPMPerSync)
    If 3 < lngPMEventCount Then lngPMEventCount = 7
    lngPMSubEventPeriod = 7 \ lngPMEventCount
    ReDim udtPMSchedules(0 To lngPMEventCount - 1)
    With udtPMSchedules(0)
      .enmPeriodicity = PMPER_Weekly
      .dteLastPMCompleted = dteStartDate - 1 - Fix(lngPMSubEventPeriod * Rnd())
      .lngMonth = 0
      .lngDay = Weekday(.dteLastPMCompleted)
      .lngPMScheduleIndex = 0
    End With
    For lngIndex = 1 To lngPMEventCount - 1
      udtPMSchedules(lngIndex) = udtPMSchedules(lngIndex - 1)
      With udtPMSchedules(lngIndex)
        .dteLastPMCompleted = .dteLastPMCompleted - lngPMSubEventPeriod
        .lngDay = Weekday(.dteLastPMCompleted)
      End With
    Next
    
    ' Apply the PM schedule
    For lngIndex = 1 To c_lngPMPeriodicity_MaxValue
      lngLastPMEvent(lngIndex) = -1
    Next
    For lngAirportEquipmentModelIndex = 0 To udtAirport.lngEquipmentCount - 1
      With udtAirport.udtEquipment(lngAirportEquipmentModelIndex)
        If udtEquipmentModels.udtEquipmentModels(.lngEquipmentModelIndex).enmPeriodicity = PMPER_Weekly Then
          For lngAirportEquipmentIndex = 0 To .lngCount - 1
            lngLastPMEvent(enmPMPeriodicity) = lngLastPMEvent(enmPMPeriodicity) + 1
            If UBound(udtPMSchedules) <= lngLastPMEvent(enmPMPeriodicity) Then lngLastPMEvent(enmPMPeriodicity) = 0
            With udtEquipment.udtEquipment(.lngEquipmentIndexes(lngAirportEquipmentIndex))
              .udtPMSchedule = udtPMSchedules(lngLastPMEvent(enmPMPeriodicity))
              .udtPMSchedule.lngPMScheduleIndex = Fix(52 * Rnd())
            End With
          Next
        End If
      End With
    Next
  
  End If
  
  ' Synchronize monthly-based PM
  For enmPMPeriodicity = PMPER_Monthly To PMPER_Annually
    If booPMPeriodicityIsPresent(enmPMPeriodicity) Then Exit For
  Next
  If enmPMPeriodicity <= PMPER_Annually Then
  
    ' Set up the PM schedule
    lngPMEventCount = 1 + Fix(udtAirport.dblPMTime(enmPMPeriodicity) / dblMaxPMPerSync)
    Select Case enmPMPeriodicity
      Case PMPER_Monthly
        If 15 < lngPMEventCount Then
          lngPMEventCount = 30
          lngPMSubEventPeriod = 1
        Else
         lngPMSubEventPeriod = 30 \ lngPMEventCount
        End If
      Case PMPER_Quarterly
        If 45 < lngPMEventCount Then
          lngPMEventCount = 90
          lngPMSubEventPeriod = 1
        Else
         lngPMSubEventPeriod = 90 \ lngPMEventCount
        End If
      Case PMPER_SemiAnnually
        If 91 < lngPMEventCount Then
          lngPMEventCount = 182
          lngPMSubEventPeriod = 1
        Else
         lngPMSubEventPeriod = 182 \ lngPMEventCount
        End If
      Case PMPER_Annually
        If 182 < lngPMEventCount Then
          lngPMEventCount = 265
          lngPMSubEventPeriod = 1
        Else
         lngPMSubEventPeriod = 365 \ lngPMEventCount
        End If
      Case Else
        Err.Raise 5
    End Select
    ReDim udtPMSchedules(0 To lngPMEventCount - 1)
    With udtPMSchedules(0)
      .enmPeriodicity = PMPER_Monthly
      .dteLastPMCompleted = dteStartDate - 1 - Fix(lngPMSubEventPeriod * Rnd())
      .lngMonth = 0
      .lngDay = Day(.dteLastPMCompleted)
      .lngPMScheduleIndex = 0
    End With
    For lngIndex = 1 To lngPMEventCount - 1
      udtPMSchedules(lngIndex) = udtPMSchedules(lngIndex - 1)
      With udtPMSchedules(lngIndex)
        If enmPMPeriodicity = PMPER_Monthly Then
          .lngDay = .lngDay - lngPMSubEventPeriod
          If .lngDay < 1 Then .lngDay = 31 + .lngDay
          .dteLastPMCompleted = .dteLastPMCompleted - lngPMSubEventPeriod
        Else
          .dteLastPMCompleted = .dteLastPMCompleted - lngPMSubEventPeriod
          .lngDay = Day(.dteLastPMCompleted)
        End If
        .lngPMScheduleIndex = 0
      End With
    Next
  
    ' Apply the PM schedule
    For lngIndex = 1 To c_lngPMPeriodicity_MaxValue
      lngLastPMEvent(lngIndex) = -1
    Next
    For lngAirportEquipmentModelIndex = 0 To udtAirport.lngEquipmentCount - 1
      With udtAirport.udtEquipment(lngAirportEquipmentModelIndex)
        Select Case udtEquipmentModels.udtEquipmentModels(.lngEquipmentModelIndex).enmPeriodicity
          Case PMPER_NoPM
            ' Do nothing
          Case PMPER_Weekly
            ' Do nothing
          Case PMPER_Monthly
            For lngAirportEquipmentIndex = 0 To .lngCount - 1
              lngLastPMEvent(PMPER_Monthly) = lngLastPMEvent(PMPER_Monthly) + 1
              If UBound(udtPMSchedules) < lngLastPMEvent(PMPER_Monthly) Then lngLastPMEvent(PMPER_Monthly) = 0
              With udtEquipment.udtEquipment(.lngEquipmentIndexes(lngAirportEquipmentIndex))
                .udtPMSchedule = udtPMSchedules(lngLastPMEvent(PMPER_Monthly))
                .udtPMSchedule.enmPeriodicity = PMPER_Monthly
                .udtPMSchedule.lngPMScheduleIndex = Fix(12 * Rnd())
              End With
            Next
          Case PMPER_Quarterly
            For lngAirportEquipmentIndex = 0 To .lngCount - 1
              lngLastPMEvent(PMPER_Quarterly) = lngLastPMEvent(PMPER_Quarterly) + 1
              If UBound(udtPMSchedules) < lngLastPMEvent(PMPER_Quarterly) Then lngLastPMEvent(PMPER_Quarterly) = 0
              With udtEquipment.udtEquipment(.lngEquipmentIndexes(lngAirportEquipmentIndex))
                .udtPMSchedule = udtPMSchedules(lngLastPMEvent(PMPER_Quarterly))
                .udtPMSchedule.enmPeriodicity = PMPER_Quarterly
                .udtPMSchedule.lngMonth = Fix(3 * Rnd())
                .udtPMSchedule.dteLastPMCompleted = moveDateMonth(.udtPMSchedule.dteLastPMCompleted, .udtPMSchedule.lngMonth - (Month(.udtPMSchedule.dteLastPMCompleted) - 1) Mod 3)
                If dteStartDate <= .udtPMSchedule.dteLastPMCompleted Then
                  .udtPMSchedule.dteLastPMCompleted = moveDateMonth(.udtPMSchedule.dteLastPMCompleted, -3)
                End If
                .udtPMSchedule.lngPMScheduleIndex = Fix(4 * Rnd())
              End With
            Next
          Case PMPER_SemiAnnually
            For lngAirportEquipmentIndex = 0 To .lngCount - 1
              lngLastPMEvent(PMPER_SemiAnnually) = lngLastPMEvent(PMPER_SemiAnnually) + 1
              If UBound(udtPMSchedules) < lngLastPMEvent(PMPER_SemiAnnually) Then lngLastPMEvent(PMPER_SemiAnnually) = 0
              With udtEquipment.udtEquipment(.lngEquipmentIndexes(lngAirportEquipmentIndex))
                .udtPMSchedule = udtPMSchedules(lngLastPMEvent(PMPER_SemiAnnually))
                .udtPMSchedule.enmPeriodicity = PMPER_SemiAnnually
                .udtPMSchedule.lngMonth = Fix(6 * Rnd())
                .udtPMSchedule.dteLastPMCompleted = moveDateMonth(.udtPMSchedule.dteLastPMCompleted, .udtPMSchedule.lngMonth - (Month(.udtPMSchedule.dteLastPMCompleted) - 1) Mod 6)
                If dteStartDate <= .udtPMSchedule.dteLastPMCompleted Then
                  .udtPMSchedule.dteLastPMCompleted = moveDateMonth(.udtPMSchedule.dteLastPMCompleted, -6)
                End If
                .udtPMSchedule.lngPMScheduleIndex = Fix(2 * Rnd())
              End With
            Next
          Case PMPER_Annually
            For lngAirportEquipmentIndex = 0 To .lngCount - 1
              lngLastPMEvent(PMPER_Annually) = lngLastPMEvent(PMPER_Annually) + 1
              If UBound(udtPMSchedules) < lngLastPMEvent(PMPER_Annually) Then lngLastPMEvent(PMPER_Annually) = 0
              With udtEquipment.udtEquipment(.lngEquipmentIndexes(lngAirportEquipmentIndex))
                .udtPMSchedule = udtPMSchedules(lngLastPMEvent(PMPER_Annually))
                .udtPMSchedule.enmPeriodicity = PMPER_Annually
                .udtPMSchedule.lngMonth = Fix(12 * Rnd())
                .udtPMSchedule.dteLastPMCompleted = moveDateMonth(.udtPMSchedule.dteLastPMCompleted, .udtPMSchedule.lngMonth - Month(.udtPMSchedule.dteLastPMCompleted))
                If dteStartDate <= .udtPMSchedule.dteLastPMCompleted Then
                  .udtPMSchedule.dteLastPMCompleted = moveDateMonth(.udtPMSchedule.dteLastPMCompleted, -12)
                End If
                .udtPMSchedule.lngPMScheduleIndex = 0
              End With
            Next
          Case Else
            Err.Raise 5
        End Select
      End With
    Next
    
  End If

End Sub


Private Function loadPMStatus_Import( _
    udtAirports As AIRPORTS_TYPE, udtEquipmentModels As EQUIPMENTMODELS_TYPE, _
    udtEquipment As EQUIPMENT_TYPE, _
    ByVal objWorkbook As Excel.Workbook) As Boolean

  Dim objWorksheet As Excel.Worksheet
  Dim varValues As Variant, lngRowCount As Long
  Dim lngAirportIndex As Long, lngAirportEquipmentModelIndex As Long, lngAirportEquipmentIndex As Long
  Dim lngEquipmentIndex As Long
  Dim strAirport As String, strMakeModel As String
  Dim lngRowNumber As Long, lngEquipmentModelIndex As Long, lngIndex As Long
  Dim udtNoPM As PMSCHEDULE_TYPE, enmPMPeriodicity As PMPERIODICITY_ENUM
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Importing equipment PM status ..."
  End If
  
  Set objWorksheet = objWorkbook.Worksheets.Item("PMStatus")
  If Not verifyColumnHeadings(objWorksheet, c_strColHeadings_PMStatus) Then
    If Not (m_objErrorRange Is Nothing) Then
      m_objErrorRange.Value = "Invalid column headings in PMStatus worksheet"
      Set objWorksheet = Nothing
    End If
    Exit Function
  End If

  ' Initialize all PM to no PM
  udtNoPM.enmPeriodicity = PMPER_NoPM
  udtNoPM.dteLastPMCompleted = 0
  udtNoPM.lngMonth = 0
  udtNoPM.lngDay = 0
  udtNoPM.lngPMScheduleIndex = 0
  For lngEquipmentIndex = 0 To udtEquipment.lngEquipmentCount - 1
    udtEquipment.udtEquipment(lngEquipmentIndex).udtPMSchedule = udtNoPM
  Next

  lngRowNumber = objWorksheet.Range("A3:A3").CurrentRegion.Rows.Count
  varValues = objWorksheet.Range("A3:H" & (2 + lngRowNumber)).Value
  For lngRowNumber = 1 To UBound(varValues, 1)
    
    If varValues(lngRowNumber, 1) <> strAirport Then
      strAirport = varValues(lngRowNumber, 1)
      lngAirportIndex = udtAirports.colAirports("A:" & strAirport)
    End If
    If varValues(lngRowNumber, 3) <> strMakeModel Then
      strMakeModel = varValues(lngRowNumber, 3)
      lngEquipmentModelIndex = udtEquipmentModels.colEquipmentModels("EM:" & strMakeModel)
    End If
    enmPMPeriodicity = decodeName(CStr(varValues(lngRowNumber, 4)), c_strPMPeriodicityList)
    
    lngEquipmentIndex = varValues(lngRowNumber, 2)
    With udtEquipment.udtEquipment(lngEquipmentIndex)
      If .lngEquipmentModelIndex <> lngEquipmentModelIndex Then Err.Raise 5
      If udtEquipmentModels.udtEquipmentModels(.lngEquipmentModelIndex).enmPeriodicity <> enmPMPeriodicity Then Err.Raise 5
      .udtPMSchedule.enmPeriodicity = enmPMPeriodicity
      .udtPMSchedule.lngMonth = varValues(lngRowNumber, 5)
      .udtPMSchedule.lngDay = varValues(lngRowNumber, 6)
      .udtPMSchedule.lngPMScheduleIndex = varValues(lngRowNumber, 7)
      .udtPMSchedule.dteLastPMCompleted = varValues(lngRowNumber, 8)
    End With
    
  Next
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Importing equipment PM status ... done"
  End If
  
  loadPMStatus_Import = True
  Set objWorksheet = Nothing

End Function


Public Function exportPMStatus( _
    udtAirports As AIRPORTS_TYPE, udtEquipmentModels As EQUIPMENTMODELS_TYPE, _
    udtEquipment As EQUIPMENT_TYPE, _
    ByVal objWorkbook As Excel.Workbook, objAfterWorksheet As Excel.Worksheet, _
    strStepParameter As String) As Excel.Worksheet

  Dim objWorksheet As Excel.Worksheet
  Dim varValues As Variant, lngRowCount As Long
  Dim lngAirportIndex As Long, lngAirportEquipmentModelIndex As Long, lngAirportEquipmentIndex As Long, _
      lngEquipmentIndex As Long
  Dim strAirport As String, strMakeModel As String
  Dim lngRowNumber As Long, lngEquipmentModelIndex As Long, lngIndex As Long
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Exporting equipment PM status ..."
  End If
  
  On Error Resume Next
  Set objWorksheet = objWorkbook.Worksheets.Item("PMStatus")
  If objWorksheet Is Nothing Then
    If objAfterWorksheet Is Nothing Then
      Set objWorksheet = objWorkbook.Worksheets.Add()
    Else
      Set objWorksheet = objWorkbook.Worksheets.Add(, objAfterWorksheet)
    End If
    objWorksheet.Name = "PMStatus"
    exportColumnHeadings objWorksheet, c_strColHeadings_PMStatus
  End If

  lngRowNumber = objWorksheet.Range("A3:A3").CurrentRegion.Rows.Count
  objWorksheet.Range("A3:H" & (2 + lngRowNumber)).Clear
  
  For lngAirportIndex = 0 To udtAirports.lngAirportCount - 1
    For lngAirportEquipmentModelIndex = 0 To udtAirports.udtAirport(lngAirportIndex).lngEquipmentCount - 1
      lngRowCount = lngRowCount + udtAirports.udtAirport(lngAirportIndex).udtEquipment(lngAirportEquipmentModelIndex).lngCount
    Next
  Next
  If lngRowCount = 0 Then GoTo Label_NoRows
  
  varValues = objWorksheet.Range("A3:H" & (2 + lngRowCount)).Value
  lngRowNumber = 1
  For lngAirportIndex = 0 To udtAirports.lngAirportCount - 1
    strAirport = udtAirports.udtAirport(lngAirportIndex).strCode
    For lngAirportEquipmentModelIndex = 0 To udtAirports.udtAirport(lngAirportIndex).lngEquipmentCount - 1
      With udtAirports.udtAirport(lngAirportIndex).udtEquipment(lngAirportEquipmentModelIndex)
        lngEquipmentModelIndex = .lngEquipmentModelIndex
        strMakeModel = udtEquipmentModels.udtEquipmentModels(lngEquipmentModelIndex).strManufacturer & "." _
            & udtEquipmentModels.udtEquipmentModels(lngEquipmentModelIndex).strModel
        For lngAirportEquipmentIndex = 0 To .lngCount - 1
          lngEquipmentIndex = .lngEquipmentIndexes(lngAirportEquipmentIndex)
          With udtEquipment.udtEquipment(lngEquipmentIndex)
            varValues(lngRowNumber, 1) = strAirport
            varValues(lngRowNumber, 2) = .lngID
            varValues(lngRowNumber, 3) = strMakeModel
            varValues(lngRowNumber, 4) = encodeName(.udtPMSchedule.enmPeriodicity, c_strPMPeriodicityList)
            varValues(lngRowNumber, 5) = .udtPMSchedule.lngMonth
            varValues(lngRowNumber, 6) = .udtPMSchedule.lngDay
            varValues(lngRowNumber, 7) = .udtPMSchedule.lngPMScheduleIndex
            varValues(lngRowNumber, 8) = .udtPMSchedule.dteLastPMCompleted
          End With
          lngRowNumber = lngRowNumber + 1
        Next
      End With
    Next
  Next
  objWorksheet.Range("A3:H" & (2 + lngRowCount)).Value = varValues
  
Label_NoRows:
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Exporting equipment PM status ... done"
  End If
  
  Set exportPMStatus = objWorksheet
  Set objWorksheet = Nothing

End Function


Private Sub planPMHelper( _
    ByRef udtEquipment As EQUIPMENT_TYPE, udtEquipmentModel As EQUIPMENTMODEL_TYPE, _
    ByVal dteStartDate As Date, ByVal dteEndDate As Date)

  Dim dtePMDate As Date, dteNextPMDate As Date, lngPMIndex As Long
  Dim lngIndex As Long, lngBasePMIndex As Long

  ' BUG: Should randomize start of PM
'  udtEquipment.lngPMActivityCount = 0
'  ReDim udtEquipment.udtPMActivities(0 To 31)
'  If udtEquipmentModel.lngPMCount = 0 Then Exit Sub
  
'  Do While True
  
'    dteNextPMDate = udtEquipment.dteNextPMDue(0)
'    If dteEndDate < dteNextPMDate Then Exit Do
    
'    For lngPMIndex = udtEquipmentModel.lngPMCount - 1 To 1 Step -1
'      If dteNextPMDate <= udtEquipment.dteNextPMDue(lngPMIndex) Then Exit For
'    Next
    
'    If UBound(udtEquipment.udtPMActivities) < udtEquipment.lngPMActivityCount Then
'      ReDim Preserve udtEquipment.udtPMActivities(0 To 31 + udtEquipment.lngPMActivityCount)
'    End If
'    With udtEquipment.udtPMActivities(udtEquipment.lngPMActivityCount)
'      .lngPMIndex = lngPMIndex
'      .dteStartTime = dteNextPMDate
'      .dteEndTime = dteNextPMDate + (udtEquipmentModel.udtPM(lngPMIndex).dblLabor_Initial _
'          + udtEquipmentModel.udtPM(lngPMIndex).dblLabor_Wait _
'          + udtEquipmentModel.udtPM(lngPMIndex).dblLabor_Final) / 24
'    End With
'    udtEquipment.lngPMActivityCount = udtEquipment.lngPMActivityCount + 1
    
'    dtePMDate = dteNextPMDate
'    For lngPMIndex = lngPMIndex To 0 Step -1
'      Select Case udtEquipmentModel.udtPM(lngPMIndex).enmPeriodicity
'        Case PMPERIODICITY_ENUM.PMPER_Daily
'          udtEquipment.dteNextPMDue(lngPMIndex) = dtePMDate + 1
'        Case PMPERIODICITY_ENUM.PMPER_Weekly
'          udtEquipment.dteNextPMDue(lngPMIndex) = dtePMDate + 7
'        Case PMPERIODICITY_ENUM.PMPER_Biweekly
'          udtEquipment.dteNextPMDue(lngPMIndex) = dtePMDate + 14
'        Case PMPERIODICITY_ENUM.PMPER_Monthly
'          udtEquipment.dteNextPMDue(lngPMIndex) = dtePMDate + 365 / 12
'        Case PMPERIODICITY_ENUM.PMPER_Quarterly
'          udtEquipment.dteNextPMDue(lngPMIndex) = dtePMDate + 365 / 4
'        Case PMPERIODICITY_ENUM.PMPER_Semiannually
'          udtEquipment.dteNextPMDue(lngPMIndex) = dtePMDate + 365 / 2
'        Case PMPERIODICITY_ENUM.PMPER_Annually
'          udtEquipment.dteNextPMDue(lngPMIndex) = dtePMDate + 365
'        Case Else
'          Err.Raise 5
'      End Select
'    Next
    
'  Loop

End Sub



Private Sub computeCM(ByVal dteStartDate As Date, ByVal dteEndDate As Date)

  Dim lngAirportIndex As Long, lngEquipmentModelIndex As Long, lngEquipmentIndex As Long
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Computing CM ..."
  End If
  
'  For lngAirportIndex = 0 To m_lngAirportCount - 1
'    For lngEquipmentModelIndex = 0 To m_udtAirports(lngAirportIndex).lngEquipmentCount - 1
'      For lngEquipmentIndex = 0 To m_udtAirports(lngAirportIndex).udtEquipment(lngEquipmentModelIndex).lngCount - 1
'        computeCMHelper m_udtAirports(lngAirportIndex).udtEquipment(lngEquipmentModelIndex).udtEquipment(lngEquipmentIndex), _
'            m_udtEquipmentModels(m_udtAirports(lngAirportIndex).udtEquipment(lngEquipmentModelIndex).lngEquipmentModelIndex), _
'            m_udtAirports(lngAirportIndex).dblCMTravelTime, dteStartDate, dteEndDate
'      Next
'    Next
'  Next
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Computing CM ... done"
  End If

End Sub


Private Sub computeCMHelper( _
    ByRef udtEquipment As EQUIPMENTITEM_TYPE, udtCMRequirements As CMREQUIREMENTS_TYPE, _
    ByVal dblCMTravelTime As Date, _
    ByVal dteStartDate As Date, ByVal dteEndDate As Date)

  Dim dblCMFrequency As Double, dblCMPercentages() As Double, dteCMTime As Date, dblP As Double
  Dim dblCMTime As Double, dblPartsTime As Double
  Dim lngCMIndex As Long
  
  Err.Raise 5
  
'  udtEquipment.lngCMActivityCount = 0
'  ReDim udtEquipment.udtCMActivities(0 To 31)
'  If udtCMRequirements.lngCMRequirementCount = 0 Then Exit Sub
  
'  ReDim dblCMPercentages(0 To udtEquipmentModel.lngCMCount - 1)
'  For lngCMIndex = 0 To udtEquipmentModel.lngCMCount - 1
'    dblCMFrequency = dblCMFrequency + udtEquipmentModel.udtCM(lngCMIndex).dblFrequency
'    dblCMPercentages(lngCMIndex) = udtEquipmentModel.udtCM(lngCMIndex).dblFrequency
'  Next
'  dblCMPercentages(0) = dblCMPercentages(0) / dblCMFrequency
'  For lngCMIndex = 1 To udtEquipmentModel.lngCMCount - 1
'    dblCMPercentages(lngCMIndex) = dblCMPercentages(lngCMIndex - 1) + dblCMPercentages(lngCMIndex) / dblCMFrequency
'  Next
  
'  dteCMTime = dteStartDate
'  Do While True
  
'    dteCMTime = dteCMTime - 365 * Log(1 - Rnd()) / dblCMFrequency
'    If dteEndDate < dteCMTime Then Exit Do
    
'    If UBound(udtEquipment.udtCMActivities) < udtEquipment.lngCMActivityCount Then
'      ReDim Preserve udtEquipment.udtCMActivities(0 To 31 + udtEquipment.lngCMActivityCount)
'    End If
    
'    dblP = Rnd()
'    For lngCMIndex = 0 To udtEquipmentModel.lngCMCount - 2
'      If dblP <= dblCMPercentages(lngCMIndex) Then Exit For
'    Next
    
'    dblCMTime = selectFromDistribution(udtEquipmentModel.udtCM(lngCMIndex).udtCMTime) / 24
'    dblPartsTime = selectFromDistribution(udtEquipmentModel.udtCM(lngCMIndex).udtPartsTime) / 24
'    With udtEquipment.udtCMActivities(udtEquipment.lngCMActivityCount)
'      .lngCMIndex = lngCMIndex
'      .dteFailureTime = dteCMTime
'      .dteCallTime = .dteFailureTime
'      .dteDispatchTime = .dteCallTime
'      .dteArrivalTime = .dteDispatchTime + dblCMTravelTime
'      .dteDiagnosisEndTime = .dteArrivalTime + dblCMTime / 2
'      .dtePartsRequestTime = .dteDiagnosisEndTime
'      .dtePartsFulfillmentTime = .dtePartsRequestTime + dblPartsTime
'      .dtePartsLocalLogisticsTime = .dtePartsFulfillmentTime
'      .dteRepairTime = .dtePartsLocalLogisticsTime + dblCMTime / 2
'      .dteTestTime = .dteRepairTime
'      .dteSignoffTime = .dteTestTime
'    End With
'    udtEquipment.lngCMActivityCount = udtEquipment.lngCMActivityCount + 1
      
'  Loop

End Sub



Public Function computeAirportDistances( _
    udtAirports As AIRPORTS_TYPE, ByRef dblAirportDistances() As Double) As Boolean

  Dim lngAirportIndex1 As Long, lngAirportIndex2 As Long
  Dim dblLongitude As Double, dblLatitude As Double
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Computing airport distances ..."
  End If
  
  ReDim dblAirportDistances(0 To udtAirports.lngAirportCount - 1, 0 To udtAirports.lngAirportCount - 1)
  For lngAirportIndex1 = 0 To udtAirports.lngAirportCount - 1
  For lngAirportIndex1 = 0 To udtAirports.lngAirportCount - 1
    dblLongitude = udtAirports.udtAirport(lngAirportIndex1).dblLongitude
    dblLatitude = udtAirports.udtAirport(lngAirportIndex1).dblLatitude
    For lngAirportIndex2 = 0 To udtAirports.lngAirportCount - 1
      dblAirportDistances(lngAirportIndex1, lngAirportIndex2) = computeDistance(dblLongitude, dblLatitude, _
          udtAirports.udtAirport(lngAirportIndex2).dblLongitude, udtAirports.udtAirport(lngAirportIndex2).dblLatitude)
    Next
  Next

  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Computing airport distances ... done"
  End If
  
  computeAirportDistances = True

End Function


Private Sub computeMetrics( _
    udtAirports As AIRPORTS_TYPE, udtEquipmentModels As EQUIPMENTMODELS_TYPE, _
    udtEquipmentTypes As EQUIPMENTTYPES_TYPE, _
    ByRef udtMetrics As METRICS_TYPE, _
    ByVal dteStartDate As Date, ByVal dteEndDate As Date)

  Dim lngAirportIndex As Long, lngAirportEquipmentModelIndex As Long
  Dim lngEquipmentModelIndex As Long, lngEquipmentTypeIndex As Long, lngEquipmentIndex As Long, _
      lngEquipmentCount As Long
  Dim lngPMIndex As Long, lngCMIndex As Long, lngPMEvents As Long, dblPMTime As Double, _
      lngCMEvents As Long, dblCMTime As Double, dblOperatingTime As Double
  Dim lngAirportEquipmentTypeIndex As Long, lngAirportStartIndex As Long, lngAirportEndIndex As Long
  Dim lngIndex As Long

  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Computing metrics (" & dteStartDate & " to " & dteEndDate & ") ..."
  End If

  For lngAirportIndex = 0 To udtAirports.lngAirportCount - 1
    For lngAirportEquipmentModelIndex = 0 To udtAirports.udtAirport(lngAirportIndex).lngEquipmentCount - 1
      lngEquipmentModelIndex = lngEquipmentModelIndex + udtAirports.udtAirport(lngAirportIndex).lngEquipmentCount
    Next
  Next
  ReDim udtMetrics.udtEquipmentModelMetrics(0 To lngEquipmentModelIndex)
  
  udtMetrics.lngEquipmentModelCount = 0
  For lngAirportIndex = 0 To udtAirports.lngAirportCount - 1
    For lngAirportEquipmentModelIndex = 0 To udtAirports.udtAirport(lngAirportIndex).lngEquipmentCount - 1
      With udtAirports.udtAirport(lngAirportIndex).udtEquipment(lngAirportEquipmentModelIndex)
        
        lngEquipmentCount = .lngCount
        lngEquipmentModelIndex = .lngEquipmentModelIndex
        lngEquipmentTypeIndex = udtEquipmentModels.udtEquipmentModels(.lngEquipmentModelIndex).lngEquipmentTypeIndex
        lngPMEvents = 0
        dblPMTime = 0#
        lngCMEvents = 0
        dblCMTime = 0#
        dblOperatingTime = (dteEndDate - dteStartDate) * .lngCount
        
        ' BUG: Does not address events that overlap boundaries
        For lngEquipmentIndex = 0 To .lngCount - 1
'          With .udtEquipment(lngEquipmentIndex)
'            For lngPMIndex = 0 To .lngPMActivityCount - 1
'              If dteStartDate <= .udtPMActivities(lngPMIndex).dteStartTime Then Exit For
'            Next
'            For lngPMIndex = lngPMIndex To .lngPMActivityCount - 1
'              If dteEndDate < .udtPMActivities(lngPMIndex).dteStartTime Then Exit For
'              lngPMEvents = lngPMEvents + 1
'              dblPMTime = dblPMTime + .udtPMActivities(lngPMIndex).dteEndTime - .udtPMActivities(lngPMIndex).dteStartTime
'            Next
'            For lngCMIndex = 0 To .lngCMActivityCount - 1
'              If dteStartDate <= .udtCMActivities(lngCMIndex).dteFailureTime Then Exit For
'            Next
'            For lngCMIndex = lngCMIndex To .lngCMActivityCount - 1
'              If dteEndDate < .udtCMActivities(lngCMIndex).dteFailureTime Then Exit For
'              lngCMEvents = lngCMEvents + 1
'              dblCMTime = dblCMTime + .udtCMActivities(lngCMIndex).dteSignoffTime - .udtCMActivities(lngCMIndex).dteFailureTime
'            Next
'          End With
        Next

        With udtMetrics.udtEquipmentModelMetrics(udtMetrics.lngEquipmentModelCount)
          .lngAirportIndex = lngAirportIndex
          .lngEquipmentModelIndex = lngEquipmentModelIndex
          .lngEquipmentTypeIndex = lngEquipmentTypeIndex
          .lngEquipmentCount = lngEquipmentCount
          .dblOperatingTime = dblOperatingTime
          .lngPMEvents = lngPMEvents
          .dblPMTime = dblPMTime
          .lngCMEvents = lngCMEvents
          .dblCMTime = dblCMTime
        End With
        udtMetrics.lngEquipmentModelCount = udtMetrics.lngEquipmentModelCount + 1
      
      End With
    Next
  Next

  udtMetrics.lngEquipmentTypeCount = 0
  ReDim udtMetrics.udtEquipmentTypeMetrics(0 To UBound(udtMetrics.udtEquipmentModelMetrics))
  lngAirportEndIndex = -1
  For lngAirportIndex = 0 To udtAirports.lngAirportCount - 1
  
    For lngAirportStartIndex = lngAirportEndIndex + 1 To udtMetrics.lngEquipmentModelCount - 1
      If lngAirportIndex <= udtMetrics.udtEquipmentModelMetrics(lngAirportStartIndex).lngAirportIndex Then Exit For
    Next
    For lngAirportEndIndex = lngAirportStartIndex To udtMetrics.lngEquipmentModelCount - 1
      If lngAirportIndex < udtMetrics.udtEquipmentModelMetrics(lngAirportEndIndex).lngAirportIndex Then Exit For
    Next
    lngAirportEndIndex = lngAirportEndIndex - 1
      
    For lngEquipmentTypeIndex = 0 To udtEquipmentTypes.lngEquipmentTypeCount - 1
    
      lngEquipmentCount = 0
      lngPMEvents = 0
      dblPMTime = 0#
      lngCMEvents = 0
      dblCMTime = 0#
      dblOperatingTime = 0#
      
      For lngIndex = lngAirportStartIndex To lngAirportEndIndex
        If udtMetrics.udtEquipmentModelMetrics(lngIndex).lngEquipmentTypeIndex = lngEquipmentTypeIndex Then
          lngEquipmentCount = lngEquipmentCount + udtMetrics.udtEquipmentModelMetrics(lngIndex).lngEquipmentCount
          lngPMEvents = lngPMEvents + udtMetrics.udtEquipmentModelMetrics(lngIndex).lngPMEvents
          dblPMTime = dblPMTime + udtMetrics.udtEquipmentModelMetrics(lngIndex).dblPMTime
          lngCMEvents = lngCMEvents + udtMetrics.udtEquipmentModelMetrics(lngIndex).lngCMEvents
          dblCMTime = dblCMTime + udtMetrics.udtEquipmentModelMetrics(lngIndex).dblCMTime
          dblOperatingTime = dblOperatingTime + udtMetrics.udtEquipmentModelMetrics(lngIndex).dblOperatingTime
        End If
      Next
      
      If 0# < dblOperatingTime Then
        With udtMetrics.udtEquipmentTypeMetrics(udtMetrics.lngEquipmentTypeCount)
          .lngAirportIndex = lngAirportIndex
          .lngEquipmentModelIndex = -1
          .lngEquipmentTypeIndex = lngEquipmentTypeIndex
          .lngEquipmentCount = lngEquipmentCount
          .lngPMEvents = lngPMEvents
          .dblPMTime = dblPMTime
          .lngCMEvents = lngCMEvents
          .dblCMTime = dblCMTime
          .dblOperatingTime = dblOperatingTime
        End With
        udtMetrics.lngEquipmentTypeCount = udtMetrics.lngEquipmentTypeCount + 1
      End If
    
    Next
  Next
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Computing metrics (" & dteStartDate & " to " & dteEndDate & ") ... done"
  End If
  
End Sub



'''''''''' Support methods
''''''
''

'''''''''' Working with name-lists
Private Function decodeName(strName As String, strValueList As String) As Long

  Dim lngStartOffset As Long, lngEndOffset As Long
  
  lngEndOffset = InStr(strValueList, ";" & strName & ";")
  lngStartOffset = 1 + InStrRev(strValueList, ";", lngEndOffset - 1)
  decodeName = CLng(Mid$(strValueList, lngStartOffset, lngEndOffset - lngStartOffset))

End Function


Public Function encodeName(lngValue As Long, strValueList As String) As String

  Dim lngStartOffset As Long, lngEndOffset As Long, strValue As String
  
  strValue = CStr(lngValue) & ";"
  lngStartOffset = InStr(strValueList, ";" & strValue)
  If lngStartOffset = 0 Then
    If Left$(strValueList, Len(strValue)) <> strValue Then
      If Not (m_objErrorRange Is Nothing) Then
        m_objErrorRange.Value = "Value " & lngValue & " not found in value list " & strValueList
      End If
      Err.Raise 5
    End If
    lngStartOffset = Len(strValue) + 1
  Else
    lngStartOffset = lngStartOffset + Len(strValue) + 1
  End If
  
  lngEndOffset = InStr(lngStartOffset, strValueList, ";")
  encodeName = Mid$(strValueList, lngStartOffset, lngEndOffset - lngStartOffset)

End Function


'''''''''' Column headings support

Public Sub exportColumnHeadings(ByVal objWorksheet As Excel.Worksheet, strColumnHeadings As String)

  Dim lngStartOffset As Long, lngEndOffset As Long, lngColumnNumber As Long
  
  lngStartOffset = 1
  lngColumnNumber = 1
  Do While lngStartOffset < Len(strColumnHeadings)
    lngEndOffset = InStr(lngStartOffset, strColumnHeadings, ";")
    objWorksheet.Range(Chr(64 + lngColumnNumber) & "1").Value = Mid$(strColumnHeadings, lngStartOffset, lngEndOffset - lngStartOffset)
    lngColumnNumber = lngColumnNumber + 1
    lngStartOffset = lngEndOffset + 1
  Loop

End Sub


Public Function verifyColumnHeadings(ByVal objWorksheet As Excel.Worksheet, strColumnHeadings As String, _
    Optional ByVal lngRowNumber As Long = 1) As Boolean

  Dim lngStartOffset As Long, lngEndOffset As Long, lngColumnNumber As Long
  
  lngStartOffset = 1
  lngColumnNumber = 1
  Do While lngStartOffset < Len(strColumnHeadings)
    lngEndOffset = InStr(lngStartOffset, strColumnHeadings, ";")
    If objWorksheet.Range(Chr(64 + lngColumnNumber) & lngRowNumber).Value <> Mid$(strColumnHeadings, lngStartOffset, lngEndOffset - lngStartOffset) Then Exit Function
    lngColumnNumber = lngColumnNumber + 1
    lngStartOffset = lngEndOffset + 1
  Loop
  verifyColumnHeadings = True

End Function




'''''''''' Statistical calculations

Private Function selectFromDistribution(udtDistribution As DISTRIBUTION_TYPE) As Double

  Dim dblValue As Double
  
  If udtDistribution.dblStndDev = 0# Then
    selectFromDistribution = udtDistribution.dblAvg
    Exit Function
  End If
  
  Do While True
    dblValue = Application.WorksheetFunction.NormInv(udtDistribution.dblAvg, udtDistribution.dblStndDev, Rnd())
    If udtDistribution.dblMin <= dblValue Then
      If dblValue <= udtDistribution.dblMax Then
        Exit Do
      End If
    End If
  Loop

  selectFromDistribution = dblValue

End Function





Private Sub loadTravelCosts(objWorkbook As Excel.Workbook)

  Dim objWorksheet As Excel.Worksheet
  Dim varValues As Variant, strFacility As String
  Dim lngRowNumber As Long
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Loading travel costs ..."
  End If
  
  Set objWorkbook = Application.ActiveWorkbook
  Set objWorksheet = objWorkbook.Worksheets.Item("TravelCosts")
  
  varValues = objWorksheet.Range("A3:E3").CurrentRegion.Value
  m_lngTravelCostCount = UBound(varValues, 1)
  ReDim m_udtTravelCosts(0 To m_lngTravelCostCount - 1)
  Set m_colTravelCosts = New Collection
  For lngRowNumber = 1 To UBound(varValues, 1)

    strFacility = varValues(lngRowNumber, 1)
    
    With m_udtTravelCosts(lngRowNumber - 1)
      .strID = strFacility
      .dblLodging = varValues(lngRowNumber, 2)
      .dblPerDiem = varValues(lngRowNumber, 3)
      .dblRentalCar = varValues(lngRowNumber, 4)
      .dblFacilityParking = varValues(lngRowNumber, 5)
    End With
    
  Next
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Loading travel costs ... done"
  End If
  
End Sub


Private Sub parseParameters( _
    strParameterString As String, _
    ByRef strParameters() As String, ByRef lngParameterCount As Long)

  Dim lngStartOffset As Long, lngEndOffset As Long
  
  lngParameterCount = 0
  lngStartOffset = 1
  Do While lngStartOffset < Len(strParameterString)
    lngEndOffset = InStr(lngStartOffset, strParameterString, ",")
    If lngEndOffset = 0 Then lngEndOffset = Len(strParameterString) + 1
    If UBound(strParameters) < lngParameterCount Then
      ReDim Preserve strParameters(0 To 7 + lngParameterCount)
    End If
    strParameters(lngParameterCount) = Trim(Mid$(strParameterString, lngStartOffset, lngEndOffset - lngStartOffset))
    lngParameterCount = lngParameterCount + 1
    lngStartOffset = lngEndOffset + 1
  Loop

End Sub


Private Function monthAdjustDate(ByVal dteStartDate As Date, ByVal lngMonthShift As Long) As Date

  Dim lngMonth As Long, lngDay As Long, lngYear As Long
  
  lngYear = Year(dteStartDate)
  lngMonth = Month(dteStartDate) + lngMonthShift
  lngDay = Day(dteStartDate)
  Do While lngMonth < 1
    lngMonth = lngMonth + 12
    lngYear = lngYear - 1
  Loop
  Do While 12 < lngMonth
    lngMonth = lngMonth - 12
    lngYear = lngYear + 1
  Loop
  
  dteStartDate = CDate(lngMonth & "/1/" & lngYear)
  lngMonth = Month(dteStartDate)
  dteStartDate = dteStartDate + lngDay - 1
  Do While Month(dteStartDate) <> lngMonth
    dteStartDate = dteStartDate - 1
  Loop

  monthAdjustDate = dteStartDate

End Function


' Computes the data of the PM event that is lngStepCount events forward/back from the specified date
Private Function computePMEventDate( _
    ByVal dteFromDate As Date, udtPMSchedule As PMSCHEDULE_TYPE, _
    ByVal lngStepCount As Long) As Date

  Dim lngMonth As Long

  Select Case udtPMSchedule.enmPeriodicity
  
    Case PMPER_Weekly
      dteFromDate = dteFromDate - Weekday(dteFromDate) + udtPMSchedule.lngDay + 7 * lngStepCount
      
    Case PMPER_Monthly
      If lngStepCount <> 0 Then dteFromDate = moveDateMonth(dteFromDate, lngStepCount)
      lngMonth = Month(dteFromDate)
      dteFromDate = dteFromDate - Day(dteFromDate) + udtPMSchedule.lngDay
      Do While Month(dteFromDate) <> lngMonth
        dteFromDate = dteFromDate - 1
      Loop
        
    Case PMPER_Quarterly
      lngMonth = (Month(dteFromDate) - 1) Mod 3
      If lngMonth <> udtPMSchedule.lngMonth Then dteFromDate = moveDateMonth(dteFromDate, udtPMSchedule.lngMonth - lngMonth)
      If lngStepCount <> 0 Then dteFromDate = moveDateMonth(dteFromDate, 3 * lngStepCount)
      lngMonth = Month(dteFromDate)
      dteFromDate = dteFromDate - Day(dteFromDate) + udtPMSchedule.lngDay
      Do While Month(dteFromDate) <> lngMonth
        dteFromDate = dteFromDate - 1
      Loop
      
    Case PMPER_SemiAnnually
      lngMonth = (Month(dteFromDate) - 1) Mod 6
      If lngMonth <> udtPMSchedule.lngMonth Then dteFromDate = moveDateMonth(dteFromDate, udtPMSchedule.lngMonth - lngMonth)
      If lngStepCount <> 0 Then dteFromDate = moveDateMonth(dteFromDate, 6 * lngStepCount)
      lngMonth = Month(dteFromDate)
      dteFromDate = dteFromDate - Day(dteFromDate) + udtPMSchedule.lngDay
      Do While Month(dteFromDate) <> lngMonth
        dteFromDate = dteFromDate - 1
      Loop
      
    Case PMPER_Annually
      lngMonth = (Month(dteFromDate) - 1) - 1
      If lngMonth <> udtPMSchedule.lngMonth Then dteFromDate = moveDateMonth(dteFromDate, udtPMSchedule.lngMonth - lngMonth)
      If lngStepCount <> 0 Then dteFromDate = moveDateMonth(dteFromDate, 12 * lngStepCount)
      lngMonth = Month(dteFromDate)
      dteFromDate = dteFromDate - Day(dteFromDate) + udtPMSchedule.lngDay
      Do While Month(dteFromDate) <> lngMonth
        dteFromDate = dteFromDate - 1
      Loop
      
    Case Else
      Err.Raise 5
      
  End Select

  computePMEventDate = dteFromDate

End Function


Private Function moveDateMonth(ByVal dteFromDate As Date, ByVal lngMonthCount As Long) As Date

  Dim lngDay As Long, lngMonth As Long
  
  lngDay = Day(dteFromDate)
  dteFromDate = dteFromDate - lngDay + 15
  If lngMonthCount < 0 Then
    For lngMonthCount = -1 To lngMonthCount Step -1
      dteFromDate = dteFromDate - 32
      dteFromDate = dteFromDate - Day(dteFromDate) + 15
    Next
  End If
  If 0 < lngMonthCount Then
    For lngMonthCount = 0 To lngMonthCount - 1
      dteFromDate = dteFromDate + 32
      dteFromDate = dteFromDate - Day(dteFromDate) + 15
    Next
  End If
  
  lngMonth = Month(dteFromDate)
  dteFromDate = dteFromDate - Day(dteFromDate) + lngDay
  Do While Month(dteFromDate) <> lngMonth
    dteFromDate = dteFromDate - 1
  Loop
  
  moveDateMonth = dteFromDate

End Function


Public Function createPMItems( _
    udtEquipment As EQUIPMENT_TYPE, _
    strParameters As String) As Boolean

  Dim lngEquipmentIndex As Long, lngPMCount As Long, dteNextPMDate As Date, _
      lngPMRequirementIndex As Long, lngPMYears As Long, lngEquipmentPMItemIndex As Long
  Dim dteFromDate As Date, dteToDate As Date

  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Creating PM items ..."
  End If
  
  dteFromDate = m_dteModelStartDate
  dteToDate = m_dteModelEndDate

  For lngEquipmentIndex = 0 To udtEquipment.lngEquipmentCount - 1
    Select Case udtEquipment.udtEquipment(lngEquipmentIndex).udtPMSchedule.enmPeriodicity
      Case PMPER_NoPM
        ' Do nothing
      Case PMPER_Weekly
        lngPMCount = lngPMCount + 52
      Case PMPER_Monthly
        lngPMCount = lngPMCount + 12
      Case PMPER_Quarterly
        lngPMCount = lngPMCount + 4
      Case PMPER_SemiAnnually
        lngPMCount = lngPMCount + 2
      Case PMPER_Annually
        lngPMCount = lngPMCount + 1
      Case Else
        Err.Raise 5
    End Select
  Next
  lngPMCount = CLng(1.1 * (dteToDate - dteFromDate) / 365 * lngPMCount)
  ReDim m_udtPMItems(0 To lngPMCount)
      
  m_lngPMItemCount = 0
  lngPMYears = Fix((1 + dteToDate - dteFromDate) / 365)
  For lngEquipmentIndex = 0 To udtEquipment.lngEquipmentCount - 1
    With udtEquipment.udtEquipment(lngEquipmentIndex)
      If .udtPMSchedule.enmPeriodicity = PMPER_NoPM Then
        .udtPMSchedule.lngPMItemCount = 0
        
      Else
      
        Select Case .udtPMSchedule.enmPeriodicity
          Case PMPER_Weekly
            ReDim .udtPMSchedule.lngPMItemIndexes(0 To lngPMYears * 52)
          Case PMPER_Monthly
            ReDim .udtPMSchedule.lngPMItemIndexes(0 To lngPMYears * 12)
          Case PMPER_Quarterly
            ReDim .udtPMSchedule.lngPMItemIndexes(0 To lngPMYears * 4)
          Case PMPER_SemiAnnually
            ReDim .udtPMSchedule.lngPMItemIndexes(0 To lngPMYears * 2)
          Case PMPER_Annually
            ReDim .udtPMSchedule.lngPMItemIndexes(0 To lngPMYears)
          Case Else
            Err.Raise 5
        End Select
        lngEquipmentPMItemIndex = 0
      
        lngPMRequirementIndex = .udtPMSchedule.lngPMScheduleIndex
        dteNextPMDate = .udtPMSchedule.dteLastPMCompleted
        Do While True
          
          dteNextPMDate = computePMEventDate(dteNextPMDate, .udtPMSchedule, 1)
          If dteToDate < dteNextPMDate Then Exit Do
          
          lngPMRequirementIndex = lngPMRequirementIndex + 1
          Select Case .udtPMSchedule.enmPeriodicity
            Case PMPER_Weekly
              If 52 <= lngPMRequirementIndex Then lngPMRequirementIndex = 0
            Case PMPER_Monthly
              If 12 <= lngPMRequirementIndex Then lngPMRequirementIndex = 0
            Case PMPER_Quarterly
              If 4 <= lngPMRequirementIndex Then lngPMRequirementIndex = 0
            Case PMPER_SemiAnnually
              If 2 <= lngPMRequirementIndex Then lngPMRequirementIndex = 0
            Case PMPER_Annually
              If 1 <= lngPMRequirementIndex Then lngPMRequirementIndex = 0
            Case Else
              Err.Raise 5
          End Select
          
          If UBound(m_udtPMItems) < m_lngPMItemCount Then
            ReDim Preserve m_udtPMItems(0 To 2000 + m_lngPMItemCount)
          End If
          
          With m_udtPMItems(m_lngPMItemCount)
            .lngEquipmentIndex = lngEquipmentIndex
            .lngPMRequirementIndex = lngPMRequirementIndex
            .lngTripIndex = -1
            .dteScheduledStart = dteNextPMDate
            .dteStartTime = 0#
            .dteEndTime = 0#
          End With
          .udtPMSchedule.lngPMItemIndexes(lngEquipmentPMItemIndex) = m_lngPMItemCount
          lngEquipmentPMItemIndex = lngEquipmentPMItemIndex + 1
          m_lngPMItemCount = m_lngPMItemCount + 1
          
        Loop
        .udtPMSchedule.lngPMItemCount = lngEquipmentPMItemIndex
      End If
    End With
  Next

  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Creating PM items ... done"
  End If

  createPMItems = True

End Function


Public Function exportDailyPMTimes( _
    udtEquipmentModels As EQUIPMENTMODELS_TYPE, udtEquipment As EQUIPMENT_TYPE, _
    ByVal objWorkbook As Excel.Workbook, ByVal objAfterWorksheet As Excel.Worksheet, _
    strParameterString As String) As Excel.Worksheet

  Dim objWorksheet As Excel.Worksheet, varValues As Variant, lngRowNumber As Long
  Dim lngPMItemIndex As Long, lngPMIndex As Long, lngEquipmentModelIndex As Long
  Dim dblPMTimes() As Double, lngTimeIndex As Long
  Dim lngParameterCount As Long, strParameters() As String, dteFromDate As Date, dteToDate As Date
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Exporting daily PM times ..."
  End If
  
  ReDim strParameters(0 To 3)
  parseParameters strParameterString, strParameters, lngParameterCount
  If lngParameterCount = 0 Then
    dteFromDate = m_dteModelStartDate
    dteToDate = m_dteModelEndDate
  ElseIf lngParameterCount = 2 Then
    dteFromDate = CDate(strParameters(0))
    dteToDate = CDate(strParameters(1))
  Else
    If Not (m_objErrorRange Is Nothing) Then m_objErrorRange.Value = "Invalid parameters"
    Exit Function
  End If

  ReDim dblPMTimes(0 To CLng(dteToDate - dteFromDate))
  For lngPMItemIndex = 0 To m_lngPMItemCount - 1
    With m_udtPMItems(lngPMItemIndex)
    
      If .dteScheduledStart < dteFromDate Then
        ' Do nothing
      ElseIf dteToDate < .dteScheduledStart Then
        ' Do nothing
      Else
        lngTimeIndex = CLng(.dteScheduledStart - dteFromDate)
        lngEquipmentModelIndex = udtEquipment.udtEquipment(.lngEquipmentIndex).lngEquipmentModelIndex
        lngPMIndex = udtEquipmentModels.udtEquipmentModels(lngEquipmentModelIndex).lngPMSchedule(.lngPMRequirementIndex)
        dblPMTimes(lngTimeIndex) = dblPMTimes(lngTimeIndex) _
            + udtEquipmentModels.udtEquipmentModels(lngEquipmentModelIndex).udtPM(lngPMIndex).dblLabor_Initial _
            + udtEquipmentModels.udtEquipmentModels(lngEquipmentModelIndex).udtPM(lngPMIndex).dblLabor_Wait _
            + udtEquipmentModels.udtEquipmentModels(lngEquipmentModelIndex).udtPM(lngPMIndex).dblLabor_Final
      End If
    
    End With
  Next
  
  On Error Resume Next
  Set objWorksheet = objWorkbook.Worksheets.Item("DailyPMTimes")
  If objWorksheet Is Nothing Then
    If objAfterWorksheet Is Nothing Then
      Set objWorksheet = objWorkbook.Worksheets.Add()
    Else
      Set objWorksheet = objWorkbook.Worksheets.Add(, , objAfterWorksheet)
    End If
    objWorksheet.Name = "DailyPMTimes"
    exportColumnHeadings objWorksheet, c_strColHeadings_DailyPMTimes
  End If
  On Error GoTo 0
  
  ' Clear current contents (if any)
  lngRowNumber = objWorksheet.Range("A3:A3").CurrentRegion.Rows.Count
  objWorksheet.Range("A3:J" & (2 + lngRowNumber)).Clear
  
  ' Export daily PM Times
  varValues = objWorksheet.Range("A3:B" & (3 + UBound(dblPMTimes))).Value
  For lngTimeIndex = 0 To UBound(dblPMTimes)
    lngRowNumber = lngTimeIndex + 1
    varValues(lngRowNumber, 1) = dteFromDate + lngTimeIndex
    varValues(lngRowNumber, 2) = dblPMTimes(lngTimeIndex)
  Next
  objWorksheet.Range("A3:B" & (2 + (3 + UBound(dblPMTimes)))).Value = varValues
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Exporting daily PM times ... done"
  End If
  
  Set exportDailyPMTimes = objWorksheet
  Set objWorksheet = Nothing

End Function


Public Function exportPMSchedule( _
    udtAirports As AIRPORTS_TYPE, udtEquipmentModels As EQUIPMENTMODELS_TYPE, _
    udtEquipment As EQUIPMENT_TYPE, _
    ByVal objWorkbook As Excel.Workbook, ByVal objAfterWorksheet As Excel.Worksheet, _
    strParameterString As String) As Excel.Worksheet

  Dim objWorksheet As Excel.Worksheet, varValues As Variant, lngRowNumber As Long
  Dim lngEquipmentIndex As Long, lngPMIndex As Long, lngPMItemIndex As Long
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Exporting PM schedule ..."
  End If
  
  If strParameterString <> "" Then
    If Not (m_objErrorRange Is Nothing) Then m_objErrorRange.Value = "Invalid parameters"
    Exit Function
  End If
  
  On Error Resume Next
  Set objWorksheet = objWorkbook.Worksheets.Item("PMSchedule")
  If objWorksheet Is Nothing Then
    If objAfterWorksheet Is Nothing Then
      Set objWorksheet = objWorkbook.Worksheets.Add()
    Else
      Set objWorksheet = objWorkbook.Worksheets.Add(, , objAfterWorksheet)
    End If
    objWorksheet.Name = "PMSchedule"
    exportColumnHeadings objWorksheet, c_strColHeadings_PMSchedule
  End If
  On Error GoTo 0
  
  ' Clear current contents (if any)
  lngRowNumber = objWorksheet.Range("A3:A3").CurrentRegion.Rows.Count
  objWorksheet.Range("A3:P" & (2 + lngRowNumber)).Clear
  
  ' Export daily PM schedule
  varValues = objWorksheet.Range("A3:P" & (2 + udtEquipment.lngEquipmentCount)).Value
  For lngEquipmentIndex = 0 To udtEquipment.lngEquipmentCount - 1
    With udtEquipment.udtEquipment(lngEquipmentIndex)
      lngRowNumber = lngEquipmentIndex + 1
      varValues(lngRowNumber, 1) = udtAirports.udtAirport(.lngAirportIndex).strCode
      varValues(lngRowNumber, 2) = .lngID
      varValues(lngRowNumber, 3) = udtEquipmentModels.udtEquipmentModels(.lngEquipmentModelIndex).strManufacturer _
          & "." & udtEquipmentModels.udtEquipmentModels(.lngEquipmentModelIndex).strModel
      For lngPMIndex = 0 To .udtPMSchedule.lngPMItemCount - 1
        If lngPMIndex = 12 Then
          varValues(lngRowNumber, 16) = "etc."
          Exit For
        Else
          varValues(lngRowNumber, 4 + lngPMIndex) = m_udtPMItems(.udtPMSchedule.lngPMItemIndexes(lngPMIndex)).dteScheduledStart
        End If
      Next
    End With
  Next
  objWorksheet.Range("A3:P" & (2 + udtEquipment.lngEquipmentCount)).Value = varValues
  
  If Not (m_objStatusRange Is Nothing) Then
    m_objStatusRange.Value = "Exporting daily PM times ... done"
  End If
  
  Set exportPMSchedule = objWorksheet
  Set objWorksheet = Nothing

End Function


Private Sub scheduleAirportPM( _
    ByRef udtAirport As AIRPORT_TYPE, udtEquipment As EQUIPMENT_TYPE)

  Dim udtPMItemsByDate() As ITEMLIST_TYPE
  Dim lngAirportEquipmentModelIndex As Long, lngAirportEquipmentIndex As Long, _
      lngEquipmentPMIndex As Long, lngPMItemIndex As Long, lngDayIndex As Long
  Dim lngIndex As Long
  
  ' Create structure to store PM items that occur each day
  ReDim udtPMItemsByDate(0 To CLng(m_dteModelEndDate - m_dteModelStartDate))
  For lngAirportEquipmentModelIndex = 0 To udtAirport.lngEquipmentCount - 1
    With udtAirport.udtEquipment(lngAirportEquipmentModelIndex)
      For lngAirportEquipmentIndex = 0 To .lngCount - 1
        With udtEquipment.udtEquipment(.lngEquipmentIndexes(lngAirportEquipmentIndex))
          For lngEquipmentPMIndex = 0 To .udtPMSchedule.lngPMItemCount - 1
            lngPMItemIndex = .udtPMSchedule.lngPMItemIndexes(lngEquipmentPMIndex)
            With m_udtPMItems(lngPMItemIndex)
              lngDayIndex = Fix(.dteScheduledStart) - m_dteModelStartDate
              With udtPMItemsByDate(lngDayIndex)
                If .lngCount = 0 Then
                  ReDim .lngItemIndexes(0 To 31)
                ElseIf UBound(.lngItemIndexes) < .lngCount Then
                  ReDim .lngItemIndexes(0 To 31 + .lngCount)
                End If
                .lngItemIndexes(.lngCount) = lngPMItemIndex
                .lngCount = .lngCount + 1
              End With
            End With
          Next
        End With
      Next
    End With
  Next
              
'  For lngDayIndex = 0 To UBound(udtPMItemsByDate)
'    If 0 < udtPMItemsByDate.lngCount Then
'
'
'
'    End If
'  Next
Err.Raise 5



End Sub



Public Function computeDistance( _
    ByVal dblLongitude1 As Double, ByVal dblLatitude1 As Double, _
    ByVal dblLongitude2 As Double, ByVal dblLatitude2 As Double) As Double

  computeDistance = Sqr((c_dblEarthRadius_Equitorial * Cos(0.5 * (dblLatitude1 + dblLatitude2) * c_dblPiOver180) * c_dblPiOver180 * (dblLongitude2 - dblLongitude1)) ^ 2 _
      + (c_dblEarthRadius_Polar * c_dblPiOver180 * (dblLatitude2 - dblLatitude1)) ^ 2)

End Function


Private Function recordModelStep( _
    ByVal enmStepType As MODELSTEPTYPE_ENUM, strStepParameter As String) As Long

  If UBound(m_udtModelSteps) < m_lngModelStepCount Then
    ReDim Preserve m_udtModelSteps(0 To 63 + m_lngModelStepCount)
  End If
  With m_udtModelSteps(m_lngModelStepCount)
    .enmStepType = enmStepType
    ReDim .strParameters(0 To 0)
    parseParameters strStepParameter, .strParameters, .lngParameterCount
    .strStatus = ""
  End With
  m_lngModelStepCount = m_lngModelStepCount + 1
  
  recordModelStep = m_lngModelStepCount - 1
  
End Function


Private Sub recordError( _
    ByRef udtModelStep As MODELSTEP_TYPE, strErrorMessage As String)
    
  If Not (m_objErrorRange Is Nothing) Then m_objErrorRange.Value = strErrorMessage
  udtModelStep.strStatus = strErrorMessage

End Sub


Public Sub runServiceAreaModel( _
    ByVal lngServiceAreaCount As Long, ByVal dblMaximumMileage As Double, _
    ByVal lngCommunitySize As Long, ByVal lngRetentionCount As Long, _
    ByVal dblMatePercent As Double, strMateApproach1 As String, strMateApproach2 As String, _
    ByVal dblInsertDeleteWeight As Double)
    
  Dim dblTripCounts() As Double, dblTripTravelTimes() As Double, dblTripRepairTimes() As Double
    
'  If Not loadAirports(objWorkbook, "") Then Err.Raise 5
'  If Not loadEquipmentModels(objWorkbook, "") Then Err.Raise 5
'  If Not loadEquipmentPM(objWorkbook, "") Then Err.Raise 5
'  If Not loadEquipmentCM(objWorkbook, "") Then Err.Raise 5
'  If Not loadEquipment(objWorkbook, "ByCount") Then Err.Raise 5
  
  ' Create airport trip counts
  
  
  ' Compute airport distances
  
  
  
  



End Sub



Public Sub serviceAreaModelMate( _
    udtAirports As AIRPORTS_TYPE, _
    ByVal lngServiceAreaCount As Long, _
    ByRef udtServiceArea1 As ITEMINDEXLIST_TYPE, ByRef udtServiceArea2 As ITEMINDEXLIST_TYPE, _
    ByRef udtResult As ITEMINDEXLIST_TYPE)

  Dim lngIndex1 As Long, lngIndex2 As Long
  Dim lngAirportIndex As Long
  
  udtResult.lngItemCount = 0
  If UBound(udtResult.lngItemIndexes) < udtServiceArea1.lngItemCount + udtServiceArea2.lngItemCount - 1 Then
    ReDim udtResult.lngItemIndexes(0 To udtServiceArea1.lngItemCount + udtServiceArea2.lngItemCount - 1)
  End If

  Do While True

    If udtServiceArea1.lngItemIndexes(lngIndex1) = udtServiceArea2.lngItemIndexes(lngIndex2) Then
      udtResult.lngItemIndexes(udtResult.lngItemCount) = udtServiceArea1.lngItemIndexes(lngIndex1)
      udtResult.lngItemCount = udtResult.lngItemCount + 1
      lngIndex1 = lngIndex1 + 1
      lngIndex2 = lngIndex2 + 1
      If lngIndex1 = udtServiceArea1.lngItemCount Then Exit Do
      If lngIndex2 = udtServiceArea2.lngItemCount Then Exit Do
      
    ElseIf udtServiceArea1.lngItemIndexes(lngIndex1) < udtServiceArea2.lngItemIndexes(lngIndex2) Then
      If Rnd() < 0.5 Then
        udtResult.lngItemIndexes(udtResult.lngItemCount) = udtServiceArea1.lngItemIndexes(lngIndex1)
        udtResult.lngItemCount = udtResult.lngItemCount + 1
      End If
      lngIndex1 = lngIndex1 + 1
      If lngIndex1 = lngServiceAreaCount Then Exit Do
      
    Else
      If Rnd() < 0.5 Then
        udtResult.lngItemIndexes(udtResult.lngItemCount) = udtServiceArea2.lngItemIndexes(lngIndex2)
        udtResult.lngItemCount = udtResult.lngItemCount + 1
      End If
      lngIndex2 = lngIndex2 + 1
      If lngIndex2 = lngServiceAreaCount Then Exit Do
      
    End If
    
  Loop
  
  For lngIndex1 = lngIndex1 To lngServiceAreaCount - 1
    If Rnd() < 0.5 Then
      udtResult.lngItemIndexes(udtResult.lngItemCount) = udtServiceArea1.lngItemIndexes(lngIndex1)
      udtResult.lngItemCount = udtResult.lngItemCount + 1
    End If
  Next
  For lngIndex2 = lngIndex2 To lngServiceAreaCount - 1
    If Rnd() < 0.5 Then
      udtResult.lngItemIndexes(udtResult.lngItemCount) = udtServiceArea2.lngItemIndexes(lngIndex2)
      udtResult.lngItemCount = udtResult.lngItemCount + 1
    End If
  Next

  Do While udtResult.lngItemCount < lngServiceAreaCount
    insertItemIndex udtResult, Fix(udtAirports.lngAirportCount * Rnd())
  Loop
  Do While lngServiceAreaCount < udtResult.lngItemCount
    removeItemIndex udtResult, udtResult.lngItemIndexes(Fix(udtResult.lngItemCount * Rnd()))
  Loop

End Sub

Public Sub serviceAreaModelDelete( _
    ByRef udtServiceArea As ITEMINDEXLIST_TYPE, _
    ByRef udtResult As ITEMINDEXLIST_TYPE)

  udtResult = udtServiceArea
  removeItemIndex udtResult, udtResult.lngItemIndexes(Fix(udtResult.lngItemCount * Rnd()))

End Sub

Public Sub serviceAreaModelInsert( _
    udtAirports As AIRPORTS_TYPE, _
    ByRef udtServiceArea As ITEMINDEXLIST_TYPE, _
    ByRef udtResult As ITEMINDEXLIST_TYPE)

  udtResult = udtServiceArea
  Do While Not insertItemIndex(udtResult, Fix(udtAirports.lngAirportCount * Rnd()))
  Loop

End Sub

Public Sub serviceModelGenerateRandom( _
    udtAirports As AIRPORTS_TYPE, _
    ByVal lngServiceAreaCount As Long, _
    ByRef udtResult As ITEMINDEXLIST_TYPE)

  udtResult.lngItemCount = 0
  If UBound(udtResult.lngItemIndexes) < lngServiceAreaCount Then
    ReDim udtResult.lngItemIndexes(0 To lngServiceAreaCount - 1)
  End If
  Do While udtResult.lngItemCount < lngServiceAreaCount
    insertItemIndex udtResult, Fix(udtAirports.lngAirportCount * Rnd())
  Loop

End Sub

Private Function findItemIndex( _
    udtItemIndexList As ITEMINDEXLIST_TYPE, ByVal lngItemIndex As Long, _
    ByRef lngListIndex As Long) As Boolean

  Dim lngIndexLow As Long, lngIndexMid As Long, lngIndexHigh As Long
  
  lngIndexHigh = udtItemIndexList.lngItemCount - 1
  Do While lngIndexLow + 1 < lngIndexHigh
    lngIndexMid = (lngIndexLow + lngIndexHigh) \ 2
    If udtItemIndexList.lngItemIndexes(lngIndexMid) < lngItemIndex Then
      lngIndexLow = lngIndexMid
    Else
      lngIndexHigh = lngIndexMid
    End If
  Loop
  
  If lngItemIndex <= udtItemIndexList.lngItemIndexes(lngIndexLow) Then
    lngListIndex = lngIndexLow
    findItemIndex = (lngItemIndex = udtItemIndexList.lngItemIndexes(lngIndexLow))
  ElseIf lngItemIndex <= udtItemIndexList.lngItemIndexes(lngIndexHigh) Then
    lngListIndex = lngIndexHigh
    findItemIndex = (lngItemIndex = udtItemIndexList.lngItemIndexes(lngIndexHigh))
  Else
    lngListIndex = lngIndexHigh + 1
  End If

End Function

Private Function insertItemIndex( _
    udtItemIndexList As ITEMINDEXLIST_TYPE, ByVal lngItemIndex As Long) As Boolean
    
  Dim lngListIndex As Long
  
  If Not findItemIndex(udtItemIndexList, lngItemIndex, lngListIndex) Then
    If UBound(udtItemIndexList.lngItemIndexes) < udtItemIndexList.lngItemCount Then
      ReDim Preserve udtItemIndexList.lngItemIndexes(0 To 7 + udtItemIndexList.lngItemCount)
    End If
    ' OPTIMIZE: Replace with MemCopy
    For lngListIndex = udtItemIndexList.lngItemCount To lngListIndex + 1 Step -1
      udtItemIndexList.lngItemIndexes(lngListIndex) = udtItemIndexList.lngItemIndexes(lngListIndex - 1)
    Next
    udtItemIndexList.lngItemIndexes(lngListIndex) = lngItemIndex
    udtItemIndexList.lngItemCount = udtItemIndexList.lngItemCount + 1
    insertItemIndex = True
  End If
End Function


Private Function removeItemIndex( _
    udtItemIndexList As ITEMINDEXLIST_TYPE, ByVal lngItemIndex As Long) As Boolean
    
  Dim lngListIndex As Long
  
  If findItemIndex(udtItemIndexList, lngItemIndex, lngListIndex) Then
    ' OPTIMIZE: Replace with MemCopy
    For lngListIndex = lngListIndex To udtItemIndexList.lngItemCount - 2
      udtItemIndexList.lngItemIndexes(lngListIndex) = udtItemIndexList.lngItemIndexes(lngListIndex + 1)
    Next
    udtItemIndexList.lngItemCount = udtItemIndexList.lngItemCount - 1
    removeItemIndex = True
  End If

End Function
