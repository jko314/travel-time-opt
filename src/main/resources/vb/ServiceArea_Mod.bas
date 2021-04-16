Attribute VB_Name = "ServiceArea_Mod"
Option Explicit


Private Type AIRPORTMAINTDATA_TYPE
  dblPMTripCount As Double
  dblPMTime As Double
  dblCMTripCount As Double
  dblCMTime As Double
  dblDepotTripCount As Double
  dblDepotTime As Double
  dblTripCount As Double
  dblTime As Double
End Type


Private Type SERVICEAREADATA_TYPE
  lngAirportIndex As Long
  dblLongitude As Double
  dblLatitude As Double
  dblTravelMiles As Double
  dblPMTripCount As Double
  dblPMTime As Double
  dblCMTripCount As Double
  dblCMTime As Double
  '
  dblFSTTime As Double
  lngFSTCount As Long
End Type

Private Type AIRPORTSERVICEAREADATA_TYPE
  lngServiceAreaIndex As Long
  dblTravelDistance As Double
End Type

Public Type POSSIBLESOLUTION_TYPE
  booResultsAreValid As Boolean
  lngServiceAreaCount As Long
  udtServiceAreaData() As SERVICEAREADATA_TYPE
  booServiceAreasAreSorted As Boolean
  udtAirportData() As AIRPORTSERVICEAREADATA_TYPE
  dblOptimizeOn As Double
  '
  dblTravelMiles As Double
  dblPMTripCount As Double
  dblPMTime As Double
  dblCMTripCount As Double
  dblCMTime As Double
  dblFSTHours As Double
  lngFSTCount As Long
  dblFSTCost As Double
End Type
  
Private Enum EVOLUTIONAPPROACH_ENUM
  EVOLAPP_RetainTop = 1
  EVOLAPP_RetainRandom = 2
  EVOLAPP_Mate = 3
  EVOLAPP_InsertDelete = 4
  EVOLAPP_Random = 5
End Enum
Private Type EVOLUTION_TYPE
  enmApproach As EVOLUTIONAPPROACH_ENUM
  lngAppliesTo As Long
  lngParameterCount As Long
  VarParameters() As Variant
End Type
  


Private Type PIECEWISELINEARITEM_TYPE
  dblFromValue As Double
  dblToValue As Double
  dblB As Double
  dblM As Double
End Type

Private Type PIECEWISELINEAR_TYPE
  lngItemCount As Long
  udtItems() As PIECEWISELINEARITEM_TYPE
End Type


Public Type SERVICEAREAMODEL_TYPE
  udtAirports As AIRPORTS_TYPE
  udtServiceAreas As SERVICEAREAS_TYPE
  udtEquipmentTypes As EQUIPMENTTYPES_TYPE
  udtEquipmentModels As EQUIPMENTMODELS_TYPE
  udtCMRequirements As CMREQUIREMENTS_TYPE
  dblAirportDistances() As Double
  ' Model Parameters
  lngServiceAreaCount As Long
  lngServiceAreaCount_Min As Long
  lngServiceAreaCount_Max As Long
  dblMaximumTravelMiles As Double
  lngCommunitySize As Long
  lngEvolutionCount As Long
  udtEvolutions() As EVOLUTION_TYPE
  dblAirportSelectionExponent As Double
  dblSolutionSelectionExponent As Double
  dblMaxPMTimePerTrip As Double
  '
  lngIterationNumber As Long
  lngIterationCount As Long
  ' Work Variables
  dblSortValues() As Double
  ' Model Variables
  dblAirportSelectionP() As Double
  udtEquipment As EQUIPMENT_TYPE
  udtAirportData() As AIRPORTMAINTDATA_TYPE
  lngSolutionCount As Long
  udtSolutions() As POSSIBLESOLUTION_TYPE
  booSolutionSelectionValid As Boolean
  dblSolutionSelectionP() As Double
  booSolutionSortIsValid As Boolean
  lngSolutionSortIndexes() As Long
  ' Cost factors
  dblCostPerMile As Double
  dblCostPerFSTHour As Double
  dblFSTHoursPerYear As Double
  udtFSTTimeToCount As PIECEWISELINEAR_TYPE
End Type






Private Function getWorksheet(ByVal colWorkbooks As Collection, strWorkbook As String, strWorksheet As String) As Excel.Worksheet

  Dim objWorkbook As Excel.Workbook

  On Error Resume Next
  Set objWorkbook = Nothing
  Set objWorkbook = colWorkbooks.Item(UCase$(strWorkbook))
  On Error GoTo 0
  If objWorkbook Is Nothing Then
    Set objWorkbook = Application.Workbooks.Open(strWorkbook)
    colWorkbooks.Add objWorkbook, UCase$(strWorkbook)
  End If
  
  On Error Resume Next
  Set getWorksheet = objWorkbook.Worksheets.Item(strWorksheet)
  On Error GoTo 0
  Set objWorkbook = Nothing
  
End Function


Private Sub openWorkbooks(ByRef colWorkbooks As Collection)

  Set colWorkbooks = New Collection
  colWorkbooks.Add Application.ActiveWorkbook, "[[ME]]"
  
End Sub


Private Sub closeWorkbooks(ByVal colWorkbooks As Collection)

  Dim objWorkbook As Excel.Workbook

  Do While 1 < colWorkbooks.Count
    Set objWorkbook = colWorkbooks.Item(colWorkbooks.Count)
    colWorkbooks.Remove colWorkbooks.Count
    objWorkbook.Close
    Set objWorkbook = Nothing
  Loop
  Set colWorkbooks = Nothing
  
End Sub



Public Sub runNewModel_Click()

  Dim udtServiceAreaModel As SERVICEAREAMODEL_TYPE
  Dim objModelWorksheet As Excel.Worksheet
  Dim objWorkbook As Excel.Workbook, objWorksheet As Excel.Worksheet, colWorkbooks As Collection
  Dim lngIndex As Long

  Set objModelWorksheet = Application.ActiveSheet
  openWorkbooks colWorkbooks

  Set objWorksheet = getWorksheet(colWorkbooks, CStr(objModelWorksheet.Range("B6").Value), CStr(objModelWorksheet.Range("C6").Value))
  If Not readOptimizationParameters(udtServiceAreaModel, objWorksheet) Then Err.Raise 5
  
  initializeModel CDate("1/1/2013"), CDate("1/1/2014")
  
  Set objWorksheet = getWorksheet(colWorkbooks, CStr(objModelWorksheet.Range("B7").Value), CStr(objModelWorksheet.Range("C7").Value))
  If Not loadAirports(udtServiceAreaModel.udtAirports, objWorksheet, "") Then Err.Raise 5
  
  Set objWorksheet = getWorksheet(colWorkbooks, CStr(objModelWorksheet.Range("B8").Value), CStr(objModelWorksheet.Range("C8").Value))
  If Not loadEquipmentModels(udtServiceAreaModel.udtEquipmentModels, udtServiceAreaModel.udtEquipmentTypes, objWorksheet, "") Then Err.Raise 5
  
  Set objWorksheet = getWorksheet(colWorkbooks, CStr(objModelWorksheet.Range("B9").Value), CStr(objModelWorksheet.Range("C9").Value))
  If Not loadEquipment(udtServiceAreaModel.udtAirports, udtServiceAreaModel.udtEquipmentModels, udtServiceAreaModel.udtEquipment, objWorksheet, "ByCount") Then Err.Raise 5
  
  Set objWorksheet = getWorksheet(colWorkbooks, CStr(objModelWorksheet.Range("B10").Value), CStr(objModelWorksheet.Range("C10").Value))
  If Not loadEquipmentPM(udtServiceAreaModel.udtAirports, udtServiceAreaModel.udtEquipmentModels, objWorksheet, "ByType") Then Err.Raise 5
  
  Set objWorksheet = getWorksheet(colWorkbooks, CStr(objModelWorksheet.Range("B11").Value), CStr(objModelWorksheet.Range("C11").Value))
  If Not loadCMRequirements(udtServiceAreaModel.udtCMRequirements, objWorksheet, "") Then Err.Raise 5




'zzzzz
  closeWorkbooks colWorkbooks
  
  Set objModelWorksheet = Nothing

End Sub


Public Sub runServiceAreaModel_Click()

  Dim udtServiceAreaModel As SERVICEAREAMODEL_TYPE
  Dim objWorkbook As Excel.Workbook, objWorksheet As Excel.Worksheet, _
      objIterationNumberRange As Excel.Range, _
      objCurrentSolutionsRange As Excel.Range, varCurrentSolutionValue As Variant
  Dim lngSolutionIndex As Long, lngIterationNumber As Long, lngSACount As Long
  Dim lngIndex As Long, dblSortValues() As Double

  Set objWorkbook = Application.ActiveWorkbook
  Set objWorksheet = objWorkbook.ActiveSheet
  Set objIterationNumberRange = objWorksheet.Range("B22")
  Set objCurrentSolutionsRange = objWorksheet.Range("B23:B27")
  varCurrentSolutionValue = objCurrentSolutionsRange.Value
  
  ' BUG: Hard-coded values
  udtServiceAreaModel.dblCostPerMile = 0.9
  udtServiceAreaModel.dblCostPerFSTHour = 60
  udtServiceAreaModel.dblFSTHoursPerYear = 2080
  setFSTTimeToCountModel udtServiceAreaModel.udtFSTTimeToCount, CStr(objWorksheet.Range("H6").Value)
    
  
  ' Read parameters
  udtServiceAreaModel.lngServiceAreaCount = objWorksheet.Range("B2").Value
  udtServiceAreaModel.lngServiceAreaCount_Min = objWorksheet.Range("C2").Value
  udtServiceAreaModel.lngServiceAreaCount_Max = objWorksheet.Range("D2").Value
  udtServiceAreaModel.dblMaximumTravelMiles = objWorksheet.Range("B3").Value
  udtServiceAreaModel.lngCommunitySize = objWorksheet.Range("B4").Value
  ReDim udtServiceAreaModel.udtEvolutions(0 To 3)
  With udtServiceAreaModel.udtEvolutions(0)
    .enmApproach = EVOLAPP_RetainTop
    .lngAppliesTo = objWorksheet.Range("B5")
    .lngParameterCount = 0
  End With
  With udtServiceAreaModel.udtEvolutions(1)
    .enmApproach = EVOLAPP_Mate
    .lngAppliesTo = objWorksheet.Range("B6")
    .lngParameterCount = 2
    ReDim .VarParameters(0 To 1)
    .VarParameters(0) = objWorksheet.Range("B7")
    .VarParameters(1) = objWorksheet.Range("B8")
  End With
  With udtServiceAreaModel.udtEvolutions(2)
    .enmApproach = EVOLAPP_InsertDelete
    .lngAppliesTo = objWorksheet.Range("B9")
    .lngParameterCount = 0
  End With
  With udtServiceAreaModel.udtEvolutions(3)
    .enmApproach = EVOLAPP_Random
    .lngAppliesTo = objWorksheet.Range("B10")
    .lngParameterCount = 0
  End With
  udtServiceAreaModel.lngEvolutionCount = 4
  udtServiceAreaModel.dblAirportSelectionExponent = objWorksheet.Range("B11")
  udtServiceAreaModel.dblSolutionSelectionExponent = objWorksheet.Range("B12")
  udtServiceAreaModel.dblMaxPMTimePerTrip = objWorksheet.Range("B13")
  udtServiceAreaModel.lngIterationCount = objWorksheet.Range("B17")
  
  ' Load airport/equipment variables
  initializeModel CDate("1/1/2013"), CDate("1/1/2014")
  loadAirports udtServiceAreaModel.udtAirports, objWorkbook.Worksheets.Item("Airports"), ""
  loadEquipmentModels udtServiceAreaModel.udtEquipmentModels, udtServiceAreaModel.udtEquipmentTypes, objWorkbook.Worksheets.Item("EquipmentModels"), ""
  loadEquipment udtServiceAreaModel.udtAirports, udtServiceAreaModel.udtEquipmentModels, _
      udtServiceAreaModel.udtEquipment, objWorkbook.Worksheets.Item("Airport_Equipment"), "ByCount"
  loadEquipmentPM udtServiceAreaModel.udtAirports, udtServiceAreaModel.udtEquipmentModels, _
      objWorkbook.Worksheets.Item("EquipmentPM"), "ByType"
  loadCMRequirements udtServiceAreaModel.udtCMRequirements, objWorkbook.Worksheets.Item("EquipmentCM"), ""
  
  applyCMRequirements udtServiceAreaModel.udtCMRequirements, udtServiceAreaModel.udtAirports, udtServiceAreaModel.udtEquipmentModels
  computeAirportDistances udtServiceAreaModel.udtAirports, udtServiceAreaModel.dblAirportDistances
  
  If udtServiceAreaModel.udtAirports.lngAirportCount < udtServiceAreaModel.lngCommunitySize Then
    ReDim udtServiceAreaModel.dblSortValues(0 To udtServiceAreaModel.lngCommunitySize - 1)
  Else
    ReDim udtServiceAreaModel.dblSortValues(0 To udtServiceAreaModel.udtAirports.lngAirportCount - 1)
  End If
  ReDim udtServiceAreaModel.lngSolutionSortIndexes(0 To udtServiceAreaModel.lngCommunitySize - 1)
  computeAirportData udtServiceAreaModel, udtServiceAreaModel.dblMaxPMTimePerTrip

  ' Generate starting set of solutions
  ReDim udtServiceAreaModel.dblAirportSelectionP(0 To udtServiceAreaModel.udtAirports.lngAirportCount - 1)
  assignAirportSelectionP udtServiceAreaModel, udtServiceAreaModel.dblAirportSelectionExponent
  ReDim udtServiceAreaModel.udtSolutions(0 To udtServiceAreaModel.lngCommunitySize - 1)
  For lngSolutionIndex = 0 To udtServiceAreaModel.lngCommunitySize - 1
    ReDim udtServiceAreaModel.udtSolutions(lngSolutionIndex).udtServiceAreaData(0 To udtServiceAreaModel.lngServiceAreaCount_Max - 1), _
          udtServiceAreaModel.udtSolutions(lngSolutionIndex).udtAirportData(0 To udtServiceAreaModel.udtAirports.lngAirportCount - 1)
    lngSACount = udtServiceAreaModel.lngServiceAreaCount_Min + Fix(Rnd() * (1 + udtServiceAreaModel.lngServiceAreaCount_Max - udtServiceAreaModel.lngServiceAreaCount_Min))
    randomSolutionServiceArea udtServiceAreaModel, udtServiceAreaModel.udtSolutions(lngSolutionIndex), lngSACount
    evaluateSolution udtServiceAreaModel, udtServiceAreaModel.udtSolutions(lngSolutionIndex), udtServiceAreaModel.dblMaximumTravelMiles
  Next
  ReDim udtServiceAreaModel.dblSolutionSelectionP(0 To udtServiceAreaModel.lngCommunitySize - 1)
  
  ' Iterate
  For lngIterationNumber = 1 To udtServiceAreaModel.lngIterationCount
  
    ' Evolve solution set
    evolveSolutions udtServiceAreaModel
    replaceDuplicateSolutions udtServiceAreaModel
    udtServiceAreaModel.booSolutionSortIsValid = False
    
    ' Update display
    objIterationNumberRange.Value = lngIterationNumber
    If Not udtServiceAreaModel.booSolutionSortIsValid Then sortSolutions udtServiceAreaModel
    For lngIndex = 1 To 5
      lngSolutionIndex = udtServiceAreaModel.lngSolutionSortIndexes(lngIndex)
      varCurrentSolutionValue(lngIndex, 1) = udtServiceAreaModel.udtSolutions(lngSolutionIndex).dblOptimizeOn
    Next
    objCurrentSolutionsRange.Value = varCurrentSolutionValue
    DoEvents
   
  Next

  displaySolution udtServiceAreaModel, udtServiceAreaModel.udtSolutions(udtServiceAreaModel.lngSolutionSortIndexes(0)), objWorksheet, "F", 22, "N"
  
  Set objIterationNumberRange = Nothing
  Set objCurrentSolutionsRange = Nothing
  Set objWorksheet = Nothing
  Set objWorkbook = Nothing
  
End Sub



Public Sub runServiceAreaEstimate_Click()

  Dim udtServiceAreaModel As SERVICEAREAMODEL_TYPE
  Dim objWorkbook As Excel.Workbook, objWorksheet As Excel.Worksheet, _
      objIterationNumberRange As Excel.Range, _
      objCurrentSolutionsRange As Excel.Range, varCurrentSolutionValue As Variant
  Dim lngSolutionIndex As Long, lngIterationNumber As Long
  Dim lngIndex As Long, dblSortValues() As Double, varValues As Variant

  Set objWorkbook = Application.ActiveWorkbook
  Set objWorksheet = objWorkbook.ActiveSheet
  
  ' Read parameters
  udtServiceAreaModel.lngServiceAreaCount = objWorksheet.Range("A12").CurrentRegion.Rows.Count
  udtServiceAreaModel.dblMaximumTravelMiles = objWorksheet.Range("B4").Value
  udtServiceAreaModel.dblMaxPMTimePerTrip = objWorksheet.Range("B5")
  udtServiceAreaModel.lngIterationCount = 0
  
  ' BUG: Hard-coded values
  udtServiceAreaModel.dblCostPerMile = 0.9
  udtServiceAreaModel.dblCostPerFSTHour = 60
  udtServiceAreaModel.dblFSTHoursPerYear = 2080
  setFSTTimeToCountModel udtServiceAreaModel.udtFSTTimeToCount, CStr(objWorkbook.Worksheets.Item("ServiceAreaModel").Range("H6").Value)

  
  ' Load airport/equipment variables
  initializeModel CDate("1/1/2013"), CDate("1/1/2014")
  loadAirports udtServiceAreaModel.udtAirports, objWorkbook, ""
  loadEquipmentModels udtServiceAreaModel.udtEquipmentModels, udtServiceAreaModel.udtEquipmentTypes, objWorkbook, ""
  loadEquipment udtServiceAreaModel.udtAirports, udtServiceAreaModel.udtEquipmentModels, _
      udtServiceAreaModel.udtEquipment, objWorkbook, "ByCount"
  loadEquipmentPM udtServiceAreaModel.udtAirports, udtServiceAreaModel.udtEquipmentModels, _
      objWorkbook, "ByType"
  loadCMRequirements udtServiceAreaModel.udtCMRequirements, objWorkbook, ""
  applyCMRequirements udtServiceAreaModel.udtCMRequirements, udtServiceAreaModel.udtAirports, udtServiceAreaModel.udtEquipmentModels
  computeAirportDistances udtServiceAreaModel.udtAirports, udtServiceAreaModel.dblAirportDistances
  
  computeAirportData udtServiceAreaModel, udtServiceAreaModel.dblMaxPMTimePerTrip

  ' Generate starting set of solutions
  varValues = objWorksheet.Range("A12:A" & (11 + udtServiceAreaModel.lngServiceAreaCount)).Value
  ReDim udtServiceAreaModel.udtSolutions(0 To 0), _
        udtServiceAreaModel.udtSolutions(0).udtServiceAreaData(0 To udtServiceAreaModel.lngServiceAreaCount), _
        udtServiceAreaModel.udtSolutions(0).udtAirportData(0 To udtServiceAreaModel.udtAirports.lngAirportCount)
  For lngIndex = 1 To udtServiceAreaModel.lngServiceAreaCount
    insertSolutionServiceArea udtServiceAreaModel, udtServiceAreaModel.udtSolutions(0), udtServiceAreaModel.udtAirports.colAirports.Item("A:" & varValues(lngIndex, 1))
  Next
  evaluateSolution udtServiceAreaModel, udtServiceAreaModel.udtSolutions(0), udtServiceAreaModel.dblMaximumTravelMiles

  displaySolution udtServiceAreaModel, udtServiceAreaModel.udtSolutions(0), objWorksheet, "F", 12, "N"
  
  Set objIterationNumberRange = Nothing
  Set objCurrentSolutionsRange = Nothing
  Set objWorksheet = Nothing
  Set objWorkbook = Nothing
  
End Sub


Private Sub displaySolution( _
    udtServiceAreaModel As SERVICEAREAMODEL_TYPE, _
    udtSolution As POSSIBLESOLUTION_TYPE, _
    objWorksheet As Excel.Worksheet, strStartingColumn As String, lngStartingRow As Long, strSAColumn As String)

  Dim strRange As String, varValue As Variant, lngAirportIndex As Long, _
      lngServiceAreaIndex As Long, lngServiceAreaAirportIndex As Long
  Dim dblPMTripCount As Double, dblPMTime As Double, dblCMTripCount As Double, dblCMTime As Double
  Dim lngIndex1 As Long, lngIndex2 As Long, lngRowIndex As Long, strServiceAreaList As String
  Dim dblTrips0 As Double, dblTrips60 As Double, dblTrips120 As Double, dblTrips200 As Double, dblTripsMore As Double, _
      dblTripsAvg As Double, lngTripIndex As Long
  Dim dblNoTravelTripCount() As Double, dblTravelDistance As Double

  ' Evaluate solution, with details
  ' Note: Must correct trips counts for no-travel trips, which are ignored in evaluateSolution
  evaluateSolution udtServiceAreaModel, udtSolution, udtServiceAreaModel.dblMaximumTravelMiles
  
  ' Correct trips counts for no-travel trips
  ReDim dblNoTravelTripCount(0 To udtSolution.lngServiceAreaCount - 1)
  For lngAirportIndex = 0 To udtServiceAreaModel.udtAirports.lngAirportCount - 1
    lngServiceAreaIndex = udtSolution.udtAirportData(lngAirportIndex).lngServiceAreaIndex
    With udtSolution.udtServiceAreaData(lngServiceAreaIndex)
      If 0 = udtSolution.udtAirportData(lngAirportIndex).dblTravelDistance Then
        dblNoTravelTripCount(lngServiceAreaIndex) = dblNoTravelTripCount(lngServiceAreaIndex) _
            + udtServiceAreaModel.udtAirportData(lngAirportIndex).dblPMTripCount + udtServiceAreaModel.udtAirportData(lngAirportIndex).dblCMTripCount
      End If
    End With
  Next
 
  
  objWorksheet.Range(strStartingColumn & lngStartingRow).CurrentRegion.Clear
  strRange = strStartingColumn & lngStartingRow & ":" & colNumberToName(colNameToNumber(strStartingColumn) + 6) & (lngStartingRow + udtServiceAreaModel.udtAirports.lngAirportCount - 1)
  varValue = objWorksheet.Range(strRange).Value

  For lngAirportIndex = 0 To udtServiceAreaModel.udtAirports.lngAirportCount - 1
    lngServiceAreaIndex = udtSolution.udtAirportData(lngAirportIndex).lngServiceAreaIndex
    lngServiceAreaAirportIndex = udtSolution.udtServiceAreaData(lngServiceAreaIndex).lngAirportIndex
    varValue(1 + lngAirportIndex, 1) = udtServiceAreaModel.udtAirports.udtAirport(lngAirportIndex).strCode
    varValue(1 + lngAirportIndex, 2) = udtServiceAreaModel.udtAirports.udtAirport(lngServiceAreaAirportIndex).strCode
    varValue(1 + lngAirportIndex, 3) = udtServiceAreaModel.dblAirportDistances(lngAirportIndex, lngServiceAreaAirportIndex)
    varValue(1 + lngAirportIndex, 4) = udtServiceAreaModel.udtAirportData(lngAirportIndex).dblPMTripCount
    varValue(1 + lngAirportIndex, 5) = udtServiceAreaModel.udtAirportData(lngAirportIndex).dblPMTime
    varValue(1 + lngAirportIndex, 6) = udtServiceAreaModel.udtAirportData(lngAirportIndex).dblCMTripCount
    varValue(1 + lngAirportIndex, 7) = udtServiceAreaModel.udtAirportData(lngAirportIndex).dblCMTime
  Next
  objWorksheet.Range(strRange).Value = varValue
  
  objWorksheet.Range(strSAColumn & lngStartingRow).CurrentRegion.Clear
  strRange = strSAColumn & lngStartingRow & ":" & colNumberToName(colNameToNumber(strSAColumn) + 15) & (lngStartingRow + udtSolution.lngServiceAreaCount - 1)
  varValue = objWorksheet.Range(strRange).Value
  lngRowIndex = 1
  strServiceAreaList = ":"
  For lngServiceAreaIndex = 0 To udtSolution.lngServiceAreaCount - 1
    With udtSolution.udtServiceAreaData(lngServiceAreaIndex)

      If 0 < .dblFSTTime Then
        varValue(lngRowIndex, 1) = udtServiceAreaModel.udtAirports.udtAirport(.lngAirportIndex).strCode
        varValue(lngRowIndex, 2) = udtServiceAreaModel.udtAirports.udtAirport(.lngAirportIndex).strCity & ", " & udtServiceAreaModel.udtAirports.udtAirport(.lngAirportIndex).strState
        varValue(lngRowIndex, 3) = .dblPMTripCount
        varValue(lngRowIndex, 4) = .dblPMTime
        varValue(lngRowIndex, 5) = .dblCMTripCount
        varValue(lngRowIndex, 6) = .dblCMTime
        varValue(lngRowIndex, 7) = .dblTravelMiles
        varValue(lngRowIndex, 8) = .dblTravelMiles / 60
        varValue(lngRowIndex, 9) = .dblPMTime + .dblCMTime + .dblTravelMiles / 60
        varValue(lngRowIndex, 10) = .lngFSTCount
        If .lngFSTCount = 0 Then
          varValue(lngRowIndex, 11) = ""
        Else
          varValue(lngRowIndex, 11) = varValue(lngRowIndex, 9) / .lngFSTCount / udtServiceAreaModel.dblFSTHoursPerYear
        End If

        lngRowIndex = lngRowIndex + 1
        strServiceAreaList = strServiceAreaList & .lngAirportIndex & ":"
      End If
    End With
  Next
  objWorksheet.Range(strRange).Value = varValue
  
  objWorksheet.Range("L3").Value = lngRowIndex - 1
  objWorksheet.Range("M3").Value = udtServiceAreaModel.lngIterationCount
  objWorksheet.Range("N3").Value = udtSolution.dblTravelMiles
  objWorksheet.Range("O3").Value = udtSolution.dblFSTHours
  objWorksheet.Range("P3").Value = udtSolution.lngFSTCount
  objWorksheet.Range("Q3").Value = udtSolution.dblFSTHours / udtSolution.lngFSTCount / udtServiceAreaModel.dblFSTHoursPerYear
  objWorksheet.Range("R3").Value = udtSolution.dblFSTCost

End Sub


Private Sub replaceDuplicateSolutions(udtServiceAreaModel As SERVICEAREAMODEL_TYPE)

  Dim lngLastSolutionIndex As Long, lngSolutionIndex As Long
  Dim lngIndex As Long
  
  If Not udtServiceAreaModel.booSolutionSortIsValid Then sortSolutions udtServiceAreaModel
  
  lngLastSolutionIndex = udtServiceAreaModel.lngSolutionSortIndexes(0)
  For lngIndex = 1 To udtServiceAreaModel.lngCommunitySize - 1
  
    lngSolutionIndex = udtServiceAreaModel.lngSolutionSortIndexes(lngIndex)
    If compareSolutions(udtServiceAreaModel.udtSolutions(lngLastSolutionIndex), _
         udtServiceAreaModel.udtSolutions(lngSolutionIndex)) Then
      randomSolutionServiceArea udtServiceAreaModel, udtServiceAreaModel.udtSolutions(lngSolutionIndex), udtServiceAreaModel.lngServiceAreaCount
      evaluateSolution udtServiceAreaModel, udtServiceAreaModel.udtSolutions(lngSolutionIndex), udtServiceAreaModel.dblMaximumTravelMiles
      udtServiceAreaModel.booSolutionSortIsValid = False
      sortSolutions udtServiceAreaModel
    Else
      lngLastSolutionIndex = lngSolutionIndex
    End If
    
  Next

End Sub


Private Function compareSolutions( _
    udtSolution1 As POSSIBLESOLUTION_TYPE, udtSolution2 As POSSIBLESOLUTION_TYPE) As Boolean

  Dim lngServiceAreaIndex As Long
  
  If udtSolution1.lngServiceAreaCount <> udtSolution2.lngServiceAreaCount Then Exit Function
  For lngServiceAreaIndex = 0 To udtSolution1.lngServiceAreaCount - 1
    If udtSolution1.udtServiceAreaData(lngServiceAreaIndex).lngAirportIndex _
        <> udtSolution2.udtServiceAreaData(lngServiceAreaIndex).lngAirportIndex Then Exit Function
  Next
  compareSolutions = True
    
End Function



Private Sub evolveSolutions(udtServiceAreaModel As SERVICEAREAMODEL_TYPE)

  Dim lngEvolutionIndex As Long
  Dim lngIndex As Long, lngSolutionSortIndex As Long
  Dim lngSolutionIndex1 As Long, lngSolutionIndex2 As Long, lngSolutionIndex As Long
  
  If Not udtServiceAreaModel.booSolutionSortIsValid Then sortSolutions udtServiceAreaModel
  
  lngSolutionSortIndex = 0
  For lngEvolutionIndex = 0 To udtServiceAreaModel.lngEvolutionCount - 1
    With udtServiceAreaModel.udtEvolutions(lngEvolutionIndex)
    
      Select Case .enmApproach
      
        Case EVOLUTIONAPPROACH_ENUM.EVOLAPP_RetainTop
          lngSolutionSortIndex = lngSolutionSortIndex + .lngAppliesTo
        
        Case EVOLUTIONAPPROACH_ENUM.EVOLAPP_RetainRandom
          Err.Raise 5
          
        Case EVOLUTIONAPPROACH_ENUM.EVOLAPP_Mate
          For lngSolutionSortIndex = lngSolutionSortIndex To lngSolutionSortIndex + .lngAppliesTo - 1
            lngSolutionIndex = udtServiceAreaModel.lngSolutionSortIndexes(lngSolutionSortIndex)
            lngSolutionIndex1 = lngSolutionIndex
            Do While lngSolutionIndex1 = lngSolutionIndex
              lngSolutionIndex1 = selectSolution(udtServiceAreaModel, CStr(.VarParameters(0)))
            Loop
            lngSolutionIndex2 = lngSolutionIndex
            Do While lngSolutionIndex2 = lngSolutionIndex
              lngSolutionIndex2 = selectSolution(udtServiceAreaModel, CStr(.VarParameters(1)))
            Loop
            mateSolutions udtServiceAreaModel, udtServiceAreaModel.udtSolutions(lngSolutionIndex1), _
                udtServiceAreaModel.udtSolutions(lngSolutionIndex2), _
                udtServiceAreaModel.udtSolutions(lngSolutionIndex)
            If udtServiceAreaModel.udtSolutions(lngSolutionIndex).lngServiceAreaCount < udtServiceAreaModel.lngServiceAreaCount_Min Then
              setSolutionServiceAreaCount udtServiceAreaModel, udtServiceAreaModel.udtSolutions(lngSolutionIndex), udtServiceAreaModel.lngServiceAreaCount_Min
            ElseIf udtServiceAreaModel.lngServiceAreaCount_Max < udtServiceAreaModel.udtSolutions(lngSolutionIndex).lngServiceAreaCount Then
              setSolutionServiceAreaCount udtServiceAreaModel, udtServiceAreaModel.udtSolutions(lngSolutionIndex), udtServiceAreaModel.lngServiceAreaCount_Max
            End If
            evaluateSolution udtServiceAreaModel, udtServiceAreaModel.udtSolutions(lngSolutionIndex), udtServiceAreaModel.dblMaximumTravelMiles
          Next

        Case EVOLUTIONAPPROACH_ENUM.EVOLAPP_InsertDelete
          For lngSolutionSortIndex = lngSolutionSortIndex To lngSolutionSortIndex + .lngAppliesTo - 1
            lngSolutionIndex = udtServiceAreaModel.lngSolutionSortIndexes(lngSolutionSortIndex)
            lngSolutionIndex1 = selectSolution(udtServiceAreaModel, "RandomByMileage")
            udtServiceAreaModel.udtSolutions(lngSolutionIndex) = udtServiceAreaModel.udtSolutions(lngSolutionIndex1)
            lngIndex = Fix(Rnd() * 3)
            If (lngIndex = 0) And (udtServiceAreaModel.lngServiceAreaCount_Min < udtServiceAreaModel.udtSolutions(lngSolutionIndex).lngServiceAreaCount) Then
              removeSolutionServiceArea udtServiceAreaModel.udtSolutions(lngSolutionIndex), Fix(Rnd() * udtServiceAreaModel.udtSolutions(lngSolutionIndex).lngServiceAreaCount)
            ElseIf lngIndex <= 1 Then
              removeSolutionServiceArea udtServiceAreaModel.udtSolutions(lngSolutionIndex), Fix(Rnd() * udtServiceAreaModel.udtSolutions(lngSolutionIndex).lngServiceAreaCount)
              setSolutionServiceAreaCount udtServiceAreaModel, udtServiceAreaModel.udtSolutions(lngSolutionIndex), udtServiceAreaModel.udtSolutions(lngSolutionIndex).lngServiceAreaCount + 1
            ElseIf udtServiceAreaModel.udtSolutions(lngSolutionIndex).lngServiceAreaCount < udtServiceAreaModel.lngServiceAreaCount_Max Then
              setSolutionServiceAreaCount udtServiceAreaModel, udtServiceAreaModel.udtSolutions(lngSolutionIndex), udtServiceAreaModel.udtSolutions(lngSolutionIndex).lngServiceAreaCount + 1
            Else
              removeSolutionServiceArea udtServiceAreaModel.udtSolutions(lngSolutionIndex), Fix(Rnd() * udtServiceAreaModel.udtSolutions(lngSolutionIndex).lngServiceAreaCount)
              setSolutionServiceAreaCount udtServiceAreaModel, udtServiceAreaModel.udtSolutions(lngSolutionIndex), udtServiceAreaModel.udtSolutions(lngSolutionIndex).lngServiceAreaCount + 1
            End If
            evaluateSolution udtServiceAreaModel, udtServiceAreaModel.udtSolutions(lngSolutionIndex), udtServiceAreaModel.dblMaximumTravelMiles
          Next
        
        Case EVOLUTIONAPPROACH_ENUM.EVOLAPP_Random
          For lngSolutionSortIndex = lngSolutionSortIndex To lngSolutionSortIndex + .lngAppliesTo - 1
            lngSolutionIndex = udtServiceAreaModel.lngSolutionSortIndexes(lngSolutionSortIndex)
            lngIndex = udtServiceAreaModel.lngServiceAreaCount_Min + Fix(Rnd() * (1 + udtServiceAreaModel.lngServiceAreaCount_Max - udtServiceAreaModel.lngServiceAreaCount_Min))
            randomSolutionServiceArea udtServiceAreaModel, udtServiceAreaModel.udtSolutions(lngSolutionIndex), lngIndex
            evaluateSolution udtServiceAreaModel, udtServiceAreaModel.udtSolutions(lngSolutionIndex), udtServiceAreaModel.dblMaximumTravelMiles
          Next
          
        Case Else
          Err.Raise 5
          
      End Select
   
    End With
  Next
  

End Sub


Private Sub sortSolutions(udtServiceAreaModel As SERVICEAREAMODEL_TYPE)

  Dim lngIndex As Long

  If udtServiceAreaModel.booSolutionSortIsValid Then Exit Sub
  
  For lngIndex = 0 To udtServiceAreaModel.lngCommunitySize - 1
    udtServiceAreaModel.dblSortValues(lngIndex) = udtServiceAreaModel.udtSolutions(lngIndex).dblOptimizeOn
  Next
  sortValues udtServiceAreaModel.dblSortValues, udtServiceAreaModel.lngCommunitySize, udtServiceAreaModel.lngSolutionSortIndexes
  udtServiceAreaModel.booSolutionSortIsValid = True

End Sub


Private Sub computeSolutionP(udtServiceAreaModel As SERVICEAREAMODEL_TYPE)

  Dim lngSolutionIndex As Long, dblTravelValue As Double, dblSum As Double
  
  If Not udtServiceAreaModel.booSolutionSortIsValid Then sortSolutions udtServiceAreaModel
  dblTravelValue = 4 * udtServiceAreaModel.udtSolutions(udtServiceAreaModel.lngSolutionSortIndexes(0)).dblTravelMiles
  For lngSolutionIndex = 0 To udtServiceAreaModel.lngCommunitySize - 1
    If dblTravelValue + udtServiceAreaModel.udtSolutions(lngSolutionIndex).dblTravelMiles = 0 Then
      udtServiceAreaModel.dblSolutionSelectionP(lngSolutionIndex) = 1
    Else
    udtServiceAreaModel.dblSolutionSelectionP(lngSolutionIndex) = _
        (1# / (dblTravelValue + udtServiceAreaModel.udtSolutions(lngSolutionIndex).dblTravelMiles)) ^ udtServiceAreaModel.dblSolutionSelectionExponent
    End If
    dblSum = dblSum + udtServiceAreaModel.dblSolutionSelectionP(lngSolutionIndex)
  Next
  
  udtServiceAreaModel.dblSolutionSelectionP(0) _
      = udtServiceAreaModel.dblSolutionSelectionP(0) / dblSum
  For lngSolutionIndex = 1 To udtServiceAreaModel.lngCommunitySize - 1
    udtServiceAreaModel.dblSolutionSelectionP(lngSolutionIndex) _
        = udtServiceAreaModel.dblSolutionSelectionP(lngSolutionIndex - 1) _
            + udtServiceAreaModel.dblSolutionSelectionP(lngSolutionIndex) / dblSum
  Next

End Sub


Private Function selectSolution( _
    udtServiceAreaModel As SERVICEAREAMODEL_TYPE, _
    strSelectionMethod As String) As Long
    
  Dim lngIndex As Long, lngSolutionIndex As Long
    
  If Left$(strSelectionMethod, 3) = "Top" Then
    If Not udtServiceAreaModel.booSolutionSortIsValid Then sortSolutions udtServiceAreaModel
    lngIndex = Fix(CLng(Mid$(strSelectionMethod, 4)) * Rnd())
    selectSolution = udtServiceAreaModel.lngSolutionSortIndexes(lngIndex)
  
  ElseIf strSelectionMethod = "Random" Then
    selectSolution = Fix(udtServiceAreaModel.lngCommunitySize * Rnd())
  
  ElseIf strSelectionMethod = "RandomByMileage" Then
    If Not udtServiceAreaModel.booSolutionSelectionValid Then computeSolutionP udtServiceAreaModel
    selectSolution = selectByP(udtServiceAreaModel.dblSolutionSelectionP, udtServiceAreaModel.lngCommunitySize)
  
  Else
    Err.Raise 5
    
  End If

End Function


Private Sub computeAirportData( _
    udtServiceAreaModel As SERVICEAREAMODEL_TYPE, _
    ByVal dblMaxPMTimePerTrip As Double)

  Dim lngAirportIndex As Long, lngPMIndex As Long, dblPMTime As Double
  Dim lngAirportEquipmentIndex As Long, lngEquipmentModelIndex As Long, _
      lngEquipmentCount As Long, lngCMRequirementIndex As Long
  Dim dblCMCount As Double, dblCMTime As Double, lngCMIndex As Long
  Dim lngPMTripCount As Long
  
  ReDim udtServiceAreaModel.udtAirportData(0 To udtServiceAreaModel.udtAirports.lngAirportCount - 1)
  For lngAirportIndex = 0 To udtServiceAreaModel.udtAirports.lngAirportCount - 1
    With udtServiceAreaModel.udtAirports.udtAirport(lngAirportIndex)
  
      Select Case .enmPMPeriodicity
        Case PMPER_Weekly
          udtServiceAreaModel.udtAirportData(lngAirportIndex).dblPMTripCount = 52
        Case PMPER_Monthly
          udtServiceAreaModel.udtAirportData(lngAirportIndex).dblPMTripCount = 12
        Case PMPER_Quarterly
          udtServiceAreaModel.udtAirportData(lngAirportIndex).dblPMTripCount = 4
        Case PMPER_SemiAnnually
          udtServiceAreaModel.udtAirportData(lngAirportIndex).dblPMTripCount = 2
        Case PMPER_Annually
          udtServiceAreaModel.udtAirportData(lngAirportIndex).dblPMTripCount = 1
        Case PMPER_NoPM
          udtServiceAreaModel.udtAirportData(lngAirportIndex).dblPMTripCount = 0
        Case Else
          Err.Raise 5
      End Select
      If udtServiceAreaModel.udtAirportData(lngAirportIndex).dblPMTripCount = 0 Then
        udtServiceAreaModel.udtAirportData(lngAirportIndex).dblPMTime = 0#
      Else
        dblPMTime = 0#
        For lngPMIndex = LBound(.dblPMTime) To UBound(.dblPMTime)
          dblPMTime = dblPMTime + .dblPMTime(lngPMIndex)
        Next
        lngPMTripCount = Fix((dblPMTime + dblMaxPMTimePerTrip - 1) / dblMaxPMTimePerTrip)
        If udtServiceAreaModel.udtAirportData(lngAirportIndex).dblPMTripCount < lngPMTripCount Then
          udtServiceAreaModel.udtAirportData(lngAirportIndex).dblPMTripCount = lngPMTripCount
        End If
        udtServiceAreaModel.udtAirportData(lngAirportIndex).dblPMTime = dblPMTime
      End If
      
      dblCMCount = 0#
      dblCMTime = 0#
      For lngAirportEquipmentIndex = 0 To .lngEquipmentCount - 1
        lngEquipmentModelIndex = .udtEquipment(lngAirportEquipmentIndex).lngEquipmentModelIndex
        lngEquipmentCount = .udtEquipment(lngAirportEquipmentIndex).lngCount
        lngCMRequirementIndex = .udtEquipment(lngAirportEquipmentIndex).lngCMRequirementIndex
        
        With udtServiceAreaModel.udtCMRequirements.udtCMRequirements(lngCMRequirementIndex)
          dblCMCount = dblCMCount + lngEquipmentCount * .dblFrequency
          dblCMTime = dblCMTime + lngEquipmentCount * .dblFrequency * .udtCMTime.dblAvg
        End With
      
      Next
      udtServiceAreaModel.udtAirportData(lngAirportIndex).dblCMTripCount = dblCMCount
      udtServiceAreaModel.udtAirportData(lngAirportIndex).dblCMTime = dblCMTime
        
      ' BUG: Do Depot time
      
      With udtServiceAreaModel.udtAirportData(lngAirportIndex)
        .dblTripCount = .dblPMTripCount + .dblCMTripCount
        .dblTime = .dblPMTime + .dblCMTime
      End With
      
    End With
  Next

End Sub


Private Sub assignAirportServiceAreas( _
    udtServiceAreaModel As SERVICEAREAMODEL_TYPE, _
    udtSolution As POSSIBLESOLUTION_TYPE)

  Dim lngServiceAreaIndex As Long, lngServiceAreaAirportIndex As Long
  Dim lngAirportIndex As Long, dblAirportDistances() As Double, dblDistance As Double
  
  lngServiceAreaAirportIndex = udtSolution.udtServiceAreaData(0).lngAirportIndex
  For lngAirportIndex = 0 To udtServiceAreaModel.udtAirports.lngAirportCount - 1
    udtSolution.udtAirportData(lngAirportIndex).lngServiceAreaIndex = 0
    udtSolution.udtAirportData(lngAirportIndex).dblTravelDistance = _
        udtServiceAreaModel.dblAirportDistances(lngAirportIndex, lngServiceAreaAirportIndex)
  Next
      
  For lngServiceAreaIndex = 1 To udtSolution.lngServiceAreaCount - 1
    lngServiceAreaAirportIndex = udtSolution.udtServiceAreaData(lngServiceAreaIndex).lngAirportIndex
    For lngAirportIndex = 0 To udtServiceAreaModel.udtAirports.lngAirportCount - 1
      dblDistance = udtServiceAreaModel.dblAirportDistances(lngAirportIndex, lngServiceAreaAirportIndex)
      If dblDistance < udtSolution.udtAirportData(lngAirportIndex).dblTravelDistance Then
        udtSolution.udtAirportData(lngAirportIndex).dblTravelDistance = dblDistance
        udtSolution.udtAirportData(lngAirportIndex).lngServiceAreaIndex = lngServiceAreaIndex
      End If
    Next
    
  Next

End Sub


Private Sub evaluateSolution( _
   udtServiceAreaModel As SERVICEAREAMODEL_TYPE, _
   udtPossibleSolution As POSSIBLESOLUTION_TYPE, _
   ByVal dblMaxTravelDistance As Double)

  Dim lngServiceAreaIndex As Long, dblTravelDistance As Double, lngFinalServiceAreaIndex As Long, booKeepItem As Boolean
  Dim lngAirportIndex As Long, lngFSTCount As Long

  assignAirportServiceAreas udtServiceAreaModel, udtPossibleSolution

  For lngServiceAreaIndex = 0 To udtPossibleSolution.lngServiceAreaCount - 1
    With udtPossibleSolution.udtServiceAreaData(lngServiceAreaIndex)
      .dblPMTripCount = 0#
      .dblPMTime = 0#
      .dblCMTripCount = 0#
      .dblCMTime = 0#
      .dblTravelMiles = 0#
    End With
  Next
  For lngAirportIndex = 0 To udtServiceAreaModel.udtAirports.lngAirportCount - 1
    dblTravelDistance = udtPossibleSolution.udtAirportData(lngAirportIndex).dblTravelDistance
    lngServiceAreaIndex = udtPossibleSolution.udtAirportData(lngAirportIndex).lngServiceAreaIndex
    With udtPossibleSolution.udtServiceAreaData(lngServiceAreaIndex)
      If 0 < dblTravelDistance Then
        .dblPMTripCount = .dblPMTripCount + udtServiceAreaModel.udtAirportData(lngAirportIndex).dblPMTripCount
        .dblCMTripCount = .dblCMTripCount + udtServiceAreaModel.udtAirportData(lngAirportIndex).dblCMTripCount
        If dblMaxTravelDistance < dblTravelDistance Then
          dblTravelDistance = dblMaxTravelDistance
        End If
        .dblTravelMiles = .dblTravelMiles + (udtServiceAreaModel.udtAirportData(lngAirportIndex).dblPMTripCount _
            + udtServiceAreaModel.udtAirportData(lngAirportIndex).dblCMTripCount) * 2 * dblTravelDistance
      End If
      .dblPMTime = .dblPMTime + udtServiceAreaModel.udtAirportData(lngAirportIndex).dblPMTime
      .dblCMTime = .dblCMTime + udtServiceAreaModel.udtAirportData(lngAirportIndex).dblCMTime
    End With
  Next
  
  udtPossibleSolution.dblPMTripCount = 0#
  udtPossibleSolution.dblPMTime = 0#
  udtPossibleSolution.dblCMTripCount = 0#
  udtPossibleSolution.dblCMTime = 0#
  udtPossibleSolution.dblTravelMiles = 0#
  udtPossibleSolution.dblFSTHours = 0#
  udtPossibleSolution.lngFSTCount = 0
'  lngServiceAreaIndex = 0
'  Do While lngServiceAreaIndex < udtPossibleSolution.lngServiceAreaCount
  
'    With udtPossibleSolution.udtServiceAreaData(lngServiceAreaIndex)
'      .dblFSTTime = .dblPMTime + .dblCMTime + .dblTravelMiles / 60
'      .lngFSTCount = Fix(computePiecewiseLinear(.dblFSTTime, udtServiceAreaModel.udtFSTTimeToCount))
'      If .dblFSTTime = 0 Then
'        booKeepItem = False
'        lngServiceAreaIndex = lngServiceAreaIndex + 1
'      Else
'        booKeepItem = True
'        If .lngFSTCount = 0 Then .lngFSTCount = 1
'        udtPossibleSolution.dblPMTripCount = udtPossibleSolution.dblPMTripCount + .dblPMTripCount
'        udtPossibleSolution.dblPMTime = udtPossibleSolution.dblPMTime + .dblPMTime
'        udtPossibleSolution.dblCMTripCount = udtPossibleSolution.dblCMTripCount + .dblCMTripCount
'        udtPossibleSolution.dblCMTime = udtPossibleSolution.dblCMTime + .dblCMTime
'        udtPossibleSolution.dblTravelMiles = udtPossibleSolution.dblTravelMiles + .dblTravelMiles
'        udtPossibleSolution.dblFSTHours = udtPossibleSolution.dblFSTHours + .dblFSTTime
'        udtPossibleSolution.lngFSTCount = udtPossibleSolution.lngFSTCount + .lngFSTCount
'      End If
'    End With
    
'    If booKeepItem Then
'      If lngFinalServiceAreaIndex < lngServiceAreaIndex Then
'        udtPossibleSolution.udtServiceAreaData(lngFinalServiceAreaIndex) = udtPossibleSolution.udtServiceAreaData(lngServiceAreaIndex)
'      End If
'      lngFinalServiceAreaIndex = lngFinalServiceAreaIndex + 1
'    End If
'    lngServiceAreaIndex = lngServiceAreaIndex + 1
    
'  Loop
'  udtPossibleSolution.lngServiceAreaCount = lngFinalServiceAreaIndex
      
  For lngServiceAreaIndex = 0 To udtPossibleSolution.lngServiceAreaCount - 1
    With udtPossibleSolution.udtServiceAreaData(lngServiceAreaIndex)
    
      .dblFSTTime = .dblPMTime + .dblCMTime + .dblTravelMiles / 60
      .lngFSTCount = Fix(computePiecewiseLinear(.dblFSTTime, udtServiceAreaModel.udtFSTTimeToCount))
      If (0 < .dblFSTTime) And (.lngFSTCount = 0) Then .lngFSTCount = 1
    
      udtPossibleSolution.dblPMTripCount = udtPossibleSolution.dblPMTripCount + .dblPMTripCount
      udtPossibleSolution.dblPMTime = udtPossibleSolution.dblPMTime + .dblPMTime
      udtPossibleSolution.dblCMTripCount = udtPossibleSolution.dblCMTripCount + .dblCMTripCount
      udtPossibleSolution.dblCMTime = udtPossibleSolution.dblCMTime + .dblCMTime
      udtPossibleSolution.dblTravelMiles = udtPossibleSolution.dblTravelMiles + .dblTravelMiles
      udtPossibleSolution.dblFSTHours = udtPossibleSolution.dblFSTHours + .dblFSTTime
      udtPossibleSolution.lngFSTCount = udtPossibleSolution.lngFSTCount + .lngFSTCount
      
    End With
  Next
  
  udtPossibleSolution.dblFSTCost = udtPossibleSolution.dblTravelMiles * udtServiceAreaModel.dblCostPerMile _
      + udtPossibleSolution.lngFSTCount * udtServiceAreaModel.dblFSTHoursPerYear * udtServiceAreaModel.dblCostPerFSTHour
  udtPossibleSolution.dblOptimizeOn = udtPossibleSolution.dblFSTCost

End Sub


Private Sub removeNullServiceAreas(udtSolution As POSSIBLESOLUTION_TYPE)

  Dim lngServiceAreaIndex As Long, lngNewServiceAreaIndex As Long

  Do While lngServiceAreaIndex < udtSolution.lngServiceAreaCount
  
    If 0 < udtSolution.udtServiceAreaData(lngServiceAreaIndex).dblFSTTime Then
      If lngNewServiceAreaIndex < lngServiceAreaIndex Then
        udtSolution.udtServiceAreaData(lngNewServiceAreaIndex) = udtSolution.udtServiceAreaData(lngServiceAreaIndex)
        lngNewServiceAreaIndex = lngNewServiceAreaIndex + 1
      End If
    End If
    lngServiceAreaIndex = lngServiceAreaIndex + 1
    
  Loop
  udtSolution.lngServiceAreaCount = lngNewServiceAreaIndex

End Sub


' Mating retains any service areas that are in both solutions and gives a 50% chance of selecting service areas that are not shared
Private Sub mateSolutions( _
    udtServiceAreaModel As SERVICEAREAMODEL_TYPE, _
    udtSolution1 As POSSIBLESOLUTION_TYPE, udtSolution2 As POSSIBLESOLUTION_TYPE, _
    ByRef udtResult As POSSIBLESOLUTION_TYPE)

  Dim lngIndex1 As Long, lngIndex2 As Long
  Dim lngServiceAreaIndex As Long
  
  udtResult.lngServiceAreaCount = 0
  If UBound(udtResult.udtServiceAreaData) < udtSolution1.lngServiceAreaCount + udtSolution2.lngServiceAreaCount - 1 Then
    ReDim udtResult.udtServiceAreaData(0 To udtSolution1.lngServiceAreaCount + udtSolution2.lngServiceAreaCount - 1)
  End If
  
  Do While True

    If udtResult.lngServiceAreaCount = udtServiceAreaModel.lngServiceAreaCount_Max Then Exit Do
    If lngIndex1 = udtSolution1.lngServiceAreaCount Then Exit Do
    If lngIndex2 = udtSolution2.lngServiceAreaCount Then Exit Do

    If udtSolution1.udtServiceAreaData(lngIndex1).lngAirportIndex = udtSolution2.udtServiceAreaData(lngIndex2).lngAirportIndex Then
      udtResult.udtServiceAreaData(udtResult.lngServiceAreaCount) = udtSolution1.udtServiceAreaData(lngIndex1)
      udtResult.lngServiceAreaCount = udtResult.lngServiceAreaCount + 1
      lngIndex1 = lngIndex1 + 1
      lngIndex2 = lngIndex2 + 1
     
    ElseIf udtSolution1.udtServiceAreaData(lngIndex1).lngAirportIndex < udtSolution2.udtServiceAreaData(lngIndex2).lngAirportIndex Then
      If Rnd() < 0.5 Then
        udtResult.udtServiceAreaData(udtResult.lngServiceAreaCount) = udtSolution1.udtServiceAreaData(lngIndex1)
        udtResult.lngServiceAreaCount = udtResult.lngServiceAreaCount + 1
      End If
      lngIndex1 = lngIndex1 + 1
      
    Else
      If Rnd() < 0.5 Then
        udtResult.udtServiceAreaData(udtResult.lngServiceAreaCount) = udtSolution2.udtServiceAreaData(lngIndex2)
        udtResult.lngServiceAreaCount = udtResult.lngServiceAreaCount + 1
      End If
      lngIndex2 = lngIndex2 + 1
      
    End If
    
  Loop
  
  For lngIndex1 = lngIndex1 To udtSolution1.lngServiceAreaCount - 1
    If udtResult.lngServiceAreaCount = udtServiceAreaModel.lngServiceAreaCount_Max Then Exit For
    If Rnd() < 0.5 Then
      udtResult.udtServiceAreaData(udtResult.lngServiceAreaCount) = udtSolution1.udtServiceAreaData(lngIndex1)
      udtResult.lngServiceAreaCount = udtResult.lngServiceAreaCount + 1
    End If
  Next
  For lngIndex2 = lngIndex2 To udtSolution2.lngServiceAreaCount - 1
    If udtResult.lngServiceAreaCount = udtServiceAreaModel.lngServiceAreaCount_Max Then Exit For
    If Rnd() < 0.5 Then
      udtResult.udtServiceAreaData(udtResult.lngServiceAreaCount) = udtSolution2.udtServiceAreaData(lngIndex2)
      udtResult.lngServiceAreaCount = udtResult.lngServiceAreaCount + 1
    End If
  Next
  
  udtResult.booResultsAreValid = False

End Sub


Private Sub setSolutionServiceAreaCount( _
    udtServiceAreaModel As SERVICEAREAMODEL_TYPE, _
    udtSolution As POSSIBLESOLUTION_TYPE, _
    ByVal lngServiceAreaCount As Long)
    
  Dim lngServiceAreaIndex As Long
  
  Do While udtSolution.lngServiceAreaCount < lngServiceAreaCount
    insertSolutionServiceArea udtServiceAreaModel, udtSolution, Fix(udtServiceAreaModel.udtAirports.lngAirportCount * Rnd())
  Loop
  Do While lngServiceAreaCount < udtSolution.lngServiceAreaCount
    removeSolutionServiceArea udtSolution, Fix(udtSolution.lngServiceAreaCount * Rnd())
  Loop

End Sub


Private Function insertSolutionServiceArea( _
    udtServiceAreaModel As SERVICEAREAMODEL_TYPE, _
    udtSolution As POSSIBLESOLUTION_TYPE, _
    ByVal lngAirportIndex As Long) As Boolean
    
  Dim lngListIndex As Long
  
  If Not findSolutionServiceArea(udtSolution, lngAirportIndex, lngListIndex) Then
    For lngListIndex = udtSolution.lngServiceAreaCount To lngListIndex + 1 Step -1
      udtSolution.udtServiceAreaData(lngListIndex) = udtSolution.udtServiceAreaData(lngListIndex - 1)
    Next
    With udtSolution.udtServiceAreaData(lngListIndex)
      .lngAirportIndex = lngAirportIndex
      .dblLongitude = udtServiceAreaModel.udtAirports.udtAirport(lngAirportIndex).dblLongitude
      .dblLatitude = udtServiceAreaModel.udtAirports.udtAirport(lngAirportIndex).dblLatitude
    End With
    udtSolution.lngServiceAreaCount = udtSolution.lngServiceAreaCount + 1
    udtSolution.booResultsAreValid = False
    insertSolutionServiceArea = True
  End If

End Function


Private Sub removeSolutionServiceArea( _
    udtSolution As POSSIBLESOLUTION_TYPE, _
    ByVal lngListIndex As Long)
    
  For lngListIndex = lngListIndex + 1 To udtSolution.lngServiceAreaCount - 1
    udtSolution.udtServiceAreaData(lngListIndex - 1) = udtSolution.udtServiceAreaData(lngListIndex)
  Next
  udtSolution.booResultsAreValid = False
  udtSolution.lngServiceAreaCount = udtSolution.lngServiceAreaCount - 1

End Sub



Public Sub randomSolutionServiceArea( _
    udtServiceAreaModel As SERVICEAREAMODEL_TYPE, _
    udtSolution As POSSIBLESOLUTION_TYPE, _
    ByVal lngServiceAreaCount As Long)

  udtSolution.lngServiceAreaCount = 0
  If UBound(udtSolution.udtServiceAreaData) < lngServiceAreaCount - 1 Then
    ReDim udtSolution.udtServiceAreaData(0 To lngServiceAreaCount)
  End If
  Do While udtSolution.lngServiceAreaCount < lngServiceAreaCount
    insertSolutionServiceArea udtServiceAreaModel, udtSolution, selectAirport(udtServiceAreaModel)
  Loop
  udtSolution.booResultsAreValid = False
  udtSolution.booServiceAreasAreSorted = False

End Sub



Private Function findSolutionServiceArea( _
    udtSolution As POSSIBLESOLUTION_TYPE, _
    ByVal lngAirportIndex As Long, _
    ByRef lngListIndex As Long) As Boolean

  Dim lngIndexLow As Long, lngIndexMid As Long, lngIndexHigh As Long
  
  If udtSolution.lngServiceAreaCount = 0 Then
    lngListIndex = 0
    Exit Function
  End If
  
  lngIndexHigh = udtSolution.lngServiceAreaCount - 1
  Do While lngIndexLow + 1 < lngIndexHigh
    lngIndexMid = (lngIndexLow + lngIndexHigh) \ 2
    If udtSolution.udtServiceAreaData(lngIndexMid).lngAirportIndex < lngAirportIndex Then
      lngIndexLow = lngIndexMid
    Else
      lngIndexHigh = lngIndexMid
    End If
  Loop
  
  If lngAirportIndex <= udtSolution.udtServiceAreaData(lngIndexLow).lngAirportIndex Then
    lngListIndex = lngIndexLow
    findSolutionServiceArea = (lngListIndex = udtSolution.udtServiceAreaData(lngIndexLow).lngAirportIndex)
  ElseIf lngAirportIndex <= udtSolution.udtServiceAreaData(lngIndexHigh).lngAirportIndex Then
    lngListIndex = lngIndexHigh
    findSolutionServiceArea = (lngListIndex = udtSolution.udtServiceAreaData(lngIndexHigh).lngAirportIndex)
  Else
    lngListIndex = lngIndexHigh + 1
  End If

End Function


Private Sub assignAirportSelectionP( _
    udtServiceAreaModel As SERVICEAREAMODEL_TYPE, _
    ByVal dblExponent As Double)

  Dim lngAirportIndex As Long, dblWeightTotal As Double
  
  udtServiceAreaModel.dblAirportSelectionExponent = dblExponent
  For lngAirportIndex = 0 To udtServiceAreaModel.udtAirports.lngAirportCount - 1
    udtServiceAreaModel.dblAirportSelectionP(lngAirportIndex) _
        = udtServiceAreaModel.udtAirportData(lngAirportIndex).dblTripCount ^ dblExponent
    If Left$(udtServiceAreaModel.udtAirports.udtAirport(lngAirportIndex).strCat, 3) <> "Cat" Then
      udtServiceAreaModel.dblAirportSelectionP(lngAirportIndex) = udtServiceAreaModel.dblAirportSelectionP(lngAirportIndex) / 1000
    End If
    dblWeightTotal = dblWeightTotal + udtServiceAreaModel.dblAirportSelectionP(lngAirportIndex)
  Next
  
  udtServiceAreaModel.dblAirportSelectionP(0) = udtServiceAreaModel.dblAirportSelectionP(0) / dblWeightTotal
  For lngAirportIndex = 1 To udtServiceAreaModel.udtAirports.lngAirportCount - 1
    udtServiceAreaModel.dblAirportSelectionP(lngAirportIndex) _
        = udtServiceAreaModel.dblAirportSelectionP(lngAirportIndex - 1) _
          + udtServiceAreaModel.dblAirportSelectionP(lngAirportIndex) / dblWeightTotal
  Next

End Sub


Private Function selectAirport(udtServiceAreaModel As SERVICEAREAMODEL_TYPE) As Long

  Dim dblP As Double, lngIndexLow As Long, lngIndexMed As Long, lngIndexHigh As Long

  If udtServiceAreaModel.dblAirportSelectionExponent = 0# Then
    selectAirport = Fix(udtServiceAreaModel.udtAirports.lngAirportCount * Rnd())
    Exit Function
  End If

  selectAirport = selectByP(udtServiceAreaModel.dblAirportSelectionP, udtServiceAreaModel.udtAirports.lngAirportCount)

End Function


Public Function selectByP(dblPValues() As Double, ByVal lngPCount As Long) As Long

  Dim dblP As Double, lngIndexLow As Long, lngIndexMed As Long, lngIndexHigh As Long
  
  dblP = Rnd()
  lngIndexHigh = lngPCount - 1
  Do While lngIndexLow + 1 < lngIndexHigh
    lngIndexMed = (lngIndexLow + lngIndexHigh) \ 2
    If dblPValues(lngIndexMed) < dblP Then
      lngIndexLow = lngIndexMed
    Else
      lngIndexHigh = lngIndexMed
    End If
  Loop
  
  If dblP < dblPValues(lngIndexLow) Then
    selectByP = lngIndexLow
  ElseIf dblP < dblPValues(lngIndexHigh) Then
    selectByP = lngIndexHigh
  Else
    Err.Raise 5
  End If

End Function


Public Sub sortValues( _
    dblValues() As Double, lngValueCount As Long, _
    ByRef lngSortedIndexes() As Long)

  Dim lngIndex As Long
  
  For lngIndex = 0 To lngValueCount - 1
    lngSortedIndexes(lngIndex) = lngIndex
  Next
  sortValues_Helper1 dblValues, lngSortedIndexes, 0, lngValueCount - 1

End Sub

Private Sub sortValues_Helper1( _
    dblValues() As Double, ByRef lngSortedIndexes() As Long, _
    lngIndexLow As Long, lngIndexHigh As Long)
    
  Dim dblSplitValue As Double
  Dim lngIndexLow_Top As Long, lngIndexHigh_Bottom As Long, lngIndex As Long
  
  If lngIndexHigh - lngIndexLow < 6 Then
    sortValues_Helper2 dblValues, lngSortedIndexes, lngIndexLow, lngIndexHigh
    Exit Sub
  End If
  
  dblSplitValue = (dblValues(lngSortedIndexes(lngIndexLow)) _
      + dblValues(lngSortedIndexes((lngIndexLow + lngIndexHigh) \ 2)) _
      + dblValues(lngSortedIndexes(lngIndexHigh))) / 3
  lngIndexLow_Top = lngIndexLow
  lngIndexHigh_Bottom = lngIndexHigh
  Do While lngIndexLow_Top < lngIndexHigh_Bottom
    If dblValues(lngSortedIndexes(lngIndexLow_Top)) <= dblSplitValue Then
      lngIndexLow_Top = lngIndexLow_Top + 1
    ElseIf dblValues(lngSortedIndexes(lngIndexHigh_Bottom)) <= dblSplitValue Then
      lngIndex = lngSortedIndexes(lngIndexLow_Top)
      lngSortedIndexes(lngIndexLow_Top) = lngSortedIndexes(lngIndexHigh_Bottom)
      lngSortedIndexes(lngIndexHigh_Bottom) = lngIndex
      lngIndexLow_Top = lngIndexLow_Top + 1
      lngIndexHigh_Bottom = lngIndexHigh_Bottom - 1
    Else
      lngIndex = lngSortedIndexes(lngIndexLow_Top)
      lngSortedIndexes(lngIndexLow_Top) = lngSortedIndexes(lngIndexHigh_Bottom - 1)
      lngSortedIndexes(lngIndexHigh_Bottom - 1) = lngIndex
      lngIndexHigh_Bottom = lngIndexHigh_Bottom - 2
    End If
   Loop
   If lngIndexLow_Top = lngIndexHigh_Bottom Then
     If dblValues(lngSortedIndexes(lngIndexLow_Top)) <= dblSplitValue Then
       lngIndexLow_Top = lngIndexLow_Top + 1
     Else
       lngIndexHigh_Bottom = lngIndexHigh_Bottom - 1
     End If
   End If
   lngIndexLow_Top = lngIndexLow_Top - 1
   lngIndexHigh_Bottom = lngIndexHigh_Bottom + 1
      
   If lngIndexLow_Top = lngIndexHigh Then
     sortValues_Helper2 dblValues, lngSortedIndexes, lngIndexLow, lngIndexHigh
   ElseIf lngIndexHigh_Bottom = lngIndexLow Then
     sortValues_Helper2 dblValues, lngSortedIndexes, lngIndexLow, lngIndexHigh
   Else
     sortValues_Helper1 dblValues, lngSortedIndexes, lngIndexLow, lngIndexLow_Top
     sortValues_Helper1 dblValues, lngSortedIndexes, lngIndexHigh_Bottom, lngIndexHigh
   End If
    
End Sub


Private Sub sortValues_Helper2( _
    dblValues() As Double, ByRef lngSortedIndexes() As Long, _
    lngIndexLow As Long, lngIndexHigh As Long)

  Dim lngIndex1 As Long, lngIndex2 As Long, lngIndex As Double
  
  For lngIndex1 = lngIndexLow + 1 To lngIndexHigh
    For lngIndex2 = lngIndexHigh To lngIndex1 Step -1
      If dblValues(lngSortedIndexes(lngIndex2)) < dblValues(lngSortedIndexes(lngIndex2 - 1)) Then
        lngIndex = lngSortedIndexes(lngIndex2)
        lngSortedIndexes(lngIndex2) = lngSortedIndexes(lngIndex2 - 1)
        lngSortedIndexes(lngIndex2 - 1) = lngIndex
      End If
    Next
  Next

End Sub


Private Function validateSolution(udtServiceAreaModel As SERVICEAREAMODEL_TYPE, _
    udtSolution As POSSIBLESOLUTION_TYPE) As Boolean

  Dim lngAirportIndex As Long, lngServiceAreaIndex As Long, lngServiceAreaAirportIndex As Long
  
  For lngAirportIndex = 0 To udtServiceAreaModel.udtAirports.lngAirportCount - 1
    lngServiceAreaIndex = udtSolution.udtAirportData(lngAirportIndex).lngServiceAreaIndex
    lngServiceAreaAirportIndex = udtSolution.udtServiceAreaData(lngServiceAreaIndex).lngAirportIndex
    If udtSolution.udtAirportData(lngAirportIndex).dblTravelDistance <> udtServiceAreaModel.dblAirportDistances(lngAirportIndex, lngServiceAreaAirportIndex) Then Exit Function
  Next
  validateSolution = True

End Function


Private Function computePiecewiseLinear(ByVal dblX As Double, udtPiecewiseLinear As PIECEWISELINEAR_TYPE) As Double

  Dim lngIndex As Long

  For lngIndex = 0 To udtPiecewiseLinear.lngItemCount - 1
    If (udtPiecewiseLinear.udtItems(lngIndex).dblFromValue <= dblX) And (dblX < udtPiecewiseLinear.udtItems(lngIndex).dblToValue) Then
      computePiecewiseLinear = udtPiecewiseLinear.udtItems(lngIndex).dblB + udtPiecewiseLinear.udtItems(lngIndex).dblM * dblX
      Exit Function
    End If
  Next
  
  Err.Raise 5

End Function



Private Sub setFSTTimeToCountModel(ByRef udtFSTTimeToCount As PIECEWISELINEAR_TYPE, strModelName As String)

  Dim objWorksheet As Excel.Worksheet, lngRowNumber As Long
  
  Set objWorksheet = Application.ActiveWorkbook.Worksheets.Item("FSTUtil")
  lngRowNumber = objWorksheet.Range("B2").Value + 2
  
  udtFSTTimeToCount.lngItemCount = 0
  ReDim udtFSTTimeToCount.udtItems(0 To 7)
  Do While objWorksheet.Range("B" & lngRowNumber).Value <> ""
  
    If objWorksheet.Range("B" & lngRowNumber).Value = strModelName Then
      
      If UBound(udtFSTTimeToCount.udtItems) < udtFSTTimeToCount.lngItemCount Then
        ReDim Preserve udtFSTTimeToCount.udtItems(0 To 15 + udtFSTTimeToCount.lngItemCount)
      End If
      
      With udtFSTTimeToCount.udtItems(udtFSTTimeToCount.lngItemCount)
        .dblFromValue = objWorksheet.Range("E" & lngRowNumber).Value
        .dblToValue = objWorksheet.Range("F" & lngRowNumber).Value
        .dblB = objWorksheet.Range("G" & lngRowNumber).Value
        .dblM = objWorksheet.Range("H" & lngRowNumber).Value
      End With
      udtFSTTimeToCount.lngItemCount = udtFSTTimeToCount.lngItemCount + 1
      
    End If
    
    lngRowNumber = lngRowNumber + 1
    
  Loop
  Set objWorksheet = Nothing
    
End Sub



Private Function readOptimizationParameters(ByRef udtServiceAreaModel As SERVICEAREAMODEL_TYPE, ByVal objWorksheet As Excel.Worksheet)

  Dim lngRowNumber As Long, lngIndex As Long

  If objWorksheet.Range("A1").Value <> "Worksheet Type:" Then Exit Function
  If objWorksheet.Range("B1").Value <> "Optimization Parameters" Then Exit Function

  udtServiceAreaModel.lngCommunitySize = objWorksheet.Range("B10").Value
  udtServiceAreaModel.lngIterationCount = objWorksheet.Range("B11").Value
  udtServiceAreaModel.dblAirportSelectionExponent = objWorksheet.Range("B12").Value
  udtServiceAreaModel.dblSolutionSelectionExponent = objWorksheet.Range("B13").Value
  
  lngRowNumber = 17
  udtServiceAreaModel.lngEvolutionCount = 0
  ReDim udtServiceAreaModel.udtEvolutions(0 To 7)
  Do While objWorksheet.Range("A" & lngRowNumber).Value <> ""
  
    If UBound(udtServiceAreaModel.udtEvolutions) < udtServiceAreaModel.lngEvolutionCount Then
      ReDim Preserve udtServiceAreaModel.udtEvolutions(0 To 15 + udtServiceAreaModel.lngEvolutionCount)
    End If
    
    With udtServiceAreaModel.udtEvolutions(udtServiceAreaModel.lngEvolutionCount)
      Select Case CStr(objWorksheet.Range("A" & lngRowNumber).Value)
        Case "Retain"
          .enmApproach = EVOLAPP_RetainTop
          .lngAppliesTo = objWorksheet.Range("B" & lngRowNumber).Value
          .lngParameterCount = 0
        Case "Mate"
          .enmApproach = EVOLAPP_Mate
          .lngAppliesTo = objWorksheet.Range("B" & lngRowNumber).Value
          .lngParameterCount = 2
          ReDim .VarParameters(0 To 1)
          .VarParameters(0) = CStr(objWorksheet.Range("D" & lngRowNumber).Value)
          lngIndex = InStr(.VarParameters(0), ";")
          If lngIndex = 0 Then Exit Function
          .VarParameters(1) = Trim$(Mid$(.VarParameters(0), lngIndex + 1))
          .VarParameters(0) = Trim$(Left$(.VarParameters(0), lngIndex - 1))
        Case "Insertion-Deletion"
          .enmApproach = EVOLAPP_InsertDelete
          .lngAppliesTo = objWorksheet.Range("B" & lngRowNumber).Value
          .lngParameterCount = 0
        Case "Random"
          .enmApproach = EVOLAPP_Random
          .lngAppliesTo = objWorksheet.Range("B" & lngRowNumber).Value
          .lngParameterCount = 0
        Case Else
          Exit Function
      End Select
    End With
    udtServiceAreaModel.lngEvolutionCount = udtServiceAreaModel.lngEvolutionCount + 1
    
    lngRowNumber = lngRowNumber + 1

  Loop
  
  readOptimizationParameters = True

End Function

