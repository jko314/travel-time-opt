Attribute VB_Name = "Macros_Mod"
Option Explicit



Public Sub ModelRun_Click()

  Dim objModelWorkbook As Excel.Workbook, objModelWorksheet As Excel.Worksheet, _
      objModelDataWorkbook As Excel.Workbook, objOutputWorkbook As Excel.Workbook, _
      objLastOutputWorksheet As Excel.Worksheet
  Dim udtAirportsModel As AIRPORTSMODEL_TYPE
  
  Dim dteStartDate As Date, dteEndDate As Date, strTripOutputFile As String, strAirportOutputFile As String
  Dim objStatusRange As Excel.Range, objErrorRange As Excel.Range, objStepRange As Excel.Range, _
      varSteps As Variant, strStepName As String, strStepParameters As String
  Dim lngStartOffset As Long
  Dim lngStepNumber As Long
  
  Set objModelWorkbook = Application.ActiveWorkbook
  Set objModelWorksheet = objModelWorkbook.Worksheets.Item("Model")
  Set objStatusRange = objModelWorksheet.Range("B7")
  Set objErrorRange = objModelWorksheet.Range("B8")
  
  objStatusRange.Value = ""
  objErrorRange.Value = ""
  
  objStatusRange.Value = "Reading model parameters ..."
  dteStartDate = objModelWorksheet.Range("B2").Value
  dteEndDate = objModelWorksheet.Range("B3").Value
  If IsEmpty(objModelWorksheet.Range("B4").Value) Then
    objErrorRange.Value = "You must specify Model Data Workbook, or THIS to use this workbook"
    GoTo Err_Abort
  End If
  If objModelWorksheet.Range("B4").Value = "THIS" Then
    Set objModelDataWorkbook = objModelWorkbook
  Else
    On Error Resume Next
    Set objModelDataWorkbook = Application.Workbooks.Open(objModelWorksheet.Range("B4").Value)
    On Error GoTo 0
    If objModelDataWorkbook Is Nothing Then
      objErrorRange.Value = "Invalid Model Data File"
      GoTo Err_Abort
    End If
  End If
  If Not IsEmpty(objModelWorksheet.Range("B5").Value) Then
    If objModelWorksheet.Range("B5").Value = "THIS" Then
      Set objOutputWorkbook = objModelWorkbook
    Else
      On Error Resume Next
      Set objOutputWorkbook = Application.Workbooks.Open(objModelWorksheet.Range("B5").Value)
      On Error GoTo 0
      If objOutputWorkbook Is Nothing Then
        Set objOutputWorkbook = Application.Workbooks.Add()
      End If
    End If
  End If
  
  Set objStepRange = objModelWorksheet.Range("A13").CurrentRegion
  varSteps = objStepRange.Value
  objModelWorksheet.Range("B13:B" & (12 + objStepRange.Rows.Count)).Clear
  
  initializeModel CDate(objModelWorksheet.Range("B2").Value), _
      CDate(objModelWorksheet.Range("B3").Value), objStatusRange, objErrorRange
  
  ' Run steps
  ' BUG: Reset model
  For lngStepNumber = 1 To UBound(varSteps, 1)
  
    strStepName = varSteps(lngStepNumber, 1)
    lngStartOffset = InStr(strStepName, "(")
    If 0 = lngStartOffset Then
      strStepParameters = ""
      lngStartOffset = 1 + Len(strStepName)
    ElseIf Right(strStepName, 1) <> ")" Then
      objStatusRange.Value = "Aborted: Invalid step name: " & strStepName
      GoTo Err_Abort
    Else
      strStepParameters = Trim$(Mid$(strStepName, lngStartOffset + 1, Len(strStepName) - lngStartOffset - 1))
      strStepName = Trim$(Left$(strStepName, lngStartOffset - 1))
    End If
    strStepName = Trim$(Left$(strStepName, lngStartOffset - 1))
      
    Select Case strStepName
    
      ' Model management
      Case "Randomize"
        If Not randomizeModel(strStepParameters) Then
          objStatusRange.Value = "Aborted: Error randomizing"
          GoTo Err_Abort
        End If
    
      ' Methods for loading and initialing data
      Case "Load airports"
        If Not loadAirports(udtAirportsModel.udtAirports, objModelWorkbook, strStepParameters) Then
          objStatusRange.Value = "Aborted: Error loading airports"
          GoTo Err_Abort
        End If
      Case "Load service areas"
        If Not loadServiceAreas(udtAirportsModel.udtServiceAreas, objModelWorkbook, strStepParameters) Then
          objStatusRange.Value = "Aborted: Error loading service areas"
          GoTo Err_Abort
        End If
      Case "Load airport service areas"
        If Not loadAirportServiceAreas(udtAirportsModel.udtAirports, udtAirportsModel.udtServiceAreas, _
            objModelWorkbook, strStepParameters) Then
          objStatusRange.Value = "Aborted: Error loading airport service areas"
          GoTo Err_Abort
        End If
      Case "Load equipment models"
        If Not loadEquipmentModels(udtAirportsModel.udtEquipmentModels, udtAirportsModel.udtEquipmentTypes, objModelWorkbook, strStepParameters) Then
          objStatusRange.Value = "Aborted: Error loading equipment models"
          GoTo Err_Abort
        End If
      Case "Load equipment"
        If Not loadEquipment(udtAirportsModel.udtAirports, udtAirportsModel.udtEquipmentModels, udtAirportsModel.udtEquipment, objModelWorkbook, strStepParameters) Then
          objStatusRange.Value = "Aborted: Error loading equipment"
          GoTo Err_Abort
        End If
      Case "Load PM requirements"
        If Not loadEquipmentPM(udtAirportsModel.udtAirports, udtAirportsModel.udtEquipmentModels, objModelWorkbook, strStepParameters) Then
          objStatusRange.Value = "Aborted: Error loading PM requirements"
          GoTo Err_Abort
        End If
      Case "Load CM requirements"
        If Not loadCMRequirements(udtAirportsModel.udtCMRequirements, objModelWorkbook, strStepParameters) Then
          objStatusRange.Value = "Aborted: Error loading CM requirements"
          GoTo Err_Abort
        End If
        If Not applyCMRequirements(udtAirportsModel.udtCMRequirements, udtAirportsModel.udtAirports, udtAirportsModel.udtEquipmentModels) Then
          objStatusRange.Value = "Aborted: Error applying CM requirements"
          GoTo Err_Abort
        End If
      Case "Load PM status"
        If Not loadPMStatus(udtAirportsModel.udtAirports, udtAirportsModel.udtEquipmentModels, udtAirportsModel.udtEquipment, objModelWorkbook, strStepParameters) Then
          objStatusRange.Value = "Aborted: Error loading PM status"
          GoTo Err_Abort
        End If
      Case "Create PM status"
        If Not createPMStatus(udtAirportsModel.udtAirports, udtAirportsModel.udtEquipmentModels, udtAirportsModel.udtEquipment, objModelWorkbook, strStepParameters) Then
          objStatusRange.Value = "Aborted: Error creating PM status"
          GoTo Err_Abort
        End If
        
        
      ' Methods for running model calculations
      Case "Create PM Items"
        If Not createPMItems(udtAirportsModel.udtEquipment, strStepParameters) Then
          objStatusRange.Value = "Aborted: Error creating PM items"
          GoTo Err_Abort
        End If
        
      Case "Compute Airport Distances"
        If Not computeAirportDistances(udtAirportsModel.udtAirports, udtAirportsModel.dblAirportDistances) Then
          objStatusRange.Value = "Aborted: Error compute airport distances"
          GoTo Err_Abort
        End If
        
      ' Methods for exporting data and results
      Case "Export airports"
        If objOutputWorkbook Is Nothing Then
          objStatusRange.Value = "Aborted: Export file not specified - cannot export airports"
          GoTo Err_Abort
        End If
        Set objLastOutputWorksheet = exportAirports(udtAirportsModel.udtAirports, objOutputWorkbook, objLastOutputWorksheet, strStepParameters)
        If objLastOutputWorksheet Is Nothing Then GoTo Err_Abort
      Case "Export service areas"
        If objOutputWorkbook Is Nothing Then
          objStatusRange.Value = "Aborted: Export file not specified - cannot export service areas"
          GoTo Err_Abort
        End If
        Set objLastOutputWorksheet = exportServiceAreas(udtAirportsModel.udtServiceAreas, objOutputWorkbook, objLastOutputWorksheet, strStepParameters)
        If objLastOutputWorksheet Is Nothing Then GoTo Err_Abort
      Case "Export airport service areas"
        If objOutputWorkbook Is Nothing Then
          objStatusRange.Value = "Aborted: Export file not specified - cannot export airport service areas"
          GoTo Err_Abort
        End If
        Set objLastOutputWorksheet = exportAirportServiceAreas(udtAirportsModel.udtAirports, udtAirportsModel.udtServiceAreas, _
            objOutputWorkbook, objLastOutputWorksheet, strStepParameters)
        If objLastOutputWorksheet Is Nothing Then GoTo Err_Abort
      Case "Export equipment models"
        If objOutputWorkbook Is Nothing Then
          objStatusRange.Value = "Aborted: Export file not specified - cannot export equipment models"
          GoTo Err_Abort
        End If
        Set objLastOutputWorksheet = exportEquipmentModels(udtAirportsModel.udtEquipmentModels, udtAirportsModel.udtEquipmentTypes, objOutputWorkbook, objLastOutputWorksheet, strStepParameters)
        If objLastOutputWorksheet Is Nothing Then GoTo Err_Abort
      Case "Export equipment"
        If objOutputWorkbook Is Nothing Then
          objStatusRange.Value = "Aborted: Export file not specified - cannot export equipment"
          GoTo Err_Abort
        End If
        Set objLastOutputWorksheet = exportEquipment(udtAirportsModel.udtAirports, udtAirportsModel.udtEquipmentModels, udtAirportsModel.udtEquipment, objOutputWorkbook, objLastOutputWorksheet, strStepParameters)
        If objLastOutputWorksheet Is Nothing Then GoTo Err_Abort
      Case "Export PM requirements"
        If objOutputWorkbook Is Nothing Then
          objStatusRange.Value = "Aborted: Export file not specified - cannot export PM requirements"
          GoTo Err_Abort
        End If
        Set objLastOutputWorksheet = exportEquipmentPM(udtAirportsModel.udtEquipmentModels, objOutputWorkbook, objLastOutputWorksheet, strStepParameters)
        If objLastOutputWorksheet Is Nothing Then GoTo Err_Abort
      Case "Export CM requirements"
        If objOutputWorkbook Is Nothing Then
          objStatusRange.Value = "Aborted: Export file not specified - cannot export CM requirements"
          GoTo Err_Abort
        End If
        Set objLastOutputWorksheet = exportEquipmentCM(udtAirportsModel.udtCMRequirements, objOutputWorkbook, objLastOutputWorksheet, strStepParameters)
        If objLastOutputWorksheet Is Nothing Then GoTo Err_Abort
      Case "Export PM status"
        If objOutputWorkbook Is Nothing Then
          objStatusRange.Value = "Aborted: Export file not specified - cannot export PM status"
          GoTo Err_Abort
        End If
        Set objLastOutputWorksheet = exportPMStatus(udtAirportsModel.udtAirports, udtAirportsModel.udtEquipmentModels, udtAirportsModel.udtEquipment, objOutputWorkbook, objLastOutputWorksheet, strStepParameters)
        If objLastOutputWorksheet Is Nothing Then GoTo Err_Abort
      Case "Export daily PM times"
        If objOutputWorkbook Is Nothing Then
          objStatusRange.Value = "Aborted: Export file not specified - cannot export daily PM times"
          GoTo Err_Abort
        End If
        Set objLastOutputWorksheet = exportDailyPMTimes(udtAirportsModel.udtEquipmentModels, udtAirportsModel.udtEquipment, objOutputWorkbook, objLastOutputWorksheet, strStepParameters)
        If objLastOutputWorksheet Is Nothing Then GoTo Err_Abort
      Case "Export PM schedule"
        If objOutputWorkbook Is Nothing Then
          objStatusRange.Value = "Aborted: Export file not specified - cannot export PM schedule"
          GoTo Err_Abort
        End If
        Set objLastOutputWorksheet = exportPMSchedule(udtAirportsModel.udtAirports, udtAirportsModel.udtEquipmentModels, udtAirportsModel.udtEquipment, objOutputWorkbook, objLastOutputWorksheet, strStepParameters)
        If objLastOutputWorksheet Is Nothing Then GoTo Err_Abort
      
      Case Else
        objStatusRange.Value = "Aborted: Invalid step name"
        GoTo Err_Abort
        
    End Select
    objModelWorksheet.Range("B" & (12 + lngStepNumber)).Value = "Done"
  Next
        
  If Not (objOutputWorkbook Is Nothing) Then
    If objOutputWorkbook.Name = "" Then
      objOutputWorkbook.SaveAs objModelWorksheet.Range("B5").Value
    Else
      objOutputWorkbook.Save
    End If
  End If
        
  objStatusRange.Value = "Completed: " & Now()
  
Err_Abort:
  Set objStatusRange = Nothing
  Set objErrorRange = Nothing
  If objModelDataWorkbook Is objModelWorkbook Then
    Set objModelDataWorkbook = Nothing
  Else
    objModelDataWorkbook.Close
    Set objModelDataWorkbook = Nothing
  End If
  Set objLastOutputWorksheet = Nothing
  If Not (objOutputWorkbook Is Nothing) Then
    If Not (objOutputWorkbook Is objModelWorkbook) Then objOutputWorkbook.Close
    Set objOutputWorkbook = Nothing
  End If
  Set objModelWorkbook = Nothing

End Sub


Public Sub SortTest_Click()

  Dim objRange As Excel.Range, varValues As Variant, dblValues() As Double, lngValueIndexes() As Long
  Dim lngIndex As Long
  
  Set objRange = Application.ActiveSheet.Range("J7").CurrentRegion
  varValues = objRange.Value
  ReDim dblValues(0 To UBound(varValues, 1) - 1), lngValueIndexes(0 To UBound(varValues, 1) - 1)
  For lngIndex = 1 To UBound(varValues, 1)
    dblValues(lngIndex - 1) = varValues(lngIndex, 1)
  Next
  
  sortValues dblValues, lngIndex - 1, lngValueIndexes
  
  For lngIndex = 1 To UBound(varValues, 1)
    varValues(lngIndex, 1) = dblValues(lngValueIndexes(lngIndex - 1))
  Next
  objRange.Value = varValues
  
  


End Sub


Public Function colNameToNumber(strLabel As String) As Long

  If Len(strLabel) = 1 Then
    colNameToNumber = Asc(strLabel) - 64
  Else
    colNameToNumber = 26 * (Asc(Left$(strLabel, 1)) - 64) + Asc(Mid$(strLabel, 2, 1)) - 64
  End If

End Function

Public Function colNumberToName(ByVal lngColNumber As Long) As String

  If lngColNumber < 27 Then
    colNumberToName = Chr(64 + lngColNumber)
  Else
    colNumberToName = Chr(64 + (lngColNumber - 1) \ 26) & Chr(65 + (lngColNumber - 1) Mod 26)
  End If

End Function
