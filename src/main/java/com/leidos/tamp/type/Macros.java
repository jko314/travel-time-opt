package com.leidos.tamp.type;

public class Macros {
    // TODO: Option Explicit ... Warning!!! not translated


//    public final void ModelRun_Click() {
//        Excel.Workbook objModelWorkbook;
//        Excel.Worksheet objModelWorksheet;
//        Excel.Workbook objModelDataWorkbook;
//        Excel.Workbook objOutputWorkbook;
//        Excel.Worksheet objLastOutputWorksheet;
//        AIRPORTSMODEL_TYPE udtAirportsModel;
//        Date dteStartDate;
//        Date dteEndDate;
//        String  strTripOutputFile;
//        String  strAirportOutputFile;
//        Excel.Range objStatusRange;
//        Excel.Range objErrorRange;
//        Excel.Range objStepRange;
//        object varSteps;
//        ((String)(strStepName));
//        ((String)(strStepParameters));
//        long lngStartOffset;
//        long lngStepNumber;
//        objModelWorkbook = Application.ActiveWorkbook;
//        objModelWorksheet = objModelWorkbook.Worksheets.Item["Model"];
//        objStatusRange = objModelWorksheet.Range("B7");
//        objErrorRange = objModelWorksheet.Range("B8");
//        objStatusRange.Value = "";
//        objErrorRange.Value = "";
//        objStatusRange.Value = "Reading model parameters ...";
//        dteStartDate = objModelWorksheet.Range("B2").Value;
//        dteEndDate = objModelWorksheet.Range("B3").Value;
//        if (IsEmpty(objModelWorksheet.Range("B4").Value)) {
//            objErrorRange.Value = "You must specify Model Data Workbook, or THIS to use this workbook";
//            goto Err_Abort;
//        }
//
//        if ((objModelWorksheet.Range("B4").Value == "THIS")) {
//            objModelDataWorkbook = objModelWorkbook;
//        }
//        else {
//            // TODO: On Error Resume Next Warning!!!: The statement is not translatable
//            objModelDataWorkbook = Application.Workbooks.Open(objModelWorksheet.Range("B4").Value);
//            // TODO: On Error GoTo 0 Warning!!!: The statement is not translatable
//            if ((objModelDataWorkbook == null)) {
//                objErrorRange.Value = "Invalid Model Data File";
//                goto Err_Abort;
//            }
//
//        }
//
//        if (!IsEmpty(objModelWorksheet.Range("B5").Value)) {
//            if ((objModelWorksheet.Range("B5").Value == "THIS")) {
//                objOutputWorkbook = objModelWorkbook;
//            }
//            else {
//                // TODO: On Error Resume Next Warning!!!: The statement is not translatable
//                objOutputWorkbook = Application.Workbooks.Open(objModelWorksheet.Range("B5").Value);
//                // TODO: On Error GoTo 0 Warning!!!: The statement is not translatable
//                if ((objOutputWorkbook == null)) {
//                    objOutputWorkbook = Application.Workbooks.Add();
//                }
//
//            }
//
//        }
//
//        objStepRange = objModelWorksheet.Range("A13").CurrentRegion;
//        varSteps = objStepRange.Value;
//        objModelWorksheet.Range(("B13:B" + (12 + objStepRange.Rows.Count))).Clear;
//        initializeModel;
//        DateTime.Parse(objModelWorksheet.Range("B2").Value);
//        DateTime.Parse(objModelWorksheet.Range("B3").Value);
//        objStatusRange;
//        objErrorRange;
//        //  Run steps
//        //  BUG: Reset model
//        for (lngStepNumber = 1; (lngStepNumber <= UBound(varSteps, 1)); lngStepNumber++) {
//            strStepName = varSteps[lngStepNumber, 1];
//            lngStartOffset = (strStepName.IndexOf("(") + 1);
//            if ((0 == lngStartOffset)) {
//                strStepParameters = "";
//                lngStartOffset = (1 + strStepName.Length);
//            }
//            else if ((strStepName.Substring((strStepName.Length - 1)) != ")")) {
//                objStatusRange.Value = ("Aborted: Invalid step name: " + strStepName);
//                goto Err_Abort;
//            }
//            else {
//                strStepParameters = strStepName.Substring(lngStartOffset, (strStepName.Length
//                        - (lngStartOffset - 1))).Trim();
//                strStepName = strStepName.Substring(0, (lngStartOffset - 1)).Trim();
//            }
//
//            strStepName = strStepName.Substring(0, (lngStartOffset - 1)).Trim();
//            switch (strStepName) {
//                case "Randomize":
//                    if (!randomizeModel(strStepParameters)) {
//                        objStatusRange.Value = "Aborted: Error randomizing";
//                        goto Err_Abort;
//                    }
//
//                    //  Methods for loading and initialing data
//                    break;
//                case "Load airports":
//                    if (!loadAirports(udtAirportsModel.udtAirports, objModelWorkbook, strStepParameters)) {
//                        objStatusRange.Value = "Aborted: Error loading airports";
//                        goto Err_Abort;
//                    }
//
//                    break;
//                case "Load service areas":
//                    if (!loadServiceAreas(udtAirportsModel.udtServiceAreas, objModelWorkbook, strStepParameters)) {
//                        objStatusRange.Value = "Aborted: Error loading service areas";
//                        goto Err_Abort;
//                    }
//
//                    break;
//                case "Load airport service areas":
//                    if (!loadAirportServiceAreas(udtAirportsModel.udtAirports, udtAirportsModel.udtServiceAreas, objModelWorkbook, strStepParameters)) {
//                        objStatusRange.Value = "Aborted: Error loading airport service areas";
//                        goto Err_Abort;
//                    }
//
//                    break;
//                case "Load equipment models":
//                    if (!loadEquipmentModels(udtAirportsModel.udtEquipmentModels, udtAirportsModel.udtEquipmentTypes, objModelWorkbook, strStepParameters)) {
//                        objStatusRange.Value = "Aborted: Error loading equipment models";
//                        goto Err_Abort;
//                    }
//
//                    break;
//                case "Load equipment":
//                    if (!loadEquipment(udtAirportsModel.udtAirports, udtAirportsModel.udtEquipmentModels, udtAirportsModel.udtEquipment, objModelWorkbook, strStepParameters)) {
//                        objStatusRange.Value = "Aborted: Error loading equipment";
//                        goto Err_Abort;
//                    }
//
//                    break;
//                case "Load PM requirements":
//                    if (!loadEquipmentPM(udtAirportsModel.udtAirports, udtAirportsModel.udtEquipmentModels, objModelWorkbook, strStepParameters)) {
//                        objStatusRange.Value = "Aborted: Error loading PM requirements";
//                        goto Err_Abort;
//                    }
//
//                    break;
//                case "Load CM requirements":
//                    if (!loadCMRequirements(udtAirportsModel.udtCMRequirements, objModelWorkbook, strStepParameters)) {
//                        objStatusRange.Value = "Aborted: Error loading CM requirements";
//                        goto Err_Abort;
//                    }
//
//                    if (!applyCMRequirements(udtAirportsModel.udtCMRequirements, udtAirportsModel.udtAirports, udtAirportsModel.udtEquipmentModels)) {
//                        objStatusRange.Value = "Aborted: Error applying CM requirements";
//                        goto Err_Abort;
//                    }
//
//                    break;
//                case "Load PM status":
//                    if (!loadPMStatus(udtAirportsModel.udtAirports, udtAirportsModel.udtEquipmentModels, udtAirportsModel.udtEquipment, objModelWorkbook, strStepParameters)) {
//                        objStatusRange.Value = "Aborted: Error loading PM status";
//                        goto Err_Abort;
//                    }
//
//                    break;
//                case "Create PM status":
//                    if (!createPMStatus(udtAirportsModel.udtAirports, udtAirportsModel.udtEquipmentModels, udtAirportsModel.udtEquipment, objModelWorkbook, strStepParameters)) {
//                        objStatusRange.Value = "Aborted: Error creating PM status";
//                        goto Err_Abort;
//                    }
//
//                    //  Methods for running model calculations
//                    break;
//                case "Create PM Items":
//                    if (!createPMItems(udtAirportsModel.udtEquipment, strStepParameters)) {
//                        objStatusRange.Value = "Aborted: Error creating PM items";
//                        goto Err_Abort;
//                    }
//
//                    break;
//                case "Compute Airport Distances":
//                    if (!computeAirportDistances(udtAirportsModel.udtAirports, udtAirportsModel.dblAirportDistances)) {
//                        objStatusRange.Value = "Aborted: Error compute airport distances";
//                        goto Err_Abort;
//                    }
//
//                    //  Methods for exporting data and results
//                    break;
//                case "Export airports":
//                    if ((objOutputWorkbook == null)) {
//                        objStatusRange.Value = "Aborted: Export file not specified - cannot export airports";
//                        goto Err_Abort;
//                    }
//
//                    objLastOutputWorksheet = exportAirports(udtAirportsModel.udtAirports, objOutputWorkbook, objLastOutputWorksheet, strStepParameters);
//                    if ((objLastOutputWorksheet == null)) {
//                        goto Err_Abort;
//                    }
//
//                    break;
//                case "Export service areas":
//                    if ((objOutputWorkbook == null)) {
//                        objStatusRange.Value = "Aborted: Export file not specified - cannot export service areas";
//                        goto Err_Abort;
//                    }
//
//                    objLastOutputWorksheet = exportServiceAreas(udtAirportsModel.udtServiceAreas, objOutputWorkbook, objLastOutputWorksheet, strStepParameters);
//                    if ((objLastOutputWorksheet == null)) {
//                        goto Err_Abort;
//                    }
//
//                    break;
//                case "Export airport service areas":
//                    if ((objOutputWorkbook == null)) {
//                        objStatusRange.Value = "Aborted: Export file not specified - cannot export airport service areas";
//                        goto Err_Abort;
//                    }
//
//                    objLastOutputWorksheet = exportAirportServiceAreas(udtAirportsModel.udtAirports, udtAirportsModel.udtServiceAreas, objOutputWorkbook, objLastOutputWorksheet, strStepParameters);
//                    if ((objLastOutputWorksheet == null)) {
//                        goto Err_Abort;
//                    }
//
//                    break;
//                case "Export equipment models":
//                    if ((objOutputWorkbook == null)) {
//                        objStatusRange.Value = "Aborted: Export file not specified - cannot export equipment models";
//                        goto Err_Abort;
//                    }
//
//                    objLastOutputWorksheet = exportEquipmentModels(udtAirportsModel.udtEquipmentModels, udtAirportsModel.udtEquipmentTypes, objOutputWorkbook, objLastOutputWorksheet, strStepParameters);
//                    if ((objLastOutputWorksheet == null)) {
//                        goto Err_Abort;
//                    }
//
//                    break;
//                case "Export equipment":
//                    if ((objOutputWorkbook == null)) {
//                        objStatusRange.Value = "Aborted: Export file not specified - cannot export equipment";
//                        goto Err_Abort;
//                    }
//
//                    objLastOutputWorksheet = exportEquipment(udtAirportsModel.udtAirports, udtAirportsModel.udtEquipmentModels, udtAirportsModel.udtEquipment, objOutputWorkbook, objLastOutputWorksheet, strStepParameters);
//                    if ((objLastOutputWorksheet == null)) {
//                        goto Err_Abort;
//                    }
//
//                    break;
//                case "Export PM requirements":
//                    if ((objOutputWorkbook == null)) {
//                        objStatusRange.Value = "Aborted: Export file not specified - cannot export PM requirements";
//                        goto Err_Abort;
//                    }
//
//                    objLastOutputWorksheet = exportEquipmentPM(udtAirportsModel.udtEquipmentModels, objOutputWorkbook, objLastOutputWorksheet, strStepParameters);
//                    if ((objLastOutputWorksheet == null)) {
//                        goto Err_Abort;
//                    }
//
//                    break;
//                case "Export CM requirements":
//                    if ((objOutputWorkbook == null)) {
//                        objStatusRange.Value = "Aborted: Export file not specified - cannot export CM requirements";
//                        goto Err_Abort;
//                    }
//
//                    objLastOutputWorksheet = exportEquipmentCM(udtAirportsModel.udtCMRequirements, objOutputWorkbook, objLastOutputWorksheet, strStepParameters);
//                    if ((objLastOutputWorksheet == null)) {
//                        goto Err_Abort;
//                    }
//
//                    break;
//                case "Export PM status":
//                    if ((objOutputWorkbook == null)) {
//                        objStatusRange.Value = "Aborted: Export file not specified - cannot export PM status";
//                        goto Err_Abort;
//                    }
//
//                    objLastOutputWorksheet = exportPMStatus(udtAirportsModel.udtAirports, udtAirportsModel.udtEquipmentModels, udtAirportsModel.udtEquipment, objOutputWorkbook, objLastOutputWorksheet, strStepParameters);
//                    if ((objLastOutputWorksheet == null)) {
//                        goto Err_Abort;
//                    }
//
//                    break;
//                case "Export daily PM times":
//                    if ((objOutputWorkbook == null)) {
//                        objStatusRange.Value = "Aborted: Export file not specified - cannot export daily PM times";
//                        goto Err_Abort;
//                    }
//
//                    objLastOutputWorksheet = exportDailyPMTimes(udtAirportsModel.udtEquipmentModels, udtAirportsModel.udtEquipment, objOutputWorkbook, objLastOutputWorksheet, strStepParameters);
//                    if ((objLastOutputWorksheet == null)) {
//                        goto Err_Abort;
//                    }
//
//                    break;
//                case "Export PM schedule":
//                    if ((objOutputWorkbook == null)) {
//                        objStatusRange.Value = "Aborted: Export file not specified - cannot export PM schedule";
//                        goto Err_Abort;
//                    }
//
//                    objLastOutputWorksheet = exportPMSchedule(udtAirportsModel.udtAirports, udtAirportsModel.udtEquipmentModels, udtAirportsModel.udtEquipment, objOutputWorkbook, objLastOutputWorksheet, strStepParameters);
//                    if ((objLastOutputWorksheet == null)) {
//                        goto Err_Abort;
//                    }
//
//                    break;
//                default:
//                    objStatusRange.Value = "Aborted: Invalid step name";
//                    goto Err_Abort;
//                    break;
//            }
//            objModelWorksheet.Range(("B" + (12 + lngStepNumber))).Value = "Done";
//        }
//
//        if (!(objOutputWorkbook == null)) {
//            if ((objOutputWorkbook.Name == "")) {
//                objOutputWorkbook.SaveAs;
//                objModelWorksheet.Range("B5").Value;
//            }
//            else {
//                objOutputWorkbook.Save;
//            }
//
//        }
//
//        objStatusRange.Value = ("Completed: " + Now());
//        Err_Abort:
//        objStatusRange = null;
//        objErrorRange = null;
//        if ((objModelDataWorkbook == objModelWorkbook)) {
//            objModelDataWorkbook = null;
//        }
//        else {
//            objModelDataWorkbook.Close;
//            objModelDataWorkbook = null;
//        }
//
//        objLastOutputWorksheet = null;
//        if (!(objOutputWorkbook == null)) {
//            if (!(objOutputWorkbook == objModelWorkbook)) {
//                objOutputWorkbook.Close;
//            }
//
//            objOutputWorkbook = null;
//        }
//
//        objModelWorkbook = null;
//    }
//
//    public final void SortTest_Click() {
//        Excel.Range objRange;
//        object varValues;
//        ((double)(dblValues()));
//        ((long)(lngValueIndexes()));
//        long lngIndex;
//        objRange = Application.ActiveSheet.Range("J7").CurrentRegion;
//        varValues = objRange.Value;
//        object dblValues;
//        -1;
//        lngValueIndexes(0, To, (UBound(varValues, 1) - 1));
//        for (lngIndex = 1; (lngIndex <= UBound(varValues, 1)); lngIndex++) {
//            dblValues[(lngIndex - 1)] = varValues[lngIndex, 1];
//        }
//
//        sortValues;
//        dblValues;
//        (lngIndex - 1);
//        lngValueIndexes;
//        for (lngIndex = 1; (lngIndex <= UBound(varValues, 1)); lngIndex++) {
//            varValues[lngIndex, 1] = dblValues[lngValueIndexes((lngIndex - 1))];
//        }
//
//        objRange.Value = varValues;
//    }
//
//    public final long colNameToNumber(String  strLabel) {
//        if ((strLabel.Length == 1)) {
//            colNameToNumber = (Asc(strLabel) - 64);
//        }
//        else {
//            colNameToNumber = ((26
//                    * (Asc(strLabel.Substring(0, 1)) - 64))
//                    + (Asc(strLabel.Substring(1, 1)) - 64));
//        }
//
//    }
//
//    public final String  colNumberToName(long lngColNumber) {
//        if ((lngColNumber < 27)) {
//            colNumberToName = ((char)((64 + lngColNumber)));
//        }
//        else {
//            colNumberToName = (Chr((64
//                    + (lngColNumber - 1)), 26) + ((char)((65
//                    + ((lngColNumber - 1)
//                    % 26)))));
//        }
//
//    }
}
