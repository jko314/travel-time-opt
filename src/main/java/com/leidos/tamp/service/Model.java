package com.leidos.tamp.service;

import com.leidos.tamp.beans.Airport;
import com.leidos.tamp.beans.EquipmentModel;
import com.leidos.tamp.type.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.PropertySource;
import org.springframework.core.env.Environment;
import org.springframework.stereotype.Component;

import java.time.LocalDate;
import java.util.*;

import static com.leidos.tamp.type.ModelEnum.MODELSTEPTYPE_ENUM.*;

@Configuration
@Component
@PropertySource("classpath:service-area-model.properties")
public class Model {

    @Autowired
    private Environment env;

    List m_objStatusRange;
    List m_objErrorRange;

    long m_lngModelStepCount;
    List<MODELSTEP_TYPE> m_udtModelSteps = new ArrayList<>();

    LocalDate m_dteModelStartDate;
    LocalDate m_dteModelEndDate;
    long m_lngTravelCostCount;
    TRAVELCOSTS_TYPE[] m_udtTravelCosts;
    Collection m_colTravelCosts;
    long m_lngPMItemCount;
    PMITEM_TYPE[] m_udtPMItems;

    static double c_dblEarthRadius_Equitorial;
    static double c_dblEarthRadius_Polar;
    static double c_dblPiOver180;

    public Model() {
    }

    // // Private m_dblAirportDistances() As Double
// // ''''''''' Methods for running the model
// // '''''
// // '
// // '''''''''' Import / Export Methods
// // '''''
// // '
// // '''''''''' Model Management
    public final void initializeModel(LocalDate dteStartDate, LocalDate dteEndDate) {

        c_dblEarthRadius_Equitorial = env.getProperty("c_dblEarthRadius_Equitorial", Double.class);
        c_dblEarthRadius_Polar = env.getProperty("c_dblEarthRadius_Polar", Double.class);
        c_dblPiOver180 = env.getProperty("c_dblPiOver180", Double.class);

        initializeModel(dteStartDate, dteEndDate, null, null);
    }

    public final void initializeModel(LocalDate dteStartDate, LocalDate dteEndDate, List objStatusRange, List objErrorRange) {
        m_dteModelStartDate = dteStartDate;
        m_dteModelEndDate = dteEndDate;
        m_objStatusRange = objStatusRange;
        m_objErrorRange = objErrorRange;

        m_lngModelStepCount = 1;
        m_udtModelSteps = new ArrayList<>(64);

        for (MODELSTEP_TYPE step : m_udtModelSteps) {
            step.setStepType(MDLSTEP_initialize);
            List<String> l = new ArrayList<>();
            l.add(dteStartDate.toString());
            l.add(dteEndDate.toString());
            step.setParameters(l);
            step.setStatus("Done");
        }
    }

    public final boolean randomizeModel(String strStepParameter) {

//        Dim lngStepNumber As Long
//
//        lngStepNumber = recordModelStep(MDLSTEP_randomize, strStepParameter)
//
//        If m_udtModelSteps(lngStepNumber).lngParameterCount = 0 Then
//                Randomize
//        ElseIf m_udtModelSteps(lngStepNumber).lngParameterCount = 1 Then
//        If IsNumeric(m_udtModelSteps(lngStepNumber).strParameters(0)) Then
//        Randomize -1
//        Randomize CLng(m_udtModelSteps(lngStepNumber).strParameters(0))
//        Else
//        m_udtModelSteps(lngStepNumber).strStatus = "Error: Invalid parameter"
//        Exit Function
//        End If
//        Else
//        m_udtModelSteps(lngStepNumber).strStatus = "Error: Invalid parameter"
//        Exit Function
//        End If
//
//        m_udtModelSteps(lngStepNumber).strStatus = "Done"
//        randomizeModel = True
        return false;
    }

    //'''''''''' Equipment Models

//    public final boolean loadEquipmentModels(EQUIPMENTMODELS_TYPE udtEquipmentModels, EQUIPMENTTYPES_TYPE udtEquipmentTypes, String strStepParameter)
//    {
//        List<Double> varValues = new ArrayList<>();
//        int lngRowNumber = 0;
//        int lngEquipmentTypeIndex = 0;
//        String strEquipmentTypeName = null;
//
//        if (m_objStatusRange != null)
//        {
//            m_objStatusRange.add("Loading equipment models ...");
//        }
//
//        if (!(objWorksheet.Range("A1").Value.equals("Worksheet Type:")))
//        {
//            return false;
//        }
//        if (!(objWorksheet.Range("B1").Value.equals("Equipment Models")))
//        {
//            return false;
//        }
//        if (!verifyColumnHeadings(objWorksheet, c_strColHeadings_EquipmentModels, 8))
//        {
//            if (m_objErrorRange != null)
//            {
//                m_objErrorRange.Value = "Invalid column headings in EquipmentModels";
//            }
//            return false;
//        }
//
//        lngRowNumber = 10;
//        lngRowNumber = objWorksheet.Range("A10:A10").CurrentRegion.Rows.Count;
//        varValues = objWorksheet.Range("A10:C" + (9 + lngRowNumber)).Value;
//        udtEquipmentTypes.lngEquipmentTypeCount = 0;
////VB TO JAVA CONVERTER TODO TASK: The following 'ReDim' could not be resolved. A possible reason may be that the object of the ReDim was not declared as an array:
//        ReDim udtEquipmentTypes.udtEquipmentTypes(15);
////VB TO JAVA CONVERTER WARNING: VB to Java Converter converts from VB(.NET), not VB6:
////VB TO JAVA CONVERTER TODO TASK: There is no Java equivalent to the legacy VB6 'Collection' class:
//        Set udtEquipmentTypes.colEquipmentTypes = new Collection();
//
//        udtEquipmentModels.lngEquipmentModelCount = varValues.GetUpperBound(0);
////VB TO JAVA CONVERTER TODO TASK: The following 'ReDim' could not be resolved. A possible reason may be that the object of the ReDim was not declared as an array:
//        ReDim udtEquipmentModels.udtEquipmentModels(udtEquipmentModels.lngEquipmentModelCount - 1);
////VB TO JAVA CONVERTER WARNING: VB to Java Converter converts from VB(.NET), not VB6:
////VB TO JAVA CONVERTER TODO TASK: There is no Java equivalent to the legacy VB6 'Collection' class:
//        Set udtEquipmentModels.colEquipmentModels = new Collection();
//
//        for (lngRowNumber = 1; lngRowNumber <= udtEquipmentModels.lngEquipmentModelCount; lngRowNumber++)
//        {
//
//            strEquipmentTypeName = varValues(lngRowNumber, 1);
////VB TO JAVA CONVERTER TODO TASK: The 'On Error Resume Next' statement is not converted by VB to Java Converter:
//            On Error Resume Next
//            lngEquipmentTypeIndex = -1;
//            lngEquipmentTypeIndex = udtEquipmentTypes.colEquipmentTypes.Item("ET:" + strEquipmentTypeName);
////VB TO JAVA CONVERTER TODO TASK: The following 'On Error GoTo' statement cannot be converted by VB to Java Converter:
//            On Error GoTo 0
//            if (lngEquipmentTypeIndex == -1)
//            {
//                lngEquipmentTypeIndex = udtEquipmentTypes.lngEquipmentTypeCount;
//                if (udtEquipmentTypes.udtEquipmentTypes.GetUpperBound(0) < udtEquipmentTypes.lngEquipmentTypeCount)
//                {
////VB TO JAVA CONVERTER TODO TASK: The following 'ReDim' could not be resolved. A possible reason may be that the object of the ReDim was not declared as an array:
//                    ReDim Preserve udtEquipmentTypes.udtEquipmentTypes(7 + udtEquipmentTypes.lngEquipmentTypeCount);
//                }
////VB TO JAVA CONVERTER TODO TASK: The return type of the tempVar variable must be corrected.
////ORIGINAL LINE: With udtEquipmentTypes.udtEquipmentTypes(lngEquipmentTypeIndex)
//                Object tempVar = udtEquipmentTypes.udtEquipmentTypes(lngEquipmentTypeIndex);
//                tempVar.lngID = lngEquipmentTypeIndex + 1;
//                tempVar.strName = strEquipmentTypeName;
//                udtEquipmentTypes.lngEquipmentTypeCount = udtEquipmentTypes.lngEquipmentTypeCount + 1;
//                udtEquipmentTypes.colEquipmentTypes.Add lngEquipmentTypeIndex, "ET:" + strEquipmentTypeName;
//            }
//
////VB TO JAVA CONVERTER TODO TASK: The return type of the tempVar2 variable must be corrected.
////ORIGINAL LINE: With udtEquipmentModels.udtEquipmentModels(lngRowNumber - 1)
//            Object tempVar2 = udtEquipmentModels.udtEquipmentModels(lngRowNumber - 1);
//            tempVar2.lngID = lngRowNumber;
//            tempVar2.lngEquipmentTypeIndex = lngEquipmentTypeIndex;
//            tempVar2.strManufacturer = varValues(lngRowNumber, 2);
//            tempVar2.strModel = varValues(lngRowNumber, 3);
//            udtEquipmentModels.colEquipmentModels.Add lngRowNumber - 1, "EM:" + tempVar2.strModel;
//        }
//
//        if (m_objStatusRange != null)
//        {
//            m_objStatusRange.Value = "Loading equipment models ... done";
//        }
//
//        return true;
//
//    }

//    // // '''''''''' Airports
//    public final boolean loadAirports(AIRPORTS_TYPE udtAirports, String strStepParameter) {
//
//        Variant varValues = null;
//        long lngRowNumber = 0;
//        long lngStepNumber = 0;
//
//        if (m_objStatusRange != null) {
//            m_objStatusRange.Value = "Loading airports ...";
//        }
//
//        lngStepNumber = recordModelStep(MDLSTEP_loadAirports, strStepParameter);
//
//        if (!verifyColumnHeadings(objWorksheet, c_strColHeadings_Airports, 8)) {
//            recordError m_udtModelSteps (lngStepNumber), "Incorrect headings in Airports worksheet";
//            return false;
//        }
//
//        lngRowNumber = objWorksheet.Range("A10:A10").CurrentRegion.Rows.Count;
//        varValues = objWorksheet.Range("A10:J" + (9 + lngRowNumber)).Value;
//
//        udtAirports.argValue.lngAirportCount = varValues.GetUpperBound(0);
////VB TO JAVA CONVERTER TODO TASK: The following 'ReDim' could not be resolved. A possible reason may be that the object of the ReDim was not declared as an array:
//        ReDim udtAirports.argValue.udtAirport(udtAirports.argValue.lngAirportCount - 1);
////VB TO JAVA CONVERTER WARNING: VB to Java Converter converts from VB(.NET), not VB6:
////VB TO JAVA CONVERTER TODO TASK: There is no Java equivalent to the legacy VB6 'Collection' class:
//        Set udtAirports.argValue.colAirports = new Collection();
//        for (lngRowNumber = 1; lngRowNumber <= udtAirports.argValue.lngAirportCount; lngRowNumber++) {
////VB TO JAVA CONVERTER TODO TASK: The return type of the tempVar variable must be corrected.
////ORIGINAL LINE: With udtAirports.udtAirport(lngRowNumber - 1)
//            Object tempVar = udtAirports.udtAirport(lngRowNumber - 1);
//            tempVar.lngID = lngRowNumber;
//            tempVar.strCode = varValues(lngRowNumber, 1);
//            tempVar.strCat = varValues(lngRowNumber, 2);
//            tempVar.strCity = varValues(lngRowNumber, 3);
//            tempVar.strState = varValues(lngRowNumber, 4);
//            tempVar.strName = varValues(lngRowNumber, 5);
//            tempVar.dblLatitude = varValues(lngRowNumber, 6);
//            tempVar.dblLongitude = varValues(lngRowNumber, 7);
//            tempVar.lngOperatingStartHour = varValues(lngRowNumber, 8);
//            tempVar.lngOperatingHours = varValues(lngRowNumber, 9);
//            if (!varValues(lngRowNumber, 10).equals("")) {
//                tempVar.lngTimeZoneAdjustment = varValues(lngRowNumber, 10);
//            }
//            udtAirports.argValue.colAirports.Add lngRowNumber -1, "A:" + tempVar.strCode;
//        }
//
//        if (m_objStatusRange != null) {
//            m_objStatusRange.Value = "Loading airports ... done";
//        }
//        m_udtModelSteps(lngStepNumber).strStatus = "Done";
//
//        return true;
//    }

    public static final double computeDistance(double Longitude1, double Latitude1, double Longitude2, double Latitude2) {
        double v = (Math.pow(c_dblEarthRadius_Equitorial
                * Math.cos(0.5 * (Latitude1 + Latitude2) * c_dblPiOver180) * c_dblPiOver180 * (Longitude2 - Longitude1), 2)
                + Math.pow(c_dblEarthRadius_Polar * c_dblPiOver180 * (Latitude2 - Latitude1), 2));

        return v * v;
    }


    private long recordModelStep(ModelEnum.MODELSTEPTYPE_ENUM enmStepType, String strStepParameter) {
//        MODELSTEP_TYPE step = new MODELSTEP_TYPE();
//        step.setEnmStepType(enmStepType);
//        step.setStrParameters(Arrays.asList(strStepParameter));
//        m_udtModelSteps.add(step);
//
//        return m_udtModelSteps.size();
//
//        if (m_udtModelSteps.length < m_lngModelStepCount) {
////VB TO JAVA CONVERTER TODO TASK: The following 'ReDim' could not be resolved. A possible reason may be that the object of the ReDim was not declared as an array:
//            ReDim Preserve m_udtModelSteps(63 + m_lngModelStepCount);
//        }
////VB TO JAVA CONVERTER TODO TASK: The return type of the tempVar variable must be corrected.
////ORIGINAL LINE: With m_udtModelSteps(m_lngModelStepCount)
//        Object tempVar = m_udtModelSteps(m_lngModelStepCount);
//        tempVar.enmStepType = enmStepType;
////VB TO JAVA CONVERTER TODO TASK: The following 'ReDim' could not be resolved. A possible reason may be that the object of the ReDim was not declared as an array:
//        ReDim tempVar.strParameters(0);
//        parseParameters strStepParameter, tempVar.strParameters, tempVar.lngParameterCount;
//        tempVar.strStatus = "";
//        m_lngModelStepCount = m_lngModelStepCount + 1;
//
//        return m_lngModelStepCount - 1;
        return -1;

    }

    private void parseParameters(String strParameterString, String[] strParameters, long lngParameterCount)
    {
//
//        long lngStartOffset = 0;
//        long lngEndOffset = 0;
//
//        lngParameterCount.argValue = 0;
//        lngStartOffset = 1;
//        while (lngStartOffset < (strParameterString == null ? 0 : strParameterString.length()))
//        {
//            lngEndOffset = strParameterString.indexOf(",", lngStartOffset - 1) + 1;
//            if (lngEndOffset == 0)
//            {
//                lngEndOffset = (strParameterString == null ? 0 : strParameterString.length()) + 1;
//            }
//            if (strParameters.argValue.length - 1 < lngParameterCount.argValue)
//            {
////VB TO JAVA CONVERTER NOTE: The following block reproduces what 'ReDim Preserve' does behind the scenes in VB:
////ORIGINAL LINE: ReDim Preserve strParameters(0 To 7 + lngParameterCount)
//                String[] tempVar = new String[(7 + lngParameterCount.argValue) + 1];
//                if (strParameters.argValue != null)
//                    System.arraycopy(strParameters.argValue, 0, tempVar, 0, Math.min(strParameters.argValue.length, tempVar.length));
//                strParameters.argValue = tempVar;
//            }
//            strParameters.argValue[lngParameterCount.argValue] = tangible.StringHelper.trim(tangible.StringHelper.substring(strParameterString, (int)(lngStartOffset - 1), (int)(lngEndOffset - lngStartOffset)), ' ');
//            lngParameterCount.argValue = lngParameterCount.argValue + 1;
//            lngStartOffset = lngEndOffset + 1;
//        }

    }

//    public final boolean applyCMRequirements(Map<String, CMREQUIREMENT_TYPE> udtCMRequirements,
//                                             List<Airport> udtAirports, List<EquipmentModel> udtEquipmentModels)
//    {
//
//        long lngAirportIndex = 0;
//        long lngAirportEquipmentIndex = 0;
//        long lngCMRequirementIndex = 0;
//        String strLocation = null;
//        String strCategory = null;
//        String strModelNum = null;
//        boolean booFail;
//        booFail = false;
//        for (lngAirportIndex = 0; lngAirportIndex < udtAirports.getLngAirportCount(); lngAirportIndex++)
//        {
////VB TO JAVA CONVERTER TODO TASK: The return type of the tempVar variable must be corrected.
////ORIGINAL LINE: With udtAirports.udtAirport(lngAirportIndex)
//            AIRPORT_TYPE tempVar = udtAirports.getUdtAirport()[(int) lngAirportIndex];
//            strLocation = tempVar.getStrCode();
//            strCategory = tempVar.getStrCat();
//
//            for (lngAirportEquipmentIndex = 0; lngAirportEquipmentIndex < tempVar.getLngEquipmentCount(); lngAirportEquipmentIndex++)
//            {
////VB TO JAVA CONVERTER TODO TASK: The return type of the tempVar2 variable must be corrected.
////ORIGINAL LINE: With .udtEquipment(lngAirportEquipmentIndex)
//                AIRPORTEQUIPMENT_TYPE tempVar2 = tempVar.getUdtEquipment()[(int) lngAirportEquipmentIndex];
//                strModelNum = udtEquipmentModels.getUdtEquipmentModels()[(int) tempVar2.getLngEquipmentModelIndex()].getStrModel();
//
////VB TO JAVA CONVERTER TODO TASK: The 'On Error Resume Next' statement is not converted by VB to Java Converter:
////                On Error Resume Next
//                boolean done = false;
//                done = udtCMRequirements.getColCMRequirement().contains("I:" + strModelNum + ":" + strLocation);
//                if (!done)
//                {
//                    done = udtCMRequirements.getColCMRequirement().contains("I:" + strModelNum + ":" + strCategory);
//                }
//                if (!done)
//                {
//                    done = udtCMRequirements.getColCMRequirement().contains("I:" + strModelNum);
//                }
////VB TO JAVA CONVERTER TODO TASK: The following 'On Error GoTo' statement cannot be converted by VB to Java Converter:
////                On Error GoTo 0
//                tempVar2.setLngCMRequirementIndex(lngCMRequirementIndex);
//                if (!done)
//                {
////VB TO JAVA CONVERTER TODO TASK: Calls to the VB 'Err' function are not converted by VB to Java Converter:
////                    Microsoft.VisualBasic.Information.Err().Raise 5;
//                }
//
//            }
//        }
//        return !booFail;
//    }

//    public final double[][] computeAirportDistances(AIRPORTS_TYPE udtAirports)
//    {
//        double[][] dblAirportDistances;
//
//        int lngAirportIndex1 = 0;
//        int lngAirportIndex2 = 0;
//        double Longitude = 0;
//        double Latitude = 0;
//
//        dblAirportDistances = new double[(int) udtAirports.getLngAirportCount()][(int) udtAirports.getLngAirportCount()];
//        for (lngAirportIndex1 = 0; lngAirportIndex1 < udtAirports.getLngAirportCount(); lngAirportIndex1++)
//        {
//            Longitude = udtAirports.getUdtAirport()[lngAirportIndex1].getDblLongitude();
//            Latitude = udtAirports.getUdtAirport()[lngAirportIndex1].getDblLatitude();
//            for (lngAirportIndex2 = 0; lngAirportIndex2 < udtAirports.getLngAirportCount(); lngAirportIndex2++)
//            {
//                dblAirportDistances[lngAirportIndex1][lngAirportIndex2] =
//                        computeDistance(dblLongitude, dblLatitude,
//                                udtAirports.getUdtAirport()[lngAirportIndex2].getDblLongitude(),
//                                udtAirports.getUdtAirport()[lngAirportIndex2].getDblLatitude());
//            }
//        }
//
//        return dblAirportDistances;
//    }

    public static final Map<String, Map<String, Double>> computeAirportDistances(Map<String, AIRPORT_TYPE> airportsMap)
    {
        Map<String, Map<String, Double>> airportDistancesMap = new HashMap<>();

        double lon = 0;
        double lat = 0;

        int i=0;
        for (AIRPORT_TYPE airport : airportsMap.values()) {
            lon = airport.getLongitude();
            lat = airport.getLatitude();
            Map<String, Double> distMap = new HashMap<>();
            for (AIRPORT_TYPE airportX : airportsMap.values()) {
                if (!airport.getCode().equals(airportX.getCode())) {
                    double distance = computeDistance(lon, lat,
                            airportX.getLongitude(),
                            airportX.getLatitude());
                    distMap.put(airportX.getCode(), distance);
                }
            }

            airportDistancesMap.put(airport.getCode(), distMap);
        }

        return airportDistancesMap;
    }

}
