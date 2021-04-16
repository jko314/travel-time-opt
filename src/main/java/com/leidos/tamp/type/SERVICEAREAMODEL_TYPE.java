package com.leidos.tamp.type;

import lombok.Data;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Data
public class SERVICEAREAMODEL_TYPE {

    Map<String, AIRPORT_TYPE> airportMap = new HashMap<>();
    SERVICEAREAS_TYPE serviceAreas;
    Map<String, EQUIPMENTTYPE_TYPE> equipmentTypes = new HashMap<>();
    Map<String, EQUIPMENTMODEL_TYPE> equipmentModels = new HashMap<>();
    Map<String, CMREQUIREMENT_TYPE> cmRequirements = new HashMap<>();

    Map<String, Map<String, Double>> airportDistances = new HashMap<>();
    int serviceAreaCount;
    int serviceAreaCount_Min;
    int serviceAreaCount_Max;
    double MaximumTravelMiles;
    int communitySize;
    int evolutionCount;
    List<EVOLUTION_TYPE> evolutions;
    double AirportSelectionExponent;
    double SolutionSelectionExponent;
    double MaxPMTimePerTrip;
    int iterationNumber;
    int iterationCount;
    List<Double> sortValues = new ArrayList<>();
    Map<String, Double> airportSelectionPMap = new HashMap<>();
    Map<String, EQUIPMENTITEM_TYPE> equipmentMap = new HashMap<>();
    Map<String, AIRPORTMAINTDATA_TYPE> airportDataMap = new HashMap<>();
    int solutionCount;
    POSSIBLESOLUTION_TYPE[] solutions;
    boolean booSolutionSelectionValid;
    List<POSSIBLESOLUTION_TYPE> sortedSolutions;
    Map<Integer, Double> solutionSelectionP = new HashMap<>();
    boolean isSolutionSortIsValid;
//    List<Integer> solutionSortIndexes;
    double CostPerMile;
    double CostPerFSTHour;
    double FSTHoursPerYear;
    PIECEWISELINEAR_TYPE fSTTimeToCount = new PIECEWISELINEAR_TYPE();
}
