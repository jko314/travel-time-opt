package com.leidos.tamp.type;

import lombok.Data;

import java.util.Map;

@Data
public class POSSIBLESOLUTION_TYPE {
    boolean isResultsAreValid;
//    int serviceAreaCount;
//    SERVICEAREADATA_TYPE[] serviceAreaData;
    Map<String, SERVICEAREADATA_TYPE> serviceAreaData;
    boolean booServiceAreasAreSorted;
    Map<String, AIRPORTSERVICEAREADATA_TYPE> airportData;
//    AIRPORTSERVICEAREADATA_TYPE[] airportData;
    double OptimizeOn;
    double TravelMiles;
    double PMTripCount;
    double PMTime;
    double CMTripCount;
    double CMTime;
    double FSTHours;
    int FSTCount;
    double FSTCost;
}
