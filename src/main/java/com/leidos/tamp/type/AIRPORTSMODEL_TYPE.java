package com.leidos.tamp.type;

import lombok.Data;

import java.util.Date;
import java.util.List;
import java.util.Map;

@Data
public class AIRPORTSMODEL_TYPE {
    Date modelStartDate;
    Date modelEndDate;
//    AIRPORTS_TYPE udtAirports;
    Map<String, AIRPORT_TYPE> airportsMap;
//    SERVICEAREAS_TYPE udtServiceAreas;
    Map<String, SERVICEAREA_TYPE> serviceAreasMap;

    Map<String, EQUIPMENTTYPE_TYPE> equipmentTypes;
    Map<String, EQUIPMENTMODEL_TYPE> equipmentModels;
    Map<String, CMREQUIREMENT_TYPE> cmRequirements;
    List<EQUIPMENTITEM_TYPE> equipment;
    double[] airportDistances;
}
