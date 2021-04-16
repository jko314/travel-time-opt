package com.leidos.tamp.type;

import lombok.Data;

@Data
public class METRICITEM_TYPE {
    String airportCode;
    int equipmentModelIndex;
    int equipmentTypeIndex;
    int equipmentCount;
    double operatingTime;
    int pmEvents;
    double pmTime;
    int cmEvents;
    double cmTime;
}
