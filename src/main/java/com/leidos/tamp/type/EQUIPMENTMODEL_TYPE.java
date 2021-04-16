package com.leidos.tamp.type;

import lombok.Data;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Data
public class EQUIPMENTMODEL_TYPE {
    long id;
    int equipmentTypeIndex;
    String type;
    String manufacturer;
    String model;
    ModelEnum.PMPERIODICITY_ENUM enmPeriodicity;
    List<PMREQUIREMENTS_TYPE> pmList = new ArrayList<>();; //(0, To, c_lngPMPeriodicity_MaxValue)));

    //  Lists in sequential order the PM requirements for a year.
    //  Each entry is an index into udtPM
    Map<Integer, ModelEnum.PMPERIODICITY_ENUM> pmSchedule = new HashMap<>();

    List<Double> pmTime =  new ArrayList<>(); //(0, To, c_lngPMPeriodicity_MaxValue)));
    //  Total PM time in each category
    //  CM Requirements
    //   lngCMCount As Long
    //   udtCM() As CMREQUIREMENT_TYPE
    //
}
