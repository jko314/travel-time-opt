package com.leidos.tamp.type;

import lombok.Data;

import java.util.Date;
import java.util.List;

@Data
public class EQUIPMENTITEM_TYPE {
    long id;
    int equipmentModelIndex;
    String airportCode;
    int airportEquipmentModelIndex;
    int airportEquipmentIndex;
    List<Integer> pmRequirementsIndexes;
    PMSCHEDULE_TYPE pmSchedule;
    Date nextPMDue;
    int nextPMRequirementIndex;
    List<Integer> pmIndexes;
    List<CMACTIVITY_TYPE> cmActivities;
}
