package com.leidos.tamp.type;

import lombok.Data;

import java.util.HashMap;
import java.util.Map;

@Data
public class AIRPORTEQUIPMENT_TYPE {
    int equipmentModelIndex;
    int cmRequirementIndex;
    Map<String, Integer> modelsCountMap = new HashMap<>();
}

