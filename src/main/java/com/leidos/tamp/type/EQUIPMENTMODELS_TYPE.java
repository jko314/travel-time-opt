package com.leidos.tamp.type;

import lombok.Data;

import java.util.Collection;

@Data
//TODO remove, replace with hashMap
public class EQUIPMENTMODELS_TYPE {
    long lngEquipmentModelCount;
    EQUIPMENTMODEL_TYPE[] udtEquipmentModels;
    Collection colEquipmentModels;
}
