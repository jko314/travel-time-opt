package com.leidos.tamp.type;

import lombok.Data;

import java.util.Date;

@Data
public class TRAVEL_TYPE {
    ModelEnum.LOCATIONTYPE_ENUM enmFromType;
    int fromIndex;
    ModelEnum.LOCATIONTYPE_ENUM enmToType;
    int toIndex;
    Date scheduleStart;
    Date scheduleEnd;
    Date actualStart;
    Date actualEnd;
}
