package com.leidos.tamp.type;


import lombok.Data;

import java.util.Date;
import java.util.List;

@Data
public class MAINTENANCETRIP_TYPE {
    long id;
    ModelEnum.MTRIP_STATUS_ENUM enmTripStatus;
    Date scheduledStart;
    double scheduledLength;
    int itemCurrent;
    List<MTRIP_ITEM_TYPE> items;
}
