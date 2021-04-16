package com.leidos.tamp.type;


import lombok.Data;

import java.util.Date;

@Data
public class MTRIP_ITEM_TYPE {
    ModelEnum.MTRIP_ITEMTYPE_ENUM enmItemType;
    int itemIndex;
    Date startTime;
    Date endTime;
}
