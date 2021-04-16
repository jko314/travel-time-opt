package com.leidos.tamp.type;

import lombok.Data;

import java.util.Date;

@Data
public class PMITEM_TYPE {
    int equipmentIndex;
    int pmRequirementIndex;
    int tripIndex;
    Date scheduledStart;
    Date startTime;
    Date endTime;
}
