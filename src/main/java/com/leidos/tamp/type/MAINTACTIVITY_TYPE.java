package com.leidos.tamp.type;

import lombok.Data;

import java.util.Date;

@Data
public class MAINTACTIVITY_TYPE {
    long equipmentIndex;
    long pmIndex;
    Date scheduledStart;
    Date actualStartDate;
    Date actualEndDate;
}
