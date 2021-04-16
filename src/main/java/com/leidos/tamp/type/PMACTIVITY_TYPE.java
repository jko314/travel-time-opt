package com.leidos.tamp.type;

import lombok.Data;

import java.util.Date;

@Data
public class PMACTIVITY_TYPE {
    int pmIndex;
    Date startTime;
    Date endTime;
}
