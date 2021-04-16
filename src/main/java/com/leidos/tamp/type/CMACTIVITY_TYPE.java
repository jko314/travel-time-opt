package com.leidos.tamp.type;

import lombok.Data;

import java.util.Date;

@Data
public class CMACTIVITY_TYPE {
    int cmIndex;
    Date failureTime;
    Date callTime;
    Date dispatchTime;
    Date arrivalTime;
    Date diagnosisEndTime;
    Date partsRequestTime;
    Date partsFulfillmentTime;
    Date partsLocalLogisticsTime;
    Date repairTime;
    Date testTime;
    Date signoffTime;
}
