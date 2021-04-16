package com.leidos.tamp.type;

import lombok.Data;

@Data
public class CMREQUIREMENT_TYPE {
    String modelNum;
    String name;
    double frequency;
    DISTRIBUTION_TYPE cmTime;
    double partsCost;
    DISTRIBUTION_TYPE partsTime;
    double consumablesCost;
    int technicianCount;
}
