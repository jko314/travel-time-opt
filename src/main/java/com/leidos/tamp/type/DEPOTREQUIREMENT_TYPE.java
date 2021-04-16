package com.leidos.tamp.type;

import lombok.Data;

@Data
public class DEPOTREQUIREMENT_TYPE {
    String name;
    double frequency;
    DISTRIBUTION_TYPE diagnosisTime;
    DISTRIBUTION_TYPE reinstallTime;
}
