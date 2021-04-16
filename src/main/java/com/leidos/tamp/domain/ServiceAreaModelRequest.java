package com.leidos.tamp.domain;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class ServiceAreaModelRequest {
    private int serviceAreaCount;
    private int maxMileage;
    private int communitySize;
    private int retain;
    private int mate;
    private int mateSelApproach1;
    private int mateSelApproach2;
    private int insertionDeletion;
    private int airportSelectionExponent;
    private int solutionSelectionExponent;
    private int maxPMTimePerTrip;
    private String cmRateSource;
    private int P1;
    private int P2;
    private int P1P2;
    private int fstUtilization;
}
