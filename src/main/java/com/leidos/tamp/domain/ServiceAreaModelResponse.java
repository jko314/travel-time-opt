package com.leidos.tamp.domain;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonRawValue;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.extern.slf4j.Slf4j;

import static com.fasterxml.jackson.annotation.JsonInclude.Include.NON_NULL;

@Data
@Slf4j
@AllArgsConstructor
@NoArgsConstructor
public class ServiceAreaModelResponse {
    private int saCount;
    private int iterCount;
    private double travelMiles;
    private int fstHours;
    private int fstCount;
    private double fstUtil;
    private double cost;
}
