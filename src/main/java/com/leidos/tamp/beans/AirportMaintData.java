package com.leidos.tamp.beans;

import lombok.Data;

@Data
public class AirportMaintData {
    double pmTripCount;
    double pmTime;
    double cmTripCount;
    double cmTime;
    double depotTripCount;
    double depotTime;
    double tripCount;
    double time;
}
