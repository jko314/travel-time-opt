package com.leidos.tamp.type;

import com.leidos.tamp.beans.Airport;
import lombok.Data;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Data
public class AIRPORT_TYPE {
    String code;
    String cat;
    String city;
    String state;
    String name;
    double latitude;
    double longitude;
    int operatingStartHour;
    int operatingHours;
    int timeZoneAdjustment;
    int baseServiceAreaIndex;
    List<AIRPORTSERVICEAREA_TYPE> serviceAreaIndexes = new ArrayList<>();
    double cmTravelTime;

    Map<String, Integer> equipmentCountMap = new HashMap<>();

    ModelEnum.PMPERIODICITY_ENUM pmPeriodicity;

    //  Most frequent PM at airport
    List<Double> pmTime = new ArrayList<>(); // (1, To, c_lngPMPeriodicity_MaxValue)));

    public AIRPORT_TYPE(Airport airport) {
        this.code = airport.getCode();
        this.cat = airport.getCat();
        this.city = airport.getCity();
        this.state = airport.getState();
        this.latitude = airport.getLatitude();
        this.longitude = airport.getLongitude();
        this.operatingHours = airport.getOp_hrs();
        this.operatingStartHour = airport.getOp_start();
        this.timeZoneAdjustment = airport.getTime_zone();
    }
}
