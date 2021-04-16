package com.leidos.tamp.type;

import lombok.Data;

import java.util.Collection;

@Data
//TODO remove, replaced by HashMap
public class AIRPORTS_TYPE {
    long lngAirportCount;
    AIRPORT_TYPE[] udtAirport;
    Collection colAirports;
}
