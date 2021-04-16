package com.leidos.tamp.type;

import lombok.Data;
import java.util.List;

@Data
public class SERVICEAREA_TYPE {

    long id;
    String name;
    String city;
    String state;
    double latitude;
    double longitude;

    List<Integer> airportIndexes;

    // Maintenance activities
    List<Integer> scheduledMaintIndexes;
    List<Integer> completedMaintIndexes;
    List<Integer> interuptedMaintIndexes;

}
