package com.leidos.tamp.type;

import lombok.Data;

@Data
public class PMREQUIREMENTS_TYPE {
// ''''''''' PM data structures
    String name;
    int allowedSlack;
    int eventsPerYear;
    double labor_Initial;
    double labor_Wait;
    double labor_Final;
    double consumables;
    int lngTechnicianCount;
}
