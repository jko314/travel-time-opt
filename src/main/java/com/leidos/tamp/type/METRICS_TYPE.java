package com.leidos.tamp.type;

import lombok.Data;

import java.util.List;

@Data
public class METRICS_TYPE {
    List<METRICITEM_TYPE> udtEquipmentModelMetrics;
    List<METRICITEM_TYPE> udtEquipmentTypeMetrics;
}
