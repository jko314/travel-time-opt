package com.leidos.tamp.type;

import com.fasterxml.jackson.databind.annotation.JsonDeserialize;
import com.leidos.tamp.beans.FstUtilDeserializer;
import lombok.Data;

@Data
@JsonDeserialize(using = FstUtilDeserializer.class)
public class PIECEWISELINEARITEM_TYPE {
    String modelName;
    double FromValue;
    double ToValue;
    double B;
    double M;
}
