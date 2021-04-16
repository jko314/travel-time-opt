package com.leidos.tamp.type;

import lombok.Data;

@Data
public class PIECEWISELINEAR_TYPE {
    long lngItemCount;
    PIECEWISELINEARITEM_TYPE[] udtItems;
}
