package com.leidos.tamp.service;

import com.leidos.tamp.type.PIECEWISELINEAR_TYPE;
import org.junit.Test;

import static org.junit.jupiter.api.Assertions.*;

public class ServiceAreaUtilTest {
    @Test
    public void setFSTTimeToCountModel() {
        PIECEWISELINEAR_TYPE udtFSTTimeToCount = new PIECEWISELINEAR_TYPE();
        ServiceAreaUtil.setFSTTimeToCountModel(udtFSTTimeToCount, "Flat 80% Utilization");
        assertEquals(udtFSTTimeToCount.getUdtItems().length, 1);
    }
}