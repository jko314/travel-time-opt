package com.leidos.tamp.type;

import lombok.Data;

import java.util.List;

@Data
public class MODELSTEP_TYPE {
    ModelEnum.MODELSTEPTYPE_ENUM stepType;
    List<String> parameters;
    String  status;
}
