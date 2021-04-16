package com.leidos.tamp.type;

import lombok.Data;

import java.util.List;

@Data
public class EVOLUTION_TYPE {
    ModelEnum.EVOLUTIONAPPROACH_ENUM enmApproach;
    int appliesTo;
    int parameterCount;
    List<String> VarParameters;
}
