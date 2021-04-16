package com.leidos.tamp.type;

import lombok.Data;

import java.util.Collection;
import java.util.List;

@Data
// TODO remove, use map
public class CMREQUIREMENTS_TYPE {
    long lngCMRequirementCount;
    CMREQUIREMENT_TYPE[] udtCMRequirements;
    List<String> colCMRequirement;
}
