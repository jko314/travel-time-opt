package com.leidos.tamp.type;

import lombok.Data;

import java.util.Date;
import java.util.List;

@Data
public class PMSCHEDULE_TYPE {
    ModelEnum.PMPERIODICITY_ENUM enmPeriodicity;

    //    ' Specifies the month within a cycle in which PM is performed.
    //    ' If weekly or monthly, this is 0
    //    ' If quarterly, this is 0 to 2, indicating the month within the quarter
    //    ' If semiannually, this is 0 to 5,  indicating the month within the semiannual cycle
    //    ' If annually, this is 0 to 11
    long month;

    //    ' Specifies on which day PM is performed.
    //    ' If weekly, it is the day of the week (1-7)
    //    ' Otherwise, it is the day of the month (1-31)
    long day;

    long lngPMScheduleIndex; //' Index into lngPMSchedule array of the EquipmentModel structure
    Date lastPMCompleted;

    List<Integer> lngPMItemIndexes; // ' Indexes into m_udtPMItems
}
