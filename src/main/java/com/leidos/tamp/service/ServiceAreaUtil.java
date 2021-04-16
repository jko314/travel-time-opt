package com.leidos.tamp.service;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.leidos.tamp.beans.*;
import com.leidos.tamp.type.*;
import org.springframework.util.ResourceUtils;

import java.io.File;
import java.io.IOException;
import java.util.*;

import static com.leidos.tamp.service.ServiceAreaModelService.*;
import static com.leidos.tamp.type.ModelEnum.PMPERIODICITY_ENUM.*;

public class ServiceAreaUtil {

    public static void loadEquipmentModels(SERVICEAREAMODEL_TYPE serviceAreaModel, List<EquipmentModel> equipmentModels) {
        Map<String, EQUIPMENTMODEL_TYPE> equipmentModelMap = serviceAreaModel.getEquipmentModels();
        for (EquipmentModel model : equipmentModels) {
            EQUIPMENTMODEL_TYPE equipmentModel = new EQUIPMENTMODEL_TYPE();
            equipmentModel.setType(model.getEquipmentType());
            equipmentModel.setManufacturer(model.getManufacturer());
            equipmentModel.setModel(model.getModel());

            equipmentModelMap.put(model.getModel(), equipmentModel);
        }
    }

    public static void loadAirportEquipments(SERVICEAREAMODEL_TYPE serviceAreaModelType, List<AirportEquipment> airportEquipments) {
        for (AirportEquipment equipment : airportEquipments) {

            AIRPORT_TYPE airport = serviceAreaModelType.getAirportMap().get(equipment.getAirport());

            if (airport != null) {
                airport.getEquipmentCountMap().put(equipment.getModel(), equipment.getCount());
            } else {
                // TODO log it
            }
        }
    }


    public static void loadAirports(SERVICEAREAMODEL_TYPE serviceAreaModelType, List<Airport> airports) {
        Map<String, AIRPORT_TYPE> airportTypeMap = serviceAreaModelType.getAirportMap();
        if (airportTypeMap == null) {
            airportTypeMap = new HashMap<>();
        }

        for (Airport airport : airports) {
            AIRPORT_TYPE airportType = new AIRPORT_TYPE(airport);
            airportTypeMap.put(airport.getCode(), airportType);
        }
    }

    public static void setFSTTimeToCountModel(PIECEWISELINEAR_TYPE udtFSTTimeToCount, String strModelName)
    {
        try {
            // create object mapper instance
            ObjectMapper mapper = new ObjectMapper();


            final File file = ResourceUtils.getFile("classpath:fstutil.json");

            List<PIECEWISELINEARITEM_TYPE> fstList =
                    Arrays.asList(mapper.readValue(file, PIECEWISELINEARITEM_TYPE[].class));

            udtFSTTimeToCount.setLngItemCount(
                    fstList.stream().filter(item -> strModelName.equals(item.getModelName())).count());
            udtFSTTimeToCount.setUdtItems((PIECEWISELINEARITEM_TYPE[]) fstList.toArray());
            fstList.forEach(System.out::println);
        } catch (IOException ioe) {
            System.out.println(ioe);
            // TODO
        }

    }

    public static void loadEquipmentPM(Map<String, AIRPORT_TYPE> airportMap,
                                       Map<String, EQUIPMENTMODEL_TYPE> equipmentModels, List<EquipmentPM> equipmentPM) {

        long lngRowNumber = 0;
        String strMakeModel = null;
        long lngEquipmentModelIndex = 0;
        long lngPMIndex = 0;
        long lngAirportIndex = 0;
        long lngAirportEquipmentModelIndex = 0;
        PMREQUIREMENTS_TYPE udtNullPMRequirement = null;
        long lngIndex = 0;

        // Initialize structures to indicate empty PM
        for (EquipmentPM itemPM : equipmentPM) {
            String modelName = itemPM.getMakeModel();
            EQUIPMENTMODEL_TYPE model = equipmentModels.get(modelName);
            model.setEnmPeriodicity(NoPM);

            for (int i=0; i <= c_lngPMPeriodicity_MaxValue; i++) {
                model.getPmList().add(i, new PMREQUIREMENTS_TYPE());
                model.getPmTime().add(i, 0D);
            }


            ModelEnum.PMPERIODICITY_ENUM enmPeriodicity = ModelEnum.PMPERIODICITY_ENUM.valueOf(itemPM.getPeriodicity());
            PMREQUIREMENTS_TYPE pmrequirementsType = model.getPmList().get(enmPeriodicity.value());

            switch (enmPeriodicity) {
                case Weekly:
                    pmrequirementsType.setEventsPerYear(52);
                    break;
                case Monthly:
                    pmrequirementsType.setEventsPerYear(12);
                    break;
                case Quarterly:
                    pmrequirementsType.setEventsPerYear(4);
                    break;
                case SemiAnnually:
                    pmrequirementsType.setEventsPerYear(2);
                    break;
                case Annually:
                    pmrequirementsType.setEventsPerYear(1);
                    break;
                case NoPM:
                    pmrequirementsType.setEventsPerYear(0);
                    break;
                default:
                    throw new RuntimeException("Unknown enmPeriodicity: " + enmPeriodicity);
            }
            model.getPmList().set(enmPeriodicity.value(), pmrequirementsType);

            if (model.getEnmPeriodicity() == NoPM ||
                    enmPeriodicity.value() < model.getEnmPeriodicity().value()) {
                model.setEnmPeriodicity(enmPeriodicity);
            }
        }

        // Correct PM events per year for less frequent events
        for (EQUIPMENTMODEL_TYPE model : equipmentModels.values()) {
            for (int i = c_lngPMPeriodicity_MaxValue-1; i >= 0; i--) {
                PMREQUIREMENTS_TYPE type = model.getPmList().get(i);
                if (0 < type.getEventsPerYear()) {
                    for (int j = i+1; j < c_lngPMPeriodicity_MaxValue; j++) {
                        PMREQUIREMENTS_TYPE type2 = model.getPmList().get(j);
                        type.setEventsPerYear(type.getEventsPerYear() - type2.getEventsPerYear());
                    }
                    model.getPmTime().set(i,
                            type.getEventsPerYear() * (
                                    type.getLabor_Initial() + type.getLabor_Wait() + type.getLabor_Final()
                                    )
                    );
                }
            }
        }

        // Create PM schedule
        for (EQUIPMENTMODEL_TYPE model : equipmentModels.values()) {
            switch (model.getEnmPeriodicity()) {
                case NoPM:
                    //  Do nothing
                    break;
                case Weekly:
                    for (int i = 0; i < 51; i++) {
                        model.getPmSchedule().put(i, Weekly);
                    }

                    if ((0 < model.getPmList().get(Monthly.value()).getEventsPerYear())) {
                        model.getPmSchedule().put(48, Monthly);
                        model.getPmSchedule().put(44, Monthly);
                        model.getPmSchedule().put(39, Monthly);
                        model.getPmSchedule().put(35, Monthly);
                        model.getPmSchedule().put(31, Monthly);
                        model.getPmSchedule().put(26, Monthly);
                        model.getPmSchedule().put(22, Monthly);
                        model.getPmSchedule().put(18, Monthly);
                        model.getPmSchedule().put(13, Monthly);
                        model.getPmSchedule().put(9, Monthly);
                        model.getPmSchedule().put(5, Monthly);
                        model.getPmSchedule().put(0, Monthly);
                    }

                    if ((0 < model.getPmList().get(Quarterly.value()).getEventsPerYear())) {
                        model.getPmSchedule().put(39, Quarterly);
                        model.getPmSchedule().put(26, Quarterly);
                        model.getPmSchedule().put(13, Quarterly);
                        model.getPmSchedule().put(0, Quarterly);
                    }

                    if ((0 < model.getPmList().get(SemiAnnually.value()).getEventsPerYear())) {
                        model.getPmSchedule().put(26, SemiAnnually);
                        model.getPmSchedule().put(0, SemiAnnually);
                    }

                    if ((0 < model.getPmList().get(Annually.value()).getEventsPerYear())) {
                        model.getPmSchedule().put(0, Annually);
                    }

                    break;
                case Monthly:
                    for (int i = 0; i < 11; i++) {
                        model.getPmSchedule().put(i, Monthly);
                    }

                    if (0 < model.getPmList().get(Quarterly.value()).getEventsPerYear()) {
                        model.getPmSchedule().put(0, Quarterly);
                        model.getPmSchedule().put(3, Quarterly);
                        model.getPmSchedule().put(6, Quarterly);
                        model.getPmSchedule().put(9, Quarterly);
                    }

                    if (0 < model.getPmList().get(SemiAnnually.value()).getEventsPerYear()) {
                        model.getPmSchedule().put(0, SemiAnnually);
                        model.getPmSchedule().put(6, SemiAnnually);
                    }

                    if (0 < model.getPmList().get(Annually.value()).getEventsPerYear()) {
                        model.getPmSchedule().put(0, Annually);
                    }

                    break;
                case Quarterly:
                    for (int i = 0; i < 3; i++) {
                        model.getPmSchedule().put(i, Quarterly);
                    }

                    if (0 < model.getPmList().get(SemiAnnually.value()).getEventsPerYear()) {
                        model.getPmSchedule().put(0, SemiAnnually);
                        model.getPmSchedule().put(2, SemiAnnually);
                    }

                    if (0 < model.getPmList().get(Annually.value()).getEventsPerYear()) {
                        model.getPmSchedule().put(0, Annually);
                    }

                    break;
                case SemiAnnually:
                    for (int i = 0; i < 1; i++) {
                        model.getPmSchedule().put(i, SemiAnnually);
                    }

                    if (0 < model.getPmList().get(Annually.value()).getEventsPerYear()) {
                        model.getPmSchedule().put(0, Annually);
                    }

                    break;
                case Annually:
                    for (int i = 0; i < 1; i++) {
                        model.getPmSchedule().put(i, Annually);
                    }
                    break;
                default:
                    throw new RuntimeException("Unknown enmPeriodicity: " + model.getEnmPeriodicity());
            }
        }

        //  Set airport PM periodicity
        for (AIRPORT_TYPE airport : airportMap.values()) {

            airport.setPmPeriodicity(null);
            for (int i = 0; i < c_lngPMPeriodicity_MaxValue; i++) {
                airport.getPmTime().add(i, 0d);
            }

            for (String modelName :  airport.getEquipmentCountMap().keySet()) {
                EQUIPMENTMODEL_TYPE model = equipmentModels.get(modelName);
                if (model.getEnmPeriodicity() != NoPM) {
                    if (airport.getPmPeriodicity() == null) {
                        airport.setPmPeriodicity(model.getEnmPeriodicity());
                    }

                    Integer count = airport.getEquipmentCountMap().get(model.getModel());
                    EQUIPMENTMODEL_TYPE equipmentmodel_type = equipmentModels.get(model.getModel());
                    for (int i = 0; i < c_lngPMPeriodicity_MaxValue; i++) {
                        PMREQUIREMENTS_TYPE pmrequirements = equipmentmodel_type.getPmList().get(i);
                        airport.getPmTime().set(i, airport.getPmTime().get(i) +
                                count * pmrequirements.getEventsPerYear() *
                                        (pmrequirements.getLabor_Initial() + pmrequirements.getLabor_Wait() +
                                                pmrequirements.getLabor_Final()
                                        )
                        );
                    }
                }
            }

            if (airport.getPmPeriodicity() == null) {
                airport.setPmPeriodicity(NoPM);
            }
        }
    }

    public static void loadEquipmentCM(Map<String, CMREQUIREMENT_TYPE> cmRequirements, List<EquipmentCM> equipmentCM) {

        for (EquipmentCM equipment : equipmentCM) {
            CMREQUIREMENT_TYPE requirement = new CMREQUIREMENT_TYPE();
            requirement.setModelNum(equipment.getMakeModel());
            requirement.setName(equipment.getName());
            requirement.setFrequency(equipment.getFrequency());
            DISTRIBUTION_TYPE distribution = new DISTRIBUTION_TYPE();
            distribution.setAvg(equipment.getCmTime());
            distribution.setStndDev(equipment.getCmsStndDev());
            distribution.setMin(equipment.getCmMin());
            distribution.setMax(equipment.getCmMax());
            requirement.setCmTime(distribution);
            requirement.setPartsCost(equipment.getPartsCost());
            DISTRIBUTION_TYPE partsTime = new DISTRIBUTION_TYPE();
            partsTime.setAvg(equipment.getPartsTime());
            partsTime.setStndDev(equipment.getPartsStndDev());
            partsTime.setMin(equipment.getCmMin());
            partsTime.setMax(equipment.getCmMax());
            requirement.setPartsTime(partsTime);
            requirement.setConsumablesCost(equipment.getConsumablesCost());
            requirement.setTechnicianCount(equipment.getTechCount());

            cmRequirements.put(equipment.getName(), requirement);
        }
    }


    public static void assignAirportSelectionP(SERVICEAREAMODEL_TYPE udtServiceAreaModel) {

        long lngAirportIndex;
        double dblWeightTotal = 0;

        double dblExponent = udtServiceAreaModel.getAirportSelectionExponent();

        for (AIRPORT_TYPE airport : udtServiceAreaModel.getAirportMap().values()) {
            udtServiceAreaModel.getAirportSelectionPMap().put(airport.getCode(),
                    Math.pow(udtServiceAreaModel.getAirportDataMap().get(airport.getCode()).getTripCount(), dblExponent));

            if (!airport.getCat().startsWith("Cat")) {
                udtServiceAreaModel.getAirportSelectionPMap().put(airport.getCode(),
                        udtServiceAreaModel.getAirportSelectionPMap().get(airport.getCode()) / 1000.0);
            }

            dblWeightTotal = dblWeightTotal + udtServiceAreaModel.getAirportSelectionPMap().get(airport.getCode());
        }

        Double prev = null;
        for (String code : udtServiceAreaModel.getAirportMap().keySet()) {
            if (prev == null) {
                prev = udtServiceAreaModel.getAirportSelectionPMap().get(code) / dblWeightTotal;
                udtServiceAreaModel.getAirportSelectionPMap().put(code, prev);
            } else {
                udtServiceAreaModel.getAirportSelectionPMap().put(code,
                        prev + udtServiceAreaModel.getAirportSelectionPMap().get(code) / dblWeightTotal);
                prev = udtServiceAreaModel.getAirportSelectionPMap().get(code);
            }
        }

    }
}
