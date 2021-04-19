package com.leidos.tamp.service;

import com.leidos.tamp.beans.*;
import com.leidos.tamp.domain.Task;
import com.leidos.tamp.type.*;
import com.squareup.okhttp.Callback;
import com.squareup.okhttp.OkHttpClient;
import com.squareup.okhttp.Request;
import com.squareup.okhttp.Response;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.PropertySource;
import org.springframework.core.env.Environment;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.time.LocalDate;
import java.util.*;
import java.util.stream.Stream;

import static com.leidos.tamp.type.ModelEnum.EVOLUTIONAPPROACH_ENUM.*;

@Slf4j
@Service
@Configuration
@PropertySource("classpath:service-area-model.properties")
public class ServiceAreaModelService {

    @Autowired
    private Environment env;

    @Autowired
    private IRADServiceImpl iradService;

    @Autowired
    private Model model;

    private OkHttpClient client = new OkHttpClient();

    public static int c_lngPMPeriodicity_MaxValue = 7;
    public static String c_strPMPeriodicityList = "0;NoPM;1;Daily;2;Weekly;3;Biweekly;4;Monthly;5;Quarterly;6;Semi-Annual;7;Annual;";
    public static Map<String, Integer> pmPeriodicityMap = new HashMap<>();
    public static Map<ModelEnum.PMPERIODICITY_ENUM, PMREQUIREMENTS_TYPE> pmMap = new HashMap<>();

    public POSSIBLESOLUTION_TYPE runServiceAreaModel_Click(int objIterationNumberRange, String flatUtilization) {

        double costPerMile = env.getProperty("costPerMile", Double.class);
        double costPerFSTHour = env.getProperty("costPerFSTHour", Double.class);
        double fstHoursPerYear = env.getProperty("fstHoursPerYear", Double.class);

        c_lngPMPeriodicity_MaxValue = env.getProperty("c_lngPMPeriodicity_MaxValue", Integer.class);
        c_strPMPeriodicityList = env.getProperty("c_strPMPeriodicityList");
        Iterator itr = Arrays.asList(c_strPMPeriodicityList.split(";")).iterator();
        while (itr.hasNext()) {
            Integer v = Integer.parseInt((String) itr.next());
            String s = (String) itr.next();
            pmPeriodicityMap.put(s, v);
        }

        SERVICEAREAMODEL_TYPE serviceAreaModel = new SERVICEAREAMODEL_TYPE();
        List<Integer> objCurrentSolutionsRange = new ArrayList<>();
        List<Double> currentSolution = new ArrayList<>(5);
        int solutionIndex = 0;
        int i = 0;
        int lngSACount = 0;
        int index = 0;
        double[] dblSortValues = null;

        serviceAreaModel.setCostPerMile(costPerMile);
        serviceAreaModel.setCostPerFSTHour(costPerFSTHour);
        serviceAreaModel.setFSTHoursPerYear(fstHoursPerYear);
        ServiceAreaUtil.setFSTTimeToCountModel(serviceAreaModel.getFSTTimeToCount(), flatUtilization);

        // Read parameters
        serviceAreaModel.setServiceAreaCount(env.getProperty("serviceAreaCount", Integer.class));
        serviceAreaModel.setServiceAreaCount_Min(env.getProperty("serviceAreaCount_Min", Integer.class));
        serviceAreaModel.setServiceAreaCount_Max(env.getProperty("serviceAreaCount_Max", Integer.class));
        serviceAreaModel.setMaximumTravelMiles(env.getProperty("maximumTravelMiles", Integer.class));
        serviceAreaModel.setCommunitySize(env.getProperty("communitySize", Integer.class));

        List<EVOLUTION_TYPE> evolution_types = new ArrayList<EVOLUTION_TYPE>();
        serviceAreaModel.setEvolutions(evolution_types);

        EVOLUTION_TYPE evolutionType = new EVOLUTION_TYPE();
        evolutionType.setEnmApproach(EVOLAPP_RetainTop);
        evolutionType.setAppliesTo(env.getProperty("retain", Integer.class));
        evolutionType.setParameterCount(0);
        evolution_types.add(evolutionType);

        evolutionType = new EVOLUTION_TYPE();
        evolutionType.setEnmApproach(EVOLAPP_Mate);
        evolutionType.setAppliesTo(env.getProperty("mate", Integer.class));
        evolutionType.setParameterCount(1);
        List<String> params = new ArrayList<>();
        params.add(env.getProperty("mateSelectionApproach1"));
        params.add(env.getProperty("mateSelectionApproach2"));
        evolutionType.setVarParameters(params);
        evolution_types.add(evolutionType);

        evolutionType = new EVOLUTION_TYPE();
        evolutionType.setEnmApproach(EVOLAPP_InsertDelete);
        evolutionType.setAppliesTo(env.getProperty("insertion-deletion", Integer.class));
        evolutionType.setParameterCount(0);
        evolution_types.add(evolutionType);

        evolutionType = new EVOLUTION_TYPE();
        evolutionType.setEnmApproach(EVOLAPP_Random);
        evolutionType.setAppliesTo(env.getProperty("random", Integer.class));
        evolutionType.setParameterCount(0);
        evolution_types.add(evolutionType);


        serviceAreaModel.setEvolutionCount(4);
        serviceAreaModel.setAirportSelectionExponent(env.getProperty("airportSelectionExponent", Long.class));
        serviceAreaModel.setSolutionSelectionExponent(env.getProperty("solutionSelectionExponent", Long.class));
        serviceAreaModel.setMaxPMTimePerTrip(env.getProperty("maxPMTimePerTrip", Double.class));
        serviceAreaModel.setIterationCount(objIterationNumberRange);

        // Load airport/equipment variables
        model.initializeModel(LocalDate.parse("2013-01-01"), LocalDate.parse("2014-01-01"));

        // Not needed since we're using DB
//        loadAirports serviceAreaModel.udtAirports, objWorkbook.Worksheets.Item("Airports"), "";
//        loadEquipmentModels serviceAreaModel.udtEquipmentModels, serviceAreaModel.udtEquipmentTypes, objWorkbook.Worksheets.Item("EquipmentModels"), "";
//        loadEquipment serviceAreaModel.udtAirports, serviceAreaModel.udtEquipmentModels, serviceAreaModel.udtEquipment, objWorkbook.Worksheets.Item("Airport_Equipment"), "ByCount";
//        loadEquipmentPM serviceAreaModel.udtAirports, serviceAreaModel.udtEquipmentModels, objWorkbook.Worksheets.Item("EquipmentPM"), "ByType";
//        loadCMRequirements serviceAreaModel.udtCMRequirements, objWorkbook.Worksheets.Item("EquipmentCM"), "";

        List<Airport> airports = iradService.getAirports();
        List<EquipmentModel> equipmentModel = iradService.getEquipmentModel();
        List<AirportEquipment> airportEquipments = iradService.getAirportEquipments();
        List<EquipmentPM> equipmentPM = iradService.getEquipmentPM();
        List<EquipmentCM> equipmentCM = iradService.getEquipmentCM();

        ServiceAreaUtil.loadAirports(serviceAreaModel, airports);
        ServiceAreaUtil.loadEquipmentModels(serviceAreaModel, equipmentModel);
        ServiceAreaUtil.loadAirportEquipments(serviceAreaModel, airportEquipments);
        ServiceAreaUtil.loadEquipmentPM(serviceAreaModel.getAirportMap(), serviceAreaModel.getEquipmentModels(), equipmentPM);
        ServiceAreaUtil.loadEquipmentCM(serviceAreaModel.getCmRequirements(), equipmentCM);

        // following are not needed
//        applyCMRequirements serviceAreaModel.udtCMRequirements, serviceAreaModel.udtAirports, serviceAreaModel.udtEquipmentModels;
//        computeAirportDistances serviceAreaModel.udtAirports, serviceAreaModel.dblAirportDistances;

        // TODO need data validation during load and save phase
//        model.applyCMRequirements(serviceAreaModel.getUdtCMRequirements(), serviceAreaModel.getAirports(),
//                serviceAreaModel.getEquipmentModels());

        serviceAreaModel.setAirportDistances(Model.computeAirportDistances(serviceAreaModel.getAirportMap()));

        computeAirportData(serviceAreaModel.getAirportMap(), serviceAreaModel.getAirportDataMap(),
                serviceAreaModel.getCmRequirements(), serviceAreaModel.getMaxPMTimePerTrip());

        // Generate starting set of solutions
        ServiceAreaUtil.assignAirportSelectionP(serviceAreaModel);

        POSSIBLESOLUTION_TYPE[] sols = Stream.iterate(0, x-> x+1 ).limit(serviceAreaModel.getCommunitySize())
                .map(s -> new POSSIBLESOLUTION_TYPE()).toArray(POSSIBLESOLUTION_TYPE[]::new);

        serviceAreaModel.setSolutions(sols);
        for (solutionIndex = 0; solutionIndex < serviceAreaModel.getCommunitySize(); solutionIndex++) {
            serviceAreaModel.getSolutions()[solutionIndex].setServiceAreaData(
                    new LinkedHashMap<>(serviceAreaModel.getServiceAreaCount_Max())
            );
            serviceAreaModel.getSolutions()[solutionIndex].setAirportData(
                    new LinkedHashMap<>(serviceAreaModel.getAirportMap().size())
            );

            lngSACount = serviceAreaModel.getServiceAreaCount_Min()
                    + new Random().nextInt(1 + serviceAreaModel.getServiceAreaCount_Max()
                    - serviceAreaModel.getServiceAreaCount_Min());
            randomSolutionServiceArea(serviceAreaModel, serviceAreaModel.getSolutions()[solutionIndex], lngSACount);
            evaluateSolution(serviceAreaModel, serviceAreaModel.getSolutions()[solutionIndex],
                    serviceAreaModel.getMaximumTravelMiles());
        }
        serviceAreaModel.setSolutionSelectionP(new HashMap<>(serviceAreaModel.getCommunitySize()));


        // Iterate
        for (i = 1; i <= serviceAreaModel.getIterationCount(); i++) {
            // Evolve solution set
            evolveSolutions(serviceAreaModel);
            replaceDuplicateSolutions(serviceAreaModel);
            serviceAreaModel.setBooSolutionSelectionValid(false);

            if (!serviceAreaModel.isSolutionSortIsValid()) {
                sortSolutions(serviceAreaModel);
            }
            for (index = 1; index <= 5; index++) {
                currentSolution.add(serviceAreaModel.getSortedSolutions().get(index).getOptimizeOn());
            }
//            objCurrentSolutionsRange.Value = varCurrentSolutionValue;
//            DoEvents;
        }

        return serviceAreaModel.getSortedSolutions().get(0);
    }

    private void replaceDuplicateSolutions(SERVICEAREAMODEL_TYPE serviceAreaModel) {

        if (!serviceAreaModel.isSolutionSortIsValid()) {
            sortSolutions(serviceAreaModel);
        }

        int lastIndex = 0;

        for (int i = 1; i < serviceAreaModel.getCommunitySize(); i++) {
            if (compareSolutions(serviceAreaModel.getSortedSolutions().get(lastIndex),
                    serviceAreaModel.getSortedSolutions().get(i))) {
                randomSolutionServiceArea(serviceAreaModel, serviceAreaModel.getSortedSolutions().get(i),
                        serviceAreaModel.getServiceAreaCount());
                evaluateSolution(serviceAreaModel, serviceAreaModel.getSortedSolutions().get(i),
                        serviceAreaModel.getMaximumTravelMiles());
                serviceAreaModel.setSolutionSortIsValid(false);
                sortSolutions(serviceAreaModel);
            } else {
                lastIndex = i;
            }
        }
    }

    private boolean compareSolutions(POSSIBLESOLUTION_TYPE sol1, POSSIBLESOLUTION_TYPE sol2) {
        return sol1.getServiceAreaData().keySet().equals(sol2.getServiceAreaData().keySet());
    }

    private final void computeAirportData(Map<String, AIRPORT_TYPE> airportMap,
                                          Map<String, AIRPORTMAINTDATA_TYPE> airportData,
                                          Map<String, CMREQUIREMENT_TYPE> cmrequirementTypeMap,
                                          double MaxPMTimePerTrip) {
        double PMTime;
        int equipmentCount;
        double CMCount;
        double CMTime;
        long PMTripCount;

        for (AIRPORT_TYPE airport : airportMap.values()) {
            AIRPORTMAINTDATA_TYPE airportmaintdata_type = new AIRPORTMAINTDATA_TYPE();
            airportData.put(airport.getCode(), airportmaintdata_type);

            switch (airport.getPmPeriodicity()) {
                case Weekly:
                    airportmaintdata_type.setCmTripCount(52);
                    break;
                case Monthly:
                    airportmaintdata_type.setCmTripCount(12);
                    break;
                case Quarterly:
                    airportmaintdata_type.setCmTripCount(4);
                    break;
                case SemiAnnually:
                    airportmaintdata_type.setCmTripCount(2);
                    break;
                case Annually:
                    airportmaintdata_type.setCmTripCount(1);
                    break;
                case NoPM:
                    airportmaintdata_type.setCmTripCount(0);
                    break;
                default:
                    throw new RuntimeException("airport.getEnmPMPeriodicity() enum not known");
            }

            if ((airportmaintdata_type.getPmTripCount() == 0)) {
                airportmaintdata_type.setPmTime(0);
            } else {
                PMTime = 0D;
                for (double time : airport.getPmTime()) {
                    PMTime = PMTime + time;
                }
                PMTripCount = (int) ((PMTime + MaxPMTimePerTrip - 1) / MaxPMTimePerTrip);
                if (airportmaintdata_type.getPmTripCount() < PMTripCount) {
                    airportmaintdata_type.setPmTripCount(PMTripCount);
                }
                airportmaintdata_type.setPmTime(PMTime);
            }

            CMCount = 0D;
            CMTime = 0D;
            equipmentCount = airport.getEquipmentCountMap().size();

            for (String name : airport.getEquipmentCountMap().keySet()) {
                CMREQUIREMENT_TYPE type = cmrequirementTypeMap.get(name + ":" + airport.getCode());
                if (type == null) {
                    type = cmrequirementTypeMap.get(name + ":" + airport.getCat());
                }
                if (type == null) {
                    type = cmrequirementTypeMap.get(name);
                }
                CMCount += equipmentCount * type.getFrequency();
                CMTime += equipmentCount * type.getFrequency() * type.getCmTime().getAvg();
            }

            airportmaintdata_type.setCmTripCount(CMCount);
            airportmaintdata_type.setCmTime(CMTime);

            // BUG: Do Depot time

            airportmaintdata_type.setTripCount(airportmaintdata_type.getCmTripCount() + airportmaintdata_type.getPmTripCount());
            airportmaintdata_type.setTime(airportmaintdata_type.getCmTime() + airportmaintdata_type.getPmTime());

        }
    }

    public void execute(final Task task) {
        log.info("Started task with {} urls", task.getUrls().size());
        task.start();
        for (int i = 0; i < task.getUrls().size(); i++) {
            final int index = i;
            final long time = System.currentTimeMillis();
            String url = task.getUrls().get(i);
            Request req = new Request.Builder().get().url(url).build();

            client.newCall(req).enqueue(new Callback() {
                @Override
                public void onFailure(Request request, IOException e) {
                    task.fail(index, time, request, e);
                }

                @Override
                public void onResponse(Response response) throws IOException {
                    task.success(index, time, response);
                }
            });
        }
    }


    public void randomSolutionServiceArea(SERVICEAREAMODEL_TYPE udtServiceAreaModel,
                                          POSSIBLESOLUTION_TYPE udtSolution,
                                          int count) {

//        udtSolution.setServiceAreaCount(0);

        for (int i = 0; i < count; i++) {
            insertSolutionServiceArea(udtServiceAreaModel, udtSolution, selectAirport(udtServiceAreaModel).getCode());
        }

        udtSolution.setResultsAreValid(false);
        udtSolution.setBooServiceAreasAreSorted(false);

    }

    private final boolean insertSolutionServiceArea(SERVICEAREAMODEL_TYPE udtServiceAreaModel,
                                                    POSSIBLESOLUTION_TYPE udtSolution, String airportCode) {

        if (udtSolution.getServiceAreaData().containsKey(airportCode)) {
            return false;
        }

        SERVICEAREADATA_TYPE serviceAreaDatum = new SERVICEAREADATA_TYPE();
        serviceAreaDatum.setLatitude(udtServiceAreaModel.getAirportMap().get(airportCode).getLatitude());
        serviceAreaDatum.setLongitude(udtServiceAreaModel.getAirportMap().get(airportCode).getLongitude());
        udtSolution.getServiceAreaData().put(airportCode, serviceAreaDatum);

//        udtSolution.setServiceAreaCount(udtSolution.getServiceAreaCount() + 1);
        udtSolution.setResultsAreValid(false);

        return true;
    }

    // TODO remove later
//    private final int findSolutionServiceArea(POSSIBLESOLUTION_TYPE udtSolution, String airportCode) {
//        int lngListIndex = 0;
//
//        int indexLow = 0;
//        int indexMid;
//        int indexHigh;
//        if ((udtSolution.getServiceAreaCount() == 0)) {
//            lngListIndex = 0;
//            return lngListIndex;
//        }
//
//        indexHigh = (int) udtSolution.getServiceAreaCount() - 1;
//        while (indexLow + 1 < indexHigh) {
//            indexMid = (indexLow + indexHigh) / 2;
//            if (udtSolution.getServiceAreaData()[indexMid].getAirportCode() < lngAirportIndex) {
//                indexLow = indexMid;
//            }
//            else {
//                indexHigh = indexMid;
//            }
//
//        }
//
//        if ((lngAirportIndex <= udtSolution.udtServiceAreaData(indexLow).lngAirportIndex)) {
//            lngListIndex = indexLow;
//            findSolutionServiceArea = (lngListIndex == udtSolution.udtServiceAreaData(indexLow).lngAirportIndex);
//        }
//        else if ((lngAirportIndex <= udtSolution.udtServiceAreaData(indexHigh).lngAirportIndex)) {
//            lngListIndex = indexHigh;
//            findSolutionServiceArea = (lngListIndex == udtSolution.udtServiceAreaData(indexHigh).lngAirportIndex);
//        }
//        else {
//            lngListIndex = (indexHigh + 1);
//        }
//
//    }


    private void evaluateSolution(SERVICEAREAMODEL_TYPE serviceAreaModel, POSSIBLESOLUTION_TYPE possibleSolution,
                                  double maxTravelDistance) {

        int lngServiceAreaIndex;
        double dblTravelDistance;
        int lngFinalServiceAreaIndex;
        boolean booKeepItem;

        int lngAirportIndex;
        int lngFSTCount;

        assignAirportServiceAreas(serviceAreaModel, possibleSolution);

        for (SERVICEAREADATA_TYPE areaData : possibleSolution.getServiceAreaData().values()) {
            if (areaData == null) {
                continue;
            }
            areaData.setPMTripCount(0);
            areaData.setPMTime(0);
            areaData.setCMTripCount(0);
            areaData.setCMTime(0);
            areaData.setTravelMiles(0);
        }

        for (String code : serviceAreaModel.getAirportMap().keySet()) {
            double travelDistance = possibleSolution.getAirportData().get(code).getTravelDistance();
            String serviceCode = possibleSolution.getAirportData().get(code).getServiceAreaCode();
            SERVICEAREADATA_TYPE data = possibleSolution.getServiceAreaData().get(serviceCode);
            if (data == null) {
                continue;
            }
            if (0 < travelDistance) {
                data.setPMTripCount(data.getPMTripCount() + serviceAreaModel.getAirportDataMap().get(code).getPmTripCount());
                data.setCMTripCount(data.getCMTripCount() + serviceAreaModel.getAirportDataMap().get(code).getCmTripCount());

                if (maxTravelDistance < travelDistance) {
                    travelDistance = maxTravelDistance;
                }
                data.setTravelMiles(data.getTravelMiles()
                        + (serviceAreaModel.getAirportDataMap().get(code).getPmTripCount()
                        + serviceAreaModel.getAirportDataMap().get(code).getCmTripCount()) * 2 * travelDistance);
            }

            data.setPMTime(data.getPMTime() + serviceAreaModel.getAirportDataMap().get(code).getPmTime());
            data.setCMTime(data.getCMTime() + serviceAreaModel.getAirportDataMap().get(code).getCmTime());
        }

//        possibleSolution.dblPMTripCount = 0#
//        possibleSolution.dblPMTime = 0#
//        possibleSolution.dblCMTripCount = 0#
//        possibleSolution.dblCMTime = 0#
//        possibleSolution.dblTravelMiles = 0#
//        possibleSolution.dblFSTHours = 0#
//        possibleSolution.lngFSTCount = 0

        for (SERVICEAREADATA_TYPE areaData : possibleSolution.getServiceAreaData().values()) {
            if (areaData == null) {
                continue;
            }
            areaData.setFSTTime(areaData.getPMTime() + areaData.getCMTime() + areaData.getTravelMiles() / 60.0);
            areaData.setFSTCount((int) computePiecewiseLinear(areaData.getFSTTime(), serviceAreaModel.getFSTTimeToCount()));

            if (0 < areaData.getFSTTime() && areaData.getFSTCount() == 0) {
                areaData.setFSTCount(1);
            }

            possibleSolution.setPMTripCount(possibleSolution.getPMTripCount() + areaData.getPMTripCount());
            possibleSolution.setPMTime(possibleSolution.getPMTime() + areaData.getPMTime());
            possibleSolution.setCMTripCount(possibleSolution.getCMTripCount() + areaData.getCMTripCount());
            possibleSolution.setCMTime(possibleSolution.getCMTime() + areaData.getCMTime());
            possibleSolution.setTravelMiles(possibleSolution.getTravelMiles() + areaData.getTravelMiles());
            possibleSolution.setFSTHours(possibleSolution.getFSTHours() + areaData.getFSTTime());
            possibleSolution.setFSTCount(possibleSolution.getFSTCount() + areaData.getFSTCount());
        }

        possibleSolution.setFSTCost(
                possibleSolution.getTravelMiles() * serviceAreaModel.getCostPerMile() + possibleSolution.getFSTCount()
                        * serviceAreaModel.getFSTHoursPerYear() * serviceAreaModel.getCostPerFSTHour()
        );
        possibleSolution.setOptimizeOn(possibleSolution.getFSTCost());


    }

    private final double computePiecewiseLinear(double dblX, PIECEWISELINEAR_TYPE udtPiecewiseLinear) {
        int index;
        for (index = 0; (index <= (udtPiecewiseLinear.getUdtItems().length - 1)); index++) {
            if (udtPiecewiseLinear.getUdtItems()[index].getFromValue() <= dblX
                    && dblX < udtPiecewiseLinear.getUdtItems()[index].getToValue()) {
                return (udtPiecewiseLinear.getUdtItems()[index].getB()
                        + (udtPiecewiseLinear.getUdtItems()[index].getM() * dblX));

            }
        }

        throw new RuntimeException("Unable to compute");
    }

    private final void assignAirportServiceAreas(SERVICEAREAMODEL_TYPE udtServiceAreaModel, POSSIBLESOLUTION_TYPE udtSolution) {
        long lngServiceAreaIndex;
        long lngServiceAreaAirportIndex;
        long lngAirportIndex;
        double[] dblAirportDistances;
        double dblDistance;

        Iterator<String> iterator = udtSolution.getServiceAreaData().keySet().iterator();
        String serviceAreaCode = iterator.next();
        for (String airportCode : udtServiceAreaModel.getAirportMap().keySet()) {
//            udtSolution.getAirportData().get(airportCode).lngServiceAreaIndex = 0;
            AIRPORTSERVICEAREADATA_TYPE airportData = new AIRPORTSERVICEAREADATA_TYPE();
            airportData.setServiceAreaCode(serviceAreaCode);
            if (airportCode.equals(serviceAreaCode)) {
                airportData.setTravelDistance(0);
            } else {
                airportData.setTravelDistance(udtServiceAreaModel.getAirportDistances().get(airportCode).get(serviceAreaCode));
            }
            udtSolution.getAirportData().put(airportCode, airportData);

        }

        while (iterator.hasNext()) {
            String code = iterator.next();
            for (String airportCode : udtServiceAreaModel.getAirportMap().keySet()) {
                double dist = airportCode.equals(code) ?
                        0 : udtServiceAreaModel.getAirportDistances().get(airportCode).get(code);
                if (dist < udtSolution.getAirportData().get(airportCode).getTravelDistance()) {
                    udtSolution.getAirportData().get(airportCode).setTravelDistance(dist);
//                    udtSolution.udtAirportData(lngAirportIndex).lngServiceAreaIndex = lngServiceAreaIndex;
                }
            }
        }
    }


    private AIRPORT_TYPE selectAirport(SERVICEAREAMODEL_TYPE serviceAreaModel) {

        if (serviceAreaModel.getAirportSelectionExponent() == 0) {
            Object key = randomKey(serviceAreaModel.getAirportMap().keySet());
            return serviceAreaModel.getAirportMap().get(key);
        }

        String code = (String) selectByP(serviceAreaModel.getAirportSelectionPMap());
        return serviceAreaModel.getAirportMap().get(code);
    }

    private String randomKey(Set<String> keySet) {
        Object[] a = keySet.toArray();
        Object key = a[new Random().nextInt(a.length)];
        return (String) key;
    }

    public Object selectByP(Map<?, Double> map) {
        Random random = new Random();
        List<?> keys = new ArrayList<>(map.keySet());
        if (keys.size() < 1) {
            return null;
        }
        return keys.get(random.nextInt(keys.size()));
    }

//    public int selectByP(Map<Integer, Double> pValues) {
//
//        int indexLow = 0;
//        int indexMed;
//        int indexHigh;
//
//        Double[] dblPValues = pValues.toArray(new Double[pValues.size()]);
//        double dblP = Math.random();
//        indexHigh = pValues.size() - 1;
//        while (indexLow + 1 < indexHigh) {
//            indexMed = (indexLow + indexHigh) / 2;
//            if (dblPValues[indexMed] < dblP)
//                indexLow = indexMed;
//            else
//                indexHigh = indexMed;
//        }
//
//        int selectByP;
//        if (dblP < dblPValues[indexLow])
//            selectByP = indexLow;
//        else if (dblP < dblPValues[indexHigh])
//            selectByP = indexHigh;
//        else
//            throw new RuntimeException("Select issue: " + dblP);
//
//        return selectByP;
//    }

    private final void evolveSolutions(SERVICEAREAMODEL_TYPE serviceAreaModel) {
        int index;
        int solutionSortIndex;
        int lngSolutionIndex1;
        int lngSolutionIndex2;
        int lngSolutionIndex;
        if (!serviceAreaModel.isBooSolutionSelectionValid()) {
            sortSolutions(serviceAreaModel);
        }

        solutionSortIndex = 0;
        POSSIBLESOLUTION_TYPE[] sortedSol = serviceAreaModel.getSortedSolutions().toArray(
                new POSSIBLESOLUTION_TYPE[serviceAreaModel.getSortedSolutions().size()]
        );
        POSSIBLESOLUTION_TYPE sol1 = null, sol2 = null;

        for (EVOLUTION_TYPE evolution : serviceAreaModel.getEvolutions()) {
            switch (evolution.getEnmApproach()) {
                case EVOLAPP_RetainTop:
                    solutionSortIndex = solutionSortIndex + evolution.getAppliesTo();
                    break;
                case EVOLAPP_RetainRandom:
                    throw new RuntimeException("ModelEnum.EVOLUTIONAPPROACH_ENUM.EVOLAPP_RetainRandom:");
                case EVOLAPP_Mate:
                    for (int i = solutionSortIndex;
                         i <= solutionSortIndex + evolution.getAppliesTo() - 1;
                         i++) {

                        lngSolutionIndex = i;
                        lngSolutionIndex1 = i;

                        sol1 = selectSolution(serviceAreaModel, evolution.getVarParameters().get(0), i);

                        lngSolutionIndex2 = i;
                        sol2 = selectSolution(serviceAreaModel, evolution.getVarParameters().get(1), i);

                        mateSolutions(serviceAreaModel, sol1, sol2, sortedSol[lngSolutionIndex]);
                        if ((sortedSol[lngSolutionIndex].getServiceAreaData().size()
                                < serviceAreaModel.getServiceAreaCount_Min())) {
                            setSolutionServiceAreaCount(serviceAreaModel,
                                    sortedSol[lngSolutionIndex],
                                    serviceAreaModel.getServiceAreaCount_Min());
                        } else if ((serviceAreaModel.getServiceAreaCount_Max()
                                < sortedSol[lngSolutionIndex].getServiceAreaData().size())) {
                            setSolutionServiceAreaCount(serviceAreaModel,
                                    sortedSol[lngSolutionIndex],
                                    serviceAreaModel.getServiceAreaCount_Max());
                        }

                        evaluateSolution(serviceAreaModel, sortedSol[lngSolutionIndex],
                                serviceAreaModel.getMaximumTravelMiles());
                    }

                    break;
                case EVOLAPP_InsertDelete:
                    for (int i = solutionSortIndex;
                         i <= solutionSortIndex + evolution.getAppliesTo() - 1; i++) {

                        lngSolutionIndex = i;

                        sol1 = selectSolution(serviceAreaModel, "RandomByMileage");
                        serviceAreaModel.getSolutions()[lngSolutionIndex] = sol1;
                        index = new Random().nextInt(3);
                        if (index == 0
                                && (serviceAreaModel.getServiceAreaCount_Min()
                                < sol1.getServiceAreaData().size())) {
                            removeSolutionServiceArea(serviceAreaModel.getSolutions()[lngSolutionIndex],
                            new Random().nextInt(serviceAreaModel.getSolutions()[lngSolutionIndex].getServiceAreaData().size()));
                        } else if (index <= 1) {
                            removeSolutionServiceArea(serviceAreaModel.getSolutions()[lngSolutionIndex],
                                    new Random().nextInt(serviceAreaModel.getSolutions()[lngSolutionIndex].getServiceAreaData().size()));
                            setSolutionServiceAreaCount(serviceAreaModel, serviceAreaModel.getSolutions()[lngSolutionIndex],
                                    serviceAreaModel.getSolutions()[lngSolutionIndex].getServiceAreaData().size() + 0);
                        } else if (serviceAreaModel.getSolutions()[lngSolutionIndex].getServiceAreaData().size()
                                < serviceAreaModel.getServiceAreaCount_Max()) {
                            setSolutionServiceAreaCount(serviceAreaModel, serviceAreaModel.getSolutions()[lngSolutionIndex],
                                    serviceAreaModel.getSolutions()[lngSolutionIndex].getServiceAreaData().size()  + 0);
                        } else {
                            removeSolutionServiceArea(serviceAreaModel.getSolutions()[lngSolutionIndex],
                                    new Random().nextInt(serviceAreaModel.getSolutions()[lngSolutionIndex].getServiceAreaData().size()));
                            setSolutionServiceAreaCount(serviceAreaModel, serviceAreaModel.getSolutions()[lngSolutionIndex],
                                    serviceAreaModel.getSolutions()[lngSolutionIndex].getServiceAreaData().size() + 0);
                        }

                        evaluateSolution(serviceAreaModel, serviceAreaModel.getSolutions()[lngSolutionIndex],
                                serviceAreaModel.getMaximumTravelMiles());
                    }

                    break;
                case EVOLAPP_Random:
                    for (int i = solutionSortIndex; i <= solutionSortIndex + evolution.getAppliesTo() - 1;
                         i++) {
                        index = serviceAreaModel.getServiceAreaCount_Min() +
                                new Random().nextInt(1 + (serviceAreaModel.getServiceAreaCount_Max()
                                        - serviceAreaModel.getServiceAreaCount_Min()));
                        randomSolutionServiceArea(serviceAreaModel,
                                serviceAreaModel.getSortedSolutions().get(i), index);
                        evaluateSolution(serviceAreaModel, serviceAreaModel.getSortedSolutions().get(i),
                                serviceAreaModel.getMaximumTravelMiles());
                    }

                    break;
                default:
                    throw new RuntimeException("Unknown Evol type: " + evolution.getEnmApproach());
            }
        }

    }

    private void setSolutionServiceAreaCount(SERVICEAREAMODEL_TYPE serviceAreaModel, POSSIBLESOLUTION_TYPE solution,
                                             int serviceAreaCount) {


        int lngServiceAreaIndex;

        while (solution.getServiceAreaData().size() < serviceAreaCount) {
            insertSolutionServiceArea(serviceAreaModel, solution,
                    randomKey(serviceAreaModel.getAirportMap().keySet()));
        }

        while (serviceAreaCount < solution.getServiceAreaData().size()) {
            removeSolutionServiceArea(solution,
                    randomKey(serviceAreaModel.getAirportMap().keySet()));
        }
    }

    private void removeSolutionServiceArea(POSSIBLESOLUTION_TYPE solution, Object randomKey) {
        solution.getServiceAreaData().remove(randomKey);
    }

    private void mateSolutions(SERVICEAREAMODEL_TYPE udtServiceAreaModel, POSSIBLESOLUTION_TYPE udtSolution1,
                               POSSIBLESOLUTION_TYPE udtSolution2, POSSIBLESOLUTION_TYPE udtResult) {
        int i1 = 0, i2 = 0;

//      udtResult.lngServiceAreaCount = 0
//      If UBound(udtResult.udtServiceAreaData) < udtSolution1.lngServiceAreaCount + udtSolution2.lngServiceAreaCount - 1 Then
//        ReDim udtResult.udtServiceAreaData(0 To udtSolution1.lngServiceAreaCount + udtSolution2.lngServiceAreaCount - 1)
//      End If

        String[] keys1 = (String[]) udtSolution1.getServiceAreaData().keySet().toArray(
                new String[udtSolution1.getServiceAreaData().size()]);
        String[] keys2 = (String[]) udtSolution2.getServiceAreaData().keySet().toArray(
                new String[udtSolution2.getServiceAreaData().size()]);

        while (i1 < keys1.length && i2 < keys2.length
                && udtResult.getServiceAreaData().size() < udtServiceAreaModel.getServiceAreaCount_Max()) {

            String code1 = keys1[i1];
            String code2 = keys2[i2];

            SERVICEAREADATA_TYPE sol1 = udtSolution1.getServiceAreaData().get(code1);
            SERVICEAREADATA_TYPE sol2 = udtSolution1.getServiceAreaData().get(code2);

            if (code1.equals(code2)) {
                udtResult.getServiceAreaData().put(code1, sol1);
                i1++;
                i2++;
            } else if (code1.compareTo(code2) < 0) {
                if (Math.random() < 0.5) {
                    udtResult.getServiceAreaData().put(code1, sol1);
                }
                i1++;
            } else {
                if (Math.random() < 0.5) {
                    udtResult.getServiceAreaData().put(code1, sol2);
                }
                i2++;
            }
        }

//        String[] keysR = (String[]) udtResult.getServiceAreaData().keySet().toArray();
        try {
            int size1 = udtSolution1.getServiceAreaData().size();
            for (int i = i1; i < keys1.length - 1; i++) {
                if (udtResult.getServiceAreaData().size() == udtServiceAreaModel.getServiceAreaCount_Max()) {
                    break;
                }
                if (Math.random() < 0.5) {
                    if (udtSolution1.getServiceAreaData().get(keys1[i]) == null)
                        continue;
                    udtResult.getServiceAreaData().put(keys1[i], udtSolution1.getServiceAreaData().get(keys1[i]));
                }
            }
            int size2 = udtSolution2.getServiceAreaData().size();
            for (int i = i2; i < keys2.length - 1; i++) {
                if (udtResult.getServiceAreaData().size() == udtServiceAreaModel.getServiceAreaCount_Max()) {
                    break;
                }
                if (Math.random() < 0.5) {
                    if (udtSolution2.getServiceAreaData().get(keys2[i]) == null)
                        continue;
                    udtResult.getServiceAreaData().put(keys2[i], udtSolution2.getServiceAreaData().get(keys2[i]));
                }
            }
        } catch (ArrayIndexOutOfBoundsException e) {
            e.printStackTrace();
        }
        udtResult.setResultsAreValid(false);
    }

    private final POSSIBLESOLUTION_TYPE selectSolution(SERVICEAREAMODEL_TYPE serviceAreaModel, String selectionMethod) {
        return selectSolution(serviceAreaModel, selectionMethod, -1);
    }

    private final POSSIBLESOLUTION_TYPE selectSolution(SERVICEAREAMODEL_TYPE serviceAreaModel, String selectionMethod,
                                                       int exclusion) {
        int selectSolution = exclusion;

        if (selectionMethod.startsWith("Top")) {
            if (!serviceAreaModel.isBooSolutionSelectionValid()) {
                sortSolutions(serviceAreaModel);
            }

            while (exclusion == selectSolution) {
                selectSolution = new Random().nextInt(Integer.parseInt(selectionMethod.substring(4)));
            }
        } else if (selectionMethod.equals("Random")) {
            while (exclusion == selectSolution) {
                selectSolution = new Random().nextInt(serviceAreaModel.getCommunitySize());
            }
        } else if (selectionMethod.equals("RandomByMileage")) {
            if (!serviceAreaModel.isBooSolutionSelectionValid() ||
                    serviceAreaModel.getSolutionSelectionP().isEmpty()) {
                computeSolutionP(serviceAreaModel);
            }

            selectSolution = (int) selectByP(serviceAreaModel.getSolutionSelectionP());
            return serviceAreaModel.getSolutions()[selectSolution];
        } else {
            throw new RuntimeException("Selection method not know: " + selectionMethod);
        }

        return serviceAreaModel.getSortedSolutions().get(selectSolution);
    }


    private final void sortSolutions(SERVICEAREAMODEL_TYPE serviceAreaModel) {
        int index;
        if (serviceAreaModel.isBooSolutionSelectionValid()) {
            return;
        }

        List<POSSIBLESOLUTION_TYPE> sorted = new ArrayList<>();
        sorted.addAll(Arrays.asList(serviceAreaModel.getSolutions()));
        sorted.sort((POSSIBLESOLUTION_TYPE a, POSSIBLESOLUTION_TYPE b) ->
                Double.compare(a.getOptimizeOn(), b.getOptimizeOn()));

        serviceAreaModel.setSortedSolutions(sorted);
//
//        for (index = 0; index < serviceAreaModel.getCommunitySize(); index++) {
//            serviceAreaModel.getSortValues().add(index, serviceAreaModel.getSolutions()[index].getOptimizeOn());
//        }
//
//        sortValues(serviceAreaModel.dblSortValues, serviceAreaModel.lngCommunitySize,
//
//                serviceAreaModel.lngSolutionSortIndexes);
        serviceAreaModel.setBooSolutionSelectionValid(true);
    }


    private void computeSolutionP(SERVICEAREAMODEL_TYPE serviceAreaModel) {
        double travelValue, sum = 0;

        if (!serviceAreaModel.isSolutionSortIsValid()) {
            sortSolutions(serviceAreaModel);
        }

        travelValue = 4 * serviceAreaModel.getSortedSolutions().get(0).getTravelMiles();

        for (int i = 0; i < serviceAreaModel.getCommunitySize(); i++) {

            if (travelValue + serviceAreaModel.getSolutions()[i].getTravelMiles() == 0) {
                serviceAreaModel.getSolutionSelectionP().put(i, 1.0);
            } else {
                serviceAreaModel.getSolutionSelectionP().put(i,
                        1.0 / Math.pow(travelValue + serviceAreaModel.getSolutions()[i].getTravelMiles(),
                                serviceAreaModel.getSolutionSelectionExponent()));
            }
            sum += serviceAreaModel.getSolutionSelectionP().get(i);
        }

        serviceAreaModel.getSolutionSelectionP().put(0, serviceAreaModel.getSolutionSelectionP().get(0) / sum);
        for (int i = 1; i < serviceAreaModel.getCommunitySize(); i++) {
            serviceAreaModel.getSolutionSelectionP().put(i,
                    serviceAreaModel.getSolutionSelectionP().get(i - 1)
                            + serviceAreaModel.getSolutionSelectionP().get(i) / sum
            );
        }
    }

}