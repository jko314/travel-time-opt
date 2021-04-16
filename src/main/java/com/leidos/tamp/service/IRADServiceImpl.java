package com.leidos.tamp.service;

import com.leidos.tamp.db.IRADServiceDB;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoDatabase;
import org.bson.Document;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import com.leidos.tamp.beans.*;

import java.util.ArrayList;
import java.util.List;

@Service
public class IRADServiceImpl implements IRADService {
	
    Logger logger = LoggerFactory.getLogger(IRADServiceImpl.class);
	
	MongoDatabase iraddb;

	MongoCollection mongoCollection;

	public IRADServiceImpl () {
		iraddb = IRADServiceDB.getIRADDB();
	}

	@Override
	public List<Airport> getAirports() {
		logger.debug("In getAirports");
		getAirportsCollection();
		Document query = new Document();
		List<Airport> airportsList = (List<Airport>) mongoCollection.find(query, Airport.class).into(
				new ArrayList<Airport>());
		logger.debug("fetched airports:"+airportsList.size());
		return airportsList;
	}
	
	@Override
	public Airport getAirport(String code) {
		Airport airport=null;
		getAirportsCollection();
		Document query = new Document();
		query.append("Code", code);
		List<Airport> airportsList = (List<Airport>) mongoCollection.find(query, Airport.class).into(
				new ArrayList<Airport>());
		logger.debug("Code:"+code +" Airports Array List size:"+airportsList.size());
		if(airportsList.size()>0) {
			airport=airportsList.get(0);
			logger.debug("Airport code:"+airport.getCode()+" City:"+airport.getCity());
		}
		return airport;
	}

/*	@Override
	public void updateAirport(Airport airport) {
		getAirportsCollection();
		Document query = new Document("Code", airport.getCode());
		Document newDoc = new Document();
		newDoc.append("Code", airport.getCode());
		newDoc.append("Cat", airport.getCat()).append("City", airport.getCity()).append("State", airport.getState()).
				append("Name", airport.getName()).append("Lattitude", airport.getLatitude()).
				append("Longitude", airport.getLongitude()).append("Op_Start", airport.getOp_start()).
				append("Op_Hrs",airport.getOp_hrs()).append("Timezone", airport.getTime_zone());
		Document update = new Document();
        update.append("$set", newDoc);
        mongoCollection.updateOne(query, update);		
	}*/
	
	@Override
	public void updateAirport(Airport airport) {
		deleteAirport(airport.getCode());
		createAirport(airport);
	}

	@Override
	public void createAirport(Airport airport) {
		getAirportsCollection();
		mongoCollection.insertOne(airport);		
	}

	@Override
	public void deleteAirport(String code) {
		getAirportsCollection();
		Document deleteDoc = new Document();
		deleteDoc.append("Code", code);
		mongoCollection.deleteMany(deleteDoc);
	}

	@Override
	public List<AirportEquipment> getAirportEquipments() {
		logger.debug("In getAirportEquipments");
		getAirportEquipmentCollection();
		Document query = new Document();
		List<AirportEquipment> airportEquipmentList = (List<AirportEquipment>) mongoCollection.find(query, AirportEquipment.class).into(
				new ArrayList<AirportEquipment>());
		logger.debug("fetched airportEquipment:"+airportEquipmentList.size());
		return airportEquipmentList;
	}

	@Override
	public AirportEquipment getAirportEquipment(String airport) {
		AirportEquipment airportEquipment=null;
		getAirportEquipmentCollection();
		Document query = new Document();
		query.append("Airport", airport);
		List<AirportEquipment> airportEquipmentList = (List<AirportEquipment>) mongoCollection.find(query, AirportEquipment.class).into(
				new ArrayList<AirportEquipment>());
		logger.debug("Airport Equipment:"+airportEquipment.getAirport() +" Airport Equipment Array List size:"+airportEquipmentList.size());
		if(airportEquipmentList.size()>0) {
			airportEquipment=airportEquipmentList.get(0);
			logger.debug("airportEquipment airport:"+airportEquipment.getAirport());
		}
		return airportEquipment;
	}

	@Override
	public void updateAirportEquipment(AirportEquipment airportEquipment) {
		getAirportEquipmentCollection();
		Document query = new Document("Airport", airportEquipment.getAirport());
		Document newDoc = new Document();
		deleteAirportEquipment(airportEquipment.getAirport());
		createAirportEquipment(airportEquipment);
	}

	@Override
	public void deleteAirportEquipment(String airport) {
		getAirportEquipmentCollection();
		Document deleteDoc = new Document();
		deleteDoc.append("AirportEquipment", airport);
		mongoCollection.deleteMany(deleteDoc);
	}
	
	@Override
	public void createAirportEquipment(AirportEquipment airportEquipment) {
		getAirportEquipmentCollection();
		mongoCollection.insertOne(airportEquipment);				
	}

	@Override
	public List<AirportServiceArea> getAirportServiceArea() {
		logger.debug("In getAirportServiceArea");
		getAirportServiceAreaCollection();
		Document query = new Document();
		List<AirportServiceArea> airportServiceAreaList = (List<AirportServiceArea>) mongoCollection.find(query, AirportServiceArea.class).into(
				new ArrayList<AirportServiceArea>());
		logger.debug("fetched airportServiceArea:"+airportServiceAreaList.size());
		return airportServiceAreaList;
	}

	@Override
	public AirportServiceArea getAirportServiceArea(String airportCode) {
		AirportServiceArea airportServiceArea=null;
		getAirportServiceAreaCollection();
		Document query = new Document();
		query.append("AirportCode", airportCode);
		List<AirportServiceArea> airportServiceAreaList = (List<AirportServiceArea>) mongoCollection.find(query, AirportServiceArea.class).into(
				new ArrayList<AirportServiceArea>());
		logger.debug("Airport Service Area:"+airportServiceArea.getAirportCode() +" Airport Service Area Array List size:"+airportServiceAreaList.size());
		if(airportServiceAreaList.size()>0) {
			airportServiceArea=airportServiceAreaList.get(0);
			logger.debug("airportServiceArea airport:"+airportServiceArea.getAirportCode());
		}
		return airportServiceArea;
	}

	@Override
	public void updateAirportServiceArea(AirportServiceArea airportServiceArea) {
		deleteAirportServiceArea(airportServiceArea.getAirportCode());
		createAirportServiceArea(airportServiceArea);
	}
	
	@Override
	public void deleteAirportServiceArea(String airportCode) {
		getAirportServiceAreaCollection();
		Document deleteDoc = new Document();
		deleteDoc.append("AirportServiceArea", airportCode);
		mongoCollection.deleteMany(deleteDoc);
	}

	@Override
	public void createAirportServiceArea(AirportServiceArea airportServiceArea) {
		getAirportServiceAreaCollection();
		mongoCollection.insertOne(airportServiceArea);				
	}

	@Override
	public List<AirportTime> getAirportTime() {
		logger.debug("In getAirportTime");
		getAirportTimeCollection();
		Document query = new Document();
		List<AirportTime> airportTimeList = (List<AirportTime>) mongoCollection.find(query, AirportTime.class).into(
				new ArrayList<AirportTime>());
		logger.debug("fetched airportTime:"+airportTimeList.size());
		return airportTimeList;
	}

	@Override
	public AirportTime getAirportTime(String facilityId) {
		AirportTime airportTime=null;
		getAirportTimeCollection();
		Document query = new Document();
		query.append("facilityId", facilityId);
		List<AirportTime> airportTimeList = (List<AirportTime>) mongoCollection.find(query, AirportTime.class).into(
				new ArrayList<AirportTime>());
		logger.debug("Airport Time:"+airportTime.getFacilityId() +" Airport Time Array List size:"+airportTimeList.size());
		if(airportTimeList.size()>0) {
			airportTime=airportTimeList.get(0);
			logger.debug("airportTime facilityId:"+airportTime.getFacilityId());
		}
		return airportTime;
	}

	@Override
	public void updateAirportTime(AirportTime airportTime) {
		deleteAirportServiceArea(airportTime.getFacilityId());
		createAirportTime(airportTime);
	}

	@Override
	public void deleteAirportTime(String facilityId) {
		getAirportTimeCollection();
		Document deleteDoc = new Document();
		deleteDoc.append("FacilityId", facilityId);
		mongoCollection.deleteMany(deleteDoc);
	}

	@Override
	public void createAirportTime(AirportTime airportTime) {
		getAirportTimeCollection();
		mongoCollection.insertOne(airportTime);				
	}

	@Override
	public List<AirportWithCargo> getAirportWithCargo() {
		logger.debug("In getAirportWithCargo");
		getAirportWithCargoCollection();
		Document query = new Document();
		List<AirportWithCargo> airportWithCargoList = (List<AirportWithCargo>) mongoCollection.find(query, AirportWithCargo.class).into(
				new ArrayList<AirportWithCargo>());
		logger.debug("fetched airportWithCargo:"+airportWithCargoList.size());
		return airportWithCargoList;
	}

	@Override
	public AirportWithCargo getAirportWithCargo(String code) {
		AirportWithCargo airportWithCargo=null;
		getAirportWithCargoCollection();
		Document query = new Document();
		query.append("Code", code);
		List<AirportWithCargo> airportCargoList = (List<AirportWithCargo>) mongoCollection.find(query, AirportWithCargo.class).into(
				new ArrayList<AirportWithCargo>());
		logger.debug("Airport With Cargo:"+airportWithCargo.getCode() +" Airport With Cargo Array List size:"+airportCargoList.size());
		if(airportCargoList.size()>0) {
			airportWithCargo=airportCargoList.get(0);
			logger.debug("airportWithCargo code:"+airportCargoList.size());
		}
		return airportWithCargo;
	}

	@Override
	public void updateAirportWithCargo(AirportWithCargo airportWithCargo) {
		deleteAirportServiceArea(airportWithCargo.getCode());
		createAirportWithCargo(airportWithCargo);		
	}
	
	@Override
	public void deleteAirportWithCargo(String code) {
		getAirportWithCargoCollection();
		Document deleteDoc = new Document();
		deleteDoc.append("Code", code);
		mongoCollection.deleteMany(deleteDoc);
	}

	@Override
	public void createAirportWithCargo(AirportWithCargo airportWithCargo) {
		getAirportWithCargoCollection();
		mongoCollection.insertOne(airportWithCargo);				
	}

	@Override
	public List<CMRate> getCMRate() {
		logger.debug("In getCMRate");
		getCMRateCollection();
		Document query = new Document();
		List<CMRate> cmRateList = (List<CMRate>) mongoCollection.find(query, CMRate.class).into(
				new ArrayList<CMRate>());
		logger.debug("fetched cmRate:"+cmRateList.size());
		return cmRateList;
	}

	@Override
	public CMRate getCMRate(String modelNum) {
		CMRate cmRate=null;
		getCMRateCollection();
		Document query = new Document();
		query.append("ModelNum", modelNum);
		List<CMRate> cmRateList = (List<CMRate>) mongoCollection.find(query, CMRate.class).into(
				new ArrayList<CMRate>());
		logger.debug("CMRate:"+cmRate.getModelNum() +" CMRate Array List size:"+cmRateList.size());
		if(cmRateList.size()>0) {
			cmRate=cmRateList.get(0);
			logger.debug("cmRate code:"+cmRateList.size());
		}
		return cmRate;
	}

	@Override
	public void updateCMRate(CMRate cmRate) {
		deleteCMRate(cmRate.getModelNum());
		createCMRate(cmRate);		
	}

	@Override
	public void deleteCMRate(String modelNum) {
		getCMRateCollection();
		Document deleteDoc = new Document();
		deleteDoc.append("ModelNum", modelNum);
		mongoCollection.deleteMany(deleteDoc);
	}

	@Override
	public void createCMRate(CMRate cmRate) {
		getCMRateCollection();
		mongoCollection.insertOne(cmRate);				
	}

	@Override
	public List<Equipment> geEquipment() {
		logger.debug("In getEquipment");
		getEquipmentCollection();
		Document query = new Document();
		List<Equipment> equipmentList = (List<Equipment>) mongoCollection.find(query, Equipment.class).into(
				new ArrayList<Equipment>());
		logger.debug("fetched equipment:"+equipmentList.size());
		return equipmentList;
	}

	@Override
	public Equipment getEquipment(String makeModel) {
		Equipment equipment=null;
		getEquipmentCollection();
		Document query = new Document();
		query.append("MakeModel", makeModel);
		List<Equipment> equipmentList = (List<Equipment>) mongoCollection.find(query, Equipment.class).into(
				new ArrayList<Equipment>());
		logger.debug("equipment:"+equipment.getMakeModel() +" Equipment Array List size:"+equipmentList.size());
		if(equipmentList.size()>0) {
			equipment=equipmentList.get(0);
			logger.debug("Equipment makeModel:"+equipmentList.size());
		}
		return equipment;
	}

	@Override
	public void updateEquipment(Equipment equipment) {
		deleteEquipment(equipment.getMakeModel());
		createEquipment(equipment);		
	}
	
	@Override
	public void deleteEquipment(String modelNum) {
		getEquipmentCollection();
		Document deleteDoc = new Document();
		deleteDoc.append("ModelNum", modelNum);
		mongoCollection.deleteMany(deleteDoc);
	}

	@Override
	public void createEquipment(Equipment equipment) {
		getEquipmentCollection();
		mongoCollection.insertOne(equipment);				
	}

	@Override
	public List<EquipmentCM> getEquipmentCM() {
		logger.debug("In getEquipmentCM");
		getEquipmentCMCollection();
		Document query = new Document();
		List<EquipmentCM> equipmentCMList = (List<EquipmentCM>) mongoCollection.find(query, EquipmentCM.class).into(
				new ArrayList<EquipmentCM>());
		logger.debug("fetched equipment:"+equipmentCMList.size());
		return equipmentCMList;
	}

	@Override
	public EquipmentCM getEquipmentCM(String makeModel) {
		EquipmentCM equipmentCM=null;
		getEquipmentCMCollection();
		Document query = new Document();
		query.append("MakeModel", makeModel);
		List<EquipmentCM> equipmentCMList = (List<EquipmentCM>) mongoCollection.find(query, EquipmentCM.class).into(
				new ArrayList<EquipmentCM>());
		logger.debug("equipmentCM:"+equipmentCM.getMakeModel() +" EquipmentCM Array List size:"+equipmentCMList.size());
		if(equipmentCMList.size()>0) {
			equipmentCM=equipmentCMList.get(0);
			logger.debug("EquipmentCM makeModel:"+equipmentCMList.size());
		}
		return equipmentCM;
	}

	@Override
	public void updateEquipmentCM(EquipmentCM equipmentCM) {
		deleteEquipmentCM(equipmentCM.getMakeModel());
		createEquipmentCM(equipmentCM);		
	}

	@Override
	public void deleteEquipmentCM(String makeModel) {
		getEquipmentCMCollection();
		Document deleteDoc = new Document();
		deleteDoc.append("MakeModel", makeModel);
		mongoCollection.deleteMany(deleteDoc);
	}

	@Override
	public void createEquipmentCM(EquipmentCM equipmentCM) {
		getEquipmentCMCollection();
		mongoCollection.insertOne(equipmentCM);				
	}

	@Override
	public List<EquipmentModel> getEquipmentModel() {
		logger.debug("In getEquipmentModel");
		getEquipmentModelCollection();
		Document query = new Document();
		List<EquipmentModel> equipmentModelList = (List<EquipmentModel>) mongoCollection.find(query, EquipmentModel.class).into(
				new ArrayList<EquipmentModel>());
		logger.debug("fetched equipment Model:"+equipmentModelList.size());
		return equipmentModelList;
	}

	@Override
	public EquipmentModel getEquipmentModel(String equipmentType) {
		EquipmentModel equipmentModel=null;
		getEquipmentModelCollection();
		Document query = new Document();
		query.append("EquipmentType", equipmentType);
		List<EquipmentModel> equipmentModelList = (List<EquipmentModel>) mongoCollection.find(query, EquipmentModel.class).into(
				new ArrayList<EquipmentModel>());
		logger.debug("equipmentModel:"+equipmentModel.getEquipmentType() +" EquipmentModel Array List size:"+equipmentModelList.size());
		if(equipmentModelList.size()>0) {
			equipmentModel=equipmentModelList.get(0);
			logger.debug("EquipmentModel equipmentType:"+equipmentModel.getEquipmentType());
		}
		return equipmentModel;
	}

	@Override
	public void updateEquipmentModel(EquipmentModel equipmentModel) {
		deleteEquipmentModel(equipmentModel.getEquipmentType());
		createEquipmentModel(equipmentModel);		
	}

	@Override
	public void deleteEquipmentModel(String equipmentType) {
		getEquipmentModelCollection();
		Document deleteDoc = new Document();
		deleteDoc.append("EquipmentType", equipmentType);
		mongoCollection.deleteMany(deleteDoc);
	}
	
	@Override
	public void createEquipmentModel(EquipmentModel equipmentModel) {
		getEquipmentModelCollection();
		mongoCollection.insertOne(equipmentModel);				
	}

	@Override
	public List<EquipmentPM> getEquipmentPM() {
		logger.debug("In getEquipmentPM");
		getEquipmentPMCollection();
		Document query = new Document();
		List<EquipmentPM> equipmentPMList = (List<EquipmentPM>) mongoCollection.find(query, EquipmentPM.class).into(
				new ArrayList<EquipmentPM>());
		logger.debug("fetched equipment PM:"+equipmentPMList.size());
		return equipmentPMList;
	}

	@Override
	public EquipmentPM getEquipmentPM(String makeModel) {
		EquipmentPM equipmentPM=null;
		getEquipmentPMCollection();
		Document query = new Document();
		query.append("MakeModel", makeModel);
		List<EquipmentPM> equipmentPMList = (List<EquipmentPM>) mongoCollection.find(query, EquipmentPM.class).into(
				new ArrayList<EquipmentPM>());
		logger.debug("equipmentPM:"+equipmentPM.getMakeModel() +" EquipmentPM Array List size:"+equipmentPMList.size());
		if(equipmentPMList.size()>0) {
			equipmentPM=equipmentPMList.get(0);
			logger.debug("EquipmentPM makeModel:"+equipmentPM.getMakeModel());
		}
		return equipmentPM;
	}

	@Override
	public void updateEquipmentPM(EquipmentPM equipmentPM) {
		deleteEquipmentPM(equipmentPM.getMakeModel());
		createEquipmentPM(equipmentPM);		
	}
	
	@Override
	public void deleteEquipmentPM(String makeModel) {
		getEquipmentPMCollection();
		Document deleteDoc = new Document();
		deleteDoc.append("MakeModel", makeModel);
		mongoCollection.deleteMany(deleteDoc);
	}
	
	@Override
	public void createEquipmentPM(EquipmentPM equipmentPM) {
		getEquipmentPMCollection();
		mongoCollection.insertOne(equipmentPM);				
	}

	@Override
	public List<LaborRule> getLaborRule() {
		logger.debug("In getLaborRule");
		getLaborRuleCollection();
		Document query = new Document();
		List<LaborRule> laborRuleList = (List<LaborRule>) mongoCollection.find(query, LaborRule.class).into(
				new ArrayList<LaborRule>());
		logger.debug("fetched labor Rule:"+laborRuleList.size());
		return laborRuleList;
	}

	@Override
	public LaborRule getLaborRule(String laborRuleId) {
		LaborRule laborRule=null;
		getLaborRuleCollection();
		Document query = new Document();
		query.append("LaborRuleId", laborRuleId);
		List<LaborRule> laborRuleList = (List<LaborRule>) mongoCollection.find(query, LaborRule.class).into(
				new ArrayList<LaborRule>());
		logger.debug("laborRule:"+laborRule.getLaborRuleId() +" LaborRule Array List size:"+laborRuleList.size());
		if(laborRuleList.size()>0) {
			laborRule=laborRuleList.get(0);
			logger.debug("laborRule laborRuleId:"+laborRule.getLaborRuleId());
		}
		return laborRule;
	}

	@Override
	public void updateLaborRule(LaborRule laborRule) {
		deleteLaborRule(laborRule.getLaborRuleId());
		createLaborRule(laborRule);		
	}
	
	@Override
	public void deleteLaborRule(String laborRuleId) {
		getLaborRuleCollection();
		Document deleteDoc = new Document();
		deleteDoc.append("LaborRule", laborRuleId);
		mongoCollection.deleteMany(deleteDoc);
	}
	
	@Override
	public void createLaborRule(LaborRule laborRule) {
		getLaborRuleCollection();
		mongoCollection.insertOne(laborRule);				
	}

	@Override
	public List<ServiceArea> getServiceArea() {
		logger.debug("In getServiceArea");
		getServiceAreaCollection();
		Document query = new Document();
		List<ServiceArea> serviceAreaList = (List<ServiceArea>) mongoCollection.find(query, ServiceArea.class).into(
				new ArrayList<ServiceArea>());
		logger.debug("fetched service area:"+serviceAreaList.size());
		return serviceAreaList;
	}

	@Override
	public ServiceArea getServiceArea(String name) {
		ServiceArea serviceArea=null;
		getServiceAreaCollection();
		Document query = new Document();
		query.append("Name", name);
		List<ServiceArea> serviceAreaList = (List<ServiceArea>) mongoCollection.find(query, ServiceArea.class).into(
				new ArrayList<ServiceArea>());
		logger.debug("laborRule:"+serviceArea.getName() +" LaborRule Array List size:"+serviceAreaList.size());
		if(serviceAreaList.size()>0) {
			serviceArea=serviceAreaList.get(0);
			logger.debug("laborRule laborRuleId:"+serviceArea.getName());
		}
		return serviceArea;
	}

	@Override
	public void updateServiceArea(ServiceArea serviceArea) {
		deleteServiceArea(serviceArea.getName());
		createServiceArea(serviceArea);		
	}

	@Override
	public void deleteServiceArea(String name) {
		getServiceAreaCollection();
		Document deleteDoc = new Document();
		deleteDoc.append("Name", name);
		mongoCollection.deleteMany(deleteDoc);
	}
	
	@Override
	public void createServiceArea(ServiceArea serviceArea) {
		getServiceAreaCollection();
		mongoCollection.insertOne(serviceArea);				
	}

	@Override
	public List<TravelCost> getTravelCost() {
		logger.debug("In getTravelCost");
		getTravelCostCollection();
		Document query = new Document();
		List<TravelCost> travelCostList = (List<TravelCost>) mongoCollection.find(query, TravelCost.class).into(
				new ArrayList<TravelCost>());
		logger.debug("fetched travel cost:"+travelCostList.size());
		return travelCostList;
	}

	@Override
	public TravelCost getTravelCost(String facility) {
		TravelCost travelCost=null;
		getTravelCostCollection();
		Document query = new Document();
		query.append("Facility", facility);
		List<TravelCost> travelCostList = (List<TravelCost>) mongoCollection.find(query, TravelCost.class).into(
				new ArrayList<TravelCost>());
		logger.debug("laborRule:"+travelCost.getFacility() +" LaborRule Array List size:"+travelCostList.size());
		if(travelCostList.size()>0) {
			travelCost=travelCostList.get(0);
			logger.debug("TravelCost facility:"+travelCost.getFacility());
		}
		return travelCost;
	}

	@Override
	public void updateTravelCost(TravelCost travelCost) {
		deleteTravelCost(travelCost.getFacility());
		createTravelCost(travelCost);		
	}

	@Override
	public void deleteTravelCost(String name) {
		getTravelCostCollection();
		Document deleteDoc = new Document();
		deleteDoc.append("Name", name);
		mongoCollection.deleteMany(deleteDoc);
	}
	
	@Override
	public void createTravelCost(TravelCost travelCost) {
		getTravelCostCollection();
		mongoCollection.insertOne(travelCost);				
	}

	@Override
	public List<TravelMode> getTravelMode() {
		logger.debug("In getTravelMode");
		getTravelModeCollection();
		Document query = new Document();
		List<TravelMode> travelModeList = (List<TravelMode>) mongoCollection.find(query, TravelMode.class).into(
				new ArrayList<TravelMode>());
		logger.debug("fetched travel mode:"+travelModeList.size());
		return travelModeList;
	}

	@Override
	public TravelMode getTravelMode(String origin) {
		TravelMode travelMode=null;
		getTravelModeCollection();
		Document query = new Document();
		query.append("Origin", origin);
		List<TravelMode> travelModeList = (List<TravelMode>) mongoCollection.find(query, TravelMode.class).into(
				new ArrayList<TravelMode>());
		logger.debug("travelMode:"+travelMode.getOrigin() +" TravelMode Array List size:"+travelModeList.size());
		if(travelModeList.size()>0) {
			travelMode=travelModeList.get(0);
			logger.debug("TravelMode origin:"+travelMode.getOrigin());
		}
		return travelMode;
	}

	@Override
	public void updateTravelMode(TravelMode travelMode) {
		deleteTravelMode(travelMode.getOrigin());
		createTravelMode(travelMode);		
	}
	
	@Override
	public void deleteTravelMode(String name) {
		getTravelModeCollection();
		Document deleteDoc = new Document();
		deleteDoc.append("Name", name);
		mongoCollection.deleteMany(deleteDoc);
	}

	@Override
	public void createTravelMode(TravelMode travelMode) {
		getTravelModeCollection();
		mongoCollection.insertOne(travelMode);				
	}
	
	@Override
	public List<TravelRule> getTravelRule() {
		logger.debug("In getTravelRule");
		getTravelRuleCollection();
		Document query = new Document();
		List<TravelRule> travelRuleList = (List<TravelRule>) mongoCollection.find(query, TravelRule.class).into(
				new ArrayList<TravelRule>());
		logger.debug("fetched travel cost:"+travelRuleList.size());
		return travelRuleList;
	}

	@Override
	public TravelRule getTravelRule(String travelRuleId) {
		TravelRule travelRule=null;
		getTravelRuleCollection();
		Document query = new Document();
		query.append("TravelRuleId", travelRuleId);
		List<TravelRule> travelRuleList = (List<TravelRule>) mongoCollection.find(query, TravelRule.class).into(
				new ArrayList<TravelRule>());
		logger.debug("travelRule:"+travelRule.getTravelRuleId() +" TravelRule Array List size:"+travelRuleList.size());
		if(travelRuleList.size()>0) {
			travelRule=travelRuleList.get(0);
			logger.debug("TravelRule TravelRuleId:"+travelRule.getTravelRuleId());
		}
		return travelRule;
	}

	@Override
	public void updateTravelRule(TravelRule travelRule) {
		deleteTravelRule(travelRule.getTravelRuleId());
		createTravelRule(travelRule);		
	}
	
	@Override
	public void deleteTravelRule(String name) {
		getTravelRuleCollection();
		Document deleteDoc = new Document();
		deleteDoc.append("Name", name);
		mongoCollection.deleteMany(deleteDoc);
	}

	@Override
	public void createTravelRule(TravelRule travelRule) {
		getTravelRuleCollection();
		mongoCollection.insertOne(travelRule);				
	}
	
	public void getAirportsCollection() {

		mongoCollection = iraddb.getCollection("airport", Airport.class);
	}

	public void getAirportEquipmentCollection() {

		mongoCollection = iraddb.getCollection("airportEquipment", AirportEquipment.class);
	}

	public void getAirportServiceAreaCollection() {

		mongoCollection = iraddb.getCollection("airportServiceArea", AirportServiceArea.class);
	}
	
	public void getAirportTimeCollection() {

		mongoCollection = iraddb.getCollection("airportTime", AirportTime.class);
	}
	
	public void getAirportWithCargoCollection() {

		mongoCollection = iraddb.getCollection("airportWithCargo", AirportWithCargo.class);
	}

	public void getCMRateCollection() {

		mongoCollection = iraddb.getCollection("cmRate", CMRate.class);
	}

	public void getEquipmentCollection() {

		mongoCollection = iraddb.getCollection("equipment", Equipment.class);
	}

	public void getEquipmentCMCollection() {

		mongoCollection = iraddb.getCollection("equipmentCM", EquipmentCM.class);
	}

	public void getEquipmentModelCollection() {

		mongoCollection = iraddb.getCollection("equipmentModel", EquipmentModel.class);
	}

	public void getEquipmentPMCollection() {

		mongoCollection = iraddb.getCollection("equipmentPM", EquipmentPM.class);
	}
	
	public void getLaborRuleCollection() {

		mongoCollection = iraddb.getCollection("laborRule", LaborRule.class);
	}
	
	public void getServiceAreaCollection() {

		mongoCollection = iraddb.getCollection("serviceArea", ServiceArea.class);
	}

	public void getServiceAreaModelDataCollection() {

		mongoCollection = iraddb.getCollection("serviceAreaModelData", ServiceAreaModelData.class);
	}

	public void getTravelCostCollection() {

		mongoCollection = iraddb.getCollection("travelCost", TravelCost.class);
	}

	public void getTravelModeCollection() {

		mongoCollection = iraddb.getCollection("travelMode", TravelMode.class);
	}
	
	public void getTravelRuleCollection() {

		mongoCollection = iraddb.getCollection("travelRule", TravelRule.class);
	}

}
