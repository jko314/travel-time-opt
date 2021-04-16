package com.leidos.tamp.service;

import java.util.List;
import com.leidos.tamp.beans.*;

public interface IRADService {

	public List<Airport> getAirports();
	public Airport getAirport(String Code);
	public void updateAirport(Airport airport);
	public void createAirport(Airport airport);
	public void deleteAirport(String code);
	
	public List<AirportEquipment> getAirportEquipments();
	public AirportEquipment getAirportEquipment(String airport);
	public void updateAirportEquipment(AirportEquipment airportEquipment);
	public void createAirportEquipment(AirportEquipment airportEquipment);
	void deleteAirportEquipment(String airport);
	
	public List<AirportServiceArea> getAirportServiceArea();
	public AirportServiceArea getAirportServiceArea(String airportCode);
	public void updateAirportServiceArea(AirportServiceArea airportServiceArea);
	public void createAirportServiceArea(AirportServiceArea airportServiceArea);
	void deleteAirportServiceArea(String airport);

	public List<AirportTime> getAirportTime();
	public AirportTime getAirportTime(String facilityId);
	public void updateAirportTime(AirportTime airportTime);
	public void createAirportTime(AirportTime airportTime);
	void deleteAirportTime(String facilityId);

	public List<AirportWithCargo> getAirportWithCargo();
	public AirportWithCargo getAirportWithCargo(String code);
	public void updateAirportWithCargo(AirportWithCargo airportWithCargo);
	public void createAirportWithCargo(AirportWithCargo airportWithCargo);
	void deleteAirportWithCargo(String airportWithCargo);
	
	public List<CMRate> getCMRate();
	public CMRate getCMRate(String modelNum);
	public void updateCMRate(CMRate CMRate);
	public void createCMRate(CMRate cmRate);
	void deleteCMRate(String name);

	public List<Equipment> geEquipment();
	public Equipment getEquipment(String modelNum);
	public void updateEquipment(Equipment equipment);
	public void createEquipment(Equipment equipment);
	void deleteEquipment(String equipmentId);
	
	public List<EquipmentCM> getEquipmentCM();
	public EquipmentCM getEquipmentCM(String makeModel);
	public void updateEquipmentCM(EquipmentCM equipmentCM);
	public void createEquipmentCM(EquipmentCM equipmentCM);
	void deleteEquipmentCM(String name);

	public List<EquipmentModel> getEquipmentModel();
	public EquipmentModel getEquipmentModel(String equipmentType);
	public void updateEquipmentModel(EquipmentModel equipmentModel);
	public void createEquipmentModel(EquipmentModel equipmentModel);
	void deleteEquipmentModel(String model);
	
	public List<EquipmentPM> getEquipmentPM();
	public EquipmentPM getEquipmentPM(String makeModel);
	public void updateEquipmentPM(EquipmentPM equipmentPM);
	public void createEquipmentPM(EquipmentPM equipmentPM);
	void deleteEquipmentPM(String makeModel);

	public List<LaborRule> getLaborRule();
	public LaborRule getLaborRule(String laborRuleId);
	public void updateLaborRule(LaborRule laborRule);
	public void createLaborRule(LaborRule laborRule);
	void deleteLaborRule(String laborRuleId);
	
	public List<ServiceArea> getServiceArea();
	public ServiceArea getServiceArea(String name);
	public void updateServiceArea(ServiceArea serviceArea);
	public void createServiceArea(ServiceArea serviceArea);
	void deleteServiceArea(String name);

	public List<TravelCost> getTravelCost();
	public TravelCost getTravelCost(String name);
	public void updateTravelCost(TravelCost travelCost);
	public void createTravelCost(TravelCost travelCost);
	void deleteTravelCost(String name);
	
	public List<TravelRule> getTravelRule();
	public TravelRule getTravelRule(String name);
	public void updateTravelRule(TravelRule travelRule);
	public void createTravelRule(TravelRule travelRule);
	void deleteTravelRule(String name);

	public List<TravelMode> getTravelMode();
	public TravelMode getTravelMode(String name);
	public void updateTravelMode(TravelMode travelMode);
	public void createTravelMode(TravelMode travelMode);
	void deleteTravelMode(String name);

}
