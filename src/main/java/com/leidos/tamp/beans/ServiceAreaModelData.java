package com.leidos.tamp.beans;

import com.mongodb.client.MongoCollection;

public class ServiceAreaModelData {
    MongoCollection<AirportWithCargo> airportWithCargoCollection;
    MongoCollection<ServiceArea> serviceAreaCollection;
    MongoCollection<Airport> airportCollection;
    MongoCollection<AirportServiceArea> airportServiceAreaCollection;
    MongoCollection<TravelCost> travelCostWithCollection;
    MongoCollection<TravelMode> travelModeCollection;
    MongoCollection<AirportTime> airportTimeCollection;
    MongoCollection<EquipmentModel> equipmentModelCollection;
    MongoCollection<AirportEquipment> airportEquipmentCollection;
    MongoCollection<Equipment> equipmentCollection;
    MongoCollection<EquipmentPM> equipmentPMCollection;
    MongoCollection<EquipmentCM> equipmentCMCollection;
    MongoCollection<CMRate> cmRateCollection;
    MongoCollection<LaborRule> laborRuleCollection;
    MongoCollection<TravelRule> travelRuleCollection;
	public MongoCollection<AirportWithCargo> getAirportWithCargoCollection() {
		return airportWithCargoCollection;
	}
	public void setAirportWithCargoCollection(MongoCollection<AirportWithCargo> airportWithCargoCollection) {
		this.airportWithCargoCollection = airportWithCargoCollection;
	}
	public MongoCollection<ServiceArea> getServiceAreaCollection() {
		return serviceAreaCollection;
	}
	public void setServiceAreaCollection(MongoCollection<ServiceArea> serviceAreaCollection) {
		this.serviceAreaCollection = serviceAreaCollection;
	}
	public MongoCollection<Airport> getAirportCollection() {
		return airportCollection;
	}
	public void setAirportCollection(MongoCollection<Airport> airportCollection) {
		this.airportCollection = airportCollection;
	}
	public MongoCollection<AirportServiceArea> getAirportServiceAreaCollection() {
		return airportServiceAreaCollection;
	}
	public void setAirportServiceAreaCollection(MongoCollection<AirportServiceArea> airportServiceAreaCollection) {
		this.airportServiceAreaCollection = airportServiceAreaCollection;
	}
	public MongoCollection<TravelCost> getTravelCostWithCollection() {
		return travelCostWithCollection;
	}
	public void setTravelCostWithCollection(MongoCollection<TravelCost> travelCostWithCollection) {
		this.travelCostWithCollection = travelCostWithCollection;
	}
	public MongoCollection<TravelMode> getTravelModeCollection() {
		return travelModeCollection;
	}
	public void setTravelModeCollection(MongoCollection<TravelMode> travelModeCollection) {
		this.travelModeCollection = travelModeCollection;
	}
	public MongoCollection<AirportTime> getAirportTimeCollection() {
		return airportTimeCollection;
	}
	public void setAirportTimeCollection(MongoCollection<AirportTime> airportTimeCollection) {
		this.airportTimeCollection = airportTimeCollection;
	}
	public MongoCollection<EquipmentModel> getEquipmentModelCollection() {
		return equipmentModelCollection;
	}
	public void setEquipmentModelCollection(MongoCollection<EquipmentModel> equipmentModelCollection) {
		this.equipmentModelCollection = equipmentModelCollection;
	}
	public MongoCollection<AirportEquipment> getAirportEquipmentCollection() {
		return airportEquipmentCollection;
	}
	public void setAirportEquipmentCollection(MongoCollection<AirportEquipment> airportEquipmentCollection) {
		this.airportEquipmentCollection = airportEquipmentCollection;
	}
	public MongoCollection<Equipment> getEquipmentCollection() {
		return equipmentCollection;
	}
	public void setEquipmentCollection(MongoCollection<Equipment> equipmentCollection) {
		this.equipmentCollection = equipmentCollection;
	}
	public MongoCollection<EquipmentPM> getEquipmentPMCollection() {
		return equipmentPMCollection;
	}
	public void setEquipmentPMCollection(MongoCollection<EquipmentPM> equipmentPMCollection) {
		this.equipmentPMCollection = equipmentPMCollection;
	}
	public MongoCollection<EquipmentCM> getEquipmentCMCollection() {
		return equipmentCMCollection;
	}
	public void setEquipmentCMCollection(MongoCollection<EquipmentCM> equipmentCMCollection) {
		this.equipmentCMCollection = equipmentCMCollection;
	}
	public MongoCollection<CMRate> getCmRateCollection() {
		return cmRateCollection;
	}
	public void setCmRateCollection(MongoCollection<CMRate> cmRateCollection) {
		this.cmRateCollection = cmRateCollection;
	}
	public MongoCollection<LaborRule> getLaborRuleCollection() {
		return laborRuleCollection;
	}
	public void setLaborRuleCollection(MongoCollection<LaborRule> laborRuleCollection) {
		this.laborRuleCollection = laborRuleCollection;
	}
	public MongoCollection<TravelRule> getTravelRuleCollection() {
		return travelRuleCollection;
	}
	public void setTravelRuleCollection(MongoCollection<TravelRule> travelRuleCollection) {
		this.travelRuleCollection = travelRuleCollection;
	}

}
