package com.leidos.tamp.beans;

import org.bson.codecs.pojo.annotations.BsonProperty;
import org.bson.types.ObjectId;

public class TravelCost {
	ObjectId id;
	@BsonProperty(value="Facility")
	String facility;
	@BsonProperty(value="Lodging")
	double lodging;
	@BsonProperty(value="PerDiem")
	double perDiem;
	@BsonProperty(value="RentalCar")
	double rentalCar;
	@BsonProperty(value="FacilityParking")
	double facilityParking;
	public ObjectId getId() {
		return id;
	}
	public void setId(ObjectId id) {
		this.id = id;
	}
	public String getFacility() {
		return facility;
	}
	public void setFacility(String facility) {
		this.facility = facility;
	}
	public double getLodging() {
		return lodging;
	}
	public void setLodging(double lodging) {
		this.lodging = lodging;
	}
	public double getPerDiem() {
		return perDiem;
	}
	public void setPerDiem(double perDiem) {
		this.perDiem = perDiem;
	}
	public double getRentalCar() {
		return rentalCar;
	}
	public void setRentalCar(double rentalCar) {
		this.rentalCar = rentalCar;
	}
	public double getFacilityParking() {
		return facilityParking;
	}
	public void setFacilityParking(double facilityParking) {
		this.facilityParking = facilityParking;
	}
	
}
