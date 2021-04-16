package com.leidos.tamp.beans;

import org.bson.codecs.pojo.annotations.BsonProperty;
import org.bson.types.ObjectId;

public class TravelMode {
	ObjectId id;
	@BsonProperty(value="Origin")
	String origin;
	@BsonProperty(value="Destination")
	String destination;
	@BsonProperty(value="Mode")
	String mode;
	@BsonProperty(value="Parking")
	double parking;
	@BsonProperty(value="DriveMiles")
	double driveMiles;
	@BsonProperty(value="Fare")
	double fare;
	@BsonProperty(value="TravelTime")
	double travelTime;
	public ObjectId getId() {
		return id;
	}
	public void setId(ObjectId id) {
		this.id = id;
	}
	public String getOrigin() {
		return origin;
	}
	public void setOrigin(String origin) {
		this.origin = origin;
	}
	public String getDestination() {
		return destination;
	}
	public void setDestination(String destination) {
		this.destination = destination;
	}
	public String getMode() {
		return mode;
	}
	public void setMode(String mode) {
		this.mode = mode;
	}
	public double getParking() {
		return parking;
	}
	public void setParking(double parking) {
		this.parking = parking;
	}
	public double getDriveMiles() {
		return driveMiles;
	}
	public void setDriveMiles(double driveMiles) {
		this.driveMiles = driveMiles;
	}
	public double getFare() {
		return fare;
	}
	public void setFare(double fare) {
		this.fare = fare;
	}
	public double getTravelTime() {
		return travelTime;
	}
	public void setTravelTime(double travelTime) {
		this.travelTime = travelTime;
	}
	
}
