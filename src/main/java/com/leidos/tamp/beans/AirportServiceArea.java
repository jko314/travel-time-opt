package com.leidos.tamp.beans;

import org.bson.codecs.pojo.annotations.BsonProperty;
import org.bson.types.ObjectId;

public class AirportServiceArea {
	ObjectId id;
	@BsonProperty(value="AirportCode")
	String airportCode;
	@BsonProperty(value="ServiceArea")
	String serviceArea;
	@BsonProperty(value="Priority")
	double priority;
	@BsonProperty(value="IsBase")
	boolean isBase;
	@BsonProperty(value="CMTravelTime")
	double CMTravelTimme;
	public ObjectId getId() {
		return id;
	}
	public void setId(ObjectId id) {
		this.id = id;
	}
	public String getAirportCode() {
		return airportCode;
	}
	public void setAirportCode(String airportCode) {
		this.airportCode = airportCode;
	}
	public String getServiceArea() {
		return serviceArea;
	}
	public void setServiceArea(String serviceArea) {
		this.serviceArea = serviceArea;
	}
	public double getPriority() {
		return priority;
	}
	public void setPriority(double priority) {
		this.priority = priority;
	}
	public boolean isBase() {
		return isBase;
	}
	public void setBase(boolean isBase) {
		this.isBase = isBase;
	}
	public double getCMTravelTimme() {
		return CMTravelTimme;
	}
	public void setCMTravelTimme(double cMTravelTimme) {
		CMTravelTimme = cMTravelTimme;
	}
	
}
