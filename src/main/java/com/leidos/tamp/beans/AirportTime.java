package com.leidos.tamp.beans;

import org.bson.codecs.pojo.annotations.BsonProperty;
import org.bson.types.ObjectId;

public class AirportTime {
	ObjectId id;
	@BsonProperty(value="FacilityId")
	String facilityId;
	@BsonProperty(value="Parking")
	String parking;
	@BsonProperty(value="Enter")
	String enter;
	@BsonProperty(value="GreetDayOne")
	double greetDayOne;
	@BsonProperty(value="GreetDayN")
	double greetDayN;
	@BsonProperty(value="Exit")
	double exit;
	@BsonProperty(value="PMWait")
	double pmWait;
	@BsonProperty(value="PMSignOff")
	double pmSignOff;
	@BsonProperty(value="CriticalCMWait")
	double criticalCMWait;
	@BsonProperty(value="Non-CriticalCMWait")
	double nonCriticalCMWait;
	@BsonProperty(value="CMSignOff")
	double cmSignOff;
	public ObjectId getId() {
		return id;
	}
	public void setId(ObjectId id) {
		this.id = id;
	}
	public String getFacilityId() {
		return facilityId;
	}
	public void setFacilityId(String facilityId) {
		this.facilityId = facilityId;
	}
	public String getParking() {
		return parking;
	}
	public void setParking(String parking) {
		this.parking = parking;
	}
	public String getEnter() {
		return enter;
	}
	public void setEnter(String enter) {
		this.enter = enter;
	}
	public double getGreetDayOne() {
		return greetDayOne;
	}
	public void setGreetDayOne(double greetDayOne) {
		this.greetDayOne = greetDayOne;
	}
	public double getGreetDayN() {
		return greetDayN;
	}
	public void setGreetDayN(double greetDayN) {
		this.greetDayN = greetDayN;
	}
	public double getExit() {
		return exit;
	}
	public void setExit(double exit) {
		this.exit = exit;
	}
	public double getPmWait() {
		return pmWait;
	}
	public void setPmWait(double pmWait) {
		this.pmWait = pmWait;
	}
	public double getPmSignOff() {
		return pmSignOff;
	}
	public void setPmSignOff(double pmSignOff) {
		this.pmSignOff = pmSignOff;
	}
	public double getCriticalCMWait() {
		return criticalCMWait;
	}
	public void setCriticalCMWait(double criticalCMWait) {
		this.criticalCMWait = criticalCMWait;
	}
	public double getNonCriticalCMWait() {
		return nonCriticalCMWait;
	}
	public void setNonCriticalCMWait(double nonCriticalCMWait) {
		this.nonCriticalCMWait = nonCriticalCMWait;
	}
	public double getCmSignOff() {
		return cmSignOff;
	}
	public void setCmSignOff(double cmSignOff) {
		this.cmSignOff = cmSignOff;
	}
	

	
	
}
