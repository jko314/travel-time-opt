package com.leidos.tamp.beans;

import org.bson.codecs.pojo.annotations.BsonProperty;
import org.bson.types.ObjectId;

public class TravelRule {
	ObjectId id;
	@BsonProperty(value="TravelRuleID")
	String travelRuleId;	
	@BsonProperty(value="PerDiemDistance")
	String perDiemDistance;
	@BsonProperty(value="FirstDay")
	String firstDay;
	@BsonProperty(value="LastDay")
	String lastDay;
	public ObjectId getId() {
		return id;
	}
	public void setId(ObjectId id) {
		this.id = id;
	}
	public String getTravelRuleId() {
		return travelRuleId;
	}
	public void setTravelRuleId(String travelRuleId) {
		this.travelRuleId = travelRuleId;
	}
	public String getPerDiemDistance() {
		return perDiemDistance;
	}
	public void setPerDiemDistance(String perDiemDistance) {
		this.perDiemDistance = perDiemDistance;
	}
	public String getFirstDay() {
		return firstDay;
	}
	public void setFirstDay(String firstDay) {
		this.firstDay = firstDay;
	}
	public String getLastDay() {
		return lastDay;
	}
	public void setLastDay(String lastDay) {
		this.lastDay = lastDay;
	}
	
}
