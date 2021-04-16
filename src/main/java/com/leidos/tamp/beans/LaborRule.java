package com.leidos.tamp.beans;

import org.bson.codecs.pojo.annotations.BsonProperty;
import org.bson.types.ObjectId;

public class LaborRule {
	ObjectId id;
	@BsonProperty(value="LaborRuleID")
	String laborRuleId;	
	@BsonProperty(value="OvertimeDayMin")
	String overtimeDayMin;
	@BsonProperty(value="OvertimeDayMax")
	String overtimeDayMax;
	@BsonProperty(value="WorkWeekMin")
	String workWeekMin;
	@BsonProperty(value="WorkWeekMax")
	String workWeekMax;
	@BsonProperty(value="OvertimeWeekStart")
	String overtimeWeekStart;
	@BsonProperty(value="OvertimeWeekMin")
	String overtimeWeekMin;
	@BsonProperty(value="OvertimeWeekMax")
	String overtimeWeekMax;
	@BsonProperty(value="WorkWeekDays")
	String workWeekDays;
	public ObjectId getId() {
		return id;
	}
	
	public String getLaborRuleId() {
		return laborRuleId;
	}


	public void setLanborRuleId(String laborRuleId) {
		this.laborRuleId = laborRuleId;
	}

	public void setId(ObjectId id) {
		this.id = id;
	}
	public String getOvertimeDayMin() {
		return overtimeDayMin;
	}
	public void setOvertimeDayMin(String overtimeDayMin) {
		this.overtimeDayMin = overtimeDayMin;
	}
	public String getOvertimeDayMax() {
		return overtimeDayMax;
	}
	public void setOvertimeDayMax(String overtimeDayMax) {
		this.overtimeDayMax = overtimeDayMax;
	}
	public String getWorkWeekMin() {
		return workWeekMin;
	}
	public void setWorkWeekMin(String workWeekMin) {
		this.workWeekMin = workWeekMin;
	}
	public String getWorkWeekMax() {
		return workWeekMax;
	}
	public void setWorkWeekMax(String workWeekMax) {
		this.workWeekMax = workWeekMax;
	}
	public String getOvertimeWeekStart() {
		return overtimeWeekStart;
	}
	public void setOvertimeWeekStart(String overtimeWeekStart) {
		this.overtimeWeekStart = overtimeWeekStart;
	}
	public String getOvertimeWeekMin() {
		return overtimeWeekMin;
	}
	public void setOvertimeWeekMin(String overtimeWeekMin) {
		this.overtimeWeekMin = overtimeWeekMin;
	}
	public String getOvertimeWeekMax() {
		return overtimeWeekMax;
	}
	public void setOvertimeWeekMax(String overtimeWeekMax) {
		this.overtimeWeekMax = overtimeWeekMax;
	}
	public String getWorkWeekDays() {
		return workWeekDays;
	}
	public void setWorkWeekDays(String workWeekDays) {
		this.workWeekDays = workWeekDays;
	}

}
