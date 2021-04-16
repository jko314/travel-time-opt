package com.leidos.tamp.beans;

import org.bson.codecs.pojo.annotations.BsonProperty;
import org.bson.types.ObjectId;

public class AirportWithCargo {
	ObjectId id;
	@BsonProperty(value="Code")
	String code;
	@BsonProperty(value="Cat")
	String cat;
	@BsonProperty(value="City")
	String city;
	@BsonProperty(value="State")
	String state;
	@BsonProperty(value="Name")
	String name;
	@BsonProperty(value="Latitude")
	double latitude;
	@BsonProperty(value="Longitude")
	double longitude;
	@BsonProperty(value="Op Start")
	double op_start;
	@BsonProperty(value="Op Hrs")
	double op_hrs;
	@BsonProperty(value="Time Zone")
	double time_zone;
	public ObjectId getId() {
		return id;
	}
	public void setId(ObjectId id) {
		this.id = id;
	}
	public String getCode() {
		return code;
	}
	public void setCode(String code) {
		this.code = code;
	}
	public String getCat() {
		return cat;
	}
	public void setCat(String cat) {
		this.cat = cat;
	}
	public String getCity() {
		return city;
	}
	public void setCity(String city) {
		this.city = city;
	}
	public String getState() {
		return state;
	}
	public void setState(String state) {
		this.state = state;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public double getLatitude() {
		return latitude;
	}
	public void setLatitude(double latitude) {
		this.latitude = latitude;
	}
	public double getLongitude() {
		return longitude;
	}
	public void setLongitude(double longitude) {
		this.longitude = longitude;
	}
	public double getOp_start() {
		return op_start;
	}
	public void setOp_start(double op_start) {
		this.op_start = op_start;
	}
	public double getOp_hrs() {
		return op_hrs;
	}
	public void setOp_hrs(double op_hrs) {
		this.op_hrs = op_hrs;
	}
	public double getTime_zone() {
		return time_zone;
	}
	public void setTime_zone(double time_zone) {
		this.time_zone = time_zone;
	}
	
}
