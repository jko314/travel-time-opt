package com.leidos.tamp.beans;

import org.bson.codecs.pojo.annotations.BsonProperty;
import org.bson.types.ObjectId;

import java.util.List;

public class ServiceArea {
	ObjectId id;
	@BsonProperty(value="Name")
	String name;
	@BsonProperty(value="City")
	String city;
	@BsonProperty(value="State")
	String state;
	@BsonProperty(value="Latitude")
	double latitude;
	@BsonProperty(value="Longitude")
	double longitude;

	List<Integer> airportIndexes;
	// Maintenance activities
	List<Integer> scheduledMaintIndexes;
	List<Integer> completedMaintIndexes;
	List<Integer> interuptedMaintIndexes;

	public ObjectId getId() {
		return id;
	}
	public void setId(ObjectId id) {
		this.id = id;
	}
	
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
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
}
