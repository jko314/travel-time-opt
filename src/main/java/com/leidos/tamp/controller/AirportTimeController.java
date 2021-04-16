package com.leidos.tamp.controller;

import com.leidos.tamp.beans.AirportServiceArea;
import com.leidos.tamp.beans.AirportTime;
import com.leidos.tamp.service.IRADService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@RestController
@RequestMapping("/irad/airportTime")
public class AirportTimeController {
	
	@Autowired
	private IRADService tampService;

	@GetMapping(value="/getall")
	public List<AirportTime> getAllAirportTime() {
		List<AirportTime> airportTime = tampService.getAirportTime();
		return airportTime;
	}
	
	@GetMapping(value="/get/{facilityId}")
	public AirportTime getAirportServiceArea(@PathVariable("facilityId") String facilityId) {
		AirportTime airportTime = tampService.getAirportTime(facilityId);
		return airportTime;
	}
	
	@PostMapping(value="/create")
	@ResponseStatus(HttpStatus.CREATED)
	public void create(@RequestBody AirportTime airportTime) {
		tampService.createAirportTime(airportTime);
	}
	
	@PutMapping(value="/update")
	@ResponseStatus(HttpStatus.OK)
	public void update(@RequestBody AirportServiceArea airportServiceArea) {
		tampService.updateAirportServiceArea(airportServiceArea);
	}

	@DeleteMapping(value="/delete/{facilityId}")
	@ResponseStatus(HttpStatus.OK)
	public void update(@PathVariable("facilityId") String facilityId) {
		tampService.deleteAirportServiceArea(facilityId);
	}
}
