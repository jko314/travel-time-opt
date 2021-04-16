package com.leidos.tamp.controller;

import com.leidos.tamp.beans.AirportServiceArea;
import com.leidos.tamp.service.IRADService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@RestController
@RequestMapping("/irad/airportServiceArea")
public class AirportServiceAreaController {
	
	@Autowired
	private IRADService tampService;

	@GetMapping(value="/getall")
	public List<AirportServiceArea> getAllAirportServiceArea() {
		List<AirportServiceArea> airportServiceArea = tampService.getAirportServiceArea();
		return airportServiceArea;
	}
	
	@GetMapping(value="/get/{code}")
	public AirportServiceArea getAirportServiceArea(@PathVariable("code") String airport) {
		AirportServiceArea airportServiceArea = tampService.getAirportServiceArea(airport);
		return airportServiceArea;
	}
	
	@PostMapping(value="/create")
	@ResponseStatus(HttpStatus.CREATED)
	public void create(@RequestBody AirportServiceArea airportServiceArea) {
		tampService.createAirportServiceArea(airportServiceArea);
	}
	
	@PutMapping(value="/update")
	@ResponseStatus(HttpStatus.OK)
	public void update(@RequestBody AirportServiceArea airportServiceArea) {
		tampService.updateAirportServiceArea(airportServiceArea);
	}

	@DeleteMapping(value="/delete/{airportCode}")
	@ResponseStatus(HttpStatus.OK)
	public void update(@PathVariable("airportCode") String airportCode) {
		tampService.deleteAirportServiceArea(airportCode);
	}
}
