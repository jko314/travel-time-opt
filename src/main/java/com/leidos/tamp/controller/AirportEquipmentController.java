package com.leidos.tamp.controller;

import com.leidos.tamp.beans.AirportEquipment;
import com.leidos.tamp.service.IRADService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@RestController
@RequestMapping("/irad/airportEquipment")
public class AirportEquipmentController {
	
	@Autowired
	private IRADService tampService;

	@GetMapping(value="/getall")
	public List<AirportEquipment> getAllAirportEquipment() {
		List<AirportEquipment> airportEquipment = tampService.getAirportEquipments();
		return airportEquipment;
	}
	
	@GetMapping(value="/get/{airportCode}")
	public AirportEquipment getAirportEquipment(@PathVariable("airportCode") String airportCode) {
		AirportEquipment airportEquipment = tampService.getAirportEquipment(airportCode);
		return airportEquipment;
	}
	
	@PostMapping(value="/create")
	@ResponseStatus(HttpStatus.CREATED)
	public void create(@RequestBody AirportEquipment airportEquipment) {
		tampService.createAirportEquipment(airportEquipment);
	}
	
	@PutMapping(value="/update")
	@ResponseStatus(HttpStatus.OK)
	public void update(@RequestBody AirportEquipment airportEquipment) {
		tampService.updateAirportEquipment(airportEquipment);
	}

	@DeleteMapping(value="/delete/{code}")
	@ResponseStatus(HttpStatus.OK)
	public void update(@PathVariable("airportCode") String airport) {
		tampService.deleteAirportEquipment(airport);
	}
}
