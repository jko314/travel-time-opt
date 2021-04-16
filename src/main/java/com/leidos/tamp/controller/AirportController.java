package com.leidos.tamp.controller;

import com.leidos.tamp.beans.Airport;
import com.leidos.tamp.service.IRADService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@RestController
@RequestMapping("/irad/airport")
public class AirportController {
	
	@Autowired
	private IRADService tampService;

	@GetMapping(value="/getall")
	public List<Airport> getAllAirports() {
		List<Airport> airports = tampService.getAirports();
		return airports;
	}
	
	@GetMapping(value="/get/{code}")
	public Airport getAirport(@PathVariable("code") String code) {
		Airport airport = tampService.getAirport(code);
		return airport;
	}
	
	@PostMapping(value="/create")
	@ResponseStatus(HttpStatus.CREATED)
	public void create(@RequestBody Airport airport) {
		tampService.createAirport(airport);
	}
	
	@PutMapping(value="/update")
	@ResponseStatus(HttpStatus.OK)
	public void update(@RequestBody Airport airport) {
		tampService.updateAirport(airport);
	}

	@DeleteMapping(value="/delete/{code}")
	@ResponseStatus(HttpStatus.OK)
	public void update(@PathVariable("code") String code) {
		tampService.deleteAirport(code);
	}

}
