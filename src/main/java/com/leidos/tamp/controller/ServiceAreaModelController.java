package com.leidos.tamp.controller;

import com.leidos.tamp.domain.ServiceAreaModelRequest;
import com.leidos.tamp.domain.ServiceAreaModelResponse;
import com.leidos.tamp.domain.TimeResponse;
import com.leidos.tamp.service.ServiceAreaModelService;
import com.leidos.tamp.type.POSSIBLESOLUTION_TYPE;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.context.request.async.DeferredResult;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.concurrent.atomic.AtomicInteger;

@Slf4j
@RestController
@RequestMapping(value = "/model")
public class ServiceAreaModelController {
    private final AtomicInteger counter = new AtomicInteger(0);

    @Autowired
    private ServiceAreaModelService areaModelService;

    @RequestMapping(value = "/servicearea", method = RequestMethod.POST)
    @ResponseBody
    public ResponseEntity<?> serviceArea(@RequestBody final ServiceAreaModelRequest request) {
        log.info("Response entity request");
        request.getServiceAreaCount();
        return ResponseEntity.ok(new ServiceAreaModelResponse());
    }

    @RequestMapping(value = "/servicearea", method = RequestMethod.GET)
    public DeferredResult<ResponseEntity<?>> serviceArea(@RequestParam(defaultValue = "1") int iteration,
                                                         @RequestParam(defaultValue = "Flat 80% Utilization") String fstUtilization) {
        log.info("servicearea");
        DeferredResult<ResponseEntity<?>> result = new DeferredResult<>();

        new Thread(() -> {
            try {
                POSSIBLESOLUTION_TYPE sol = areaModelService.runServiceAreaModel_Click(iteration, fstUtilization);
                result.setResult(ResponseEntity.ok("TravelMiles: " + sol.getTravelMiles() +
                        "\nFSTHours: " + sol.getFSTHours() +
                        "\nFSTCount: " + sol.getFSTCount() +
                        "\nFST Util: " + (sol.getFSTHours() / sol.getFSTCount() / sol.getFSTHours()) +
                        "\nCost: " + sol.getFSTCost()
                ));
            } catch (RuntimeException re) {
                re.printStackTrace();
            }
        }, "MyThread-" + counter.incrementAndGet()).start();

        return result;
    }
    @RequestMapping(value = "/deferred", method = RequestMethod.GET)
    public DeferredResult<ResponseEntity<?>> timeDeferred() {
        log.info("Deferred time request");
        DeferredResult<ResponseEntity<?>> result = new DeferredResult<>();

        new Thread(() -> {
            result.setResult(ResponseEntity.ok(now()));
        }, "MyThread-" + counter.incrementAndGet()).start();

        return result;
    }

    private static TimeResponse now() {
        log.info("Creating TimeResponse");
        return new TimeResponse(LocalDateTime
                .now()
                .format(DateTimeFormatter.ISO_LOCAL_DATE_TIME));
    }
}
