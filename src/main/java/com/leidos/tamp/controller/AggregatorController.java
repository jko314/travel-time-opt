package com.leidos.tamp.controller;

import com.leidos.tamp.domain.AggregateResponse;
import com.leidos.tamp.domain.ApiRequest;
import com.leidos.tamp.domain.Task;
import com.leidos.tamp.service.AggregatorService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.context.request.async.DeferredResult;

@RestController
public class AggregatorController {
    private final AggregatorService service;

    @Autowired
    public AggregatorController(final AggregatorService service) {
        this.service = service;
    }

    @RequestMapping(value = "/aggregate", method = RequestMethod.POST)
    @ResponseBody
    public DeferredResult<ResponseEntity<AggregateResponse>> call(@RequestBody final ApiRequest request) {

        DeferredResult<ResponseEntity<AggregateResponse>> result = new DeferredResult<>();
        Task task = new Task(result, request.getUrls());
        service.execute(task);

        return result;
    }
}
