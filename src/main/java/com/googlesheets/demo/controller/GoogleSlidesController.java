package com.googlesheets.demo.controller;

import com.googlesheets.demo.service.GoogleDocsService;
import com.googlesheets.demo.service.GoogleSlidesService;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.security.GeneralSecurityException;

@RestController
@RequestMapping("/api")
public class GoogleSlidesController {

    private final GoogleSlidesService googleSlidesService;

    public GoogleSlidesController(GoogleSlidesService googleSlidesService) {
        this.googleSlidesService = googleSlidesService;
    }

    @GetMapping("/UpdateGoogleSlide/{personId}")
    public String updateGoogleDoc(@PathVariable Double personId) throws GeneralSecurityException, IOException {

        return googleSlidesService.fillPlaceholders(personId);
    }
}
