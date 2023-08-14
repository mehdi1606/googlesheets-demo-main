package com.googlesheets.demo.controller;

import com.googlesheets.demo.service.AllDataGoogleSlides;
import com.googlesheets.demo.service.GoogleSlidesService;
import org.springframework.beans.factory.annotation.Autowired;
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
    @Autowired
    private final AllDataGoogleSlides allDataGoogleSlides;


    public GoogleSlidesController(GoogleSlidesService googleSlidesService, AllDataGoogleSlides allDataGoogleSlides) {
        this.googleSlidesService = googleSlidesService;
        this.allDataGoogleSlides = allDataGoogleSlides;
    }

    @GetMapping("/UpdateGoogleSlide/{personId}")
    public String updateGoogleSlide(@PathVariable Long personId) throws GeneralSecurityException, IOException {

        return googleSlidesService.fillPlaceholders(personId);
    }
    @GetMapping("/UpdateGoogleSlide/AllPresentation")
    public String updateGoogleSlideALL() throws GeneralSecurityException, IOException {

        return allDataGoogleSlides.fillPlaceholdersForAll().toString();
    }
}
