package com.googlesheets.demo.controller;

import com.googlesheets.demo.service.GoogleDocsService;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.security.GeneralSecurityException;
@RestController
@RequestMapping("/api")
public class GoogleDocsController {

    private final GoogleDocsService googleDocsService;

    public GoogleDocsController(GoogleDocsService googleDocsService) {
        this.googleDocsService = googleDocsService;
    }

    @GetMapping("/updateGoogleDoc/{personId}")
    public String updateGoogleDoc(@PathVariable Double personId) throws GeneralSecurityException, IOException {

        return googleDocsService.fillPlaceholders(personId);
    }
}
