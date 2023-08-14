package com.googlesheets.demo.controller;

import com.googlesheets.demo.Entity.Person;
import com.googlesheets.demo.service.AllDataGoogleDocs;
import com.googlesheets.demo.service.GoogleDocsService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import java.io.IOException;
import java.security.GeneralSecurityException;
@RestController
@RequestMapping("/api")
public class GoogleDocsController {

    private final GoogleDocsService googleDocsService;
    @Autowired
    private final AllDataGoogleDocs allDataService;


    public GoogleDocsController(GoogleDocsService googleDocsService, AllDataGoogleDocs allDataService) {
        this.googleDocsService = googleDocsService;
        this.allDataService = allDataService;
    }

    @GetMapping("/updateGoogleDoc/{personId}")
    public String updateGoogleDoc(@PathVariable Long personId) throws GeneralSecurityException, IOException {

        return googleDocsService.fillPlaceholders(personId);
    }
    @GetMapping("/updateGoogleDoc/AllDocuments")
    public String updateGoogleDocALL() throws GeneralSecurityException, IOException {

        return allDataService.fillPlaceholdersForAll().toString();
    }
}
