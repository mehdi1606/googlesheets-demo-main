
package com.googlesheets.demo.controller;


import com.googlesheets.demo.Entity.Person;
import com.googlesheets.demo.Repository.PersonRepository;
import com.googlesheets.demo.config.GoogleAuthorizationConfig;
import com.googlesheets.demo.service.GoogleSheetsService;

import com.googlesheets.demo.service.PersonService;
import lombok.RequiredArgsConstructor;
import org.springframework.beans.factory.annotation.Autowired;


import org.springframework.web.bind.annotation.*;

import java.io.IOException;
import java.security.GeneralSecurityException;

import java.util.*;
import java.util.List;


@RequiredArgsConstructor
@RestController
@RequestMapping(value = "/api/v1/googlesheets")
public class    GoogleSheetsController {

    @Autowired
    private GoogleSheetsService googleSheetsService;
    @Autowired
    private PersonRepository personRepository;
    @Autowired
    private PersonService  personService;






    @GetMapping()
    public String getSpreadsheetValues() throws IOException, GeneralSecurityException {
       googleSheetsService.getSpreadsheetValues();
        return "OK";
    }
    @GetMapping("/persons")
    public List<Person> allPersons() throws IOException, GeneralSecurityException {
        return  personService.getAllPersons();
    }

    @GetMapping("/persons/{id}")
    public Optional<Person> getPersonById(@PathVariable Double  id) {
        return personService.getPersonById(id);
    }
    @PostMapping("/persons")
    public Person createPerson(@RequestBody Person person) {
        return personService.createPerson(person);
    }
    @PutMapping("/persons/{id}")
    public Person updatePerson(@PathVariable Double  id, @RequestBody Person updatedPerson) {
        return  personService.updatePerson(id, updatedPerson);

    }
    @DeleteMapping("/persons/{id}")
    public String deletePerson(@PathVariable Double  id) {
       return personService.deletePerson(id);
    }
//--------------------------------------------------------//


}
