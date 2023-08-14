
package com.googlesheets.demo.controller;


import com.googlesheets.demo.Entity.Person;
import com.googlesheets.demo.Repository.PersonRepository;
import com.googlesheets.demo.service.GoogleDocsService;
import com.googlesheets.demo.service.GoogleSheetsService;

import com.googlesheets.demo.service.PersonService;
import lombok.RequiredArgsConstructor;
import org.springframework.beans.factory.annotation.Autowired;


import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.server.ResponseStatusException;

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
    @Autowired
    private  GoogleDocsService googleDocsService;
    String documentId = "10cWWJ8GIWz9zKBzboWTst67C11OBt9VrgT01c_juAB8";







    @GetMapping("/{spreadsheetId}")
    public String getSpreadsheetValues(@PathVariable String spreadsheetId) throws IOException, GeneralSecurityException {
       googleSheetsService.getSpreadsheetValues(spreadsheetId);
        return "OK";
    }
    @GetMapping("/persons")
    public List<Person> allPersons() throws IOException, GeneralSecurityException {
        return personService.getAllPersons();
    }

    @GetMapping("/persons/{id}")
    public Optional<Person> getPersonById(@PathVariable Long  id) {
        return personService.getPersonById(id);
    }
    @PostMapping("/persons")
    public Person createPerson(@RequestBody Person person) {
        return personService.createPerson(person);
    }
    @PutMapping("/persons/{id}")
    public Person updatePerson(@PathVariable Long  id, @RequestBody Person updatedPerson) {
        return  personService.updatePerson(id, updatedPerson);

    }
    @DeleteMapping("/persons/{id}")
    public ResponseEntity<Object> deletePerson(@PathVariable Long id) {
        Optional<Person> existingPersonOptional = personService.getPersonById(id);
        if (!existingPersonOptional.isPresent()) {
            throw new ResponseStatusException(HttpStatus.NOT_FOUND, "Could not find person");
        }

        Person existingPerson = existingPersonOptional.get();
        existingPerson.setDeleted(true);
        personRepository.save(existingPerson);

        return ResponseEntity.ok().build();
    }

}

