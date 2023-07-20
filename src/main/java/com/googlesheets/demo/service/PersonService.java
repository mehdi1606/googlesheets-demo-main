package com.googlesheets.demo.service;

import com.googlesheets.demo.Entity.Person;


import java.util.List;
import java.util.Optional;


    public interface PersonService {
        List<Person> getAllPersons();

        Optional<Person> getPersonById(Double id);

        Person createPerson(Person person);

        Person updatePerson(Double id, Person updatedPerson);

        String deletePerson(Double id);
    }
