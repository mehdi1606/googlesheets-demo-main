package com.googlesheets.demo.service;

import com.googlesheets.demo.Entity.Person;
import com.googlesheets.demo.Repository.PersonRepository;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;
import java.util.Optional;

@Service
public class PersonServiceImpl implements PersonService {

    @Autowired
    private PersonRepository personRepository;

    @Override
    public List<Person> getAllPersons() {
        return personRepository.findAll();
    }

    @Override
    public Optional<Person> getPersonById(Double id) {
        return personRepository.findById(id);
    }

    @Override
    public Person createPerson(Person person) {
        return personRepository.save(person);
    }

    @Override
    public Person updatePerson(Double id, Person updatedPerson) {
        updatedPerson.setId(id);
        Person person = personRepository.findById(updatedPerson.getId()).orElse(null);
        if (person != null) {
            BeanUtils.copyProperties(updatedPerson, person);
            return personRepository.save(person);

        }

        return updatedPerson;


    }

    @Override
    public String deletePerson(Double id) {
        Person person = personRepository.findById(id)
                .orElseThrow(() -> new IllegalArgumentException("Person not found with ID: " + id));
        personRepository.delete(person);
        return "Person deleted successfully.";
    }
}
