package com.googlesheets.demo.service;

import com.googlesheets.demo.Entity.Person;
import com.googlesheets.demo.Repository.PersonRepository;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;

@Service
public class PersonServiceImpl implements PersonService {

    @Autowired
    private PersonRepository personRepository;

    @Override
    public  List<Person> getAllPersons() {
        List<Person> allPersons = personRepository.findAll();
        return allPersons.stream()
                .filter(person -> !person.isDeleted())
                .collect(Collectors.toList());
    }

    @Override
    public Optional<Person> getPersonById(Long id) {
        return personRepository.findById(id);
    }

    @Override
    public Person createPerson(Person person) {
        Person savedPerson = personRepository.save(person);
        return savedPerson;
    }

    @Override
    public Person updatePerson(Long id, Person updatedPerson) {
        updatedPerson.setId(id);
        Person person = personRepository.findById(updatedPerson.getId()).orElse(null);
        if (person != null) {
            BeanUtils.copyProperties(updatedPerson, person);
            return personRepository.save(person);

        }

        return updatedPerson;


    }

    @Override
    public String deletePerson(Long id) {
        Person person = personRepository.findById(id)
                .orElseThrow(() -> new IllegalArgumentException("Person not found with ID: " + id));
        personRepository.delete(person);
        return "Person deleted successfully.";
    }
}
