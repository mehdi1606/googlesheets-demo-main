package com.googlesheets.demo.controller;

import com.googlesheets.demo.Entity.Person;
import com.googlesheets.demo.Repository.PersonRepository;
import com.googlesheets.demo.service.PersonService;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.Text;
import org.docx4j.finders.ClassFinder;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.bind.annotation.PathVariable;
import org.docx4j.TraversalUtil;
import org.docx4j.TraversalUtil.Callback;
import org.docx4j.wml.Body;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Optional;

import static java.lang.Double.*;
@Slf4j
@RestController
@RequestMapping("/api")
public class WordController {

    private final PersonRepository databaseService;
@Autowired
private PersonService  personService;
    public WordController(PersonRepository databaseService) {
        this.databaseService = databaseService;
    }

    @SneakyThrows
    @GetMapping("/populateWordDoc/{id}")
    public void populateWordDoc(@PathVariable Long id) {
        Person person = personService.getPersonById(id).orElseThrow(() -> new Exception("Person not found"));
        populateWordDocForIndividual(person);
    }

    @SneakyThrows
    @GetMapping("/populateWordDocForAll")
    public void populateWordDocForAll() {
        List<Person> allPersons = personService.getAllPersons();
        for (Person person : allPersons) {
            populateWordDocForIndividual(person);
        }
    }
    private void populateWordDocForIndividual(Person person) throws Exception {
        // Specify your Word document file path
        String inputFilePath = "C:\\Users\\hp\\Downloads\\Mehdi interface.docx";
        String outputFilePath = "C:\\Users\\hp\\Desktop\\PrintDocs\\" + person.getLastName() + "_" + person.getId() + ".docx";

        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new java.io.File(inputFilePath));

// Trouvez tous les éléments Text dans le document
        ClassFinder finder = new ClassFinder(Text.class);

        new TraversalUtil(wordMLPackage.getMainDocumentPart().getContent(), finder);

        // Parcourez tous les éléments Text et remplacez les placeholders
        for (Object obj : finder.results) {
            Text textElement = (Text) obj;
            String textValue = textElement.getValue();
            if (textValue.contains("(id)")) {
                textValue = textValue.replace("(id)", String.valueOf(person. getId()));
                textElement.setValue(textValue);
            }
            if (textValue.contains("(lastname)")) {
                textValue = textValue.replace("(lastname)", person. getLastName());
                textElement.setValue(textValue);
            }

            if (textValue.contains("(firstname)")) {
                textValue = textValue.replace("(firstname)", person. getFirstName());
                textElement.setValue(textValue);
            }
            if (textValue.contains("(genre)")) {
                textValue = textValue.replace("(genre)", person. getGender());
                textElement.setValue(textValue);
            }
            if (textValue.contains("(age)")) {
                textValue = textValue.replace("(age)", String.valueOf(person. getAge()));
                textElement.setValue(textValue);
            }
            if (textValue.contains("(titre)")) {
                textValue = textValue.replace("(titre)", person. getTitle());
                textElement.setValue(textValue);
            }
            if (textValue.contains("(zoneImplpr)")) {
                textValue = textValue.replace("(zoneImplpr)", person. getImplementationZone());
                textElement.setValue(textValue);
            }
            if (textValue.contains("(etatlocal)")) {
                textValue = textValue.replace("(etatlocal)", person. getLocaleState());
                textElement.setValue(textValue);
            }
            if (textValue.contains("(statutactual)")) {
                textValue = textValue.replace("(statutactual)", person. getCurrentLegalStatus());
                textElement.setValue(textValue);
            }
            if (textValue.contains("(specialdiplome)")) {
                textValue = textValue.replace("(specialdiplome)", person. getDegreeSpecialty());
                textElement.setValue(textValue);
            }
            if (textValue.contains("(nvetude)")) {
                textValue = textValue.replace("(nvetude)", person. getEducationLevel());
                textElement.setValue(textValue);
            }
            if (textValue.contains("(descriptionprojet)")) {
                textValue = textValue.replace("(descriptionprojet)", person. getProjectDescription());
                textElement.setValue(textValue);
            }
            if (textValue.contains("(etatduprojet)")) {
                textValue = textValue.replace("(etatduprojet)", person. getProjectState());
                textElement.setValue(textValue);
            }
            if (textValue.contains("dejafinanc")) {
                String boolAsString = String.valueOf(person.isHasReceivedFunding());
                textValue = textValue.replace("dejafinanc", boolAsString);
                textElement.setValue(textValue);
            }
            if (textValue.contains("(statutjuridique)")) {
                textValue = textValue.replace("(statutjuridique)", person. getCurrentLegalStatus());
                textElement.setValue(textValue);
            }
            if (textValue.contains("(rhactual)")) {
                textValue = textValue.replace("(rhactual)", String.valueOf(person. getCurrentHR()));
                textElement.setValue(textValue);
            }
            if (textValue.contains("(rhprevisionel)")) {
                textValue = textValue.replace("(rhprevisionel)", String.valueOf(person. getProjectedHR()));
                textElement.setValue(textValue);
            }
            if (textValue.contains("(region)")) {
                textValue = textValue.replace("(region)", person. getRegion());
                textElement.setValue(textValue);
            }

        }
        wordMLPackage.save(new java.io.File(outputFilePath));
    }
}

