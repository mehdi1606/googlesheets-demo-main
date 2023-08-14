package com.googlesheets.demo.controller;

import com.googlesheets.demo.Entity.Person;
import com.googlesheets.demo.Repository.PersonRepository;
import com.googlesheets.demo.service.PersonService;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xslf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Optional;

@Slf4j
@RestController
@RequestMapping("/api")
public class PptController {

    private final PersonRepository databaseService;
@Autowired
private PersonService personService;
    public PptController(PersonRepository databaseService) {
        this.databaseService = databaseService;
    }


    @SneakyThrows
    @GetMapping("/populatePpt/{id}")
    public void populatePpt(@PathVariable Long id) {
        Person person = databaseService.findById(id).orElseThrow(() -> new Exception("Person not found"));
        populatePptForIndividual(person);
    }

    @SneakyThrows
    @GetMapping("/populatePptForAll")
    public void populatePptForAll() {
        List<Person> allPersons = personService.getAllPersons();
        for (Person person : allPersons) {
            populatePptForIndividual(person);
        }
    }
    private void populatePptForIndividual(Person person) throws Exception {
        // Specify your PowerPoint file path
        String inputFilePath = "C:\\Users\\hp\\Downloads\\print.pptx";
        String outputFilePath = "C:\\Users\\hp\\Desktop\\pptPrint\\" + person.getLastName() + "_" + person.getId() + ".pptx";

        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(inputFilePath));

// Go through each slide in the PowerPoint
        for (XSLFSlide slide : ppt.getSlides()) {
            for (XSLFShape shape : slide) {
                if (shape instanceof XSLFTextShape) {
                    XSLFTextShape textShape = (XSLFTextShape) shape;
                    // Go through each paragraph in the text box
                    for (XSLFTextParagraph paragraph : textShape) {
                        // Go through each run in the paragraph
                        for (XSLFTextRun run : paragraph) {
                            String textValue = run.getRawText();
                            // Replace placeholders in the text
                            if (textValue.contains("id")) {
                                run.setText(textValue.replace("id", String.valueOf(person.getId())));
                            }
                            if (textValue.contains("lastname")) {
                                run.setText(textValue.replace("lastname", person. getLastName()));

                            }

                            if (textValue.contains("firstname")) {
                                run.setText(textValue.replace("firstname", person.getFirstName()));
                            }
                            if (textValue.contains("(genre)")) {
                                run.setText(textValue.replace("(genre)", person. getGender()));

                            }
                            if (textValue.contains("(age)")) {
                                run.setText(textValue.replace("(age)", String.valueOf(person. getAge())));

                            }
                            if (textValue.contains("titre")) {
                                run.setText(textValue.replace("titre", person. getTitle()));

                            }
                            if (textValue.contains("zoneImplpr")) {
                                run.setText( textValue.replace("zoneImplpr", person. getImplementationZone()));

                            }
                            if (textValue.contains("etatlocal")) {
                                run.setText( textValue.replace("etatlocal", person. getLocaleState()));

                            }
                            if (textValue.contains("statutactual")) {
                                run.setText(textValue.replace("statutactual", person. getCurrentLegalStatus()));

                            }
                            if (textValue.contains("specialdiplome")) {
                                run.setText(textValue.replace("specialdiplome", person. getDegreeSpecialty()));

                            }
                            if (textValue.contains("nvetude")) {
                                run.setText( textValue.replace("nvetude", person. getEducationLevel()));

                            }
                            if (textValue.contains("descriptionprojet")) {
                                run.setText( textValue.replace("descriptionprojet", person. getProjectDescription()));

                            }
                            if (textValue.contains("etatduprojet")) {
                                run.setText(textValue.replace("etatduprojet", person. getProjectState()));
                            }
                            if (textValue.contains("dejafinanc")) {
                                String boolAsString = String.valueOf(person.isHasReceivedFunding());
                                run.setText(textValue.replace("dejafinanc", boolAsString));
                            }
                            if (textValue.contains("statutjuridique")) {
                                run.setText(textValue.replace("statutjuridique", person. getCurrentLegalStatus()));
                            }
                            if (textValue.contains("rhactual")) {
                                run.setText(textValue.replace("rhactual", String.valueOf(person. getCurrentHR())));
                            }
                            if (textValue.contains("rhprevisionel")) {
                                run.setText( textValue.replace("rhprevisionel", String.valueOf(person. getProjectedHR())));
                            }
                            if (textValue.contains("(region)")) {
                                run.setText(textValue.replace("(region)", person. getRegion()));
                            }
                        }
                    }
                }
            }
        }
        FileOutputStream out = new FileOutputStream(outputFilePath);
        ppt.write(out);
        out.close();
        ppt.close();
    }

}
