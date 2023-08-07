package com.googlesheets.demo.controller;

import com.googlesheets.demo.Entity.Person;
import com.googlesheets.demo.Repository.PersonRepository;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xslf.usermodel.*;
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

    public PptController(PersonRepository databaseService) {
        this.databaseService = databaseService;
    }

    @SneakyThrows
    @GetMapping("/populatePpt/{id}")
    public void populatePpt(@PathVariable Double id) {
        // Retrieve your person from the database based on the id
        Optional<Person> person =  databaseService.findById(id);

        // Open your existing PowerPoint document
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("C:\\Users\\hp\\Downloads\\print.pptx"));

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
                                run.setText(textValue.replace("id", String.valueOf(person.get().getId())));
                            }
                            if (textValue.contains("lastname")) {
                                run.setText(textValue.replace("lastname", person.get().getLastName()));

                            }

                            if (textValue.contains("firstname")) {
                                run.setText(textValue.replace("firstname", person.get().getFirstName()));
                            }
                            if (textValue.contains("(genre)")) {
                                run.setText(textValue.replace("(genre)", person.get().getGender()));

                            }
                            if (textValue.contains("(age)")) {
                                run.setText(textValue.replace("(age)", String.valueOf(person.get().getAge())));

                            }
                            if (textValue.contains("titre")) {
                                run.setText(textValue.replace("titre", person.get().getTitle()));

                            }
                            if (textValue.contains("zoneImplpr")) {
                                run.setText( textValue.replace("zoneImplpr", person.get().getImplementationZone()));

                            }
                            if (textValue.contains("etatlocal")) {
                                run.setText( textValue.replace("etatlocal", person.get().getLocaleState()));

                            }
                            if (textValue.contains("statutactual")) {
                                run.setText(textValue.replace("statutactual", person.get().getCurrentLegalStatus()));

                            }
                            if (textValue.contains("specialdiplome")) {
                                run.setText(textValue.replace("specialdiplome", person.get().getDegreeSpecialty()));

                            }
                            if (textValue.contains("nvetude")) {
                                run.setText( textValue.replace("nvetude", person.get().getEducationLevel()));

                            }
                            if (textValue.contains("descriptionprojet")) {
                                run.setText( textValue.replace("descriptionprojet", person.get().getProjectDescription()));

                            }
                            if (textValue.contains("etatduprojet")) {
                                run.setText(textValue.replace("etatduprojet", person.get().getProjectState()));
                            }
           /* if (textValue.contains("dejafinanc")) {
                textValue = textValue.replace("dejafinanc", person.get().getHasReceivedFunding());
                textElement.setValue(textValue);
            }*/
                            if (textValue.contains("statutjuridique")) {
                                run.setText(textValue.replace("statutjuridique", person.get().getCurrentLegalStatus()));
                            }
                            if (textValue.contains("rhactual")) {
                                run.setText(textValue.replace("rhactual", String.valueOf(person.get().getCurrentHR())));
                            }
                            if (textValue.contains("rhprevisionel")) {
                                run.setText( textValue.replace("rhprevisionel", String.valueOf(person.get().getProjectedHR())));
                            }
                            if (textValue.contains("(region)")) {
                                run.setText(textValue.replace("(region)", person.get().getRegion()));
                            }
                        }
                    }
                }
            }
        }

        // Save the PowerPoint
        FileOutputStream out = new FileOutputStream("C:\\Users\\hp\\Desktop\\googlesheets-demo-main\\src\\main\\resources\\output.pptx");
        ppt.write(out);
        out.close();
        ppt.close();
    }
}
