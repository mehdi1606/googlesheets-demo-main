package com.googlesheets.demo.controller;

import com.googlesheets.demo.Entity.Person;
import com.googlesheets.demo.Repository.PersonRepository;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.Text;
import org.docx4j.finders.ClassFinder;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.bind.annotation.PathVariable;
import org.docx4j.TraversalUtil;
import org.docx4j.TraversalUtil.Callback;
import org.docx4j.wml.Body;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Optional;

import static java.lang.Double.*;
@Slf4j
@RestController
@RequestMapping("/api")
public class WordController {

    private final PersonRepository databaseService;

    public WordController(PersonRepository databaseService) {
        this.databaseService = databaseService;
    }

    @SneakyThrows
    @GetMapping("/populateWordDoc/{id}")
    public void populateWordDoc(@PathVariable Double id) {
        // Récupérez votre personne de la base de données en fonction de l'id
        Optional<Person> person =  databaseService.findById(id);

        // Ouvrez votre document Word existant
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new java.io.File("C:\\Users\\hp\\Downloads\\Mehdi interface.docx"));

        // Trouvez tous les éléments Text dans le document
        ClassFinder finder = new ClassFinder(Text.class);

        new TraversalUtil(wordMLPackage.getMainDocumentPart().getContent(), finder);

        // Parcourez tous les éléments Text et remplacez les placeholders
        for (Object obj : finder.results) {
            Text textElement = (Text) obj;
            String textValue = textElement.getValue();
            if (textValue.contains("(id)")) {
                textValue = textValue.replace("(id)", String.valueOf(person.get().getId()));
                textElement.setValue(textValue);
            }
            if (textValue.contains("(lastname)")) {
                textValue = textValue.replace("(lastname)", person.get().getLastName());
                textElement.setValue(textValue);
            }

            if (textValue.contains("(firstname)")) {
                textValue = textValue.replace("(firstname)", person.get().getFirstName());
                textElement.setValue(textValue);
            }
            if (textValue.contains("(genre)")) {
                textValue = textValue.replace("(genre)", person.get().getGender());
                textElement.setValue(textValue);
            }
            if (textValue.contains("(age)")) {
                textValue = textValue.replace("(age)", String.valueOf(person.get().getAge()));
                textElement.setValue(textValue);
            }
            if (textValue.contains("(titre)")) {
                textValue = textValue.replace("(titre)", person.get().getTitle());
                textElement.setValue(textValue);
            }
            if (textValue.contains("(zoneImplpr)")) {
                textValue = textValue.replace("(zoneImplpr)", person.get().getImplementationZone());
                textElement.setValue(textValue);
            }
            if (textValue.contains("(etatlocal)")) {
                textValue = textValue.replace("(etatlocal)", person.get().getLocaleState());
                textElement.setValue(textValue);
            }
            if (textValue.contains("(statutactual)")) {
                textValue = textValue.replace("(statutactual)", person.get().getCurrentLegalStatus());
                textElement.setValue(textValue);
            }
            if (textValue.contains("(specialdiplome)")) {
                textValue = textValue.replace("(specialdiplome)", person.get().getDegreeSpecialty());
                textElement.setValue(textValue);
            }
            if (textValue.contains("(nvetude)")) {
                textValue = textValue.replace("(nvetude)", person.get().getEducationLevel());
                textElement.setValue(textValue);
            }
            if (textValue.contains("(descriptionprojet)")) {
                textValue = textValue.replace("(descriptionprojet)", person.get().getProjectDescription());
                textElement.setValue(textValue);
            }
            if (textValue.contains("(etatduprojet)")) {
                textValue = textValue.replace("(etatduprojet)", person.get().getProjectState());
                textElement.setValue(textValue);
            }
           /* if (textValue.contains("(dejafinanc)")) {
                textValue = textValue.replace("(dejafinanc)", person.get().getHasReceivedFunding());
                textElement.setValue(textValue);
            }*/
            if (textValue.contains("(statutjuridique)")) {
                textValue = textValue.replace("(statutjuridique)", person.get().getCurrentLegalStatus());
                textElement.setValue(textValue);
            }
            if (textValue.contains("(rhactual)")) {
                textValue = textValue.replace("(rhactual)", String.valueOf(person.get().getCurrentHR()));
                textElement.setValue(textValue);
            }
            if (textValue.contains("(rhprevisionel)")) {
                textValue = textValue.replace("(rhprevisionel)", String.valueOf(person.get().getProjectedHR()));
                textElement.setValue(textValue);
            }
            if (textValue.contains("(region)")) {
                textValue = textValue.replace("(region)", person.get().getRegion());
                textElement.setValue(textValue);
            }

        }

        // Enregistrez le document
        wordMLPackage.save(new java.io.File("C:\\Users\\hp\\Desktop\\googlesheets-demo-main\\src\\main\\resources\\output.docx"));
    }
}
