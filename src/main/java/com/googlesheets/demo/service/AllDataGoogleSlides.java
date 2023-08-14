package com.googlesheets.demo.service;

import com.google.api.services.drive.Drive;
import com.google.api.services.slides.v1.Slides;
import com.google.api.services.slides.v1.SlidesScopes;
import com.google.api.services.slides.v1.model.*;

import com.googlesheets.demo.Entity.Person;
import com.googlesheets.demo.Repository.PersonRepository;
import com.googlesheets.demo.config.AutenificationPrintDocs;
import com.googlesheets.demo.config.AutenificationPrintSlides;
import com.googlesheets.demo.config.DriverConfig;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.*;
import java.util.List;
import java.util.stream.Collectors;

import static org.hibernate.tool.schema.SchemaToolingLogging.LOGGER;

@Service
public class AllDataGoogleSlides {
    private static final Logger LOGGER = LoggerFactory.getLogger(GoogleSlidesService.class);

    @Value("1wIOkKHF-kLd85Ul29qm7y_2vw0DWkNq7_gvjffIPofs")
    private String presentationId;
    @Autowired
    private PersonService   personService;
    @Autowired
    private PersonRepository personRepository;
    @Autowired
    private AutenificationPrintSlides autenificationPrintSlides;
    @Autowired
    private DriverConfig driverConfig;

    public Presentation getPresentation(String presentationId) throws IOException, GeneralSecurityException {
        Slides slidesService = autenificationPrintSlides.getSlidesService();
        Presentation presentation = slidesService.presentations().get(presentationId).execute();
        return presentation;
    }

    public List<String> fillPlaceholdersForAll() throws IOException, GeneralSecurityException {
        List<Person> allPersons = personService.getAllPersons();
        List<String> responses = new ArrayList<>();
        com.google.api.services.drive.Drive driveService = driverConfig.getDriveService();

        for (Person person : allPersons) {
            String response = fillPlaceholdersForIndividual(person, driveService);
            responses.add(response);
        }

        return responses;
    }

    private String fillPlaceholdersForIndividual(Person person, com.google.api.services.drive.Drive driveService) throws IOException, GeneralSecurityException {
        Presentation presentation = getPresentation(presentationId);

        com.google.api.services.drive.model.File file = new com.google.api.services.drive.model.File();
        file.setName(person.getLastName() + "_" + person.getId());
        List<String> parents = Collections.singletonList("1SBEIko84GyPLRZO_9oFsAkruTC7r9kGS");
        file.setParents(parents);

        com.google.api.services.drive.model.File copiedFile =
                driveService.files().copy(presentationId, file).execute();
        String newPresentationId = copiedFile.getId();

        Map<String, String> data = new HashMap<>();
        data.put("{id}", String.valueOf(person.getId()));

        data.put("{firstname}", person.getFirstName());
        data.put("{lastname}", person.getLastName());
        data.put("{gener}",person.getGender());
        data.put("{Title}", person.getTitle());
        data.put("{zoneImplpr}",person.getImplementationZone());
        data.put("{etatlocal}", person.getLocaleState());
        data.put("{age}", String.valueOf(person.getAge()));
        data.put("{nvetude}",person.getEducationLevel());
        data.put("{specialdiplome}", person.getDegreeSpecialty());
        data.put("{statutactual}",person.getCurrentStatus());
        data.put("{statutjuridique}",person.getCurrentLegalStatus());
        data.put("{descriptionprojet}",person.getProjectDescription());
        data.put("{etatduprojet}",person.getProjectState());
        String boolAsString = String.valueOf(person.isHasReceivedFunding());
        data.put("{dejafinanc}",boolAsString);
        data.put("{rhactual}", String.valueOf(person.getCurrentHR()));
        data.put("{rhprevisionel}", String.valueOf(person.getProjectedHR()));
        data.put("{region}",person.getRegion());

        List<Request> requests = new ArrayList<>();
        for (Map.Entry<String, String> entry : data.entrySet()) {
            ReplaceAllTextRequest replaceRequest = new ReplaceAllTextRequest()
                    .setContainsText(new SubstringMatchCriteria().setText(entry.getKey()))
                    .setReplaceText(entry.getValue());
            requests.add(new Request().setReplaceAllText(replaceRequest));
        }

        Slides slidesService = autenificationPrintSlides.getSlidesService();
        BatchUpdatePresentationRequest batchUpdateRequest = new BatchUpdatePresentationRequest().setRequests(requests);
        slidesService.presentations().batchUpdate(newPresentationId, batchUpdateRequest).execute();

        return "Presentation for " + person.getLastName() + "_" + person.getId() + " updated";
    }
}