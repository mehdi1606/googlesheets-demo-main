package com.googlesheets.demo.service;

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

@Service
public class GoogleSlidesService {
    private static final Logger LOGGER = LoggerFactory.getLogger(GoogleSlidesService.class);

    @Value("1wIOkKHF-kLd85Ul29qm7y_2vw0DWkNq7_gvjffIPofs")
    private String presentationId;
    @Autowired
    private PersonService  personService;
    @Autowired
    private PersonRepository personRepository;
    @Autowired
    private AutenificationPrintSlides autenificationPrintSlides ;
    @Autowired
    private DriverConfig driverConfig;

    List<String> scopes = Collections.singletonList(SlidesScopes.PRESENTATIONS);

    public Presentation getPresentation(String presentationId) throws IOException, GeneralSecurityException {
        Slides slidesService = autenificationPrintSlides.getSlidesService();
        Presentation presentation = slidesService.presentations().get(presentationId).execute();
        return presentation;
    }

    public String fillPlaceholders( Long personId) throws IOException, GeneralSecurityException {
        // Get the presentation
        Presentation presentation = getPresentation(presentationId);

        // Log or analyse the presentation if necessary
        LOGGER.info("Original presentation: " + presentation.getTitle());

        // Get the person and fill the placeholders as before
        Optional<Person> personData = personService.getPersonById(personId);
        if (!personData.isPresent()) {
            throw new NoSuchElementException("Person not found");
        }

        // Create a copy of the presentation
        com.google.api.services.drive.Drive driveService = driverConfig.getDriveService();
        com.google.api.services.drive.model.File file = new com.google.api.services.drive.model.File();
        file.setName("PresentationPerson");
        List<String> parents = Collections.singletonList("1SBEIko84GyPLRZO_9oFsAkruTC7r9kGS");
        file.setParents(parents);

        com.google.api.services.drive.model.File copiedFile =
                driveService.files().copy(presentationId, file).execute();

        // Get the new presentation's ID
        String newPresentationId = copiedFile.getId();

        // Prepare the data
        Map<String, String> data = new HashMap<>();
        data.put("{id}", String.valueOf(personData.get().getId()));

        data.put("{firstname}", personData.get().getFirstName());
        data.put("{lastname}", personData.get().getLastName());
        data.put("{gener}",personData.get().getGender());
        data.put("{Title}", personData.get().getTitle());
        data.put("{zoneImplpr}",personData.get().getImplementationZone());
        data.put("{etatlocal}", personData.get().getLocaleState());
        data.put("{age}", String.valueOf(personData.get().getAge()));
        data.put("{nvetude}",personData.get().getEducationLevel());
        data.put("{specialdiplome}", personData.get().getDegreeSpecialty());
        data.put("{statutactual}",personData.get().getCurrentStatus());
        data.put("{statutjuridique}",personData.get().getCurrentLegalStatus());
        data.put("{descriptionprojet}",personData.get().getProjectDescription());
        data.put("{etatduprojet}",personData.get().getProjectState());
        String boolAsString = String.valueOf(personData.get().isHasReceivedFunding());
        data.put("{dejafinanc}",boolAsString);
        data.put("{rhactual}", String.valueOf(personData.get().getCurrentHR()));
        data.put("{rhprevisionel}", String.valueOf(personData.get().getProjectedHR()));
        data.put("{region}",personData.get().getRegion());

        List<Request> requests = new ArrayList<>();
        for (Map.Entry<String, String> entry : data.entrySet()) {
            String find = entry.getKey();
            String replace = entry.getValue();
            ReplaceAllTextRequest replaceRequest = new ReplaceAllTextRequest()
                    .setContainsText(new SubstringMatchCriteria().setText(find))
                    .setReplaceText(replace);
            requests.add(new Request().setReplaceAllText(replaceRequest));
        }

        Slides slides = autenificationPrintSlides.getSlidesService();
        BatchUpdatePresentationRequest batchUpdateRequest = new BatchUpdatePresentationRequest().setRequests(requests);

        // Update the new presentation
        slides.presentations().batchUpdate(newPresentationId, batchUpdateRequest).execute();

        return "New presentation updated";
    }
}
