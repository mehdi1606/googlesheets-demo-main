package com.googlesheets.demo.service;

import com.google.api.services.docs.v1.Docs;
import com.google.api.services.docs.v1.model.*;
import com.googlesheets.demo.Entity.Person;
import com.googlesheets.demo.Repository.PersonRepository;
import com.googlesheets.demo.config.AutenificationPrintDocs;
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

import static org.docx4j.org.apache.poi.util.DocumentHelper.readDocument;

@Service
public class AllDataGoogleDocs {
    private static final Logger LOGGER = LoggerFactory.getLogger(GoogleDocsService.class);

    @Value("1S2KCBljZXwXFj-gRdv_I0f2VKNfEKZ4RdGJUDW42kc0")
    private String documentId;

    @Autowired
    private PersonRepository personRepository;
    @Autowired
    private AutenificationPrintDocs autenificationPrintDocs;
    @Autowired
    private DriverConfig driverConfig;
    public String readDocument(String DocumentId) throws IOException, GeneralSecurityException {
        Docs docsService = autenificationPrintDocs.getDocsService();
        Document document = docsService.documents().get(DocumentId).execute();

        StringBuilder textContent = new StringBuilder();
        List<StructuralElement> elements = document.getBody().getContent();
        for (StructuralElement element : elements) {
            if (element.getParagraph() != null) {
                for (ParagraphElement paragraphElement : element.getParagraph().getElements()) {
                    TextRun textRun = paragraphElement.getTextRun();
                    if (textRun != null) {
                        textContent.append(textRun.getContent());
                    }
                }
            }
        }

        return textContent.toString();
    }

    public List<String> fillPlaceholdersForAll() throws IOException, GeneralSecurityException {
        // Retrieve all persons from the database
        List<Person> allPersons = personRepository.findAll();

        // This list will hold the response message for each person's document creation
        List<String> responses = new ArrayList<>();

        // Create a Docs service instance once for efficiency
        Docs docsService = autenificationPrintDocs.getDocsService();

        // Create a Drive service instance once for efficiency
        com.google.api.services.drive.Drive driveService = driverConfig.getDriveService();

        for (Person person : allPersons) {
            String response = fillPlaceholdersForBatchedIndividual(person, docsService, driveService);
            responses.add(response);
        }

        return responses;
    }

    private String fillPlaceholdersForBatchedIndividual(Person person, Docs docsService, com.google.api.services.drive.Drive driveService) throws IOException, GeneralSecurityException {
        // Read the template document content (assuming it's same for everyone)
        String documentContent = readDocument(documentId);

        // Log or analyse the content if necessary
        LOGGER.info("Original content: " + documentContent);

        // Create a copy of the document
        com.google.api.services.drive.model.File file = new com.google.api.services.drive.model.File();
        file.setName(person.getLastName() + "_" + person.getId());
        List<String> parents = Collections.singletonList("1nULvtk_AvRaTu0lFZlVk6L0G-jAGS-Xr");
        file.setParents(parents);
        com.google.api.services.drive.model.File copiedFile = driveService.files().copy(documentId, file).execute();
        String newDocumentId = copiedFile.getId();

        Map<String, String> data = new HashMap<>();
        data.put("{id}", String.valueOf(person.getId()));

        data.put("{firstname}", person.getFirstName());
        data.put("{lastname}",person.getLastName());
        data.put("{gener}",person.getGender());
        data.put("{Title}",person.getTitle());
        data.put("{zoneImplpr}",person.getImplementationZone());
        data.put("{etatlocal}",person.getLocaleState());
        data.put("{age}", String.valueOf(person.getAge()));
        data.put("{nvetude}",person.getEducationLevel());
        data.put("{specialdiplome}",person.getDegreeSpecialty());
        data.put("{statutactual}",person.getCurrentStatus());
        data.put("{statutjuridique}",person.getCurrentLegalStatus());
        data.put("{descriptionprojet}",person.getProjectDescription());
        data.put("{etatduprojet}",person.getProjectState());
        data.put("{rhactual}", String.valueOf(person.getCurrentHR()));
        data.put("{rhprevisionel}", String.valueOf(person.getProjectedHR()));
        data.put("{region}",person.getRegion());
        List<Request> requests = data.entrySet().stream()
                .map(entry -> new Request()
                        .setReplaceAllText(new ReplaceAllTextRequest()
                                .setContainsText(new SubstringMatchCriteria()
                                        .setText(entry.getKey())
                                        .setMatchCase(true))
                                .setReplaceText(entry.getValue())))
                .collect(Collectors.toList());
        Docs docs = autenificationPrintDocs.getDocsService();
        BatchUpdateDocumentRequest batchUpdateRequest = new BatchUpdateDocumentRequest().setRequests(requests);

        // Update the new document
        docsService.documents().batchUpdate(newDocumentId, batchUpdateRequest).execute();

        return "Document for " + person.getLastName() + "_" + person.getId() + " updated";
    }

}
