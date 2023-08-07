package com.googlesheets.demo.service;
import com.google.api.services.docs.v1.Docs;
import com.google.api.services.docs.v1.DocsScopes;
import com.google.api.services.docs.v1.model.*;

import com.googlesheets.demo.Entity.Person;
import com.googlesheets.demo.Repository.PersonRepository;
import com.googlesheets.demo.config.AutenificationPrint;
import com.googlesheets.demo.config.DriverConfig;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.*;
import java.util.stream.Collectors;

import java.util.List;

@Service
public class GoogleDocsService {
    private static final Logger LOGGER = LoggerFactory.getLogger(GoogleDocsService.class);

    @Value("1S2KCBljZXwXFj-gRdv_I0f2VKNfEKZ4RdGJUDW42kc0")
    private String documentId;

    @Autowired
    private PersonRepository personRepository;
    @Autowired
    private AutenificationPrint autenificationPrint;
    @Autowired
    private DriverConfig driverConfig;





    List<String> scopes = Collections.singletonList(DocsScopes.DOCUMENTS);




    public String readDocument(String DocumentId) throws IOException, GeneralSecurityException {
        Docs docsService = autenificationPrint.getDocsService();
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

    public String fillPlaceholders( Double personId) throws IOException, GeneralSecurityException {
        // Lire le contenu du document
        String documentContent = readDocument(documentId);

        // Log ou analyse du contenu si nécessaire
        LOGGER.info("Original content: " + documentContent);

        // Récupération de la personne et remplissage des placeholders comme avant
        Optional<Person> personData = personRepository.findById(personId);
        if (!personData.isPresent()) {
            throw new NoSuchElementException("Person not found");
        }
        // Create a copy of the document
        com.google.api.services.drive.Drive driveService = driverConfig.getDriveService();
        com.google.api.services.drive.model.File file = new com.google.api.services.drive.model.File();
        file.setName("DocumentPerson");
        List<String> parents = Collections.singletonList("1nULvtk_AvRaTu0lFZlVk6L0G-jAGS-Xr");
        file.setParents(parents);

        com.google.api.services.drive.model.File copiedFile =
                driveService.files().copy(documentId, file).execute();


        // Get the new document's ID
        String newDocumentId = copiedFile.getId();


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
        data.put("{rhactual}", String.valueOf(personData.get().getCurrentHR()));
        data.put("{rhprevisionel}", String.valueOf(personData.get().getProjectedHR()));
        data.put("{region}",personData.get().getRegion());
        List<Request> requests = data.entrySet().stream()
                .map(entry -> new Request()
                        .setReplaceAllText(new ReplaceAllTextRequest()
                                .setContainsText(new SubstringMatchCriteria()
                                        .setText(entry.getKey())
                                        .setMatchCase(true))
                                .setReplaceText(entry.getValue())))
                .collect(Collectors.toList());

        Docs docs = autenificationPrint.getDocsService();
        BatchUpdateDocumentRequest batchUpdateRequest = new BatchUpdateDocumentRequest().setRequests(requests);

        // Update the new document
        docs.documents().batchUpdate(newDocumentId, batchUpdateRequest).execute();

        return "New document updated";

    }


}


