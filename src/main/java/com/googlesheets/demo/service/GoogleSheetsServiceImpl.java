package com.googlesheets.demo.service;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.BatchGetValuesResponse;
import com.google.api.services.sheets.v4.model.Sheet;
import com.google.api.services.sheets.v4.model.Spreadsheet;
import com.googlesheets.demo.Entity.Person;
import com.googlesheets.demo.Repository.PersonRepository;
import com.googlesheets.demo.config.GoogleAuthorizationConfig;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.Collections;
import java.util.List;

@Service
public class GoogleSheetsServiceImpl implements GoogleSheetsService {

    private static final Logger LOGGER = LoggerFactory.getLogger(GoogleSheetsServiceImpl.class);




    @Autowired
    private PersonRepository personRepository;
    @Autowired
    private GoogleAuthorizationConfig googleAuthorizationConfig;

    @Override
    public void getSpreadsheetValues(String spreadsheetId) throws IOException, GeneralSecurityException {
        Sheets sheetsService = googleAuthorizationConfig.getSheetsService();
        Sheets.Spreadsheets.Values.BatchGet request =
                sheetsService.spreadsheets().values().batchGet(spreadsheetId);
        request.setRanges(getSpreadSheetRange(spreadsheetId));
        request.setMajorDimension("ROWS");
        BatchGetValuesResponse response = request.execute();
        List<List<Object>> spreadSheetValues = response.getValueRanges().get(0).getValues();
        List<Object> headers = spreadSheetValues.remove(0);
        for ( List<Object> row : spreadSheetValues ) {
           LOGGER.info(row +" | ");
            // Create a new instance of the Person entity


            Person person = new Person();

            // Set the values from the row to the corresponding properties of the Person entity
            person.setId(Long.valueOf(row.get(0).toString()));
            person.setLastName(row.get(1).toString());
            person.setFirstName(row.get(1).toString());
            person.setGender(row.get(2).toString());
            person.setImplementationZone(row.get(3).toString());
            person.setTitle(row.get(4).toString());
            person.setLocaleState(row.get(5).toString());
            person.setAge(Integer.parseInt(row.get(6).toString()));
            person.setEducationLevel(row.get(7).toString());
            person.setDegreeSpecialty(row.get(8).toString());
            person.setCurrentStatus(row.get(9).toString());
            person.setCurrentLegalStatus(row.get(10).toString());
            person.setProjectDescription(row.get(11).toString());
            person.setProjectState(row.get(12).toString());
            person.setHasReceivedFunding(Boolean.parseBoolean(row.get(13).toString()));
            person.setCurrentHR(Integer.parseInt(row.get(14).toString()));
            person.setProjectedHR(Integer.parseInt(row.get(15).toString()));
            person.setRegion(row.get(16).toString());

            // Save the Person entity to the database
            personRepository.save(person);


        }



    }

    private List<String> getSpreadSheetRange(String spreadsheetId) throws IOException, GeneralSecurityException {
        Sheets sheetsService = googleAuthorizationConfig.getSheetsService();
        Sheets.Spreadsheets.Get request = sheetsService.spreadsheets().get(spreadsheetId);
        Spreadsheet spreadsheet = request.execute();
        Sheet sheet = spreadsheet.getSheets().get(0);
        int row = sheet.getProperties().getGridProperties().getRowCount();
        int col = sheet.getProperties().getGridProperties().getColumnCount();
        return Collections.singletonList("R1C1:R".concat(String.valueOf(row))
                .concat("C").concat(String.valueOf(col)));
    }


}
