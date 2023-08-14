package com.googlesheets.demo.service;

import com.googlesheets.demo.Entity.Person;
import com.googlesheets.demo.Repository.PersonRepository;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;




@Service
@Slf4j
public class ExcelService {

    private final PersonRepository personRepository;

    public ExcelService(PersonRepository personRepository) {
        this.personRepository = personRepository;
    }

    public void importDataFromExcel(InputStream inputStream) {
        try (Workbook workbook = WorkbookFactory.create(inputStream)) {
            Sheet sheet = workbook.getSheetAt(0); // Assuming data is in the first sheet

            List<Person> persons = new ArrayList<>();

            // Iterate over rows (skip the header row)
            Iterator<Row> rowIterator = sheet.iterator();
            rowIterator.next(); // Skip the header row

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                Person person = new Person();

               try {
                    person.setId((long) row.getCell(0).getNumericCellValue());
                    log.info("setId: " +  row.getCell(0).getNumericCellValue());
                } catch (NumberFormatException e) {
                    log.error("Invalid ID format at row " + (row.getRowNum() + 1));
                    continue; // Skip this row and continue with the next row
                }



                person.setLastName(getStringCellValue(row.getCell(1)));
                person.setFirstName(getStringCellValue(row.getCell(2)));
                person.setGender(getStringCellValue(row.getCell(3)));
                person.setImplementationZone(getStringCellValue(row.getCell(4)));
                person.setTitle(getStringCellValue(row.getCell(5)));
                person.setLocaleState(getStringCellValue(row.getCell(6)));

                try {
                    person.setAge((int) row.getCell(7).getNumericCellValue());
                } catch (NumberFormatException e) {
                    log.error("Invalid age format at row " + (row.getRowNum() + 1));
                    continue; // Skip this row and continue with the next row
                }

                person.setEducationLevel(getStringCellValue(row.getCell(8)));
                person.setDegreeSpecialty(getStringCellValue(row.getCell(9)));
                person.setCurrentStatus(getStringCellValue(row.getCell(10)));
                person.setCurrentLegalStatus(getStringCellValue(row.getCell(11)));
                person.setProjectDescription(getStringCellValue(row.getCell(12)));
                person.setProjectState(getStringCellValue(row.getCell(13)));

                try {
                    person.setHasReceivedFunding(Boolean.parseBoolean(getStringCellValue(row.getCell(14))));
                } catch (IllegalArgumentException e) {
                    log.error("Invalid hasReceivedFunding format at row " + (row.getRowNum() + 1));
                    continue; // Skip this row and continue with the next row
                }

                try {
                    person.setCurrentHR(row.getCell(15).getNumericCellValue());
                } catch (NumberFormatException e) {
                    log.error("Invalid currentHR format at row " + (row.getRowNum() + 1));
                    continue; // Skip this row and continue with the next row
                }

                try {
                    person.setProjectedHR(row.getCell(16).getNumericCellValue());
                } catch (NumberFormatException e) {
                    log.error("Invalid projectedHR format at row " + (row.getRowNum() + 1));
                    continue; // Skip this row and continue with the next row
                }

                person.setRegion(getStringCellValue(row.getCell(17)));

                persons.add(person);
            }

            personRepository.saveAll(persons); // Save the imported data to the database
        } catch (IOException | InvalidFormatException e) {
            log.error("Failed to import data from Excel file.", e);
            throw new RuntimeException("Failed to import data from Excel file.");
        }
    }

    private String getStringCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }

        CellType cellType = CellType.forInt(cell.getCellType());

        if (cellType == CellType.STRING) {
            return cell.getStringCellValue();
        } else if (cellType == CellType.NUMERIC) {
            if (DateUtil.isCellDateFormatted(cell)) {
                DateFormat df = new SimpleDateFormat("yyyy-MM-dd");
                return df.format(cell.getDateCellValue());
            } else {
                return NumberToTextConverter.toText(cell.getNumericCellValue());
            }
        } else if (cellType == CellType.BOOLEAN) {
            return String.valueOf(cell.getBooleanCellValue());
        } else if (cellType == CellType.FORMULA) {
            return cell.getCellFormula();
        } else if (cellType == CellType.BLANK) {
            return "";
        } else {
            return null;
        }
    }

}
