package com.googlesheets.demo.controller;

import com.googlesheets.demo.Entity.Person;
import com.googlesheets.demo.Repository.PersonRepository;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xslf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

@RestController
@Slf4j
public class PowerPointController {

    @Autowired
    private PersonRepository personRepository;

    @GetMapping("/ppt")
    public String createPresentation() {
        try {
            XMLSlideShow ppt = new XMLSlideShow();

            // First Slide (Title Slide)
            XSLFSlideMaster defaultMaster = ppt.getSlideMasters().get(0);
            XSLFSlideLayout titleLayout = defaultMaster.getLayout(SlideLayout.TITLE);
            XSLFSlide titleSlide = ppt.createSlide(titleLayout);
            XSLFTextShape titleShape = titleSlide.getPlaceholder(0);
            titleShape.setText("Presentation Title"); // Set your desired title here

            // Second Slide (Table with all person data)
            XSLFSlideLayout tableLayout = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
            XSLFSlide tableSlide = ppt.createSlide(tableLayout);
            XSLFTextShape title = tableSlide.getPlaceholder(0);
            title.setText("Person Data");

            List<Person> personList = personRepository.findAll();
            int numColumns = 4; // Assuming four columns for Lastname, Firstname, Age, and Gender
            int numRows = personList.size() + 1; // +1 for the header row
            double tableWidth = ppt.getPageSize().width * 0.8;
            double tableHeight = ppt.getPageSize().height * 0.6;
            double columnWidth = tableWidth / numColumns;

            XSLFTable table = tableSlide.createTable(numRows, numColumns);
            table.setAnchor(new Rectangle2D.Double((ppt.getPageSize().width - tableWidth) / 2, (ppt.getPageSize().height - tableHeight) / 2, tableWidth, tableHeight));

            // Populate table headers
            XSLFTableRow headerRow = table.getRows().get(0);
            headerRow.addCell().setText("Lastname");
            headerRow.addCell().setText("Firstname");
            headerRow.addCell().setText("Age");
            headerRow.addCell().setText("Gender");
            headerRow.addCell().setText("Implementation Zone");
            headerRow.addCell().setText("Title");
            headerRow.addCell().setText("Locale State");
            headerRow.addCell().setText("Education Level");
            headerRow.addCell().setText("Degree Specialty");
            headerRow.addCell().setText("Region");

// Populate table with person data
            for (int i = 0; i < personList.size(); i++) {
                Person person = personList.get(i);
                XSLFTableRow dataRow = table.getRows().get(i + 1);
                dataRow.addCell().setText(person.getLastName());
                dataRow.addCell().setText(person.getFirstName());
                dataRow.addCell().setText(String.valueOf(person.getAge()));
                dataRow.addCell().setText(person.getGender());
                dataRow.addCell().setText(person.getImplementationZone());
                dataRow.addCell().setText(person.getTitle());
                dataRow.addCell().setText(person.getLocaleState());
                dataRow.addCell().setText(person.getEducationLevel());
                dataRow.addCell().setText(person.getDegreeSpecialty());
                dataRow.addCell().setText(person.getRegion());
            }

            // Slides for each person
            XSLFSlideLayout personLayout = defaultMaster.getLayout(SlideLayout.BLANK); // Customize layout as needed

            for (Person person : personList) {
                XSLFSlide personSlide = ppt.createSlide(personLayout);
                XSLFTextShape personTextShape = personSlide.createTextBox();
                personTextShape.setText(
                        "First Name: " + person.getFirstName() +
                                "\nAge: " + person.getAge() +
                                "\nLast Name: " + person.getLastName() +
                                "\nGender: " + person.getGender());
            }

            String fileName = "powerpoint.pptx";
            FileOutputStream out = new FileOutputStream(fileName);
            ppt.write(out);
            out.close();

            return "Presentation created successfully. File saved as " + fileName;
        } catch (IOException e) {
            log.error("Error creating PowerPoint presentation: {}", e.getMessage(), e);
            return "Error creating PowerPoint presentation. Please check the logs for details.";
        }
    }
}
