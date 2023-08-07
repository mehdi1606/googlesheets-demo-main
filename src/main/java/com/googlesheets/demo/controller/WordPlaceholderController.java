    package com.googlesheets.demo.controller;// ... (imports and annotations)

    import com.googlesheets.demo.Entity.Person;
    import com.googlesheets.demo.Repository.PersonRepository;
    import lombok.extern.slf4j.Slf4j;
    import org.apache.poi.xwpf.usermodel.XWPFDocument;
    import org.apache.poi.xwpf.usermodel.XWPFParagraph;
    import org.apache.poi.xwpf.usermodel.XWPFRun;
    import org.springframework.beans.factory.annotation.Autowired;
    import org.springframework.web.bind.annotation.GetMapping;
    import org.springframework.web.bind.annotation.PathVariable;
    import org.springframework.web.bind.annotation.RestController;

    import java.io.FileInputStream;
    import java.io.FileNotFoundException;
    import java.io.FileOutputStream;
    import java.io.IOException;
    import java.util.List;
    import java.util.Optional;
    // ... (imports and annotations)
    @Slf4j
    @RestController
    public class WordPlaceholderController {

        private final PersonRepository personRepository;

        @Autowired
        public WordPlaceholderController(PersonRepository personRepository) {
            this.personRepository = personRepository;
        }

        @GetMapping("/fill-word-template/{id}")
        public String fillWordTemplate(@PathVariable Long id) {
            // Retrieve person data from the database based on personId
            Optional<Person> optionalPerson = personRepository.findById(Double.valueOf(id));
            if (optionalPerson.isEmpty()) {
                return "Person not found!";
            }

            Person person = optionalPerson.get();

            try {
                String templateFilePath = "C:\\Users\\hp\\Documents\\Bonjour.docx"; // Use a relative path or properties here
                String outputFilePath = "C:\\Users\\hp\\Desktop\\googlesheets-demo-main\\src\\main\\resources\\Xupload.docx"; // Use a relative path or properties here

                // Read the content of the template file into a XWPFDocument
                XWPFDocument doc = new XWPFDocument(new FileInputStream(templateFilePath));

                // Loop through all the paragraphs in the document
                for (XWPFParagraph paragraph : doc.getParagraphs()) {
                    // Loop through all the runs (portions of text) in the paragraph
                    for (XWPFRun run : paragraph.getRuns()) {
                        String text = run.getText(0);
                        // Check if the run contains the {{firstname}} placeholder

                        // Check if the run contains the {{lastname}} placeholder
                        if (text != null&& text.contains("$lastname") ) {
                            // Replace the {{lastname}} placeholder with the person's last name
                            text = text.replace("{{lastname}}", person.getLastName());
                            run.setText(text, 0); // Set the updated text in the run
                        }
                        else{
                            return "dosnt find {{lastname}} and {{firstname}}  ";
                        }

                        // Add more conditions for other placeholders if needed
                    }
                }

                // Save the updated content as a new Word document
                saveDocument(doc, outputFilePath);

                return "Word document filled successfully!";
            } catch (IOException e) {
                e.printStackTrace();
                return "Error filling Word document.";
            }
        }


        private void saveDocument(XWPFDocument doc, String filePath) throws IOException {
            try (FileOutputStream fileOutputStream = new FileOutputStream(filePath)) {
                doc.write(fileOutputStream);
            }
        }
    }
