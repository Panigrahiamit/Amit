package project.manualTc;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.xwpf.usermodel.*;

public class TRDocumentation {
    private static final String CSV_DELIMITER = ",";

    public static class TestCase {
        private String testName;         
        private String tcNumber;         
        private String userstoryNumber;  
        private String description;      
        private String preCondition;     
        private List<String> designSteps;

        public TestCase() {
            this.designSteps = new ArrayList<>();
        }

        public String getTestName() { return testName; }
        public void setTestName(String testName) {
            this.testName = testName;
            if (testName != null && testName.startsWith("TC")) {
                this.tcNumber = testName.split("_")[0];
            }
        }
        public String getTcNumber() { return tcNumber; }
        public String getuserstoryNumber() { return userstoryNumber; }
        public void setuserstoryNumber(String userstoryNumber) { this.userstoryNumber = userstoryNumber; }
        public String getDescription() { return description; }
        public void setDescription(String description) { this.description = description; }
        public String getPreCondition() { return preCondition; }
        public void setPreCondition(String preCondition) { this.preCondition = preCondition; }
        public List<String> getDesignSteps() { return designSteps; }
        public void addDesignStep(String step, String description) {
            if (step != null && !step.trim().isEmpty() && description != null && !description.trim().isEmpty()) {
                String fullStep = step.trim() + " - " + description.trim();
                this.designSteps.add(fullStep);
            }
        }
    }

    public static List<TestCase> readTestCasesFromCSV(String filePath) {
        List<TestCase> testCases = new ArrayList<>();
        try (BufferedReader br = new BufferedReader(new FileReader(filePath))) {
            // Skip header
            String headerLine = readCompleteCsvLine(br);
            if (headerLine == null) {
                throw new IOException("CSV file is empty");
            }

            TestCase currentTestCase = null;
            String currentTcNumber = null;

            String line;
            while ((line = readCompleteCsvLine(br)) != null) {
                String[] values = line.split(CSV_DELIMITER, -1);

                // Get the test name from column B (index 1)
                String testName = values.length > 1 ? values[1].trim() : "";
                
                if (!testName.isEmpty() && testName.startsWith("TC")) {
                    String newTcNumber = testName.split("_")[0];
                    if (!newTcNumber.equals(currentTcNumber)) {
                        currentTestCase = new TestCase();
                        currentTestCase.setTestName(testName);
                        if (values.length > 0) {
                            currentTestCase.setuserstoryNumber(values[0].trim());
                        }
                        if (values.length > 2) {
                            currentTestCase.setDescription(values[2].trim());
                        }
                        if (values.length > 10) {
                            currentTestCase.setPreCondition(values[10].trim());
                        }
                        testCases.add(currentTestCase);
                        currentTcNumber = newTcNumber;
                    }
                }

                // Add design steps
                if (currentTestCase != null && values.length > 12) {
                    String stepNumber = values[11].trim();
                    String stepDescription = values[12].trim();
                    if (!stepNumber.isEmpty() && !stepDescription.isEmpty()) {
                        currentTestCase.addDesignStep(stepNumber, stepDescription);
                    }
                }
            }
        } catch (IOException e) {
            System.err.println("Error reading CSV file: " + e.getMessage());
            e.printStackTrace();
        }
        return testCases;
    }

    // Utility to read a complete CSV record, handling quoted fields that may span multiple lines
    private static String readCompleteCsvLine(BufferedReader br) throws IOException {
        String line = br.readLine();
        if (line == null) return null;

        int quoteCount = countQuotes(line);
        while (quoteCount % 2 != 0) {
            String nextLine = br.readLine();
            if (nextLine == null) break;
            line += "\n" + nextLine;
            quoteCount += countQuotes(nextLine);
        }
        return line;
    }

    private static int countQuotes(String s) {
        int count = 0;
        for (char c : s.toCharArray()) {
            if (c == '"') count++;
        }
        return count;
    }

    public static void main(String[] args) {
        String csvFilePath = "C:\\Users\\amit.panigrahi\\Desktop\\Format - Copy (2).csv";
        
        String outputDir = "C:\\Users\\amit.panigrahi\\Documents\\Test Rail Demo Documentation\\";

        try {
            File outputDirectory = new File(outputDir);
            if (!outputDirectory.exists()) {
                outputDirectory.mkdirs();
            }
            if (!outputDirectory.canWrite()) {
                System.err.println("No write permissions for output directory.");
                return;
            }
            List<TestCase> testCases = readTestCasesFromCSV(csvFilePath);
            if (testCases.isEmpty()) {
                System.err.println("No test cases were read from the CSV file.");
                return;
            }
            for (TestCase tc : testCases) {
                String fileName = tc.getTestName() + ".docx";
                String outputPath = outputDir + fileName;
                generateWordDocument(tc, outputPath);
            }
            System.out.println("Successfully generated " + testCases.size() + " test case documents.");
        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            e.printStackTrace();
        }
    }

    public static void generateWordDocument(TestCase testCase, String outputPath) {
        try (XWPFDocument document = new XWPFDocument()) {
            // Title
            XWPFParagraph title = document.createParagraph();
            XWPFRun titleRun = title.createRun();
            titleRun.setText("Test Case: " + testCase.getTestName());
            titleRun.setBold(true);
            titleRun.setFontSize(16);
            title.addRun(titleRun);

            // Test Case Number
            XWPFParagraph tcNumberParagraph = document.createParagraph();
            XWPFRun tcNumberRun = tcNumberParagraph.createRun();
            tcNumberRun.setText("Test Case Number: " + testCase.getTcNumber());
            tcNumberRun.setFontSize(12);
            tcNumberParagraph.addRun(tcNumberRun);

            // User Story Number
            XWPFParagraph userstoryParagraph = document.createParagraph();
            XWPFRun usLabelRun = userstoryParagraph.createRun();
            usLabelRun.setText("User Story Number: ");
            usLabelRun.setBold(true);
            usLabelRun.setFontSize(12);
            userstoryParagraph.addRun(usLabelRun);

            XWPFRun usValueRun = userstoryParagraph.createRun();
            usValueRun.setText(testCase.getuserstoryNumber());
            usValueRun.setFontSize(12);

            // Description
            XWPFParagraph descriptionParagraph = document.createParagraph();
            XWPFRun descriptionRun = descriptionParagraph.createRun();
            descriptionRun.setText("Description: " + testCase.getDescription());
            descriptionRun.setFontSize(12);
            descriptionParagraph.addRun(descriptionRun);

            // Pre-condition
            XWPFParagraph preConditionParagraph = document.createParagraph();
            XWPFRun preConditionRun = preConditionParagraph.createRun();
            preConditionRun.setText("Pre-condition: " + testCase.getPreCondition());
            preConditionRun.setFontSize(12);
            preConditionParagraph.addRun(preConditionRun);

            // Test Data
            XWPFParagraph testdataTitle = document.createParagraph();
            XWPFRun testdataTitleRun = testdataTitle.createRun();
            testdataTitleRun.setText("Test Data:");
            testdataTitleRun.setBold(true);
            testdataTitleRun.setFontSize(12);
            testdataTitle.addRun(testdataTitleRun);

            // Design Steps
            if (!testCase.getDesignSteps().isEmpty()) {
                XWPFParagraph stepsTitle = document.createParagraph();
                XWPFRun stepsTitleRun = stepsTitle.createRun();
                stepsTitleRun.setText("Design Steps:");
                stepsTitleRun.setBold(true);
                stepsTitleRun.setFontSize(12);
                stepsTitle.addRun(stepsTitleRun);

                for (String step : testCase.getDesignSteps()) {
                    XWPFParagraph stepParagraph = document.createParagraph();
                    XWPFRun stepRun = stepParagraph.createRun();
                    stepRun.setText(step);
                    stepRun.setFontSize(12);
                    stepParagraph.addRun(stepRun);
                }
            }
            try (FileOutputStream out = new FileOutputStream(outputPath)) {
                document.write(out);
            }
        } catch (IOException e) {
            System.err.println("Error while generating the Word document: " + e.getMessage());
            e.printStackTrace();
        }
    }
}