import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;
import java.util.StringTokenizer;

public class readDocx {
    public static void main(String[] args) {
        //FILE NAME COLLECT
        File folder = new File("D://filecount//");
        File[] listOfFiles = folder.listFiles();
        String[] filenames=new String[listOfFiles.length];
        for (int i = 0; i < listOfFiles.length; i++) {
            if (listOfFiles[i].isFile()) {
                //System.out.println("File " + listOfFiles[i].getName());
                //filenames[i]="D://filecount//" + listOfFiles[i].getName();
                filenames[i]= listOfFiles[i].getName();

            } else if (listOfFiles[i].isDirectory()) {
                //System.out.println("Directory " + listOfFiles[i].getName());
            }
        }
        /*for (int i = 0; i < filenames.length; i++) {
            System.out.println(filenames[i]);
        }*/

        //FILE READ
        int[] wordCount=new int[listOfFiles.length];
        try {
            for(int i=0; i<wordCount.length; i++) {
                FileInputStream fis = new FileInputStream("D://filecount//" + filenames[i]);
                XWPFDocument xdoc = new XWPFDocument(OPCPackage.open(fis));
                XWPFWordExtractor extractor = new XWPFWordExtractor(xdoc);
                String needToCount = extractor.getText();
                StringTokenizer st = new StringTokenizer(needToCount);
                int count = st.countTokens();
                wordCount[i] = count;
                //System.out.println(count);
            }
        } catch(Exception ex) {
            ex.printStackTrace();
        }


        // creating workbook
        XSSFWorkbook workbook = new XSSFWorkbook();
        // creating sheet with name "Report" in workbook
        XSSFSheet sheet = workbook.createSheet("Report");
        // this method creates header for our table


        int rowCount = 0;
        for (int i=0; i<filenames.length; i++) {
            // creating row
            Row row = sheet.createRow(++rowCount);

            // adding first cell to the row
            Cell idCell = row.createCell(0);
            idCell.setCellValue(filenames[i]);

            // adding second cell to the row
            Cell nameCell = row.createCell(1);
            nameCell.setCellValue(wordCount[i]);
            /*
            //adding third cell to the row
            Cell statusCell = row.createCell(2);
            statusCell.setCellValue(user.lastName);*/

        }
        String FILE_SAVE_LOCATION = "D:\\reports\\";
        String FILE_NAME = "UserReport.xlsx";
        try (FileOutputStream outputStream = new FileOutputStream(FILE_SAVE_LOCATION + FILE_NAME)) {
            workbook.write(outputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            // don't forget to close workbook to prevent memory leaks
           // workbook.close();
        }



    }
}
