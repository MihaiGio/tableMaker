import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Testing {

    public static XSSFWorkbook readWorkbook() {

        XSSFWorkbook wb = null;
        try {
            //wb = XSSFWorkbookFactory.create(new File("C:\\Users\\Mihai\\Desktop\\texting.xlsx"));
            FileInputStream file = new FileInputStream("C:\\Users\\mgiosan\\OneDrive - Signant Health\\Desktop\\newTexts.xlsx");
            wb = new XSSFWorkbook(file);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return wb;
    }

    public static void writeToSheet(List<String> languageID, List<String> formID, List<String> formName, List<String> textID,
                                    List<String> comment, List<String> text, int sheetCounter, List<Sheet> outputSheets) {


        for (int i = 0; i < languageID.size(); i++) {

            Row row = outputSheets.get(sheetCounter).createRow(i);

            Cell languageIDCell = row.createCell(0);
            languageIDCell.setCellValue(languageID.get(i));

            Cell formIDCell = row.createCell(1);
            formIDCell.setCellValue(formID.get(i));

            Cell formNameCell = row.createCell(2);
            formNameCell.setCellValue(formName.get(i));

            Cell textIDCell = row.createCell(3);
            textIDCell.setCellValue(textID.get(i));

            Cell commentCell = row.createCell(4);
            commentCell.setCellValue(comment.get(i));

            Cell textCell = row.createCell(5);
            textCell.setCellValue(text.get(i));

        }
    }

    public static void writeToWorkbook(Workbook wb) {

        try {
            FileOutputStream out = new FileOutputStream("C:\\Users\\mgiosan\\OneDrive - Signant Health\\Desktop\\newTextsFile.xlsx");
            wb.write(out);
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {

        XSSFWorkbook inputWb = readWorkbook();
        Sheet inputWs = inputWb.getSheet("Sheet1");


        List<String> languageID = new ArrayList<>();
        List<String> formID = new ArrayList<>();
        List<String> formName = new ArrayList<>();
        List<String> textID = new ArrayList<>();
        List<String> comment = new ArrayList<>();
        List<String> text = new ArrayList<>();

        XSSFWorkbook outputWb = new XSSFWorkbook();

        //Styling------------------------------------------------
        XSSFCellStyle outputStyle = outputWb.createCellStyle();
        outputStyle.setWrapText(true);
        outputStyle.setVerticalAlignment(VerticalAlignment.TOP);
        //-----------------------------------------------------
        //Why does it bold everything?
        //XSSFCellStyle headerStyle = outputWb.createCellStyle();
        //headerStyle.getFont().setBold(true);
        //-------------------------------------------------------

        List<Sheet> outputSheets = new ArrayList<>();

        int rowIndex = inputWs.getLastRowNum() + 1;
        int sheetCounter = 0;

        for (int i = 1; i < rowIndex - 1; i++) {
            Row outerRow = inputWs.getRow(i);
            Row innerRow = null;
            Cell outerCell = CellUtil.getCell(outerRow, 2); // Column to filter comparison
            Cell innerCell = null;


            int j = 0;
            for (j = i + 1; j < rowIndex; j++) { //j = i + 1 -> second row moved at the bottom of the page
                // j = i -> second row copied at the bottom of the page -> maybe delete after?

                innerRow = inputWs.getRow(j);
                // Cell innerCell = innerRow.getCell(2);
                innerCell = CellUtil.getCell(innerRow, 2); // Column to filter comparison

                if (outerCell.getStringCellValue().equals(innerCell.getStringCellValue())) {
                    languageID.add(CellUtil.getCell(innerRow, 0).getStringCellValue());
                    formID.add(CellUtil.getCell(innerRow, 1).getStringCellValue());
                    formName.add(CellUtil.getCell(innerRow, 2).getStringCellValue());
                    textID.add(CellUtil.getCell(innerRow, 3).getStringCellValue());
                    comment.add(CellUtil.getCell(innerRow, 4).getStringCellValue());
                    text.add(CellUtil.getCell(innerRow, 5).getStringCellValue());
                }


                if (!outerCell.getStringCellValue().equals(innerCell.getStringCellValue())) {

                    break;
                }
            }
            languageID.add(CellUtil.getCell(outerRow, 0).getStringCellValue());
            formID.add(CellUtil.getCell(outerRow, 1).getStringCellValue());
            formName.add(CellUtil.getCell(outerRow, 2).getStringCellValue());
            textID.add(CellUtil.getCell(outerRow, 3).getStringCellValue());
            comment.add(CellUtil.getCell(outerRow, 4).getStringCellValue());
            text.add(CellUtil.getCell(outerRow, 5).getStringCellValue());
            i = j;
            outputSheets.add(outputWb.createSheet("(" + sheetCounter + ")" + outerCell));
            writeToSheet(languageID, formID, formName, textID, comment, text, sheetCounter, outputSheets);
            sheetCounter++;
            languageID.clear();
            formID.clear();
            formName.clear();
            textID.clear();
            comment.clear();
            text.clear();

            Row tempRow = inputWs.getRow(i);
            try {
                languageID.add(tempRow.getCell(0).getStringCellValue());
                formID.add(tempRow.getCell(1).getStringCellValue());
                formName.add(tempRow.getCell(2).getStringCellValue());
                textID.add(tempRow.getCell(3).getStringCellValue());
                comment.add(tempRow.getCell(4).getStringCellValue());
                text.add(tempRow.getCell(5).getStringCellValue());
            } catch (Exception e) {
                continue;
            }
        }

        // iterating through sheets and formatting/styling them
        for (int i = 0; i < outputWb.getNumberOfSheets(); i++) {
            Sheet sheet = outputWb.getSheetAt(i);
            //adding header to all sheets
            sheet.shiftRows(0, sheet.getLastRowNum(), 1);
            Row header = sheet.createRow(0);
            //Header cells hardcoded text------------------
            CellUtil.getCell(header, 0).setCellValue("Language ID");
            CellUtil.getCell(header, 1).setCellValue("Form ID");
            CellUtil.getCell(header, 2).setCellValue("Form name");
            CellUtil.getCell(header, 3).setCellValue("Text ID");
            CellUtil.getCell(header, 4).setCellValue("Comment");
            CellUtil.getCell(header, 5).setCellValue("Text");
            //--------------------------------------------------
//            CellUtil.getCell(header,0).setCellStyle(headerStyle);
//            CellUtil.getCell(header,1).setCellStyle(headerStyle);
//            CellUtil.getCell(header,2).setCellStyle(headerStyle);
//            CellUtil.getCell(header,3).setCellStyle(headerStyle);
//            CellUtil.getCell(header,4).setCellStyle(headerStyle);
//            CellUtil.getCell(header,5).setCellStyle(headerStyle);
            //-----------------------------------------------------
            sheet.autoSizeColumn(0);
            sheet.autoSizeColumn(1);
            sheet.autoSizeColumn(2);
            sheet.autoSizeColumn(3);
            sheet.setColumnWidth(4, 7000);
            sheet.setColumnWidth(5, 10000);

            // iterate over rows and cells to wrap text
            for (int j = 0; j < sheet.getLastRowNum() + 1; j++) {
                Row row = sheet.getRow(j);
                CellUtil.getCell(row, 0).setCellStyle(outputStyle);
                CellUtil.getCell(row, 1).setCellStyle(outputStyle);
                CellUtil.getCell(row, 2).setCellStyle(outputStyle);
                CellUtil.getCell(row, 3).setCellStyle(outputStyle);
                CellUtil.getCell(row, 4).setCellStyle(outputStyle);
                CellUtil.getCell(row, 5).setCellStyle(outputStyle);
            }

        }
        writeToWorkbook(outputWb);
    }
}
