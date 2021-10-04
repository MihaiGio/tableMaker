import org.apache.commons.io.input.BOMInputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.StandardCharsets;

public class ConvertToExcel {
    public static void main(String[] args) throws IOException {
        String fileName = "C:\\Users\\mgiosan\\OneDrive - Signant Health\\Desktop\\Texts.txt";
        String excelFileName = "C:\\Users\\mgiosan\\OneDrive - Signant Health\\Desktop\\newTexts.xlsx";
        BOMInputStream bomIn = new BOMInputStream(new FileInputStream(fileName));

// Create a Workbook and a sheet in it
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sheet1");
        XSSFCellStyle style = workbook.createCellStyle();
        style.setWrapText(true);
        style.setVerticalAlignment(VerticalAlignment.TOP);

// Read your input file and make cells into the workbook
        try (BufferedReader br = new BufferedReader(new InputStreamReader(bomIn, StandardCharsets.UTF_16))) {
            String line;
            Row row;
            Cell cell;
            int rowIndex = 0;
            while ((line = br.readLine()) != null) {
                row = sheet.createRow(rowIndex);
                String[] tokens = line.split("[\\t]"); // delimit by TAB regex
                for (int iToken = 0; iToken < tokens.length; iToken++) {
                    cell = row.createCell(iToken);
                    cell.setCellValue(tokens[iToken]);
                    cell.setCellStyle(style);
                }
                rowIndex++;
            }
            sheet.autoSizeColumn(0);
            sheet.autoSizeColumn(1);
            sheet.autoSizeColumn(2);
            sheet.autoSizeColumn(3);
            sheet.setColumnWidth(4, 7000);
            sheet.setColumnWidth(5, 10000);
        } catch (Exception e) {
            e.printStackTrace();
        }


// Write your xlsx file
        try (FileOutputStream outputStream = new FileOutputStream(excelFileName)) {
            workbook.write(outputStream);
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
