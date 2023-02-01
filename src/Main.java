import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class Main {
    public static void main(String[] args) throws IOException {


        String excelFilePath=".\\datafiles\\collecte.xlsx";

        FileInputStream inputStream = new FileInputStream(excelFilePath);

        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

//        XSSFSheet sheet = workbook.getSheet("Accueil");
        XSSFSheet sheet = workbook.getSheetAt(0);

        Iterator iterator = sheet.iterator();
         while(iterator.hasNext()) {
            XSSFRow row = (XSSFRow) iterator.next();

            Iterator cellIterator =  row.cellIterator();

            while(cellIterator.hasNext()) {
                XSSFCell cell = (XSSFCell) cellIterator.next();

                switch (cell.getCellType()) {
                    case STRING:
                        System.out.println(cell.getStringCellValue());
                        break;
                    case NUMERIC:
                        System.out.println(cell.getNumericCellValue());
                        break;
                    case BOOLEAN:
                        System.out.println(cell.getBooleanCellValue());
                        break;
                }
                System.out.print(" | ");
            }
             System.out.println();
         }
    }
}