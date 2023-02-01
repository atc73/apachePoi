import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

//Workbook-->Sheet-->Rows-->Cells
public class WritingExcel {
    public static void main(String[] args) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Emp Info");

        ArrayList<Object[]> empdata = new ArrayList<Object[]>();

        empdata.add(new Object[]{"Empid", "Name", "Job"});
        empdata.add(new Object[]{101, "David", "Enginner"});
        empdata.add(new Object[]{102, "Smith", "Manager"});
        empdata.add(new Object[]{103, "Scott", "Analyst"});

        //Using for Loop
//        int rows = empdata.length;
//        int cols = empdata[0].length;
//
//        System.out.println(rows);
//        System.out.println(cols);
//
//        for(int r = 0; r < rows; r++) {
//
//            XSSFRow row = sheet.createRow(r);
//
//            for(int c = 0; c < cols; c++) {
//
//                XSSFCell cell = row.createCell(c);
//                Object value = empdata[r][c];
//
//                if(value instanceof String)
//                    cell.setCellValue((String) value);
//                if(value instanceof Integer)
//                    cell.setCellValue((Integer) value);
//                if(value instanceof Boolean)
//                    cell.setCellValue((Boolean) value);
//            }
//        }

        //Using for...each loop
        int rowNum = 0;
        for(Object[] emp: empdata) {
            XSSFRow row = sheet.createRow(rowNum++);
            int cellNum = 0;

            for (Object value : emp) {
                XSSFCell cell = row.createCell(cellNum++);

                if(value instanceof String)
                    cell.setCellValue((String) value);
                if(value instanceof Integer)
                    cell.setCellValue((Integer) value);
                if(value instanceof Boolean)
                    cell.setCellValue((Boolean) value);
            }
        }

        String filePath = ".\\datafiles\\employee.xlsx";
        FileOutputStream outputStream = new FileOutputStream(filePath);
        workbook.write(outputStream );

        outputStream.close();

        System.out.println("Employee.xlsx file written successfully");

    }
}
