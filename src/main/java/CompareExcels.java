import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class CompareExcels {

    public static void main(String[] args) {
        BufferedReader br = new BufferedReader(new InputStreamReader(System.in));
        while (true) {
            try {
                System.out.println("enter the path to the first file");
                String input = br.readLine();
                File file1 = new File(input);

                System.out.println("enter the path to the second file");
                input = br.readLine();
                File file2 = new File(input);
                FileInputStream excellFile1 = new FileInputStream(file1);
                FileInputStream excellFile2 = new FileInputStream(file2);

                XSSFWorkbook workbook1 = new XSSFWorkbook(excellFile1);
                XSSFWorkbook workbook2 = new XSSFWorkbook(excellFile2);

                System.out.println("enter the sheet number for the first excel file (starts from 0)");
                input = br.readLine();
                XSSFSheet sheet1 = workbook1.getSheetAt(Integer.parseInt(input));

                System.out.println("enter the sheet number for the second excel file (starts from 0)");
                input = br.readLine();
                XSSFSheet sheet2 = workbook2.getSheetAt(Integer.parseInt(input));


                System.out.println("enter the column number of the first file (starts from 0)");
                input = br.readLine();
                List<String> columnValues1 = populatingCellValues(sheet1, Integer.parseInt(input));
                System.out.println(columnValues1);

                System.out.println("enter the column number of the second file (starts from 0)");
                input = br.readLine();
                List<String> columnValues2 = populatingCellValues(sheet2, Integer.parseInt(input));
                System.out.println(columnValues2);

                List<String> commonValues = mergeArrays(columnValues1, columnValues2);
                System.out.println(commonValues);


                System.out.println("enter the location for the merged file");
                input = br.readLine();
                FileOutputStream out = new FileOutputStream(new File(input));
                XSSFWorkbook mergedWorkbook = new XSSFWorkbook();
                XSSFSheet mergedSheet = mergedWorkbook.createSheet("merged");

                Row row = mergedSheet.createRow(0);

                row.createCell(0).setCellValue("Merged");
                row.createCell(1).setCellValue(file1.getName());
                row.createCell(2).setCellValue(file2.getName());

                int rowNum = 1;
                for (String s : commonValues) {

                    Row row1 = mergedSheet.createRow(rowNum++);
                    row1.createCell(0).setCellValue(s);
                    if (columnValues1.contains(s)) {
                        row1.createCell(1).setCellValue("+");
                    } else {
                        row1.createCell(1).setCellValue("-");
                    }
                    if (columnValues2.contains(s)) {
                        row1.createCell(2).setCellValue("+");
                    } else {
                        row1.createCell(2).setCellValue("-");
                    }
                }
                mergedWorkbook.write(out);
                out.close();
                System.out.println("Files were merged successfully");


                excellFile1.close();
                excellFile2.close();

            } catch (Exception e) {
                e.printStackTrace();
            }
        }

    }

    private static List<String> populatingCellValues(XSSFSheet sheet, int number) {
        List<String> columnValues = new ArrayList<String>();
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                Cell cell = row.getCell(number);
                if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK) {
                    columnValues.add(String.valueOf(cell));
                }
            }
        }
        return columnValues;
    }

    private static List<String> mergeArrays(List<String> columnValues1, List<String> columnValues2) {
        List<String> commonValues = new ArrayList<String>();
        commonValues.addAll(columnValues2);
        for (String o : columnValues1) {
            if (!commonValues.contains(o)) {
                commonValues.add(o);
            }
        }
        return commonValues;
    }

}
