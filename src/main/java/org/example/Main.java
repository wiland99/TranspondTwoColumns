package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.Month;
import java.util.Arrays;
import java.util.Iterator;
import java.util.stream.IntStream;

public class Main {
    public static void main(String[] args) throws IOException {
        String desktopUrl = System.getProperty("user.home") + "/Desktop/";
        Iterator<Row> rows = new XSSFWorkbook(new FileInputStream(desktopUrl + "Calendar.xlsx")).getSheetAt(0).rowIterator();
        Workbook workBookFinal = new XSSFWorkbook();
        Sheet sheetFinal = workBookFinal.createSheet("final");
        rows.next();
        int[] i = new int[3];
        i[2] = 1;
        while (rows.hasNext()) {
            Row row = rows.next();
            Iterator<Cell> cells = row.cellIterator();
            cells.next();
            int year = Integer.parseInt(cells.next().getStringCellValue());
            LocalDate[] startDate = new LocalDate[1]
            startDate[0] = LocalDate.of(year, Month.JANUARY, 1);


cells.forEachRemaining(cell -> {
for(char c : cell.toCharArray()) {

                    i[1] = 0;
                    i[2] += Character.getNumericValue(c);
                    Row row2 = sheetFinal.createRow(i[0]++);
                    for (char c : cell) {
                    i[1] = 0;
                    i[2] += Character.getNumericValue(c);
                    Row row2 = sheetFinal.createRow(i[0]++);
                    for (Object n : Arrays.asList(startDate[0], c, i[2])) {
                        row2.createCell(i[1]++).setCellValue(String.valueOf(n));
                    }
                    startDate[0] = startDate[0].plusDays(1);
                }
}
            while (cells.hasNext()) {
                char[] cell = cells.next().getStringCellValue().toCharArray();
                
            }
        }
        IntStream.range(0, 3).forEach(sheetFinal::autoSizeColumn);
        workBookFinal.write(new FileOutputStream(desktopUrl + "final.xlsx"));
        workBookFinal.close();
    }
}
