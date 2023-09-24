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
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Iterator;
import java.util.stream.IntStream;

public class Main {

    public static void main(String[] args) throws IOException, RuntimeException {
        String desktopUrl = System.getProperty("user.home") + "/Desktop/";
        Iterator<Row> rows = new XSSFWorkbook(new FileInputStream(desktopUrl + "Calendar.xlsx")).getSheetAt(0).rowIterator();
        Workbook workBookFinal = new XSSFWorkbook();
        Sheet sheetFinal = workBookFinal.createSheet("final");
        rows.next();
        int[] i = new int[3];
        i[2] = 1;
        rows.forEachRemaining(row -> {
            Iterator<Cell> cells = row.cellIterator();
            IntStream.range(0, 2).forEach(n -> cells.next());

            final LocalDate[] startDate = new LocalDate[]{LocalDate.of(Integer.parseInt(row.getCell(1).getStringCellValue()), Month.JANUARY, 1)};
            cells.forEachRemaining(cell -> {
                for (char number : cell.getStringCellValue().toCharArray()) {
                    i[1] = 0;
                    i[2] = i[2] + Integer.parseInt(String.valueOf(number));
                    Row row2 = sheetFinal.createRow(i[0]++);
                    Arrays.asList(startDate[0].toString(), String.valueOf(number), String.valueOf(i[2])).forEach(n -> row2.createCell(i[1]++).setCellValue(n));
                    startDate[0] = startDate[0].plusDays(1);
                }
            });
        });
        IntStream.range(0, 3).forEach(sheetFinal::autoSizeColumn);
        workBookFinal.write(new FileOutputStream(desktopUrl + "final.xlsx"));

    }
}
