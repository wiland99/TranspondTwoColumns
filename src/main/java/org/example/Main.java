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

    public static void main(String[] args) throws IOException, RuntimeException {
        String desktopUrl = System.getProperty("user.home") + "/Desktop/";
        Iterator<Row> rows = new XSSFWorkbook(new FileInputStream(desktopUrl + "Calendar.xlsx")).getSheetAt(0).rowIterator();
        Workbook workBookFinal = new XSSFWorkbook();
        Sheet sheetFinal = workBookFinal.createSheet("final");
        rows.next();
        int[] i = new int[3];
        rows.forEachRemaining(row -> {
            Iterator<Cell> cells = row.cellIterator();
            IntStream.range(0, 2).forEach(n-> cells.next());
            final LocalDate[] startDate = {LocalDate.of(Integer.parseInt(row.getCell(1).getStringCellValue()), Month.JANUARY, 1)};
            cells.forEachRemaining(cell -> Arrays.stream(cell.getStringCellValue().split("")).forEach(workDayOrNot -> {
                i[2] = 0;
                Arrays.asList(new Object[]{startDate[0].toString(), i[1], startDate[0]}).forEach(n -> sheetFinal.createRow(i[0]++).createCell(i[2]++).setCellValue(String.valueOf(n)));
                i[1] = i[1] + Integer.parseInt(workDayOrNot);
                startDate[0] = startDate[0].plusDays(1);
            }));
        });
        IntStream.range(0, 3).forEach(sheetFinal::autoSizeColumn);
        workBookFinal.write(new FileOutputStream(desktopUrl + "final.xlsx"));
    }
}

