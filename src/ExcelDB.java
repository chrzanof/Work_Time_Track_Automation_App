import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.OffsetDateTime;
import java.util.Iterator;

public class ExcelDB {
    private Path path;

    public ExcelDB(String path) throws Exception {
        if(!path.endsWith(".xlsx")) {
            throw new Exception("Wrong File format");
        } else {
            this.path = Paths.get(path);

            if(!Files.exists(this.path))
            {
                Files.createFile(this.path);
            }
            File file = new File(this.path.toString());
            if(file.length() == 0) {
                createFirstSheet(OffsetDateTime.now().getYear() + "-"+ OffsetDateTime.now().getMonthValue());
            }
        }
    }

    public int returnMonthFromLastSheetName() throws Exception {
        FileInputStream file = new FileInputStream(this.path.toString());
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        String sheetName = workbook.getSheetAt(workbook.getNumberOfSheets() - 1).getSheetName();
        String[] splitSheetName = sheetName.split("-");
        workbook.close();
        file.close();
        return Integer.parseInt(splitSheetName[1]);
    }

    public boolean isNowTheSameMonthAsInLastSheetName (WorkingDay workingDay) throws Exception {
        if(returnMonthFromLastSheetName() != workingDay.getStart().getMonthValue() || returnYearFromLastSheetName() != workingDay.getStart().getYear()) {
            return false;
        }
        return true;
    }

    public int returnYearFromLastSheetName() throws Exception {
        FileInputStream file = new FileInputStream(this.path.toString());
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        String sheetName = workbook.getSheetAt(workbook.getNumberOfSheets() - 1).getSheetName();
        String[] splitSheetName = sheetName.split("-");
        workbook.close();
        file.close();
        return Integer.parseInt(splitSheetName[0]);
    }

    public void createNextSheet(String sheetName) throws Exception {
        FileInputStream file = new FileInputStream(this.path.toString());
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        FileOutputStream output = new FileOutputStream(this.path.toString());
        XSSFSheet sheet1 = workbook.createSheet(sheetName);
        workbook.write(output);
        workbook.close();
        file.close();
        output.close();
    }

    public void createFirstSheet(String sheetName) throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook();
        FileOutputStream output = new FileOutputStream(this.path.toString());
        XSSFSheet sheet = workbook.createSheet(sheetName);
        workbook.write(output);
        workbook.close();
        output.close();
    }

    public int returnIndexOfEmptyRow() {

        int i = 0;
        try {
            FileInputStream file = new FileInputStream(this.path.toString());
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(workbook.getNumberOfSheets() - 1);
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Cell cell = row.getCell(0);
                String str = cell.getStringCellValue();
                if (str == null) {
                    break;
                } else {
                    i++;
                }
            }
            workbook.close();
            file.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return i;
    }

    public void addHeader() throws Exception {
        FileInputStream file = new FileInputStream(this.path.toString());
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(workbook.getNumberOfSheets() - 1);
        FileOutputStream outputStream = new FileOutputStream(this.path.toString());
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        Cell cell1 = row.createCell(1);
        Cell cell2 = row.createCell(2);
        Cell cell3 = row.createCell(3);
        cell.setCellValue("Start");
        cell1.setCellValue("End");
        cell2.setCellValue("Task Description");
        cell3.setCellValue("Duration");
        workbook.write(outputStream);
        workbook.close();
        file.close();
        outputStream.close();
    }

    public void addRecord(WorkingDay workingDay) throws Exception {
        int indexOfEmptyRow = returnIndexOfEmptyRow();
        FileInputStream file = new FileInputStream(this.path.toString());
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(workbook.getNumberOfSheets() - 1);
        FileOutputStream outputStream = new FileOutputStream(this.path.toString());
        Row row = sheet.createRow(indexOfEmptyRow);
        Cell cell = row.createCell(0);
        Cell cell1 = row.createCell(1);
        Cell cell2 = row.createCell(2);
        Cell cell3 = row.createCell(3);
        cell.setCellValue(workingDay.getFormattedTimeStamp(workingDay.getStart()));
        cell1.setCellValue(workingDay.getFormattedTimeStamp(workingDay.getEnd()));
        cell2.setCellValue(workingDay.getTask());
        cell3.setCellValue(workingDay.getFormattedDuration());
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
    }

    public Path getPath() {
        return path;
    }

    public void setPath(Path path) {
        this.path = path;
    }
}
