import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.channels.Channel;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.OffsetDateTime;
import java.util.List;

public class ExcelDB {
    private Path path;

    public ExcelDB(String path) throws Exception {
        if(!path.endsWith(".xlsx")) {
            throw new Exception("Wrong File format. It has to be.xlsx");
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
    public boolean isNowTheSameMonthAsInLastSheetName (WorkingDay workingDay) throws Exception {
        if(getMonthFromLastSheetName() != workingDay.getStart().getMonthValue() || getYearFromLastSheetName() != workingDay.getStart().getYear()) {
            return false;
        }
        return true;
    }

    public String getLastSheetName() throws Exception{
        FileInputStream file = new FileInputStream(this.path.toString());
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        return workbook.getSheetAt(workbook.getNumberOfSheets() - 1).getSheetName();
    }

    public int getMonthFromLastSheetName() throws Exception {
        String[] splitSheetName = getLastSheetName().split("-");
        return Integer.parseInt(splitSheetName[1]);
    }

    public int getYearFromLastSheetName() throws Exception {
        String[] splitSheetName = getLastSheetName().split("-");
        return Integer.parseInt(splitSheetName[0]);
    }

    public void createNextSheet(String sheetName) throws Exception {
        FileInputStream file = new FileInputStream(this.path.toString());
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        FileOutputStream output = new FileOutputStream(this.path.toString());
        workbook.createSheet(sheetName);
        workbook.write(output);
        workbook.close();
        file.close();
        output.close();
    }

    public void createFirstSheet(String sheetName) throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook();
        FileOutputStream output = new FileOutputStream(this.path.toString());
        workbook.createSheet(sheetName);
        workbook.write(output);
        workbook.close();
        output.close();
    }

    public int getIndexOfEmptyRowFromLastSheet() {

//        int i = 0;
//        try {
//            FileInputStream file = new FileInputStream(this.path.toString());
//            XSSFWorkbook workbook = new XSSFWorkbook(file);
//            XSSFSheet sheet = workbook.getSheetAt(workbook.getNumberOfSheets() - 1);
//            Iterator<Row> rowIterator = sheet.iterator();
//            while (rowIterator.hasNext()) {
//                Row row = rowIterator.next();
//                Cell cell = row.getCell(0);
//                String str = cell.getStringCellValue();
//                if (str == null) {
//                    break;
//                } else {
//                    i++;
//                }
//            }
//            workbook.close();
//            file.close();
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
        int i = 0;
        try {
            FileInputStream file = new FileInputStream(this.path.toString());
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(workbook.getNumberOfSheets() - 1);
            for (i = 0; i < 1048576; i++) {
                String str = sheet.getRow(i).getCell(0).getStringCellValue();
            }
            workbook.close();
            file.close();
        } catch (Exception e) {

        }
        return i;
    }

    public void addRecord(int indexOfRow, List<String> cellValues) throws Exception {
        FileInputStream file = new FileInputStream(this.path.toString());
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(workbook.getNumberOfSheets() - 1);
        FileOutputStream outputStream = new FileOutputStream(this.path.toString());
        Row row = sheet.createRow(indexOfRow);
        for(int i = 0; i < cellValues.size(); i++){
            Cell cell = row.createCell(i);
            cell.setCellValue(cellValues.get(i));
        }
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
    }
    public boolean isFileClosed() {
        boolean closed;
        Channel channel = null;
        try {
            File file = new File(this.path.toString());
            channel = new RandomAccessFile(file, "rw").getChannel();
            closed = true;
        } catch(Exception ex) {
            closed = false;
        } finally {
            if(channel!=null) {
                try {
                    channel.close();
                } catch (IOException ex) {
                    // exception handling
                    ex.printStackTrace();
                }
            }
        }
        return closed;
    }

    public Path getPath() {
        return path;
    }

    public void setPath(Path path) {
        this.path = path;
    }
}
