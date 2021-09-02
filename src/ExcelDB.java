import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.channels.Channel;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.OffsetDateTime;
import java.util.Date;
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
        int i = 0;
        try {
            FileInputStream file = new FileInputStream(this.path.toString());
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(workbook.getNumberOfSheets() - 1);
            for (i = 0; i < 1048576; i++) {
                XSSFRow row = sheet.getRow(i);
                if(row == null) {
                    break;
                }
            }
            workbook.close();
            file.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return i;
    }

    public void addRecord(int indexOfRow, List<Object> cellValues) throws Exception {
        FileInputStream file = new FileInputStream(this.path.toString());
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(workbook.getNumberOfSheets() - 1);
        FileOutputStream outputStream = new FileOutputStream(this.path.toString());
        Row row = sheet.createRow(indexOfRow);
        CellStyle dateStyle = workbook.createCellStyle();
        CellStyle durationStyle = workbook.createCellStyle();
        CreationHelper createHelper = workbook.getCreationHelper();
        dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy-MM-dd hh:mm:ss"));
        durationStyle.setDataFormat(createHelper.createDataFormat().getFormat("hh:mm:ss"));
        int cellnum = 0;
        for(Object obj : cellValues) {
            Cell cell = row.createCell(cellnum);
            if(obj instanceof OffsetDateTime) {
                cell.setCellValue(((OffsetDateTime) obj).toLocalDateTime());
                cell.setCellStyle(dateStyle);
            } else if(obj instanceof String) {
                cell.setCellValue((String)obj);
            }
            sheet.autoSizeColumn(cellnum);
            cellnum++;
        }
        if (indexOfRow != 0) {
            String formula = "MOD(B" + (indexOfRow+1) + "-A" + (indexOfRow+1) + ",1)";
            Cell cell1 = row.createCell(cellnum);
            cell1.setCellFormula(formula);
            cell1.setCellStyle(durationStyle);
            sheet.autoSizeColumn(cellnum);
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
