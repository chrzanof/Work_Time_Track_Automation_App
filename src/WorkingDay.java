import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.Duration;
import java.time.OffsetDateTime;
import java.util.Iterator;

public class WorkingDay {
    private OffsetDateTime start;
    private OffsetDateTime end;
    private Duration duration;
    private String task;

    public WorkingDay() {
    }

    public void writeToExcel(String path) {
        int indexOfEmptyRow = returnIndexOfEmptyRow(path);
        try
        {
            FileInputStream file = new FileInputStream(path);
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(workbook.getNumberOfSheets()-1);
            FileOutputStream outputStream = new FileOutputStream(path);
            if(indexOfEmptyRow == 0) {
                Row row = sheet.createRow(indexOfEmptyRow);
                Cell cell = row.createCell(0);
                Cell cell1 = row.createCell(1);
                Cell cell2 = row.createCell(2);
                cell.setCellValue("Start");
                cell1.setCellValue("Finish");
                cell2.setCellValue("Duration");
                indexOfEmptyRow ++;
            }
            Row row = sheet.createRow(indexOfEmptyRow);
            Cell cell = row.createCell(0);
            Cell cell1 = row.createCell(1);
            Cell cell2 = row.createCell(2);
            cell.setCellValue(returnTimeStampString(start));
            cell1.setCellValue(returnTimeStampString(end));
            cell2.setCellValue(getDurationAsString(duration));
            workbook.write(outputStream);
            workbook.close();
            file.close();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }

    public String getFormattedDuration() {
        if (this.duration == null) {
            return "Can not be calculated, start and end are required!";
        } else {
            return this.getDurationAsString(this.duration);
        }
    }

    private int returnIndexOfEmptyRow(String path) {
        int i = 0;
        try {
            FileInputStream file = new FileInputStream(path);
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


        } catch (Exception e) {
            e.printStackTrace();
        }
        return i;
    }

    public String returnTimeStampString(OffsetDateTime timeStamp) {
        int day = timeStamp.getDayOfMonth();
        int hour = timeStamp.getHour();
        int minute = timeStamp.getMinute();
        int second = timeStamp.getSecond();
        return timeStamp.getYear() + "-" + timeStamp.getMonth() + "-" + (day < 10 ? "0" + day : day) + " "
                + (hour < 10 ? "0" + hour : hour) + ":" + (minute < 10 ? "0" + minute : minute) + ":" + (second < 10 ? "0" + second : second);
    }

    private String getDurationAsString(Duration duration) {
        long durationSeconds = duration.getSeconds();
        int newHours = 0;
        int newMinutes = 0 ;
        int newSeconds;
        if (durationSeconds > 59) {
            newMinutes = (int) (durationSeconds / 60);
            newSeconds = (int) (durationSeconds % 60);
        }else {
            newSeconds = (int) durationSeconds;
        }
        if(newMinutes > 59){
            newHours = newMinutes / 60;
            newMinutes = newMinutes % 60;
        }
        return newHours + ":" + (newMinutes < 10 ? "0" + newMinutes : newMinutes) + ":" + (newSeconds < 10 ? "0" + newSeconds : newSeconds) ;
    }

    public OffsetDateTime getStart() {
        return start;
    }

    public void start() {
        this.start = OffsetDateTime.now();
    }

    public OffsetDateTime getEnd() {
        return end;
    }

    public void end() {
        this.end = OffsetDateTime.now();
    }

    public Duration getDuration() {
        return duration;
    }

    public void calculateDuration() {
        if (start != null && end != null) {
            this.duration = Duration.between(start, end);
        }
    }

    public void setStart(OffsetDateTime now) {
        this.start = now;
    }

    public void setEnd(OffsetDateTime now) {
        this.end = now;
    }

    public void setDuration(Duration duration) {
        this.duration = duration;
    }

    public String getTask() {
        return task;
    }

    public void setTask(String task) {
        this.task = task;
    }
}
