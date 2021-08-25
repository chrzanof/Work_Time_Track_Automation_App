import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.Duration;
import java.time.OffsetDateTime;
import java.util.Iterator;

public class WorkingDay {
    OffsetDateTime start;
    OffsetDateTime finish;
    Duration duration;

    public OffsetDateTime getStart() {
        return start;
    }

    public void setStart() {
        this.start = OffsetDateTime.now();
    }

    public OffsetDateTime getFinish() {
        return finish;
    }

    public void setFinish() {
        this.finish = OffsetDateTime.now();
    }

    public Duration getDuration() {
        return duration;
    }

    public void calculateDuration() {
        this.duration = Duration.between(start, finish);
    }

    public void writeToExcel(String path) {
        int indexOfEmptyRow = returnIndexOfEmptyRow(path);
        try
        {
            FileInputStream file = new FileInputStream(path);
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(workbook.getNumberOfSheets()-1);
            FileOutputStream outputStream = new FileOutputStream(path);
            Row row = sheet.createRow(indexOfEmptyRow);
            Cell cell = row.createCell(0);
            Cell cell1 = row.createCell(1);
            Cell cell2 = row.createCell(2);
            cell.setCellValue(returnTimeStampString(start));
            cell1.setCellValue(returnTimeStampString(finish));
            cell2.setCellValue(returnDuration(duration));
            workbook.write(outputStream);
            workbook.close();
            file.close();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }

    private int returnIndexOfEmptyRow(String path) {
        int i = 0;
        try {
            FileInputStream file = new FileInputStream(path);
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(0);
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
                + (hour < 10 ? "0" + hour : hour) + ":" + (minute < 10 ? "0" + minute : minute) + ":" + (second < 10 ? "" + second : second);
    }
    public String returnDuration(Duration duration) {
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
            newMinutes = newMinutes / 60;
        }
        return newHours + ":" + (newMinutes < 10 ? "0" + newMinutes : newMinutes) + ":" + (newSeconds < 10 ? "0" + newSeconds : newSeconds) ;
    }
    }
