import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.time.OffsetDateTime;
import java.util.*;

public class WorkingDay {
    private OffsetDateTime start;
    private OffsetDateTime end;
    private Duration duration;
    private String task;

    public WorkingDay() {
    }

    public List<String> getHeader () {
        List<String> header = new ArrayList(Arrays.asList("Start", "End", "Task Description", "Duration"));
        return header;
    }


    public List<String> getFormattedValues() throws NullPointerException{
            List<String> values = new ArrayList();
            values.add(this.getFormattedTimeStamp(start));
            values.add(this.getFormattedTimeStamp(end));
            values.add(this.task);
            values.add(this.getFormattedDuration());
            return values;
    }

    public String getFormattedDuration() {
        if (this.duration == null) {
            return "Can not be calculated, start and end are required!";
        } else {
            return this.getDurationAsString(this.duration);
        }
    }

    public String getFormattedTimeStamp(OffsetDateTime timeStamp) {
        if(timeStamp == null) {
            return "time Stamp is not set yet!";
        } else {
            return  this.getTimeStampAsString(timeStamp);
        }
    }

    private String getTimeStampAsString(OffsetDateTime timeStamp) {
        int month = timeStamp.getMonth().getValue();
        int day = timeStamp.getDayOfMonth();
        int hour = timeStamp.getHour();
        int minute = timeStamp.getMinute();
        int second = timeStamp.getSecond();
        return timeStamp.getYear() + "-" + (month < 10 ? "0" + month : month) + "-" + (day < 10 ? "0" + day : day) + " "
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
        return (newHours < 10 ? "0" + newHours : newHours) + ":" + (newMinutes < 10 ? "0" + newMinutes : newMinutes) + ":" + (newSeconds < 10 ? "0" + newSeconds : newSeconds) ;
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
