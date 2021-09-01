

import java.io.File;
import java.util.Scanner;


public class WorkTimeTrackAutomationApp {
    public static final String DEFAULT_FILE_PATH = "Working_Hours.xlsx";
    public static final String TEXT_RED = "\033[31m";
    public static final String TEXT_DEFAULT = "\033[0m";
    public static ExcelDB db;
    public static WorkingDayDAO  workingDayDAO;
    public static void main(String[] args) {
        try {
            if(args.length == 0) {
                db = new ExcelDB(DEFAULT_FILE_PATH);
            } else {
                db = new ExcelDB(args[0]);
            }
            workingDayDAO = new WorkingDayDAO(db);
            WorkingDay workingDay = new WorkingDay();
            System.out.println("Please make sure, that your saving file is closed. Otherwise, the app won't save your data.");
            System.out.println("Type task description: ");
            Scanner scanner = new Scanner(System.in);
            workingDay.setTask(scanner.nextLine());
            workingDay.start();
            System.out.println("Started working at: " + workingDay.getFormattedTimeStamp(workingDay.getStart()));
            String userStr = "";
            while (!userStr.equals("end")) {
                System.out.println("Type 'end' to end your work day: ");
                userStr = scanner.nextLine();
            }
            workingDay.end();
            System.out.println("Finished working at: " + workingDay.getFormattedTimeStamp(workingDay.getEnd()));
            workingDay.calculateDuration();
            System.out.println("Today's duration: " + workingDay.getFormattedDuration());
            while (!db.isFileClosed()){
                    System.out.println(TEXT_RED + "Your file is open by another process! Close file and type anything to continue:" + TEXT_DEFAULT);
                    scanner.nextLine();
            }
            scanner.close();
            try {
                workingDayDAO.persist(workingDay);
                System.out.println("Data successfully saved. You can close program now.");
            } catch (Exception e) {
                e.printStackTrace();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }


    }
}
