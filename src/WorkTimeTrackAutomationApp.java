import java.io.File;
import java.util.Scanner;


public class WorkTimeTrackAutomationApp {
    public static final String FILE_PATH = "testfile.xlsx";
    public static ExcelDB db;
    public static WorkingDayDAO  workingDayDAO;
    public static void main(String[] args) {
        try {
            db = new ExcelDB(FILE_PATH);
            workingDayDAO = new WorkingDayDAO(db);

            WorkingDay workingDay = new WorkingDay();
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
            scanner.close();
            workingDay.end();
            System.out.println("Finished working at: " + workingDay.getFormattedTimeStamp(workingDay.getEnd()));
            workingDay.calculateDuration();
            System.out.println("Today's duration: " + workingDay.getFormattedDuration());
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
