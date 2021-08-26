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
        } catch (Exception e) {
            e.printStackTrace();
        }

        WorkingDay workingDay = new WorkingDay();

        // pobierz task

        workingDay.start();
        //workingDay.setStart(OffsetDateTime.now());
        System.out.println("Started working at: " + workingDay.getStart());
        try(Scanner scanner = new Scanner(System.in)) {
            String userStr = "";
            while(!userStr.equals("finish")) {
                System.out.println("Type 'finish' to end your work day: ");
                userStr = scanner.nextLine();
            }
        }catch (Exception e) {
            e.printStackTrace();
        }
        workingDay.end();

        try {
            workingDayDAO.persist(workingDay);
            System.out.println("komunikat do uzytkownika, ze zostalo zapisane w excelu i moze zamknac program");
        } catch (Exception e) {
            e.printStackTrace();
        }


//        workingDay.setEnd(OffsetDateTime.now());
//        System.out.println("Finished working at: " + workingDay.getEnd());
//        workingDay.calculateDuration();
//        System.out.println("Duration: " + workingDay.getFormattedDuration());
//        workingDay.writeToExcel(FILE_PATH);

    }
}
