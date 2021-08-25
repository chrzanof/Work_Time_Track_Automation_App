import java.util.Scanner;

public class WorkTimeTrackAutomationApp {
    public static final String FILE_PATH = "testfile.xlsx";
    public static void main(String[] args) {
        WorkingDay workingDay = new WorkingDay();
        workingDay.setStart();
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
        workingDay.setFinish();
        System.out.println("Finished working at: " + workingDay.getFinish());
        workingDay.calculateDuration();
        System.out.println("Duration:" + workingDay.getDuration());
        workingDay.writeToExcel(FILE_PATH);

    }
}
