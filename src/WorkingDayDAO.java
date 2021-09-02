import java.time.OffsetDateTime;

public class WorkingDayDAO implements IWorkingDayDAO {
    private ExcelDB db;

    public WorkingDayDAO(ExcelDB db) {
        this.db = db;
    }

    @Override
    public void persist(WorkingDay workingDay) throws Exception {
        while (!db.isFileClosed()){
            System.out.println("YOUR FILE IS OPENED BY ANOTHER PROCESS! PLEASE CLOSE THE FILE.");
            WorkTimeTrackAutomationApp.pressEnterKeyToContinue();
        }
        if(!db.isNowTheSameMonthAsInLastSheetName(workingDay))
        {
            db.createNextSheet(OffsetDateTime.now().getYear() + "-"+ OffsetDateTime.now().getMonthValue());
        }
        int indexOfEmptyRow = db.getIndexOfEmptyRowFromLastSheet();
        if(indexOfEmptyRow == 0){
            db.addRecord(indexOfEmptyRow, workingDay.getHeader());
            indexOfEmptyRow++;
        }
        db.addRecord(indexOfEmptyRow, workingDay.getValuesWithoutDuration());
    }
}
