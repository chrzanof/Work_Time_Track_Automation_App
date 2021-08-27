import java.time.OffsetDateTime;

public class WorkingDayDAO implements IWorkingDayDAO {
    private ExcelDB db;

    public WorkingDayDAO(ExcelDB db) {
        this.db = db;
    }

    @Override
    public void persist(WorkingDay workingDay) throws Exception {
        if(!db.isNowTheSameMonthAsInLastSheetName(workingDay))
        {
            db.createNextSheet(OffsetDateTime.now().getYear() + "-"+ OffsetDateTime.now().getMonthValue());
        }
        int indexOfEmptyRow = db.returnIndexOfEmptyRow();
        if(indexOfEmptyRow == 0){
            db.addRecord(indexOfEmptyRow, workingDay.getHeader());
            indexOfEmptyRow++;
        }
        db.addRecord(indexOfEmptyRow, workingDay.getFormattedValues());
    }
}
