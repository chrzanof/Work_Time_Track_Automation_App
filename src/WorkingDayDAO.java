public class WorkingDayDAO implements IWorkingDayDAO {
    private ExcelDB db;

    public WorkingDayDAO(ExcelDB db) {
        this.db = db;
    }

    @Override
    public void persist(WorkingDay workingDay) throws Exception {
        db.addRecord(workingDay);
    }
}
