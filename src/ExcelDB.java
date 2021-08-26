import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;

public class ExcelDB {
    private Path path;

    public ExcelDB(String path) throws Exception {
        // sprawdz czy xlsx? czy poprawny plik bazy excelowej
        // jezeli to nie jest poprawny plik to thorw new Excepton

        this.path = Paths.get(path);

        if (this.returnIndexOfEmptyRow() == 0) {
            this.addHeader();
        }
    }

    public int returnIndexOfEmptyRow() {
        // co jak Path nie jest zdefinowane?

        return 0;
    }

    public void addHeader() {
        // czy w excelu jest nagłówek czy nie? jezeli nie to dodadaj nagłówek
    }

    public void addRecord(WorkingDay workingDay) throws Exception {
        // dodaje nowy wiersz
        // wyrzuca wszystkei excpetion
    }
}
