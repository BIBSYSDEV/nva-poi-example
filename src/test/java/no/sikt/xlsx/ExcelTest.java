package no.sikt.xlsx;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;

class ExcelTest {

    @TempDir
    static Path temp;

    @Test
    void shouldCreateExcelSheetFromTemplate() throws IOException {
        var data = List.of(List.of("My", "First", "Row"), List.of("My", "Second", "Røw"));
        var xlsx = Excel.fromTemplate(data);
        var path = Paths.get(temp.toString(), "This_NEW_file.xlsx");
        xlsx.write(path);
        var actual = getActualWithoutHeaders(path);
        assertEquals(data, actual);
    }

    @Test
    void shouldCreateExcelSheetFromJava() throws IOException {
        var data = List.of(List.of("My", "First", "Row"), List.of("My", "Second", "Røw"));
        var xlsx = Excel.fromJava(data);
        var path = Paths.get(temp.toString(), "This_OTHER_NEW_file.xlsx");
        xlsx.write(path);
        var actual = getActualWithoutHeaders(path);
        assertEquals(data, actual);
    }


    private List<List<String>> getActualWithoutHeaders(Path path) throws IOException {
        var actual = readFromExcel(path);
        actual.remove(0);
        return actual;
    }

    private List<List<String>> readFromExcel(Path path) throws IOException {
        var file = new FileInputStream(path.toAbsolutePath().toString());
        var workbook = new XSSFWorkbook(file);
        var sheet = workbook.getSheetAt(0);

        var data = new ArrayList<List<String>>();
        for (Row row : sheet) {
            var currentRow = new ArrayList<String>();
            data.add(currentRow);
            for (Cell cell : row) {
                currentRow.add(cell.getStringCellValue());
            }
        }
        return data;
    }

}
