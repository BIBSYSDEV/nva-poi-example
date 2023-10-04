package no.sikt.xlsx;

import nva.commons.core.attempt.Try;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;

import static nva.commons.core.attempt.Try.attempt;

public class Excel {

    public static final String TEMPLATE = "/test_template.xlsx";
    public static final List<String> HEADERS = List.of("Not", "My", "Problem");
    private final Workbook workbook;

    private Excel(Workbook workbook) {
        this.workbook = workbook;
    }

    public static Excel fromJava(List<List<String>> data) {
        var excel = new Excel(new XSSFWorkbook());
        var sheet = excel.workbook.createSheet("My lovely sheet");
        addHeadersToSheet(sheet);
        addDataToSheet(data, sheet);
        return excel;
    }

    public static Excel fromTemplate(List<List<String>> data) {
         var excel = new Excel(getTemplate());
         var sheet = excel.workbook.getSheetAt(0);
         addDataToSheet(data, sheet);
         return excel;
    }

    public void write(Path path) throws IOException {
        var absolutePath = path.toAbsolutePath().toString();
        var outputStream = attempt(() -> new FileOutputStream(absolutePath))
                .orElseThrow(failure -> new RuntimeException("Could not write file: " + path));
        workbook.write(outputStream);
        workbook.close();
    }

    private static Workbook getTemplate() {
            return Optional.ofNullable(Excel.class.getResourceAsStream(TEMPLATE))
                    .map(template -> attempt(() -> new XSSFWorkbook(template)))
                    .map(Try::orElseThrow)
                    .orElseThrow(() -> new RuntimeException("Could not read template"));
    }


    private static void addHeadersToSheet(Sheet sheet) {
        var header = sheet.createRow(0);
        var headerStyle = getCellStyle(sheet);

        for (var headerCounter = 0; headerCounter < HEADERS.size(); headerCounter++) {
            var headerCell = header.createCell(headerCounter);
            headerCell.setCellStyle(headerStyle);
            headerCell.setCellValue(HEADERS.get(headerCounter));
        }
    }

    private static CellStyle getCellStyle(Sheet sheet) {
        var headerStyle = sheet.getWorkbook().createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return headerStyle;
    }

    private static void addDataToSheet(List<List<String>> data, Sheet sheet) {
        for (var counter = 0; counter < data.size(); counter++) {
            var currentRow = sheet.createRow(counter + 1);
            var rowData = data.get(counter);
            for (var subCounter = 0; subCounter < rowData.size(); subCounter++) {
                var currentCell = currentRow.createCell(subCounter);
                currentCell.setCellValue(rowData.get(subCounter));
            }
        }
    }
}
