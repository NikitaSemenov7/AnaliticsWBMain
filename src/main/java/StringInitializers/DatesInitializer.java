package StringInitializers;

import enumsMaps.Alphabet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class DatesInitializer {
    public static void initDates(String directoryReports, Sheet sheetReportConverted){
        int index = directoryReports.indexOf("Отчеты");
        String dates = directoryReports.substring(index + 6, index + 27);
        Row row = sheetReportConverted.createRow(0);
        row.createCell(Alphabet.A.ordinal()).setCellValue("Даты");
        row.createCell(Alphabet.B.ordinal()).setCellValue(dates);
    }
}
