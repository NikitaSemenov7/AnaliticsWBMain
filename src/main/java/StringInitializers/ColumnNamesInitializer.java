package StringInitializers;

import enumsMaps.ColumnNames;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ColumnNamesInitializer {
    public static int counter = 0;

    public static void initAllColumnNames(Sheet sheetReportConverted) {
        counter = 0;
        Row row = sheetReportConverted.createRow(1);
        for (ColumnNames columnName: ColumnNames.values()) {
            row.createCell(counter).setCellValue(columnName.getName());
            counter++;
        }
    }
}

