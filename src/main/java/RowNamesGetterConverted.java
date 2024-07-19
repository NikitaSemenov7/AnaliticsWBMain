import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;

//получает названия столбцов в sheetReportConverted

public class RowNamesGetterConverted {
    private static int counter = 4;

    public static ArrayList<String> getRowNamesConverted(XSSFSheet sheetReportConverted) {
        ArrayList<String> names = new ArrayList<>();
        while (true) {
            XSSFCell cell = sheetReportConverted.getRow(4).getCell(0);
            if (cell != null && !cell.getStringCellValue().equals("Общий")) {
                String name = cell.getStringCellValue();
                names.add(name);
                counter++;
            } else {
                counter += 2;
                break;
            }
        }

        return names;
    }
}
