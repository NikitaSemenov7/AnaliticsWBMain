import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;

public class ColumnNamesCreator {

    public static ArrayList<String> createColumnNames(XSSFSheet sheet) {
        ArrayList<String> columnNames = new ArrayList<>();
        int i = 0;
        while (true) {
            XSSFCell cell = sheet.getRow(0).getCell(i);
            if (cell != null) {
                String name = cell.getStringCellValue();
                columnNames.add(name);
                i++;
            } else {
                break;
            }
        }
        return columnNames;
    }
}
