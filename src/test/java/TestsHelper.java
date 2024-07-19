import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;

public class TestsHelper {

    public double createUnitEcomonicaAndGetPrihodNaScetWB(String path) throws IOException, InterruptedException {
        //ExcelConverter.mainShadow(path);
        FileInputStream fileInputStream = new FileInputStream(path + "UnitЭкономика.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);
        double prihodNaSchetWB = sheet.getRow(24).getCell(1).getNumericCellValue();
        BigDecimal bd = new BigDecimal(prihodNaSchetWB);
        bd = bd.setScale(2, RoundingMode.HALF_UP);
        prihodNaSchetWB = bd.doubleValue();
        fileInputStream.close();
        return prihodNaSchetWB;
    }
}
