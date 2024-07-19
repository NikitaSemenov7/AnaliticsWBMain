import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;

public class TestsHelper {

    public double createUnitEcomonicaAndGetPrihodNaScetWB(String path) throws IOException {
        ExcelConverter.mainShadow(path);
        FileInputStream fileInputStream = new FileInputStream(path + "UnitЭкономика.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);
        double prihodNaSchetWB = sheet.getRow(23).getCell(1).getNumericCellValue();
        BigDecimal bd = new BigDecimal(prihodNaSchetWB);
        bd = bd.setScale(2, RoundingMode.HALF_UP);
        prihodNaSchetWB = bd.doubleValue();
        fileInputStream.close();
        return prihodNaSchetWB;
    }

    public int getPrihodNaSchetWBSumFromCategories(String path) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(path);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);
        double sum = 0;
        int i = 2;
        while(true) {
            XSSFCell cell = sheet.getRow(23).getCell(i);
            if(cell == null)
                break;
            sum += cell.getNumericCellValue();
            i++;
        }
        return  (int)sum;
    }

    public int getPrihodNaSchetWBFromCategoriesExpected(String path) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(path);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);
        return (int)sheet.getRow(23).getCell(1).getNumericCellValue();
    }

    public int getPribilNaMarketingSumFromCategories(String path) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(path);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);
        double sum = 0;
        int i = 2;
        while(true) {
            XSSFCell cell = sheet.getRow(23).getCell(i);
            if(cell == null)
                break;
            sum += cell.getNumericCellValue();
            i++;
        }
        return  (int)sum;
    }

    public int getPribilNaMarketingFromCategoriesExpected(String path) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(path);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);
        return (int)sheet.getRow(23).getCell(1).getNumericCellValue();
    }
}
