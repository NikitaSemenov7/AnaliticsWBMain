package MatrixInitializer;

import org.apache.poi.xssf.usermodel.XSSFSheet;

public class MatrixCreator {
    public static double[][] createMatrix(XSSFSheet sheetReportConverted) {
        double[][] matrix = new double[sheetReportConverted.getLastRowNum() + 1][sheetReportConverted.getRow(1).getLastCellNum()];
        return matrix;
    }
}
