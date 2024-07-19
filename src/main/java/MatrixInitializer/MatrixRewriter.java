package MatrixInitializer;

import org.apache.poi.xssf.usermodel.XSSFSheet;

public class MatrixRewriter {
    public static void rewriteMatrix(double[][] matrix, XSSFSheet sheetReportConverted) {
        for (int j = 2; j < matrix.length; j++) {
            for (int i = 1; i < matrix[0].length; i++) {
                sheetReportConverted.getRow(j).createCell(i).setCellValue(matrix[j][i]);
            }
        }
    }
}
