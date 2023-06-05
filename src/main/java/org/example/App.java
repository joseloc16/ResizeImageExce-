package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;

/**
 * Hello world!
 *
 */
public class App {
    public static void main(String[] args) throws Exception {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Avengers");
        //setWidth(sheet, 107);
        Row row1 = sheet.createRow(0);
        row1.createCell(0).setCellValue("IRON-MAN");

        InputStream inputStream1 = TestClass.class.getClassLoader()
                .getResourceAsStream("ironman.png");
        InputStream inputStream2 = TestClass.class.getClassLoader()
                .getResourceAsStream("spiderman.png");

        byte[] inputImageBytes1 = IOUtils.toByteArray(inputStream1);
        int inputImagePictureID1 = workbook.addPicture(inputImageBytes1, Workbook.PICTURE_TYPE_PNG);

        XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();

        XSSFClientAnchor ironManAnchor = new XSSFClientAnchor();

        // Combinar celdas desde B1 hasta F6
        sheet.addMergedRegion(new CellRangeAddress(0, 5, 1, 5));

        ironManAnchor.setCol1(1); // Sets the column (0 based) of the first cell.
        ironManAnchor.setCol2(5); // Sets the column (0 based) of the Second cell.
        ironManAnchor.setRow1(0); // Sets the row (0 based) of the first cell.
        ironManAnchor.setRow2(5); // Sets the row (0 based) of the Second cell.

        drawing.createPicture(ironManAnchor, inputImagePictureID1);

        for (int i = 0; i < 3; i++) {
            sheet.autoSizeColumn(i);
        }

        FileOutputStream saveExcel = null;
        try {
            saveExcel = new FileOutputStream("C:\\data\\imageResize.xlsx");
            workbook.write(saveExcel);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (saveExcel != null) {
                try {
                    saveExcel.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    private static void setWidth(Sheet sheet, int numberColumn) {
        int width = (int) (256 * (1.7));
        for (int i = 0; i < numberColumn; i++) {
            sheet.setColumnWidth(i + 1, width);
        }
    }

}
