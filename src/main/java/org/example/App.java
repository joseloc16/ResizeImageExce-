package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

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
        File imageFile = new File("C:\\data\\descarga.jpg");
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Resize Image");
        int cellWidthInPixels = 200;
        int cellHeightInPixels = 150;
        BufferedImage originalImage = ImageIO.read(imageFile);

        // Calcula los factores de escala para el ancho y la altura.
        double widthScaleFactor = (double) cellWidthInPixels / originalImage.getWidth();
        double heightScaleFactor = (double) cellHeightInPixels / originalImage.getHeight();

        // Usa el menor factor de escala para mantener las proporciones de la imagen.
        double scaleFactor = Math.min(widthScaleFactor, heightScaleFactor);

        // Calcula las nuevas dimensiones manteniendo la proporción.
        int scaledWidth = (int) (originalImage.getWidth() * scaleFactor);
        int scaledHeight = (int) (originalImage.getHeight() * scaleFactor);

        // Crea una imagen nueva con las dimensiones escaladas.
        Image scaledImage = originalImage.getScaledInstance(scaledWidth, scaledHeight, Image.SCALE_SMOOTH);
        BufferedImage bufferedScaledImage = new BufferedImage(scaledWidth, scaledHeight, BufferedImage.TYPE_INT_RGB);

        // Dibuja la imagen escalada en la nueva imagen.
        bufferedScaledImage.getGraphics().drawImage(scaledImage, 0, 0, null);

        // Escribe la imagen escalada a un ByteArrayOutputStream.
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(bufferedScaledImage, "jpg", baos);

        // Agrega la imagen a la hoja de cálculo.
        int pictureIdx = workbook.addPicture(baos.toByteArray(), Workbook.PICTURE_TYPE_JPEG);
        CreationHelper helper = workbook.getCreationHelper();
        Drawing drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(0);
        anchor.setRow1(0);
        anchor.setDx1(0);
        anchor.setDy1(0);
        anchor.setDx2(Units.toEMU(scaledWidth));  // Ajusta el ancho y la altura apropiadamente
        anchor.setDy2(Units.toEMU(scaledHeight));
        Picture pict = drawing.createPicture(anchor, pictureIdx);
        pict.resize();

        // Escribe la hoja de cálculo a un archivo.
        File file = new File("C:\\data\\imageResize.xlsx");
        FileOutputStream fos = new FileOutputStream(file);
        workbook.write(fos);
        fos.close();
    }

}
