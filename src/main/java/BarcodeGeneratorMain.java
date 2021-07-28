import com.google.zxing.BarcodeFormat;
import com.google.zxing.WriterException;
import com.google.zxing.client.j2se.MatrixToImageWriter;
import com.google.zxing.common.BitMatrix;
import com.google.zxing.qrcode.QRCodeWriter;
import com.itextpdf.text.*;
import com.itextpdf.text.Font;
import com.itextpdf.text.Image;
import com.itextpdf.text.pdf.*;
import org.apache.poi.ss.usermodel.*;

import java.awt.*;
import java.awt.Color;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

public class BarcodeGeneratorMain {

    private List<String> codes = new ArrayList<String>();
    private static int[] columns = new int[] { 62, 359 };
    private static int[] lines = new int[] { 461, 40 };

    public BarcodeGeneratorMain(List<String> codes) {
        this.codes = codes;
    }

    public void execute() throws Exception {

        Document document = new Document(PageSize.A4);
        System.out.println(document.getPageSize().getHeight());
        System.out.println(document.getPageSize().getWidth());

        PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream("teste.pdf"));

        document.open();

        int line = 0;
        int column = 0;

        Font f = new Font(Font.FontFamily.COURIER, 8.0f, Font.BOLD, BaseColor.BLACK);

        for(String code : codes) {

            QRCodeWriter barcodeWriter = new QRCodeWriter();
            BitMatrix bitMatrix =
                    barcodeWriter.encode(code, BarcodeFormat.QR_CODE, 170, 170);

            BufferedImage bufimage = MatrixToImageWriter.toBufferedImage(bitMatrix);

            int x = columns[column];
            int y = lines[line];

            Image image = Image.getInstance(bufimage, Color.white);
            image.setAbsolutePosition(x, y);
            document.add(image);

            // imprime o codigo

            Chunk chunk = new Chunk(code, f);
            Phrase phrase = new Phrase(chunk);

            PdfContentByte content = writer.getDirectContent();
            ColumnText.showTextAligned(content, Element.ALIGN_LEFT, phrase, x+55, y+10, 0);

            column++;

            if (column == 2) {
                column = 0;
                line++;
            }
            if (line == 2) {
                document.newPage();
                line = 0;
            }

        }

        document.close();

    }

    public static void main(String[] args) throws Exception {

        Workbook wb = WorkbookFactory.create(new File("/Users/jean/Downloads/codigos.xlsx"));
        Sheet sheet = wb.getSheetAt(0);
        ArrayList<String> codes = new ArrayList<String>();

        for(int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            Cell cell = row.getCell(0);

            if (cell == null) break;

            BigDecimal number = BigDecimal.valueOf(cell.getNumericCellValue());
            codes.add(number.toPlainString());

            cell = row.getCell(2);
            number = BigDecimal.valueOf(cell.getNumericCellValue());
            codes.add(number.toPlainString());

            cell = row.getCell(4);
            number = BigDecimal.valueOf(cell.getNumericCellValue());
            codes.add(number.toPlainString());

            cell = row.getCell(6);
            number = BigDecimal.valueOf(cell.getNumericCellValue());
            codes.add(number.toPlainString());
        }

        new BarcodeGeneratorMain(codes).execute();





//        new BarcodeGeneratorMain(codes).execute();




        /*
        QRCodeWriter barcodeWriter = new QRCodeWriter();
        BitMatrix bitMatrix =
                barcodeWriter.encode(barcodeText, BarcodeFormat.QR_CODE, 200, 200);

        return MatrixToImageWriter.toBufferedImage(bitMatrix);*/
    }

}
