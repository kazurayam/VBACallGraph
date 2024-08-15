package com.kazurayam.vba.printing;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;

/**
 * https://www.baeldung.com/java-pdf-creation
 */
public class PDFFromImageGenerator {

    private PDFFromImageGenerator() {}

    public static void generate(Path image, Path pdf) throws IOException {
        generate(image.toFile(), pdf.toFile());
    }

    public static void generate(File image, File pdf) throws IOException {
        PDDocument document = new PDDocument();
        //画像をロードしてラッパーオブジェクトで包む
        PDImageXObject pdImage = PDImageXObject.createFromFileByContent(image, document);
        //画像と同じ大きさのPDRectangleオブジェクトをひとつ持ったPDPageオブジェクトを作り
        float width = pdImage.getWidth();
        float height = pdImage.getHeight();
        PDRectangle rect = new PDRectangle(width, height);
        PDPage page = new PDPage(rect);
        //そのPDPageオブジェクトをdocumentに追加する
        document.addPage(page);
        //documentの中に画像データを流し込む
        PDPageContentStream contentStream = new PDPageContentStream(document, page);
        contentStream.drawImage(pdImage, 0.0f, 0.0f);
        contentStream.close();
        //
        document.save(pdf);
        document.close();
    }

    public static Path resolvePDFFileNameFromImage(Path image) {
        String imageFileName = image.getFileName().toString();
        if (imageFileName.contains(".")) {
            String nameWithoutExt =
                    imageFileName.substring(0, imageFileName.lastIndexOf("."));
            return Paths.get(nameWithoutExt + ".pdf");
        } else {
            return Paths.get(imageFileName + ".pdf");
        }
    }
}
