package com.kazurayam.vba;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;

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
        PDPage page = new PDPage();
        document.addPage(page);
        //
        PDPageContentStream contentStream = new PDPageContentStream(document, page);
        PDImageXObject pdImage = PDImageXObject.createFromFileByContent(image, document);
        contentStream.drawImage(pdImage, 0, 0);
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
