package com.kazurayam.vba.printing;

import com.kazurayam.subprocessj.Subprocess;
import com.kazurayam.subprocessj.Subprocess.CompletedProcess;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;

/**
 * run "mutool poster" command in a commandline subprocess
 */
public class MutoolPosterRunner extends AbstractCommandRunner {

    private static final Logger logger = LoggerFactory.getLogger(MutoolPosterRunner.class);

    private final int xDecimationFactor;
    private final int yDecimationFactor;
    private final Path original;   // PDF file
    private final Path poster;  // PDF with multiple pages with image pieces

    private MutoolPosterRunner(Builder builder) {
        this.xDecimationFactor = builder.xDecimationFactor;
        this.yDecimationFactor = builder.yDecimationFactor;
        this.original = builder.original;
        this.poster = builder.poster;
    }

    public int getX() {
        return xDecimationFactor;
    }

    public int getY() {
        return yDecimationFactor;
    }

    public void run() throws IOException, InterruptedException {
        CompletedProcess cp;
        List<String> commandline = Arrays.asList(
                findCommand("mutool"),
                "poster",
                "-x", String.valueOf(this.xDecimationFactor),
                "-y", String.valueOf(this.yDecimationFactor),
                original.toString(),
                poster.toString()
        );
        logger.info(String.join(" ", commandline));
        cp = new Subprocess().cwd(new File("."))
                .run(commandline);

        if (!cp.stderr().isEmpty()) {
            cp.stderr().forEach(System.err::println);
        }
        if (!cp.stdout().isEmpty()) {
            cp.stdout().forEach(System.out::println);
        }
        if (cp.returncode() != 0) {
            throw new IOException("mutool returned: " + cp.returncode());
        }
    }

    public Path getPoster() {
        return this.poster;
    }

    /**
     *
     */
    public static class Builder {

        private int xDecimationFactor;
        private int yDecimationFactor;
        private Path original;
        private Path poster;
        private PieceSize pieceSize;

        public Builder() {
            xDecimationFactor = 1;
            yDecimationFactor = 1;
            original = null;
            poster = null;
            pieceSize = null;
        }

        public Builder x(int xDecimationFactor) {
            if (xDecimationFactor < 1 || xDecimationFactor > 10) {
                throw new IllegalArgumentException("x must be between 1 and 10");
            }
            this.xDecimationFactor = xDecimationFactor;
            return this;
        }

        public Builder y(int yDecimationFactor) {
            if (yDecimationFactor < 1 || yDecimationFactor > 10) {
                throw new IllegalArgumentException("y must be between 1 and 10");
            }
            this.yDecimationFactor = yDecimationFactor;
            return this;
        }

        public Builder pieceSize(String pieceSize) {
            this.pieceSize = PieceSize.findByName(pieceSize);
            if (this.pieceSize == null) {
                logger.warn(String.format("pieceSize(\"%s\") is not acceptable. ", pieceSize));
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < PieceSize.values().length; i++) {
                    if (i > 0) {
                        sb.append(", ");
                    }
                    sb.append(PieceSize.values()[i].name());
                }
                logger.info(String.format("acceptable pieceSize values include %s", sb.toString()));
            }
            return this;
        }

        public Builder original(Path original) throws IOException {
            if (!Files.exists(original)) {
                throw new IOException(original + " is not present");
            }
            this.original = original;
            return this;
        }

        public Builder poster(Path poster) throws IOException {
            Path parent = poster.getParent();
            if (!Files.exists(parent)) {
                Files.createDirectories(parent);
            }
            this.poster = poster;
            return this;
        }

        public MutoolPosterRunner build() throws IOException {
            if (poster == null) {
                // If the poster is not specified, the file name will be
                // the original file name without ".pdf" appended with "-poster.pdf"
                // The poster will be saved in the parent directory of the original.
                Path originalParent = original.getParent();
                String originalName = original.getFileName().toString();
                String originalNameBody = originalName.substring(0, originalName.lastIndexOf('.'));
                String posterName = originalNameBody + "-poster.pdf";
                this.poster = originalParent.resolve(posterName);
            }
            if (pieceSize != null) {
                if (xDecimationFactor > 1 || yDecimationFactor > 1) {
                    throw new IllegalArgumentException("if you want to specify pieceSize, " +
                            "then you can not explicitly specify x and y with value > 1 together.");
                }
                PDFWrapper wrapped = new PDFWrapper(original);
                PDRectangle originalRectangle = wrapped.getRectangle(0);
                logger.info(String.format("original PDF has width=%.2f, height=%.2f in millimeter",
                        originalRectangle.getWidth(),
                        originalRectangle.getHeight() ));
                xDecimationFactor =
                        deriveDecimationFactor(originalRectangle.getWidth(), pieceSize.getWidthMM());
                yDecimationFactor =
                        deriveDecimationFactor(originalRectangle.getHeight(), pieceSize.getHeightMM());
                logger.info(String.format("pieceSize=%s was specified", pieceSize.name()));
                logger.info(String.format("%s is defined as width=%.2f, height=%.2f in millimeter",
                        pieceSize.name(), pieceSize.getWidthMM(), pieceSize.getHeightMM()));
                logger.info(String.format("derived decimation factors: -x %d -y %d",
                        xDecimationFactor, yDecimationFactor));
            }
            return new MutoolPosterRunner(this);
        }

        int deriveDecimationFactor(float source, float unit) {
            return (int)Math.ceil(source / unit);
        }
    }

    /**
     *
     */
    public static enum PieceSize {
        A2(420, 594),
        A3(297, 420),
        A4(210, 297),
        A5(148, 210),
        A6(104, 148),
        LEGAL(216, 355.6f),
        LETTER(216, 279);
        private final float widthMM;
        private final float heightMM;

        PieceSize(float widthMM, float heightMM) {
            this.widthMM = widthMM;
            this.heightMM = heightMM;
        }

        float getWidthMM() {
            return widthMM;
        }

        float getHeightMM() {
            return heightMM;
        }

        public static PieceSize findByName(String name) {
            PieceSize instance = null;
            for (PieceSize pieceSize : PieceSize.values()) {
                if (pieceSize.name().equalsIgnoreCase(name)) {
                    instance = pieceSize;
                }
            }
            return instance;
        }
    }

    /**
     *
     */
    public static class PDFWrapper {
        private final PDDocument document;
        PDFWrapper(Path pdf) throws IOException {
            document = Loader.loadPDF(pdf.toFile());
            if (document == null) {
                throw new IllegalArgumentException("unable to create PDFWrapper object with " +
                        pdf.toString());
            }
        }

        public PDPage getPage(int pageIndex) {
            return document.getPage(pageIndex);
        }

        public PDRectangle getRectangle(int pageIndex) {
            return getRectangle(getPage(pageIndex));
        }

        public PDRectangle getRectangle(PDPage page) {
            float widthMM = point2mm(page.getMediaBox().getWidth());
            float heightMM = point2mm(page.getMediaBox().getHeight());
            return new PDRectangle(widthMM, heightMM);
        }

        /*
           convert point unit to milli-miters
         */
        public static float point2mm(float points) {
            return points * 25.4f / 72;
        }
    }


}