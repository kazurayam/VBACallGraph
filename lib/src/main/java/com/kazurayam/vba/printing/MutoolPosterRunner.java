package com.kazurayam.vba.printing;

import com.kazurayam.subprocessj.Subprocess;
import com.kazurayam.subprocessj.Subprocess.CompletedProcess;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;

public class MutoolPosterRunner extends AbstractCommandRunner{

    private int x;
    private int y;
    private final Path original;   // PDF file
    private final Path poster;  // PDF with multiple pages with image pieces

    private MutoolPosterRunner(Builder builder) {
        this.x = builder.x;
        this.y = builder.y;
        this.original = builder.original;
        this.poster = builder.poster;
    }

    public void run() throws IOException, InterruptedException {
        CompletedProcess cp;
        cp = new Subprocess().cwd(new File("."))
                .run(Arrays.asList(
                        findCommand("mutool"),
                        "poster",
                        "-x", "" + this.x,
                        "-y", "" + this.y,
                        original.toString(),
                        poster.toString()
                ));
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
        private int x;
        private int y;
        private Path original;
        private Path poster;
        public Builder() {
            x = 1;
            y = 1;
            original = null;
            poster = null;
        }
        public Builder x(int x) {
            if (x < 1 || x > 4) {
                throw new IllegalArgumentException("x must be between 1 and 4");
            }
            this.x = x;
            return this;
        }
        public Builder y(int y) {
            if (y < 1 || y > 4) {
                throw new IllegalArgumentException("y must be between 1 and 4");
            }
            this.y = y;
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
        public MutoolPosterRunner build() {
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
            return new MutoolPosterRunner(this);
        }
    }
}
