package com.kazurayam.vba.printing;

import com.kazurayam.subprocessj.Subprocess;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.List;

/**
 * https://gist.github.com/GAM3RG33K/cc59290e8fe68d61c7ab2540f8471fd3
 */
public class PlantUMLRunner extends AbstractCommandRunner{

    private static final Logger logger =
            LoggerFactory.getLogger(PlantUMLRunner.class);

    private final Path workingDirectory;
    private final Path pumlFile;
    private final Path outdir;

    private PlantUMLRunner(Builder builder) {
        this.workingDirectory = builder.workingDirectory;
        this.pumlFile = builder.pumlFile;
        this.outdir = builder.outdir;
    }

    public void run() throws IOException, InterruptedException {
        Subprocess sp = new Subprocess()
                .cwd(workingDirectory.toFile());
        sp.environment().put("PLANTUML_LIMIT_SIZE", "8192");
        assert sp.environment("PLANTUML_LIMIT_SIZE").equals("8192");
        Subprocess.CompletedProcess cp;
        cp = sp.run(makeCommandLine());
        //cp.stdout().forEach(System.out::println);
        //cp.stderr().forEach(System.err::println);
        if (cp.returncode() != 0) {
            throw new IllegalStateException(cp.toString());
        }
    }

    private List<String> makeCommandLine() {
        List<String> commandline;
        if (outdir != null) {
            commandline = Arrays.asList(
                    findCommand("plantuml"),
                    pumlFile.toString(),
                    "-o", outdir.toString(),
                    "-progress", "-tpng", "--verbose");
        } else {
            commandline = Arrays.asList(
                    findCommand("plantuml"),
                    pumlFile.toString(),
                    "-progress", "-tpng", "--verbose");
        }
        return commandline;
    }


    /**
     *
     */
    public static class Builder {
        private Path workingDirectory;
        private Path pumlFile;
        private Path outdir;

        public Builder() {
            workingDirectory = Paths.get(".");
            pumlFile = null;
            outdir = null;
        }

        public Builder workingDirectory(Path workingDirectory) throws IOException {
            if (!Files.exists(workingDirectory)) {
                throw new IOException("Working directory does not exist: " +
                        workingDirectory);
            }
            this.workingDirectory = workingDirectory;
            return this;
        }

        public Builder puml(String pumlString) throws IOException{
            Path p = Paths.get(pumlString);
            return puml(p);
        }

        public Builder puml(Path p) throws IOException{
            if (!Files.exists(p)) {
                throw new IOException(p + " does not exist");
            }
            this.pumlFile = p;
            return this;
        }

        public Builder outdir(String outdirString) throws IOException {
            Path dir = Paths.get(outdirString);
            return this.outdir(dir);
        }

        public Builder outdir(Path outdir) throws IOException {
            this.outdir = outdir;
            Files.createDirectories(outdir);
            return this;
        }

        public PlantUMLRunner build() {
            if (workingDirectory == null) {
                throw new IllegalArgumentException("workingDirectory is required but not given");
            }
            if (pumlFile == null) {
                throw new IllegalArgumentException("puml file is required but not given");
            }
            return new PlantUMLRunner(this);
        }

    }
}
