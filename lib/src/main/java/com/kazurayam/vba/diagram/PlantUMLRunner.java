package com.kazurayam.vba.diagram;

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
public class PlantUMLRunner {

    private static final Logger logger =
            LoggerFactory.getLogger(PlantUMLRunner.class);

    private Path workingDirectory = null;
    private Path diagram = null;
    private Path outdir = null;

    public PlantUMLRunner() {}

    public void workingDirectory(Path workingDirectory) {
        if (!Files.exists(workingDirectory)) {
            throw new IllegalArgumentException("Working directory does not exist: " + workingDirectory);
        }
        this.workingDirectory = workingDirectory;
    }

    public void setDiagram(String pu) {
        this.setDiagram(Paths.get(pu));
    }

    public void setDiagram(Path diagram) {
        this.diagram = diagram;
        if (!Files.exists(diagram)) {
            throw new IllegalArgumentException(diagram +
                    " does not exist");
        }
    }

    public void setOutdir(String outdir) throws IOException {
        this.setOutdir(Paths.get(outdir));
    }

    public void setOutdir(Path outdir) throws IOException {
        this.outdir = outdir;
        Files.createDirectories(outdir);
    }

    public void run() throws IOException, InterruptedException {
        validateParams();
        if (workingDirectory == null) {
            workingDirectory = Paths.get(".");
        }
        Subprocess.CompletedProcess cp;
        List<String> commandline = getCommandline();
        cp = new Subprocess().cwd(workingDirectory.toFile())
                .run(commandline);
        logger.info("commandline: " + cp.commandline());
        cp.stdout().forEach(System.out::println);
        cp.stderr().forEach(System.err::println);
        if (cp.returncode() != 0) {
            throw new IllegalStateException("Subprocess returns " + cp.returncode());
        }
    }

    private List<String> getCommandline() {
        List<String> commandline;
        if (outdir != null) {
            commandline = Arrays.asList(
                    "docker", "run", "ghcr.io/plantuml/plantuml",
                    diagram.toString(),
                    "-o", outdir.toString(),
                    "-progress", "-tpdf", "--verbose");
        } else {
            commandline = Arrays.asList(
                    "docker", "run", "ghcr.io/plantuml/plantuml",
                    diagram.toString(),
                    "-progress", "-tpdf", "--verbose");
        }
        return commandline;
    }

    private void validateParams() {
        if (diagram == null) {
            throw new IllegalArgumentException("pathPuFile is required but not given");
        }
    }
}
