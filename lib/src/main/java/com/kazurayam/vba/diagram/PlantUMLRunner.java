package com.kazurayam.vba.diagram;

import com.kazurayam.subprocessj.Subprocess;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;

/**
 * https://gist.github.com/GAM3RG33K/cc59290e8fe68d61c7ab2540f8471fd3
 */
public class PlantUMLRunner {

    public static String EXE_PATH_MAC = "/usr/local/bin/plantuml";

    private Path executable = Paths.get(EXE_PATH_MAC);
    private Path workingDirectory = null;
    private Path pu = null;
    private Path output = null;

    public PlantUMLRunner() {}

    public void setExecutable(String pathExecutable) {
        this.setExecutable(Paths.get(pathExecutable));
    }

    public void setExecutable(Path pathExecutable) {
        this.executable = pathExecutable;
    }

    public void workingDirectory(Path workingDirectory) {
        if (!Files.exists(workingDirectory)) {
            throw new IllegalArgumentException("Working directory does not exist: " + workingDirectory);
        }
        this.workingDirectory = workingDirectory;
    }

    public void setPu(String pu) {
        this.setPu(Paths.get(pu));
    }

    public void setPu(Path pu) {
        this.pu = pu;
        if (!Files.exists(pu)) {
            throw new IllegalArgumentException(pu +
                    " does not exist");
        }
    }

    public void setOutput(String output) throws IOException {
        this.setOutput(Paths.get(output));
    }

    public void setOutput(Path output) throws IOException {
        this.output = output;
        Files.createDirectories(output.getParent());
    }

    public void run() throws IOException, InterruptedException {
        validateParams();
        if (workingDirectory == null) {
            workingDirectory = Paths.get(".");
        }
        Subprocess.CompletedProcess cp;
        cp = new Subprocess().cwd(workingDirectory.toFile())
                .run(Arrays.asList(executable.toString(),
                        pu.toString(),
                        "-o", output.toString(),
                        "-progress",
                        "-tpdf"));
        cp.stdout().forEach(System.out::println);
        cp.stderr().forEach(System.err::println);
        if (cp.returncode() != 0) {
            throw new IllegalStateException("Subprocess returns " + cp.returncode());
        }
    }

    private void validateParams() {
        if (Files.notExists(executable)) {
            throw new IllegalArgumentException(executable + " does not exist");
        }
        if (pu == null) {
            throw new IllegalArgumentException("pathPuFile is required but not given");
        }
        if (output == null) {
            throw new IllegalArgumentException("pathOutput is required but not given");
        }
    }
}
