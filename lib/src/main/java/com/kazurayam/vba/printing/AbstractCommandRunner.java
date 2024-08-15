package com.kazurayam.vba.printing;

import com.kazurayam.subprocessj.CommandLocator;

import java.io.File;

public abstract class AbstractCommandRunner {

    protected String findCommand(String command) {
        CommandLocator.CommandLocatingResult clr = CommandLocator.find(command);
        if (clr.returncode() == 0) {
            return clr.command();
        } else {
            throw new IllegalStateException("Could not find the full path of mutool command. " +
                    "Possibly you have not installed the mutool. " +
                    "Or the current environment variable PATH does not contain the mutool. " +
                    "PATH=" + getPATH());
        }
    }

    protected String getPATH() {
        String envPATH = System.getenv("PATH");
        String[] envPATHComponents = envPATH.split(File.pathSeparator);
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < envPATHComponents.length; i++) {
            sb.append(envPATHComponents[i]);
            sb.append(File.pathSeparator);
            sb.append("\n");
        }
        return sb.toString();
    }

}
