package com.kazurayam.littlevba;

import com.kazurayam.unittest.TestOutputOrganizer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;

public class InterpreterTest {

    private static Logger logger = LoggerFactory.getLogger(InterpreterTest.class);

    private static TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(InterpreterTest.class)
                    .subOutputDirectory(InterpreterTest.class).build();
    @BeforeClass
    public void setupClass() {
//        System.setProperty("org.slf4j.simpleLogger.defaultLogLevel", "debug");
//        System.setProperty("org.slf4j.simpleLogger.showDateTime", "true");
//        System.setProperty("org.slf4j.simpleLogger.dateTimeFormat", "yyyy-MM-dd HH:mm:ss:SSS Z");
    }

    @Test
    public void testWriteFile() throws IOException {
        Path outDir = too.cleanMethodOutputDirectory("testWriteFile");
        Path outFile = outDir.resolve("foo.txt");
        Files.writeString(outFile, "Hello, world!");
    }

    @Test
    public void testVisitingExample01bas() throws IOException {
        logger.debug("[testVisitingExample01bas]");
        System.out.println("Hello, how are you?");
        Path basDir = too.getProjectDirectory().resolve("src/test/fixtures/vba6");
        Path bas = basDir.resolve("example01.bas");
        InputStream progrIn = new FileInputStream(bas.toFile());
        Interpreter interpreter = new Interpreter(System.in, System.out, System.err);
        Value value = interpreter.run(progrIn);
    }

    @Test
    public void testVisitingModule1() throws IOException {
        System.out.println("The logger is an instance of " + logger.getClass().getName());
        logger.debug("[testVisitingModule1]");
        Path basDir = too.getProjectDirectory().resolve("src/test/fixtures/Book1");
        Path bas = basDir.resolve("Module1.bas");
        InputStream progrIn = new FileInputStream(bas.toFile());
        Interpreter interpreter = new Interpreter(System.in, System.out, System.err);
        Value value = interpreter.run(progrIn);
    }
}
