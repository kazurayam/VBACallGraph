package littlevba;

import com.kazurayam.littlevba.Interpreter;
import com.kazurayam.littlevba.Value;
import com.kazurayam.unittest.TestOutputOrganizer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

public class InterpreterTest {

    private static Logger logger = LoggerFactory.getLogger(InterpreterTest.class);

    private static TestOutputOrganizer too =
            new TestOutputOrganizer.Builder(InterpreterTest.class)
                    .subOutputDirectory(InterpreterTest.class).build();
    private static Path basDir =
            too.getProjectDirectory().resolve("src/test/fixtures/vba6");
    @Test
    public void testWriteFile() throws IOException {
        Path outDir = too.cleanMethodOutputDirectory("testWriteFile");
        Path outFile = outDir.resolve("foo.txt");
        Files.writeString(outFile, "Hello, world!");
        assertThat(outFile.toFile()).exists();
    }

    private Value interpret(Path bas) throws IOException {
        InputStream progrIn = new FileInputStream(bas.toFile());
        Interpreter interpreter = new Interpreter(System.in, System.out, System.err);
        return interpreter.run(progrIn);
    }
    @Test
    public void testVisitingExample01() throws IOException {
        logger.debug("[testVisitingExample01]");
        Path bas = basDir.resolve("example01.bas");
        Value value = interpret(bas);
    }

    @Test
    public void testVisitingExample06dateliteral() throws IOException {
        logger.debug("[testVisitingExample06dateliteral]");
        Path bas = basDir.resolve("example06dateliteral.bas");
        Value value = interpret(bas);
    }

}
