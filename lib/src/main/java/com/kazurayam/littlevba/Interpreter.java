package com.kazurayam.littlevba;

import org.antlr.v4.runtime.ANTLRInputStream;
import org.antlr.v4.runtime.BailErrorStrategy;
import org.antlr.v4.runtime.CharStream;
import org.antlr.v4.runtime.CharStreams;
import org.antlr.v4.runtime.CommonTokenStream;
import org.antlr.v4.runtime.InputMismatchException;
import org.antlr.v4.runtime.misc.ParseCancellationException;
import org.antlr.v4.runtime.tree.ParseTree;
import vba.VBALexer;
import vba.VBAParser;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintStream;

/**
 * The entry point of the VBA interpreter
 */
public class Interpreter {

    private InputStream stdin;
    private OutputStream stdout;
    private OutputStream stderr;
    private PrintStream stdoutPrint;
    private PrintStream stderrPrint;
    private Memory memory;

    public Interpreter(InputStream stdin, OutputStream stdout, OutputStream stderr) {
        this.stdin = stdin;
        this.stdout = stdout;
        this.stderr = stderr;
        this.stdoutPrint = new PrintStream(stdout, true);
        this.stderrPrint = new PrintStream(stderr, true);
    }

    public Value run(InputStream progrIn) throws IOException {
        // wrap the input stream
        CharStream input = CharStreams.fromStream(progrIn);
        // initialize the lexer with the input stream
        VBALexer lexer = new VBALexer(input);
        // initialize the token stream with the lexer
        CommonTokenStream tokens = new CommonTokenStream(lexer);
        // initialize the parser with the tokens stream
        VBAParser parser = new VBAParser(tokens);
        // set the error handling strategy
        parser.setErrorHandler(new BailErrorStrategy());
        // remove the default error listeners
        parser.removeErrorListeners();
        // set our custom error listener
        parser.addErrorListener(new ErrorListener(stderrPrint));

        try {
            // parse the input into a parse tree
            ParseTree tree = parser.prog();
            memory = new Memory();
            // init our custom visitor
            LittleVBAVisitor eval = new LittleVBAVisitor(memory, stdin, stdoutPrint, stderrPrint);
            // visit the parse tree and interpret the program
            eval.visit(tree);
        } catch (InterpreterException e) {
            // handle interpreter errors
            stderrPrint.println(e.getMessage());
        } catch (ParseCancellationException e) {
            // handle parser errors
            if (e.getCause() instanceof InputMismatchException) {
                InputMismatchException inputEx = (InputMismatchException)e.getCause();
                String msg = Utils.formatErrorMessage(
                        inputEx.getOffendingToken().getLine(),
                        inputEx.getOffendingToken().getCharPositionInLine(),
                        "Syntax error");
                stderrPrint.println(msg);
            }
        }
        return null;
    }

    public Memory getMemory() {
        return memory;
    }

    public void clear() {
        memory.free();
    }
}
