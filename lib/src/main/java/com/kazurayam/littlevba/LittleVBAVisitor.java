package com.kazurayam.littlevba;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import vba.VBABaseVisitor;
import vba.VBAParser;

import java.io.BufferedReader;
import java.io.InputStream;
import java.io.PrintStream;

public class LittleVBAVisitor extends VBABaseVisitor<Value> {

    private static Logger logger = LoggerFactory.getLogger(LittleVBAVisitor.class);

    private InputStream stdin;
    private PrintStream stdout;
    private PrintStream stderr;
    private Memory memory;

    private PrintStream printStream;
    private BufferedReader inputStream;

    public LittleVBAVisitor(Memory memory, InputStream stdin, PrintStream stdout, PrintStream stderr) {
        this.stdin = stdin;
        this.stdout = stdout;
        this.stderr = stderr;
        this.memory = memory;
    }

    @Override
    public Value visitFunctionStmt(VBAParser.FunctionStmtContext ctx) {
        logger.info("Function " + ctx.ambiguousIdentifier().getText());   // "Module" from "Public Sub Module()"
        return visitChildren(ctx);
    }


    @Override
    public Value visitSubStmt(VBAParser.SubStmtContext ctx) {
        logger.info("Sub " + ctx.ambiguousIdentifier().getText());   // "Module" from "Public Sub Module()"
        return visitChildren(ctx);
    }

}
