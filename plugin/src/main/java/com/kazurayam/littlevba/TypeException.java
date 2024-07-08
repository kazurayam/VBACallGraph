package com.kazurayam.littlevba;

/**
 * Type mismatch error.
 */
public class TypeException extends InterpreterException {

    public TypeException(String message) {
        super(message);
    }

}