package com.jg.poiet.exception;

public class ExpressionEvalException extends RuntimeException {

    private static final long serialVersionUID = 957226688249738010L;

    public ExpressionEvalException() {
    }

    public ExpressionEvalException(String message) {
        super(message);
    }

    public ExpressionEvalException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExpressionEvalException(Throwable cause) {
        super(cause);
    }

}
