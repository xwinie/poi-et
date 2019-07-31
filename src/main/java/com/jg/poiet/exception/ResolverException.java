package com.jg.poiet.exception;

public class ResolverException extends RuntimeException {

    private static final long serialVersionUID = -3575288404315397969L;

    public ResolverException() {
    }

    public ResolverException(String message) {
        super(message);
    }

    public ResolverException(String message, Throwable cause) {
        super(message, cause);
    }

    public ResolverException(Throwable cause) {
        super(cause);
    }

}
