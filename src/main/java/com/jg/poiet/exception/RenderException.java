package com.jg.poiet.exception;

public class RenderException extends RuntimeException {

    private static final long serialVersionUID = 3720983069910623272L;

    public RenderException() {
    }

    public RenderException(String message) {
        super(message);
    }

    public RenderException(String message, Throwable cause) {
        super(message, cause);
    }

    public RenderException(Throwable cause) {
        super(cause);
    }
}
