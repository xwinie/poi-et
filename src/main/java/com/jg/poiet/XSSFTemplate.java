package com.jg.poiet;

import com.jg.poiet.config.Configure;
import com.jg.poiet.exception.ResolverException;
import com.jg.poiet.render.RenderFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;

/**
 * 模板
 */
public class XSSFTemplate {

    private static Logger logger = LoggerFactory.getLogger(XSSFTemplate.class);

    private NiceXSSFWorkbook workbook;
    private Configure config;

    private XSSFTemplate() {}

    public static XSSFTemplate compile(String filePath) {
        return compile(new File(filePath));
    }

    public static XSSFTemplate compile(File file) {
        return compile(file, Configure.createDefault());
    }

    public static XSSFTemplate compile(InputStream inputStream) {
        return compile(inputStream, Configure.createDefault());
    }

    public static XSSFTemplate compile(String filePath, Configure config) {
        return compile(new File(filePath), config);
    }

    public static XSSFTemplate compile(File file, Configure config) {
        try {
            return compile(new FileInputStream(file), config);
        } catch (FileNotFoundException e) {
            logger.error("Cannot find the file", e);
            throw new ResolverException("Cannot find the file [" + file.getPath() + "]");
        }
    }

    public static XSSFTemplate compile(InputStream inputStream, Configure config) {
        try {
            XSSFTemplate instance = new XSSFTemplate();
            instance.config = config;
            instance.workbook = new NiceXSSFWorkbook(inputStream);
            return instance;
        } catch (IOException e) {
            logger.error("Compile template failed", e);
            throw new ResolverException("Compile template failed");
        }
    }

    public XSSFTemplate render(Object model) {
        RenderFactory.getRender(model, config.getElMode()).render(this);
        return this;
    }

    public void write(OutputStream out) throws IOException {
        this.workbook.write(out);
    }

    public void writeToFile(String path) throws IOException {
        FileOutputStream out = new FileOutputStream(path);
        this.write(out);
        this.close();
        out.flush();
        out.close();
    }

    public void close() throws IOException {
        this.workbook.close();
    }

    public NiceXSSFWorkbook getXSSFWorkbook() {
        return this.workbook;
    }

    public Configure getConfig() {
        return config;
    }

}
