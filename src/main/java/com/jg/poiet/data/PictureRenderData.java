package com.jg.poiet.data;

import com.jg.poiet.exception.RenderException;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;

/**
 * 图片渲染数据
 */
public class PictureRenderData implements RenderData {

    public final static int DX1_MAX = 256;
    public final static int DY1_MAX = 512;
    public final static int DX2_MAX = 268;
    public final static int DY2_MAX = 636;

    private BufferedImage bufferedImage;
    /**
     * 距离第一个单元格的x偏移量
     */
    private int dx1 = 0;
    /**
     * 距离第一个单元格的y偏移量
     */
    private int dy1 = 0;
    /**
     * 距离第二个单元格的x偏移量
     */
    private int dx2 = 268;
    /**
     * 距离第二个单元格的y偏移量
     */
    private int dy2 = 636;

    public PictureRenderData(String path) {
        this(new File(path));
    }

    public PictureRenderData(String path, int dx1, int dy1, int dx2, int dy2) {
        this(new File(path), dx1, dy1, dx2, dy2);
    }

    public PictureRenderData(File file) {
        try {
            this.bufferedImage = ImageIO.read(file);
        } catch (IOException e) {
            throw new RenderException(e);
        }
    }

    public PictureRenderData(File file, int dx1, int dy1, int dx2, int dy2) {
        try {
            this.bufferedImage = ImageIO.read(file);
            this.dx1 = dx1;
            this.dy1 = dy1;
            this.dx2 = dx2;
            this.dy2 = dy2;
        } catch (IOException e) {
            throw new RenderException(e);
        }
    }

    public PictureRenderData(BufferedImage bufferedImage) {
        this(bufferedImage, 0, 0, DX2_MAX, DY2_MAX);
    }

    public PictureRenderData(BufferedImage bufferedImage, int dx1, int dy1, int dx2, int dy2) {
        this.bufferedImage = bufferedImage;
        this.dx1 = dx1;
        this.dy1 = dy1;
        this.dx2 = dx2;
        this.dy2 = dy2;
    }

    public BufferedImage getBufferedImage() {
        return bufferedImage;
    }

    public PictureRenderData setBufferedImage(BufferedImage bufferedImage) {
        this.bufferedImage = bufferedImage;
        return this;
    }

    public int getDx1() {
        return dx1;
    }

    public PictureRenderData setDx1(int dx1) {
        this.dx1 = dx1;
        return this;
    }

    public int getDy1() {
        return dy1;
    }

    public PictureRenderData setDy1(int dy1) {
        this.dy1 = dy1;
        return this;
    }

    public int getDx2() {
        return dx2;
    }

    public PictureRenderData setDx2(int dx2) {
        this.dx2 = dx2;
        return this;
    }

    public int getDy2() {
        return dy2;
    }

    public PictureRenderData setDy2(int dy2) {
        this.dy2 = dy2;
        return this;
    }
}
