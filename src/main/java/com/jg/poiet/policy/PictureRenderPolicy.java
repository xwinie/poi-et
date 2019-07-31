package com.jg.poiet.policy;

import com.jg.poiet.XSSFTemplate;
import com.jg.poiet.data.PictureRenderData;
import com.jg.poiet.exception.RenderException;
import com.jg.poiet.template.cell.CellTemplate;
import com.sun.org.apache.bcel.internal.generic.DUP_X1;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;

import javax.imageio.ImageIO;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

import static org.apache.poi.ss.usermodel.ClientAnchor.AnchorType.MOVE_AND_RESIZE;

/**
 * 图片处理
 */
public class PictureRenderPolicy extends AbstractRenderPolicy<PictureRenderData> {

    @Override
    protected boolean validate(PictureRenderData data) {
        return (null != data.getBufferedImage());
    }

    @Override
    public void doRender(CellTemplate cellTemplate, PictureRenderData picture, XSSFTemplate template) throws Exception {
        XSSFCell cell = cellTemplate.getCell();
        Helper.renderPicture(cell, picture, template);
    }

    public static class Helper {

        public static void renderPicture(XSSFCell cell, PictureRenderData picture, XSSFTemplate template) throws IOException {
            checkPictureData(picture);
            cell.setCellValue("");
            ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
            ImageIO.write(picture.getBufferedImage(), "jpg", byteArrayOut);
            //画图的顶级管理器，一个sheet只能获取一个（一定要注意这点）
            XSSFDrawing patriarch = cell.getSheet().createDrawingPatriarch();
            //anchor主要用于设置图片的属性
            int firstColumn = cell.getColumnIndex();
            int firstRow = cell.getRowIndex();
            int secondColumn = cell.getColumnIndex();
            int secondRow = cell.getRowIndex();

            int dx1 = cell.getSheet().getColumnWidth(cell.getColumnIndex()) * picture.getDx1();
            int dy1 = cell.getRow().getHeight() * picture.getDy1();
            int dx2 = cell.getSheet().getColumnWidth(cell.getColumnIndex()) * picture.getDx2();
            int dy2 = cell.getRow().getHeight() * picture.getDy2();
            CellRangeAddress cellAddresses = template.getXSSFWorkbook().getCellRangeAddress(cell);
            if (null != cellAddresses) {
                secondColumn = cellAddresses.getLastColumn();
                secondRow = cellAddresses.getLastRow();
            } else {
                if (dx1 > dx2 || dy1 > dy2) {
                    throw new RenderException("dx1 must less than dx2 and dy1 must less than dy2 when cell is not a merged cells");
                }
            }
            XSSFClientAnchor anchor = new XSSFClientAnchor(dx1, dy1, dx2, dy2, firstColumn, firstRow, secondColumn, secondRow);
            anchor.setAnchorType(MOVE_AND_RESIZE);
            //插入图片
            patriarch.createPicture(anchor, template.getXSSFWorkbook().addPicture(byteArrayOut.toByteArray(), HSSFWorkbook.PICTURE_TYPE_JPEG));
        }

        private static void checkPictureData(PictureRenderData picture) {
            if (picture.getDx1() < 0) {
                picture.setDx1(0);
            }
            if (picture.getDx1() > PictureRenderData.DX1_MAX) {
                picture.setDx1(PictureRenderData.DX1_MAX);
            }
            if (picture.getDy1() < 0) {
                picture.setDy1(0);
            }
            if (picture.getDy1() > PictureRenderData.DY1_MAX) {
                picture.setDy1(PictureRenderData.DY1_MAX);
            }
            if (picture.getDx2() < 0) {
                picture.setDx2(0);
            }
            if (picture.getDx2() > PictureRenderData.DX2_MAX) {
                picture.setDx2(PictureRenderData.DX2_MAX);
            }
            if (picture.getDy2() < 0) {
                picture.setDy2(0);
            }
            if (picture.getDy2() > PictureRenderData.DY2_MAX) {
                picture.setDy2(PictureRenderData.DY2_MAX);
            }
        }

    }

}
