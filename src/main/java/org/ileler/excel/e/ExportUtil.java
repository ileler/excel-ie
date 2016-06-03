/**
 * @(#)ExportUtil.java 1.0 2015年10月29日
 * @Copyright:  Copyright 2007 - 2015 MPR Tech. Co. Ltd. All Rights Reserved.
 * @Description: 
 * 
 * Modification History:
 * Date:        2015年10月29日
 * Author:      zhangle
 * Version:     1.0.0.0
 * Description: (Initialize)
 * Reviewer:    
 * Review Date: 
 */
package org.ileler.excel.e;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.ileler.excel.util.LangCode;
import org.ileler.excel.wrapper.Wrapper;

import java.awt.Color;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * Created by ileler@qq.com on 2016/5/12.
 */
public final class ExportUtil {

    private static Map<String, Short> colorShort;
    private static Map<String, Short> borderStyle;

    static {
        colorShort = new HashMap<>(0);
        colorShort.put("BLACK", HSSFColor.BLACK.index);
        colorShort.put("BROWN", HSSFColor.BROWN.index);
        colorShort.put("OLIVE_GREEN", HSSFColor.OLIVE_GREEN.index);
        colorShort.put("DARK_GREEN", HSSFColor.DARK_GREEN.index);
        colorShort.put("DARK_TEAL", HSSFColor.DARK_TEAL.index);
        colorShort.put("DARK_BLUE", HSSFColor.DARK_BLUE.index);
        colorShort.put("INDIGO", HSSFColor.INDIGO.index);
        colorShort.put("GREY_80_PERCENT", HSSFColor.GREY_80_PERCENT.index);
        colorShort.put("ORANGE", HSSFColor.ORANGE.index);
        colorShort.put("DARK_YELLOW", HSSFColor.DARK_YELLOW.index);
        colorShort.put("GREEN", HSSFColor.GREEN.index);
        colorShort.put("TEAL", HSSFColor.TEAL.index);
        colorShort.put("BLUE", HSSFColor.BLUE.index);
        colorShort.put("BLUE_GREY", HSSFColor.BLUE_GREY.index);
        colorShort.put("GREY_50_PERCENT", HSSFColor.GREY_50_PERCENT.index);
        colorShort.put("RED", HSSFColor.RED.index);
        colorShort.put("LIGHT_ORANGE", HSSFColor.LIGHT_ORANGE.index);
        colorShort.put("LIME", HSSFColor.LIME.index);
        colorShort.put("SEA_GREEN", HSSFColor.SEA_GREEN.index);
        colorShort.put("AQUA", HSSFColor.AQUA.index);
        colorShort.put("LIGHT_BLUE", HSSFColor.LIGHT_BLUE.index);
        colorShort.put("VIOLET", HSSFColor.VIOLET.index);
        colorShort.put("GREY_40_PERCENT", HSSFColor.GREY_40_PERCENT.index);
        colorShort.put("PINK", HSSFColor.PINK.index);
        colorShort.put("GOLD", HSSFColor.GOLD.index);
        colorShort.put("YELLOW", HSSFColor.YELLOW.index);
        colorShort.put("BRIGHT_GREEN", HSSFColor.BRIGHT_GREEN.index);
        colorShort.put("TURQUOISE", HSSFColor.TURQUOISE.index);
        colorShort.put("DARK_RED", HSSFColor.DARK_RED.index);
        colorShort.put("SKY_BLUE", HSSFColor.SKY_BLUE.index);
        colorShort.put("PLUM", HSSFColor.PLUM.index);
        colorShort.put("GREY_25_PERCENT", HSSFColor.GREY_25_PERCENT.index);
        colorShort.put("ROSE", HSSFColor.ROSE.index);
        colorShort.put("LIGHT_YELLOW", HSSFColor.LIGHT_YELLOW.index);
        colorShort.put("LIGHT_GREEN", HSSFColor.LIGHT_GREEN.index);
        colorShort.put("LIGHT_TURQUOISE", HSSFColor.LIGHT_TURQUOISE.index);
        colorShort.put("PALE_BLUE", HSSFColor.PALE_BLUE.index);
        colorShort.put("LAVENDER", HSSFColor.LAVENDER.index);
        colorShort.put("WHITE", HSSFColor.WHITE.index);
        colorShort.put("CORNFLOWER_BLUE", HSSFColor.CORNFLOWER_BLUE.index);
        colorShort.put("LEMON_CHIFFON", HSSFColor.LEMON_CHIFFON.index);
        colorShort.put("MAROON", HSSFColor.MAROON.index);
        colorShort.put("ORCHID", HSSFColor.ORCHID.index);
        colorShort.put("CORAL", HSSFColor.CORAL.index);
        colorShort.put("ROYAL_BLUE", HSSFColor.ROYAL_BLUE.index);
        colorShort.put("LIGHT_CORNFLOWER_BLUE",
                HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
        colorShort.put("TAN", HSSFColor.TAN.index);

        borderStyle = new HashMap<>(0);
        borderStyle.put("THIN", HSSFCellStyle.BORDER_THIN);
        borderStyle.put("MEDIUM", HSSFCellStyle.BORDER_MEDIUM);
        borderStyle.put("DASHED", HSSFCellStyle.BORDER_DASHED);
        borderStyle.put("DOTTED", HSSFCellStyle.BORDER_DOTTED);
        borderStyle.put("THICK", HSSFCellStyle.BORDER_THICK);
        borderStyle.put("DOUBLE", HSSFCellStyle.BORDER_DOUBLE);
        borderStyle.put("BORDER_HAIR", HSSFCellStyle.BORDER_HAIR);
        borderStyle.put("MEDIUM_DASHED", HSSFCellStyle.BORDER_MEDIUM_DASHED);
        borderStyle.put("DASH_DOT", HSSFCellStyle.BORDER_DASH_DOT);
        borderStyle
                .put("MEDIUM_DASH_DOT", HSSFCellStyle.BORDER_MEDIUM_DASH_DOT);
        borderStyle.put("DASH_DOT_DOT", HSSFCellStyle.BORDER_DASH_DOT_DOT);
        borderStyle.put("MEDIUM_DASH_DOT_DOT",
                HSSFCellStyle.BORDER_MEDIUM_DASH_DOT_DOT);
        borderStyle.put("SLANTED_DASH_DOT",
                HSSFCellStyle.BORDER_SLANTED_DASH_DOT);
        borderStyle.put("NONE", HSSFCellStyle.BORDER_NONE);

    }

    private HSSFWorkbook workBook;

    private ExportUtil() {
        workBook = new HSSFWorkbook();
    }

    /**
     * 获取水平对齐方式
     * 
     * @param align
     *            对齐方式
     * @return
     */
    private short getAlignment(String align) {
        if (StringUtils.isBlank(align)) {
            align = "CENTER";
        }
        short st = -1;
        if ("CENTER".equalsIgnoreCase(align)) {
            st = HSSFCellStyle.ALIGN_CENTER;
        } else if ("GENERAL".equalsIgnoreCase(align)) {
            st = HSSFCellStyle.ALIGN_GENERAL;
        } else if ("JUSTIFY".equalsIgnoreCase(align)) {
            st = HSSFCellStyle.ALIGN_JUSTIFY;
        } else if ("LEFT".equalsIgnoreCase(align)) {
            st = HSSFCellStyle.ALIGN_LEFT;
        } else if ("RIGHT".equalsIgnoreCase(align)) {
            st = HSSFCellStyle.ALIGN_RIGHT;
        } else if ("FILL".equalsIgnoreCase(align)) {
            st = HSSFCellStyle.ALIGN_FILL;
        } else {
            st = HSSFCellStyle.ALIGN_CENTER_SELECTION;
        }
        return st;
    }

    /**
     * 获取垂直对齐方式
     * 
     * @param vAlign
     *            对齐方式
     * @return
     */
    private short getVAlignment(String vAlign) {
        if (StringUtils.isBlank(vAlign)) {
            vAlign = "CENTER";
        }
        if ("CENTER".equalsIgnoreCase(vAlign)) {
            return HSSFCellStyle.VERTICAL_CENTER;
        } else if ("BOTTOM".equalsIgnoreCase(vAlign)) {
            return HSSFCellStyle.VERTICAL_BOTTOM;
        } else if ("JUSTIFY".equalsIgnoreCase(vAlign)) {
            return HSSFCellStyle.VERTICAL_JUSTIFY;
        } else {
            return HSSFCellStyle.VERTICAL_TOP;
        }
    }

    /**
     * 获取边框风格
     * 
     * @param border
     *            边框风格
     * @return
     */
    private short getBorderStyle(String border) {
        if (StringUtils.isBlank(border)) {
            border = "THIN";
        }
        return borderStyle.get(border.toUpperCase());
    }

    /**
     * 从Color对象获取颜色值
     * 
     * @param color
     *            颜色
     * @return
     */
    private Short getColorByClr(Color color) {
        if (color == null) {
            return null;
        }
        byte r = (byte) color.getRed();
        byte g = (byte) color.getGreen();
        byte b = (byte) color.getBlue();
        HSSFPalette palette = workBook.getCustomPalette();
        HSSFColor hssfColor = null;
        try {
            hssfColor = palette.findColor(r, g, b);
            if (hssfColor == null) {
                palette.setColorAtIndex(HSSFColor.WHITE.index, r, g, b);
                hssfColor = palette.getColor(HSSFColor.WHITE.index);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return hssfColor.getIndex();
    }

    /**
     * 从字符串获取颜色值
     * 
     * @param color
     *            颜色
     * @return
     */
    private Short getColorByStr(String color) {
        if (StringUtils.isBlank(color)) {
            return null;
        }
        return colorShort.get(color.toUpperCase());
    }

    /**
     * 获取颜色值
     * 
     * @param obj
     *            颜色
     * @return
     */
    private Short getColor(Object obj) {
        if (obj == null) {
            return null;
        } else if (obj instanceof Color) {
            return getColorByClr((Color) obj);
        } else {
            return getColorByStr((String) obj);
        }
    }

    /**
     * 设置边框颜色
     * 
     * @param style
     *            风格
     * @param diyColor
     *            颜色
     */
    private void setBorderColor(HSSFCellStyle style, Style.Square<Object> diyColor) {
        if (style == null || diyColor == null) {
            return;
        }
        Short st = getColor(diyColor.getTop());
        style.setTopBorderColor(st == null ? style.getTopBorderColor() : st);
        st = getColor(diyColor.getLeft());
        style.setLeftBorderColor(st == null ? style.getLeftBorderColor() : st);
        st = getColor(diyColor.getRight());
        style.setRightBorderColor(st == null ? style.getRightBorderColor() : st);
        st = getColor(diyColor.getBottom());
        style.setBottomBorderColor(st == null ? style.getBottomBorderColor()
                : st);
    }

    /**
     * 设置边框
     * 
     * @param style
     *            单元格风格
     * @param border
     *            边框对象
     */
    private void setBorder(HSSFCellStyle style, Style.Border border) {
        if (style == null) {
            return;
        }
        setBorderColor(style, border == null ? null : border.getColor());
        Style.Square<String> diyStyle = border == null ? null : border.getStyle();
        style.setBorderTop(getBorderStyle(diyStyle == null ? null : diyStyle
                .getTop()));
        style.setBorderLeft(getBorderStyle(diyStyle == null ? null : diyStyle
                .getLeft()));
        style.setBorderRight(getBorderStyle(diyStyle == null ? null : diyStyle
                .getBottom()));
        style.setBorderBottom(getBorderStyle(diyStyle == null ? null : diyStyle
                .getRight()));
    }

    /**
     * 创建一行
     * 
     * @param index
     *            行索引
     * @param style
     *            行风格
     * @param sheet
     *            行所属Sheet
     * @return
     */
    private HSSFRow createRow(int index, Style style, HSSFSheet sheet) {
        HSSFRow row = sheet.createRow(index);
        if (style != null) {
            if (style.getHeight() != null) {
                row.setHeight((short) (20 * style.getHeight()));
            }
        }
        return row;
    }

    /**
     * 创建字体
     * 
     * @param valueStyle
     *            风格
     * @return
     */
    private HSSFFont createFont(Style valueStyle) {
        // 生成一个字体
        HSSFFont font = workBook.createFont();

        short fontSize = 12;
        if (valueStyle != null) {
            // 字体颜色
            Short st = getColor(valueStyle.getFontColor());
            if (st != null) {
                font.setColor(st);
            }

            if (valueStyle.getBold()) {
                font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
            }

            if (valueStyle.getFontSize() != null) {
                fontSize = valueStyle.getFontSize().byteValue();
            }
        }
        font.setFontHeightInPoints(fontSize);
        return font;
    }

    /**
     * 创建单元格风格
     * 
     * @param index
     *            单元格索引
     * @param style
     *            风格           
     * @param fieldStyle
     *            风格
     * @param sheet
     *            单元格所属Sheet
     * @return
     */
    private HSSFCellStyle createCellStyle(int index, HSSFCellStyle style,
            Style fieldStyle, HSSFSheet sheet) {
        if (style == null) {
            return null;
        }
        style.setWrapText(true);// 回绕文本

        if (fieldStyle != null) {
            Short st = null;

            // 单元格宽度
            if (fieldStyle.getWidth() != null) {
                sheet.setColumnWidth(index, 256 * fieldStyle.getWidth());
            }

            // 单元格前景色
            st = getColor(fieldStyle.getColor());
            if (st != null) {
                style.setFillForegroundColor(st); // 设置单元格前景色
            }

        }
        setBorder(style, fieldStyle == null ? null : fieldStyle.getBorder());
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setAlignment(getAlignment(fieldStyle == null ? null : fieldStyle
                .getAlign())); // 设置对齐方式
        style.setVerticalAlignment(getVAlignment(fieldStyle == null ? null
                : fieldStyle.getVAlign()));
        return style;
    }

    /**
     * 创建单元格
     * 
     * @param index
     *            单元格索引
     * @param style
     *            风格           
     * @param fieldBean
     *            实体所属字段
     * @param value
     *            实体值
     * @param langCode
     *            语言信息
     * @param row
     *            单元格所属行
     * @param sheet
     *            单元格所属Sheet
     */
    private void createCell(int index, HSSFCellStyle style,
            FieldBean fieldBean, Object value, Object rowValue, LangCode langCode, HSSFRow row,
            HSSFSheet sheet) {
        if (style == null || fieldBean == null || langCode == null
                || row == null || sheet == null) {
            return;
        }
        try {
            HSSFCell cell = row.createCell(index);

            String textValue = null;
            if (value == null) {
                textValue = fieldBean.getTitle();
            } else {
                Wrapper valueBean = fieldBean.getWrapper();
                if (valueBean == null) {
                    // 其它数据类型都当作字符串简单处理
                    textValue = value.toString();
                } else {
                    textValue = valueBean.value(value, rowValue, langCode);
                }
            }
            cell.setCellValue(new HSSFRichTextString(textValue));
            cell.setCellStyle(style);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 创建Excel头
     * 
     * @param fields
     *            字段数组
     * @param rowStyle
     *            行风格
     * @param langCode
     *            语言信息
     * @param sheet
     *            所属Sheet
     */
    private void createHeader(FieldBean[] fields, Style rowStyle,
                              LangCode langCode, HSSFSheet sheet) {
        HSSFRow header = createRow(0, rowStyle, sheet);
        int cellIndex = 0;
        for (FieldBean field : fields) {
            HSSFCellStyle style = workBook.createCellStyle();
            // 把字体应用到当前的样式
            style.setFont(createFont(field.getStyle()));
            createCell(
                    cellIndex,
                    createCellStyle(cellIndex, style,
                            field.getStyle(), sheet), field, null, null, langCode,
                    header, sheet);
            cellIndex++;
        }
    }

    /**
     * 得到属性值
     * 
     * @param <T>
     *            导出实体的类型
     * @param t
     *            属性所在对象
     * @param fieldName
     *            属性名
     * @return
     * @throws NoSuchMethodException
     *             异常
     * @throws SecurityException
     *             异常
     * @throws IllegalAccessException
     *             异常
     * @throws IllegalArgumentException
     *             异常
     * @throws InvocationTargetException
     *             异常
     */
    @SuppressWarnings("unchecked")
    private <T> Object getValue(T t, String fieldName) {
        Object obj = null;
        try {
            if (StringUtils.isBlank(fieldName)) return null;
            if (fieldName.contains(".")) {
                String[] _array = fieldName.split("\\.");
                for (String fn : _array) {
                    obj = getValue(obj == null ? t : obj, fn);
                }
            } else {
                if (t == null) {
                    obj = null;
                } else if (t instanceof Map) {
                    Map<String, Object> map = (Map<String, Object>) t;
                    obj = map.containsKey(fieldName) ? map.get(fieldName) : null;
                } else {
                    String getMethodName = "get"
                            + fieldName.substring(0, 1).toUpperCase()
                            + fieldName.substring(1);
                    
                    Method getMethod = t.getClass().getMethod(getMethodName);
                    obj = getMethod.invoke(t);
                }
            }
        } catch (Exception e) {}
        return obj == null ? "" : obj;
    }

    /**
     * 导出数据
     * 
     * @param <T>
     *            导出实体的类型
     * @param sheetName
     *            sheet名称
     * @param out
     *            输出数据源
     * @param entityBean
     *            导出实体描述信息
     * @param dataset
     *            需要导出的数据
     * @throws IllegalAccessException
     *             异常
     * @throws IllegalArgumentException
     *             异常
     * @throws InvocationTargetException
     *             异常
     * @throws NoSuchMethodException
     *             异常
     * @throws SecurityException
     *             异常
     * @throws IOException
     *             异常
     */
    private <T> void export(String sheetName, OutputStream out,
            EntityBean entityBean, Collection<T> dataset, LangCode langCode)
            throws IllegalAccessException, IllegalArgumentException,
            InvocationTargetException, NoSuchMethodException,
            SecurityException, IOException {
        if (StringUtils.isBlank(sheetName) || out == null || entityBean == null || langCode == null) {
            return;
        }
        // 生成一个表格
        HSSFSheet sheet = workBook.createSheet(sheetName);
        // 设置默认列宽
        sheet.setDefaultColumnWidth(20);

        FieldBean[] fields = entityBean.getFields();
        Style valueStyle = entityBean.getValueStyle();
        createHeader(fields, entityBean.getStyle(), langCode, sheet);

        if (dataset != null) {
            List<HSSFCellStyle> cellStyles = new ArrayList<HSSFCellStyle>(0);
            Iterator<T> it = dataset.iterator();
            int index = 1;
            while (it.hasNext()) {
                HSSFRow row = createRow(index, valueStyle, sheet);
                T t = it.next();
                int cellIndex = 0;
                for (FieldBean field : fields) {
                    if (cellStyles.size() - 1 < cellIndex) {
                        HSSFCellStyle style = workBook.createCellStyle();
                        // 把字体应用到当前的样式
                        style.setFont(createFont(field.getValueStyle()));
                        cellStyles.add(cellIndex, createCellStyle(cellIndex, style,
                                field.getValueStyle(), sheet));
                    }
                    createCell(cellIndex, cellStyles.get(cellIndex), field,
                            getValue(t, field.getName()), t, langCode, row, sheet);
                    cellIndex++;
                }
                index++;
            }
        }
        workBook.write(out);
    }

    /**
     * 导出数据到所指定的输出源
     * 
     * <br>
     * 如果是输出到Response,需要写下面两句 <br>
     * response.setHeader("Content-Type", "application/vnd.ms-excel"); <br>
     * response.setHeader("Content-Disposition",
     * "attachment;fileName=xxxxx.xls");
     * 
     * @param <T>
     *            导出实体的类型
     * @param sheetName
     *            sheet名称
     * @param out
     *            输出源
     * @param in
     *            模板
     * @param langCode
     *            语言信息
     * @param dataset
     *            需要导出的数据
     * @throws Exception
     *             异常
     */
    public static <T> void exportExcel(String sheetName, OutputStream out,
                                       InputStream in, LangCode langCode, Collection<T> dataset)
            throws Exception {
        if (StringUtils.isBlank(sheetName) || out == null || in == null
                || langCode == null) {
            return;
        }
        new ExportUtil().export(sheetName, out, EntityBean.getEntityBean(in, langCode), dataset, langCode);
    }
    
}
