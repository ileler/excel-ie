/**
 * @(#)FieldStyle.java 1.0 2015年10月29日
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

import java.awt.Color;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Hashtable;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.builder.ToStringBuilder;

public class Style {

    private static final String COM = ",";

    private static final String SEM = ";";

    private static final String COL = ":";

    private static final String SET = "set";

    private static final String GET = "get";

    private static Map<String, Method> map;

    private static Map<String, Method> getterMap;

    static {
        Method[] methods = Style.class.getMethods();
        if (methods != null && methods.length > 0) {
            for (Method method : methods) {
                String name = method.getName();
                if (name.startsWith(SET)) {
                    name = name.substring(SET.length());
                    name = name.substring(0, 1).toLowerCase()
                            + name.substring(1);
                    if (map == null) {
                        map = new Hashtable<>();
                    }
                    map.put(name, method);
                } else if (name.startsWith(GET)) {
                    name = name.substring(GET.length());
                    name = name.substring(0, 1).toLowerCase()
                            + name.substring(1);
                    if (getterMap == null) {
                        getterMap = new Hashtable<>();
                    }
                    getterMap.put(name, method);
                } else {
                    continue;
                }
            }
        }
    }

    private String align;

    private String color;

    private Boolean bold;

    private String vAlign;

    private Integer width;

    private Integer height;

    private Integer fontSize;

    private String fontColor;

    private Border border;

    public Style() {
        super();
    }

    /** --------------------setter-stt-------------------- */

    /**
     * 设置水平对齐方式
     * 
     * @param align
     *            对齐方式
     */
    public void setAlign(String align) {
        this.align = align;
    }

    /**
     * 设置是否加粗
     * 
     * @param bold
     *            是否加粗
     */
    public void setBold(Boolean bold) {
        this.bold = bold;
    }

    /**
     * 设置垂直对齐方式
     * 
     * @param vAlign
     *            对齐方式
     */
    public void setVAlign(String vAlign) {
        this.vAlign = vAlign;
    }

    /**
     * 设置前景颜色
     * 
     * @param color
     *            颜色
     */
    public void setColor(String color) {
        this.color = color;
    }

    /**
     * 设置宽度
     * 
     * @param width
     *            宽度
     */
    public void setWidth(Integer width) {
        this.width = width;
    }

    /**
     * 设置高度
     * 
     * @param height
     *            高度
     */
    public void setHeight(Integer height) {
        this.height = height;
    }

    /**
     * 设置字体大小
     * 
     * @param fontSize
     *            字体大小
     */
    public void setFontSize(Integer fontSize) {
        this.fontSize = fontSize;
    }

    /**
     * 设置字体颜色
     * 
     * @param fontColor
     *            字体颜色
     */
    public void setFontColor(String fontColor) {
        this.fontColor = fontColor;
    }

    /**
     * 设置边框颜色
     * 
     * @param borderColor
     *            颜色
     */
    public void setBorderColor(String borderColor) {
        setBorderColorT(borderColor);
        setBorderColorL(borderColor);
        setBorderColorB(borderColor);
        setBorderColorR(borderColor);
    }

    /**
     * 设置上边框颜色
     * 
     * @param borderColorT
     *            颜色
     */
    public void setBorderColorT(String borderColorT) {
        setBorderColor(borderColorT, Direction.TOP);
    }

    /**
     * 设置左边框颜色
     * 
     * @param borderColorL
     *            颜色
     */
    public void setBorderColorL(String borderColorL) {
        setBorderColor(borderColorL, Direction.LEFT);
    }

    /**
     * 设置下边框颜色
     * 
     * @param borderColorB
     *            颜色
     */
    public void setBorderColorB(String borderColorB) {
        setBorderColor(borderColorB, Direction.BOTTOM);
    }

    /**
     * 设置右边框颜色
     * 
     * @param borderColorR
     *            颜色
     */
    public void setBorderColorR(String borderColorR) {
        setBorderColor(borderColorR, Direction.RIGHT);
    }

    /**
     * 设置边框风格
     * 
     * @param borderStyle
     *            风格
     */
    public void setBorderStyle(String borderStyle) {
        setBorderStyleT(borderStyle);
        setBorderStyleL(borderStyle);
        setBorderStyleB(borderStyle);
        setBorderStyleR(borderStyle);
    }

    /**
     * 设置上边框风格
     * 
     * @param borderStyleT
     *            风格
     */
    public void setBorderStyleT(String borderStyleT) {
        setBorderStyle(borderStyleT, Direction.TOP);
    }

    /**
     * 设置左边框风格
     * 
     * @param borderStyleL
     *            风格
     */
    public void setBorderStyleL(String borderStyleL) {
        setBorderStyle(borderStyleL, Direction.LEFT);
    }

    /**
     * 设置下边框风格
     * 
     * @param borderStyleB
     *            风格
     */
    public void setBorderStyleB(String borderStyleB) {
        setBorderStyle(borderStyleB, Direction.BOTTOM);
    }

    /**
     * 设置右边框风格
     * 
     * @param borderStyleR
     *            风格
     */
    public void setBorderStyleR(String borderStyleR) {
        setBorderStyle(borderStyleR, Direction.RIGHT);
    }

    /** --------------------setter-end-------------------- */

    /** --------------------getter-stt-------------------- */

    public String getAlign() {
        return align;
    }

    public Boolean getBold() {
        return bold == null ? false : bold;
    }

    public String getVAlign() {
        return vAlign;
    }

    public Object getColor() {
        return getColor(this.color);
    }

    public Integer getWidth() {
        return width;
    }

    public Integer getHeight() {
        return height;
    }

    public Integer getFontSize() {
        return fontSize;
    }

    public Object getFontColor() {
        return getColor(this.fontColor);
    }

    public Border getBorder() {
        return border;
    }

    /** --------------------getter-end-------------------- */

    @Override
    public String toString() {
        return ToStringBuilder.reflectionToString(this);
    }

    /**
     * 设置边框颜色
     * 
     * @param src
     *            源字符串
     * @param dir
     *            方向
     */
    private void setBorderColor(String src, Direction dir) {
        if (StringUtils.isBlank(src) || dir == null) {
            return;
        }
        if (border == null) {
            border = new Border();
        }
        if (border.getColor() == null) {
            border.setColor(new Square<>());
        }
        Object obj = getColor(src);
        switch (dir) {
            case TOP:
                border.getColor().setTop(obj);
                break;
            case LEFT:
                border.getColor().setLeft(obj);
                break;
            case BOTTOM:
                border.getColor().setBottom(obj);
                break;
            case RIGHT:
                border.getColor().setRight(obj);
                break;
            default:
                break;
        }
    }

    /**
     * 设置边框风格
     * 
     * @param src
     *            源字符串
     * @param dir
     *            方向
     */
    private void setBorderStyle(String src, Direction dir) {
        if (StringUtils.isBlank(src) || dir == null) {
            return;
        }
        if (border == null) {
            border = new Border();
        }
        if (border.getStyle() == null) {
            border.setStyle(new Square<String>());
        }
        switch (dir) {
            case TOP:
                border.getStyle().setTop(src);
                break;
            case LEFT:
                border.getStyle().setLeft(src);
                break;
            case BOTTOM:
                border.getStyle().setBottom(src);
                break;
            case RIGHT:
                border.getStyle().setRight(src);
                break;
            default:
                break;
        }
    }

    /**
     * 将字符串转为颜色 ,如果颜色由R,G,B组成则自动装换为Color对象，否则直接返回字符串
     * 
     * @param color
     *            源字符串
     * @return
     */
    private Object getColor(String color) {
        if (color == null || color.split(COM).length != 3) {
            return color;
        } else {
            String[] ss = color.split(COM);
            return new Color(Integer.valueOf(ss[0]), Integer.valueOf(ss[1]),
                    Integer.valueOf(ss[2]));
        }
    }

    /**
     * 得到对象值
     * 
     * @param obj
     *            对象值
     * @param attr
     *            属性名
     * @param parent
     *            父对象          
     * @return
     */
    private static Object getObject(Object obj, String attr, Style parent) {
        if (obj == null
                || (obj instanceof String && StringUtils
                        .isBlank(obj.toString()))) {
            Method getter = getterMap.get(attr);
            if (getter != null) {
                try {
                    obj = getter.invoke(parent);
                } catch (IllegalAccessException | IllegalArgumentException
                        | InvocationTargetException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                }
            }
        }
        return obj;
    }

    /**
     * 反射设置风格
     * 
     * @param src
     *            源字符串
     * @return
     */
    public static Style getStyle(String src) {
        return getStyle(src, null);
    }

    /**
     * 反射设置风格
     * 
     * @param src
     *            源字符串
     * @param parent
     *            父对象
     * @return
     */
    public static Style getStyle(String src, Style parent) {
        if (StringUtils.isBlank(src)) {
            return null;
        }
        Style style = null;
        try {
            String[] styles = src.split(SEM);
            for (String str : styles) {
                String[] ss = str.split(COL);
                if (ss.length != 2 || !map.containsKey(ss[0])) {
                    continue;
                }
                if (style == null) {
                    style = new Style();
                }
                Method method = map.get(ss[0]);
                Class<?> clazz = method.getParameterTypes()[0];
                Object obj = null;
                if (clazz.equals(Integer.class)) {
                    obj = Integer.valueOf(ss[1]);
                } else if (clazz.equals(Boolean.class)) {
                    obj = Boolean.valueOf(ss[1]);
                } else {
                    obj = clazz.cast(ss[1]);
                }
                method.invoke(style, getObject(obj, ss[0], parent));
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return style;
    }

    /**
     * 反射设置风格
     * 
     * @param parent
     *            父源字符串
     * @param src
     *            源字符串
     * @return
     */
    public static Style getStyles(String parent, String src) {
        if (StringUtils.isBlank(parent) && StringUtils.isBlank(src)) {
            return null;
        }
        if (StringUtils.isBlank(parent)) {
            return getStyle(src);
        } else if (StringUtils.isBlank(src)) {
            return getStyle(parent);
        } else {
            return getStyle(src, getStyle(parent, null));
        }
    }

    /**
     * Copyright: Copyright 2007 - 2015 MPR Tech. Co. Ltd. All Rights Reserved.
     * Date: 2015年10月30日 下午2:32:34 Author: zhangle Version: 1.0.0.0 Description:
     * 方向枚举类
     */
    public enum Direction {
        TOP, LEFT, BOTTOM, RIGHT;
    }

    /**
     * Copyright: Copyright 2007 - 2015 MPR Tech. Co. Ltd. All Rights Reserved.
     * Date: 2015年10月30日 下午2:31:40 Author: zhangle Version: 1.0.0.0 Description:
     * 方形类
     */
    public class Square<T> {
        /**
         * 上
         */
        private T top;
        /**
         * 左
         */
        private T left;
        /**
         * 下
         */
        private T bottom;
        /**
         * 右
         */
        private T right;

        public Square() {
        }

        public Square(T top, T left, T bottom, T right) {
            super();
            this.top = top;
            this.left = left;
            this.bottom = bottom;
            this.right = right;
        }

        public T getTop() {
            return top;
        }

        public void setTop(T top) {
            this.top = top;
        }

        public T getLeft() {
            return left;
        }

        public void setLeft(T left) {
            this.left = left;
        }

        public T getBottom() {
            return bottom;
        }

        public void setBottom(T bottom) {
            this.bottom = bottom;
        }

        public T getRight() {
            return right;
        }

        public void setRight(T right) {
            this.right = right;
        }
    }

    /**
     * Copyright: Copyright 2007 - 2015 MPR Tech. Co. Ltd. All Rights Reserved.
     * Date: 2015年10月30日 下午2:31:11 Author: zhangle Version: 1.0.0.0 Description:
     * 边框类
     */
    public class Border {

        /**
         * 边框颜色
         */
        private Square<Object> color;

        /**
         * 边框风格
         */
        private Square<String> style;

        public Square<Object> getColor() {
            return color;
        }

        public void setColor(Square<Object> color) {
            this.color = color;
        }

        public Square<String> getStyle() {
            return style;
        }

        public void setStyle(Square<String> style) {
            this.style = style;
        }

    }

}
