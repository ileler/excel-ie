/**
 * @(#)EntityBean.java 1.0 2015年10月28日
 * @Copyright:  Copyright 2007 - 2015 MPR Tech. Co. Ltd. All Rights Reserved.
 * @Description: 
 * 
 * Modification History:
 * Date:        2015年10月28日
 * Author:      zhangle
 * Version:     1.0.0.0
 * Description: (Initialize)
 * Reviewer:    
 * Review Date: 
 */
package org.ileler.excel.e;

import org.apache.commons.lang3.StringUtils;
import org.ileler.excel.util.LangCode;
import org.ileler.excel.wrapper.Wrapper;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Hashtable;
import java.util.List;
import java.util.Map;

/**
 * Copyright: Copyright 2007 - 2015 MPR Tech. Co. Ltd. All Rights Reserved.
 * Date: 2015年10月29日 上午8:46:03 Author: zhangle Version: 1.0.0.0 Description:
 * 导出实体定义类型
 */
public final class EntityBean {

    private static final String DOT = ".";

    private static final String ROW = "row";

    private static final String NAME = "name";

    private static final String TITLE = "title";

    private static final String CLASS = "class";

    private static final String FIELD = "field";

    private static final String STYLE = "style";

    private static final String PROP_KEY = "key";

    private static final String VSTYLE = "vstyle";

    private static final String PROP_VALUE = "value";

    private static final String PROPERTY = "property";

    private static final String WRAPPER = "wrapper";

    /**
     * 文档构建对象
     */
    private static ThreadLocal<DocumentBuilder> docBuildeIns = new ThreadLocal<DocumentBuilder>() {
        protected DocumentBuilder initialValue() {
            try {
                return DocumentBuilderFactory.newInstance()
                        .newDocumentBuilder();
            } catch (ParserConfigurationException e) {
                String msg = "DocumentBuilder 对象初始化失败！";
                throw new IllegalStateException(msg, e);
            }
        }
    };

    /**
     * 样式
     */
    private Style style;

    /**
     * 样式
     */
    private Style valueStyle;

    /**
     * 需要导出实体的字段数组
     */
    private FieldBean[] fields;

    private EntityBean(Style style, Style valueStyle, FieldBean[] fields) {
        super();
        this.style = style;
        this.fields = fields;
        this.valueStyle = valueStyle;
    }

    public Style getStyle() {
        return style;
    }

    public FieldBean[] getFields() {
        return fields;
    }

    public Style getValueStyle() {
        return valueStyle;
    }

    /**
     * 获取元素的属性集合
     * 
     * @param nodeList
     *            元素子节点列表
     * @return
     */
    private static Map<String, String> getProperties(NodeList nodeList) {
        if (nodeList == null) {
            return null;
        }
        Map<String, String> map = null;
        try {
            for (int a = 0, b = nodeList.getLength(); a < b
                    && nodeList.item(a) instanceof Element; a++) {
                Element propertyElement = (Element) nodeList.item(a);
                String key = propertyElement.getAttribute(PROP_KEY);
                String value = StringUtils.isBlank(key) ? null
                        : propertyElement.getAttribute(PROP_VALUE);
                if (StringUtils.isBlank(value)) {
                    continue;
                }
                if (map == null) {
                    map = new Hashtable<>();
                }
                map.put(key, value);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return map;
    }

    /**
     * 获取字段的ValueSet对象
     * 
     * @param valueElement
     *            valueset元素对象
     * @return
     */
    private static Wrapper getValueSet(Element valueElement) {
        if (valueElement == null) {
            return null;
        }
        String clazz = valueElement.getAttribute(CLASS);
        if (StringUtils.isBlank(clazz)) {
            return null;
        }
        try {
            return (Wrapper) Class
                    .forName(
                            (clazz.startsWith(DOT) ? Wrapper.class
                                    .getPackage().getName() : "") + clazz)
                    .getConstructor(Element.class, Map.class)
                    .newInstance(
                            valueElement,
                            getProperties(valueElement
                                    .getElementsByTagName(PROPERTY)));
        } catch (InstantiationException | IllegalAccessException
                | IllegalArgumentException | InvocationTargetException
                | NoSuchMethodException | SecurityException
                | ClassNotFoundException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 获取字段对象
     * 
     * @param rowNodeList
     *            行字段节点
     * @param fieldElement
     *            字段元素
     * @return
     */
    private static FieldBean getFieldBean(NodeList rowNodeList,
            Element fieldElement) {
        if (fieldElement == null) {
            return null;
        }
        try {
            String name = fieldElement.getAttribute(NAME); // 读取字段名
            String title = fieldElement.getAttribute(TITLE); // 读取字段标题
            NodeList valuesetNodeList = fieldElement
                    .getElementsByTagName(WRAPPER); // 读取字段值处理
            if (StringUtils.isBlank(name) || StringUtils.isBlank(title)) {
                // 字段没有名字或标题则无效
                return null;
            }
            Element valuesetElement = (valuesetNodeList == null
                    || valuesetNodeList.getLength() < 1 || !(valuesetNodeList
                    .item(0) instanceof Element)) ? null
                    : (Element) valuesetNodeList.item(0); // 得到 值处理 元素
            return new FieldBean(name, getStyle(rowNodeList, fieldElement,
                    STYLE), getStyle(rowNodeList, fieldElement, VSTYLE),
                    getValueSet(valuesetElement),
                    title);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    /**
     * 得到风格对象
     * 
     * @param nodeList
     *            节点列表
     * @param attr
     *            属性名称
     * @return
     */
    private static Style getStyle(NodeList nodeList, String attr) {
        Element element = (nodeList == null || nodeList.getLength() < 1 || !(nodeList
                .item(0) instanceof Element)) ? null : (Element) nodeList
                .item(0);
        return Style.getStyle(element == null || attr == null ? null : element
                .getAttribute(attr));
    }

    /**
     * 得到风格对象
     * 
     * @param parentNodeList
     *            父节点列表
     * @param element
     *            风格元素
     * @param attr
     *            属性名称
     * @return
     */
    private static Style getStyle(NodeList parentNodeList, Element element,
            String attr) {
        Element parent = (parentNodeList == null
                || parentNodeList.getLength() < 1 || !(parentNodeList.item(0) instanceof Element)) ? null
                : (Element) parentNodeList.item(0);
        return Style.getStyles(
                parent == null || attr == null ? null : parent
                        .getAttribute(attr),
                element == null || attr == null ? null : element
                        .getAttribute(attr));
    }

    /**
     * 获取EntityBean对象
     * 
     * @param <T>
     *            导出实体的类型
     * @param src
     *            数据源
     * @param langCode
     *            语言信息
     * @return
     */
    public static <T> EntityBean getEntityBean(InputStream src,
            LangCode langCode) {
        if (src == null || langCode == null) {
            return null;
        }
        try {
            Document doc = docBuildeIns.get().parse(src); // 将数据源转换为文档对象
            NodeList rowNodeList = doc.getElementsByTagName(ROW); // 得到所有行元素集合
            NodeList fieldNodeList = doc.getElementsByTagName(FIELD); // 得到所有字段节点集合
            if (fieldNodeList == null) {
                return null;
            }
            List<FieldBean> fieldBeans = new ArrayList<FieldBean>();
            // 遍历字段节点集合
            for (int i = 0, j = fieldNodeList.getLength(); i < j
                    && fieldNodeList.item(i) instanceof Element; i++) {
                FieldBean fieldBean = getFieldBean(rowNodeList,
                        (Element) fieldNodeList.item(i)); // 获取单个字段节点
                if (fieldBean != null) {
                    fieldBeans.add(fieldBean);
                }
            }
            return fieldBeans.size() < 1 ? null : new EntityBean(getStyle(
                    rowNodeList, STYLE), getStyle(rowNodeList, VSTYLE),
                    fieldBeans.toArray(new FieldBean[fieldBeans
                            .size()]));
        } catch (SAXException | IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 获取EntityBean对象
     * 
     * @param <T>
     *            导出实体的类型
     * @param src
     *            数据源
     * @param langCode
     *            语言信息
     * @param clazz
     *            导出实体类型
     * @return
     */
    public static <T> EntityBean getEntityBean(String src, LangCode langCode,
            Class<T> clazz) {
        if (StringUtils.isBlank(src)) {
            return null;
        }
        return getEntityBean(new ByteArrayInputStream(src.getBytes()), langCode);
    }

}
