package org.ileler.excel.wrapper;

import org.apache.commons.lang3.builder.ToStringBuilder;
import org.ileler.excel.util.LangCode;
import org.w3c.dom.Element;

import java.util.Map;

/**
 * Created by ileler@qq.com on 2016/5/12.
 */
public abstract class Wrapper {

    protected Element element;

    protected Map<String, String> map;

    public Wrapper(Element element, Map<String, String> map) {
        super();
        this.map = map;
        this.element = element;
    }

    @Override
    public String toString() {
        return ToStringBuilder.reflectionToString(this);
    }

    /**
     * 设置值
     * 
     * @param value
     *            原始值
     * @return
     */
    public String value(Object value) {
        return value(value, null);
    }
    
    /**
     * 设置值
     * 
     * @param value
     *            原始值
     * @param rowValue
     *            值所属对象           
     * @return
     */
    public String value(Object value, Object rowValue) {
        return value(value, rowValue, null);
    }

    /**
     * 设置值
     * 
     * @param value
     *            原始值
     * @param rowValue
     *            值所属对象              
     * @param langCode
     *            语言信息
     * @return
     */
    public abstract String value(Object value, Object rowValue, LangCode langCode);

}
