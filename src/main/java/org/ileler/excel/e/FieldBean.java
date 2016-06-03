/**
 * @(#)FieldBean.java 1.0 2015年10月28日
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

import org.ileler.excel.wrapper.Wrapper;


/**
 * Created by ileler@qq.com on 2016/5/12.
 */
public class FieldBean {

    /**
     * 字段名
     */
    private String name;
    
    /**
     * 字段样式
     */
    private Style style;
    
    /**
     * 值样式
     */
    private Style valueStyle;

    /**
     * 设置值
     */
    private Wrapper wrapper;

    /**
     * 字段标题集合
     */
    private String title;

    public FieldBean(String name, Style style, Style valueStyle, Wrapper wrapper, String title) {
        super();
        this.name = name;
        this.style = style;
        this.title = title;
        this.wrapper = wrapper;
        this.valueStyle = valueStyle;
    }

    public String getName() {
        return name;
    }
    
    public Style getStyle() {
        return style;
    }
    
    public Style getValueStyle() {
        return valueStyle;
    }

    public Wrapper getWrapper() {
        return wrapper;
    }

    public String getTitle() {
        return title;
    }

}
