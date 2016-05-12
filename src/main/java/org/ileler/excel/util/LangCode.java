package org.ileler.excel.util;

import org.apache.commons.lang3.StringUtils;

/**
 * Created by zhangle on 2016/5/11.
 */
public enum LangCode {

    ZH_CN,
    EN_US;

    /**
     * 根据字符串返回Lang对象
     * @param lang
     *          lang字符串
     * @return
     */
    public static LangCode getLang(String lang) {
        if (StringUtils.isBlank(lang)) {
            return null;
        }
        try {
            return LangCode.valueOf(lang.toUpperCase());
        } catch (Exception e) { }
        return LangCode.ZH_CN;
    }

}

