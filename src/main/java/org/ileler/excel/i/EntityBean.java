package org.ileler.excel.i;

import org.apache.commons.lang3.StringUtils;
import org.ileler.excel.util.LangCode;
import org.ileler.excel.validator.Validator;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;

/**
 * Created by ileler@qq.com on 2016/5/12.
 */
public class EntityBean {
    
    private static final String EFIELD = "emsg";
    
    private static final String ESPLIT = "/";
    
    private String efname;
    
    private String esplit;
    
    private String etitle;
    
    private FieldBean[] fieldBeans;
    
    private Validator[] validators;

    public EntityBean(String efname, String esplit, String etitle, FieldBean[] fieldBeans, Validator[] validators) {
        super();
        this.efname = StringUtils.isBlank(efname) ? EFIELD : efname;
        this.esplit = StringUtils.isBlank(esplit) ? ESPLIT : esplit;
        this.etitle = etitle;
        this.validators = validators;
        this.fieldBeans = fieldBeans;
    }

    public String getEfname() {
        return efname;
    }

    public void setEfname(String efname) {
        this.efname = efname;
    }

    public String getEsplit() {
        return esplit;
    }

    public void setEsplit(String esplit) {
        this.esplit = esplit;
    }

    public String getEtitle() {
        return etitle;
    }

    public void setEtitle(String etitle) {
        this.etitle = etitle;
    }

    public FieldBean[] getFieldBeans() {
        return fieldBeans;
    }

    public void setFieldBeans(FieldBean[] fieldBeans) {
        this.fieldBeans = fieldBeans;
    }

    public Validator[] getValidators() {
        return validators;
    }

    public void setValidators(Validator[] validators) {
        this.validators = validators;
    }
    
    private Map<String, Object> getErrorMap(Map<FieldBean, Object> map) {
        if (map == null)    return null;
        Iterator<Entry<FieldBean, Object>> iter = map.entrySet().iterator();
        Map<String, Object> _map = new HashMap<>(0);
        String errField = null;
        while (iter.hasNext()) {
            Entry<FieldBean, Object> entry = iter.next();
            if (StringUtils.isEmpty(errField = entry.getKey().getEfName()))     continue;
            _map.put(errField, entry.getValue());
        }
        return _map;
    }
    
    private Map<String, Object> getRightMap(Map<FieldBean, Object> map) {
        if (map == null)    return null;
        Iterator<Entry<FieldBean, Object>> iter = map.entrySet().iterator();
        Map<String, Object> _map = new HashMap<>(0);
        FieldBean fieldBean = null;
        String field = null;
        Object value = null;
        while (iter.hasNext()) {
            Entry<FieldBean, Object> entry = iter.next();
            if ((fieldBean = entry.getKey()) == null || StringUtils.isEmpty(field = fieldBean.getRfName()))     continue;
            _map.put(field, (value = entry.getValue()) == null ? fieldBean.getDefValue() : value);
        }
        return _map;
    }
    
    public boolean checkRow(String[] values, Map<String, Object> map, LangCode langCode, Map<String, Object> session) {
        if (values == null || values.length < 1)    return false;
        String colMsg = null, rowMsg = null;
        Map<FieldBean, Object> _map = null;
        FieldBean[] fieldBeans = getFieldBeans();
        FieldBean fieldBean = null;
        Validator[] ruleBeans = null;
        //循环列
        for (int a = 0, b = fieldBeans.length; a < b; a++) {
            String col = a < values.length ? values[a] : null;
            fieldBean = (fieldBean = fieldBeans[a]) == null ? new FieldBean(a, null, null, null, null, langCode, null) : fieldBean;
            //判断是否有格式要求、有则循环判断各个格式要求验证是否通过
            if ((ruleBeans = fieldBean.getValidators()) != null && ruleBeans.length > 0) {
                for (Validator ruleBean : ruleBeans) {
                    if (ruleBean == null)   continue;
                    Validator _ruleBean = ruleBean.clone();
                    try {
                        //如果碰到格式验证不通过则记录错误消息并退出判断
                        if (!_ruleBean.check(col, langCode, session)) {
                            colMsg = _ruleBean.getEmsg();
                            break;
                        }  
                    } catch (Exception e) {}
                }
                //叠加当前行各列格式错误信息
                rowMsg = (!StringUtils.isEmpty(rowMsg) ? rowMsg + (!StringUtils.isEmpty(colMsg) ? esplit : "") : "") + (!StringUtils.isEmpty(colMsg) ? colMsg : "");
                colMsg = null;
            }
            if (_map == null)   _map = new HashMap<>(0);
            _map.put(fieldBean, col == null ? "" : col);
        }
        Map<String, Object> __map = null;
        if (StringUtils.isEmpty(rowMsg)) {
            if ((ruleBeans = getValidators()) != null && ruleBeans.length > 0) {
                __map = getRightMap(_map);
                for (Validator ruleBean : ruleBeans) {
                    if (ruleBean == null)   continue;
                    Validator _ruleBean = ruleBean.clone();
                    try {
                        //行错误不要退出，继续判断
                        if (!_ruleBean.check(__map, langCode, session)) {
                            //叠加当前行各列格式错误信息
                            colMsg = _ruleBean.getEmsg();
                            rowMsg = (!StringUtils.isEmpty(rowMsg) ? rowMsg + (!StringUtils.isEmpty(colMsg) ? esplit : "") : "") + (!StringUtils.isEmpty(colMsg) ? colMsg : "");
                        }
                    } catch (Exception e) {}
                }
            } else {
                __map = getRightMap(_map); 
            }
        }
        if (StringUtils.isEmpty(rowMsg)) {
            if (map != null)    map.putAll(__map);
            return true;
        } else {
            if (map != null) {
                __map =  getErrorMap(_map);
                __map.put(efname, rowMsg);
                map.putAll(__map);
            }
            return false;
        }
    }
    
}
