package org.ileler.excel.i;

import org.apache.commons.lang3.StringUtils;
import org.ileler.excel.util.LangCode;
import org.ileler.excel.validator.Validator;

/**
 * Created by ileler@qq.com on 2016/5/12.
 */
public class FieldBean {

	private Integer index;
	
	private String title;
	
	private String efName;
	
	private String rfName;
	
	private String defValue;
	
	private LangCode langCode;
	
	private Validator[] validators;

	public FieldBean(Integer index, String title, String efName, String rfName,
					 String defValue, LangCode langCode, Validator[] validators) {
		super();
		this.index = index;
		this.title = title;
		this.efName = efName;
		this.rfName = rfName;
		this.langCode = langCode;
		this.validators = validators;
		this.setDefValue(defValue);
	}

	public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public Integer getIndex() {
		return index;
	}

	public void setIndex(Integer index) {
		this.index = index;
	}

	public String getEfName() {
		return efName;
	}

	public void setEfName(String efName) {
		this.efName = efName;
	}

	public String getRfName() {
		return rfName;
	}

	public void setRfName(String rfName) {
		this.rfName = rfName;
	}

	public String getDefValue() {
		return defValue;
	}

	public LangCode getLangCode() {
        return langCode;
    }

    public void setLangCode(LangCode langCode) {
        this.langCode = langCode;
    }

    public void setDefValue(String defValue) {
		String colMsg = null;
		//判断是否有格式要求、有则循环判断各个格式要求验证是否通过
		if (StringUtils.isEmpty(defValue) || validators == null || validators.length < 1) {
			this.defValue = defValue;
		} else {
			for (Validator ruleBean : validators) {
				if (ruleBean == null) 	continue;
				//如果碰到格式验证不通过则记录错误消息并退出判断
				if (!ruleBean.check(defValue, langCode) && !StringUtils.isEmpty(colMsg = ruleBean.getEmsg())) 	break;
			}
		}
		if (!StringUtils.isEmpty(colMsg)) {
			throw new RuntimeException("set default value error:" + colMsg);
		}
	}

	public Validator[] getValidators() {
		return validators;
	}

	public void setValidators(Validator[] validators) {
		this.validators = validators;
	}
	
	@Override
	public String toString() {
		String str = "{index:" + index + ",title:" + (title == null ? "" : title) + ",efname:" + (efName == null ? "" : efName)
				+ ",rfname:" + (rfName == null ? "" : rfName)
				+ ",defValue:" + (defValue == null ? "" : defValue) + ",validators:[";
		if (validators != null) {
			for (Validator erb : validators) {
				if (erb == null) 	continue;
				str += erb.toString();
			}
		}
		return str + "]}";
	}
	
}
