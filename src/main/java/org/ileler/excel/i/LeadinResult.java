/**
 * @(#)LeadinResult.java 1.0 2016年3月8日
 * @Copyright:  Copyright 2007 - 2016 MPR Tech. Co. Ltd. All Rights Reserved.
 * @Description: 
 * 
 * Modification History:
 * Date:        2016年3月8日
 * Author:      zhangle
 * Version:     1.0.0.0
 * Description: (Initialize)
 * Reviewer:    
 * Review Date: 
 */
package org.ileler.excel.i;

import java.io.Serializable;
import java.util.Map;

/**
 * Created by ileler@qq.com on 2016/5/12.
 */
public class LeadinResult implements Serializable {

    private static final long serialVersionUID = -8776390534801677291L;
    
    private Map<Integer, Map<String, Object>> rightList;
    
    private Map<Integer, Map<String, Object>> errorList;
    
    public LeadinResult() {}

    public LeadinResult(Map<Integer, Map<String, Object>> rightList,
            Map<Integer, Map<String, Object>> errorList) {
        super();
        this.rightList = rightList;
        this.errorList = errorList;
    }

    public Map<Integer, Map<String, Object>> getRightList() {
        return rightList;
    }

    public void setRightList(Map<Integer, Map<String, Object>> rightList) {
        this.rightList = rightList;
    }

    public Map<Integer, Map<String, Object>> getErrorList() {
        return errorList;
    }

    public void setErrorList(Map<Integer, Map<String, Object>> errorList) {
        this.errorList = errorList;
    }
    
    public boolean hasRight() {
        return !(rightList == null || rightList.size() < 1);
    }
    
    public boolean hasError() {
        return !(errorList == null || errorList.size() < 1);
    }

}

