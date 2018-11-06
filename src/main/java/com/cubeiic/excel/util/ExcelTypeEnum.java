package com.cubeiic.excel.util;

/**
 * @author hanxuan
 * @date 2018/7/10 15:18
 */
public enum ExcelTypeEnum {

    /**
     * 2003版Excel
     */
    EXCEL_THREE(1,"xls"),
    /**
     * 2007版Excel
     */
    EXCEL_SEVEN(2,"xlsx");


    private Integer key;

    private String text;

    ExcelTypeEnum(Integer key,String text){
        this.key=key;
        this.text=text;
    }

    public Integer getKey() {
        return key;
    }


    public String getText() {
        return text;
    }

    public static ExcelTypeEnum getText(Integer key) {
        for(ExcelTypeEnum typeEnum : ExcelTypeEnum.values()) {
            if((typeEnum.key).equals(key)) {
                return typeEnum;
            }
        }
        throw new IllegalArgumentException("没有元素匹配" + key);
    }


}