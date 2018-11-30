package com.yiqiang.learn.poi.util;


import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

/**
 * Title:
 * Description:
 *      用来存储Excel标题的对象，通过该对象可以获取标题和方法的对应关系
 * Create Time: 2018/10/22 0:17
 *
 * @author: YEEChan
 * @version: 1.0
 */
@Getter
@Setter
@ToString
public class ExcelHeader implements Comparable<ExcelHeader> {

    /**标题顺序*/
    private int order;

    /**标题所对应的实体类方法*/
    private String methodName;

    /**excel的标题名称*/
    private String title;


    public ExcelHeader() {
    }

    public ExcelHeader(int order, String title, String methodName) {
        this.order = order;
        this.title = title;
        this.methodName = methodName;
    }

    @Override
    public int compareTo(ExcelHeader o) {
        return order > o.order ? 1 : (order < o.order ? -1 : 0);
    }

}
