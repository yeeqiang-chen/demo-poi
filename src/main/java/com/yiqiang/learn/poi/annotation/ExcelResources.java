package com.yiqiang.learn.poi.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Title:
 * Description:
 *      用来在对象的get方法上加入的annotation,通过该annotation说明某个属性所对应的标题
 * Create Time: 2018/10/22 0:13
 *
 * @author: YEEChan
 * @version: 1.0
 */
@Target({ElementType.METHOD,ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelResources {

    /**属性的标题名称*/
    String title();

    /**标题所在顺序*/
    int order() default 9999;
}
