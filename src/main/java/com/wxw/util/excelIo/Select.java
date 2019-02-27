package com.wxw.util.excelIo;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Documented
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Select {
	String method();//要调用的方法
	Class<?> clazz();
	/**
	 * 0表示没有级联
	 * 	cascadeNumber大于零表示多少级级联
	 * @return
	 */
	int cascadeNumber() default 0;
}
