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
	String method();//Ҫ���õķ���
	Class<?> clazz();
	/**
	 * 0��ʾû�м���
	 * 	cascadeNumber�������ʾ���ټ�����
	 * @return
	 */
	int cascadeNumber() default 0;
}
