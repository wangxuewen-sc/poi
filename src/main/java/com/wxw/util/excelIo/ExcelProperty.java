package com.wxw.util.excelIo;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Documented
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelProperty {
	String zh();
	int sequence() default 0;
	boolean requisite() default false;
	DataType dataType() default DataType.NULL;
}
