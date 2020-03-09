package io.cruder.excellent;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * ExcelField: Annotation of excel field.
 *
 * @author cruder
 * @since 2019-04-17
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelField {

    /**
     * @return Title in excel
     */
    String title() default "";

    /**
     * @return Column order in excel
     */
    int order() default 1;
}
