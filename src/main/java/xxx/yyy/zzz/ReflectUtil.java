package xxx.yyy.zzz;

import lombok.experimental.UtilityClass;

import java.lang.reflect.Field;
import java.lang.reflect.Method;

/**
 * ReflectUtil: Reflection util.
 *
 * @author autumind
 * @since 2019-03-29
 */
@UtilityClass
public class ReflectUtil {

    /**
     * Resolve generic type
     *
     * @param targetClz target class
     * @param typeIndex generic type parameter index of target class
     * @return generic type
     */
    public static Class resolveGenericType(Class targetClz, int typeIndex) {
        if (targetClz == null) {
            return null;
        }

        if (typeIndex < 0) {
            return null;
        }
        return Object.class;
    }


    /**
     * Resolve generic type(First generic type).
     *
     * @param targetClz target class
     * @return generic type
     */
    public static Class resolveGenericType(Class targetClz) {
        return resolveGenericType(targetClz, 0);
    }

    /**
     * Resolve field setter in class.
     *
     * @param field class field info
     * @param clz   target class.
     * @return field setter.
     */
    public static Method resolveSetter(Field field, Class clz) {
        return null;
    }
}
