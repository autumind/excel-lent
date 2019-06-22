package xxx.yyy.zzz;

import lombok.experimental.UtilityClass;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.stream.Stream;

/**
 * ReflectUtil: Reflection util.
 *
 * @author autumind
 * @since 2019-03-29
 */
@UtilityClass
public class ReflectUtil {

    /**
     * Common setter name prefix string.
     */
    private static final String SET_PREFIX = "set";

    /**
     * Resolve generic type
     *
     * @param targetClz target class
     * @param typeIndex generic type parameter index of target class
     * @return generic type
     */
    @Deprecated
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
    @Deprecated
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
        String fieldName = field.getName();
        char[] chars = fieldName.toCharArray();
        if (chars[0] <= 'z' && chars[0] >= 'a') {
            chars[0] = (char) (chars[0] + 'A' - 'a');
        }
        String setterName = SET_PREFIX.concat(String.valueOf(chars));
        return Stream.of(clz.getDeclaredMethods())
                .filter(method -> method.getName().equals(setterName))
                .findAny()
                .orElseThrow(() -> new NoSuchMethodError(String.format("Field [%s] has no setter.", fieldName)));
    }
}
