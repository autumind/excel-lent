/*
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */

package io.cruder.excellent.util;

import lombok.experimental.UtilityClass;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.stream.Stream;

/**
 * Reflects: Reflection util.
 *
 * @author cruder
 * @since 2019-03-29
 */
@UtilityClass
public class Reflects {

    /**
     * Common setter name prefix string.
     */
    private static final String SET_PREFIX = "set";

    /**
     * Resolve generic type
     *
     * @param targetClazz target class
     * @param typeIndex   generic type parameter index of target class
     * @return generic type
     */
    public static Type resolveGenericType(Class targetClazz, int typeIndex) {
        if (targetClazz == null) {
            return null;
        }

        if (typeIndex < 0) {
            return null;
        }

        Type[] actualTypeArguments = ((ParameterizedType) targetClazz
                .getGenericSuperclass()).getActualTypeArguments();

        if (actualTypeArguments == null || typeIndex > actualTypeArguments.length) {
            return null;
        }
        return actualTypeArguments[typeIndex];
    }


    /**
     * Resolve generic type(First generic type).
     *
     * @param targetClz target class
     * @return generic type
     */
    public static Type resolveGenericType(Class targetClz) {
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
