/*
 * Copyright 2016 - 2022 Draco, https://github.com/draco1023
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package org.ddr.poi.util;

import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;

/**
 * @author Draco
 * @since 2022-02-20 19:15
 */
public class XmlUtils {
    public static final String NS_WORDPROCESSINGML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /**
     * 移除xml元素上声明的命名空间
     *
     * @param xmlObject xml元素
     */
    public static void removeNamespaces(XmlObject xmlObject) {
        XmlCursor cursor = xmlObject.newCursor();
        cursor.toNextToken();
        while (cursor.hasNextToken()) {
            if (cursor.isNamespace()) {
                cursor.removeXml();
            } else {
                cursor.toNextToken();
            }
        }
        cursor.dispose();
    }

}
