/*
 * Copyright 2016 - 2021 Draco, https://github.com/draco1023
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

package org.ddr.poi.html;

/**
 * HTML元素渲染器提供者
 *
 * @author Draco
 * @since 2022-10-21
 */
@FunctionalInterface
public interface ElementRendererProvider {
    /**
     * 根据HTML元素名称获取渲染器
     *
     * @param tagNormalName HTML元素名称（小写）
     * @return HTML元素渲染器
     */
    ElementRenderer get(String tagNormalName);
}
