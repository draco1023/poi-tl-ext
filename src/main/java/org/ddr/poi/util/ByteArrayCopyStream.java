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

package org.ddr.poi.util;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;

/**
 * 可以复用输出的二进制数据并转换为输入的流
 *
 * @author Draco
 * @since 2021-02-09
 */
public class ByteArrayCopyStream extends ByteArrayOutputStream {
    public ByteArrayCopyStream() {
    }

    public ByteArrayCopyStream(int size) {
        super(size);
    }

    /**
     * @return 转换为输入流
     */
    public InputStream toInput() {
        return new ByteArrayInputStream(buf, 0, count);
    }
}
