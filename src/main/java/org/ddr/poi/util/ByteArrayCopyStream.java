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
