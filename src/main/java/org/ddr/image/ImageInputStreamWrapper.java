package org.ddr.image;

import javax.imageio.stream.IIOByteBuffer;
import javax.imageio.stream.ImageInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.ByteOrder;

public class ImageInputStreamWrapper extends InputStream implements ImageInputStream {

    private final ImageInputStream input;

    public ImageInputStreamWrapper(ImageInputStream input) {
        this.input = input;
    }

    @Override
    public void setByteOrder(ByteOrder byteOrder) {
        input.setByteOrder(byteOrder);
    }

    @Override
    public ByteOrder getByteOrder() {
        return input.getByteOrder();
    }

    @Override
    public int read() throws IOException {
        return input.read();
    }

    @Override
    public void readBytes(IIOByteBuffer buf, int len) throws IOException {
        input.readBytes(buf, len);
    }

    @Override
    public boolean readBoolean() throws IOException {
        return input.readBoolean();
    }

    @Override
    public byte readByte() throws IOException {
        return input.readByte();
    }

    @Override
    public int readUnsignedByte() throws IOException {
        return input.readUnsignedByte();
    }

    @Override
    public short readShort() throws IOException {
        return input.readShort();
    }

    @Override
    public int readUnsignedShort() throws IOException {
        return input.readUnsignedShort();
    }

    @Override
    public char readChar() throws IOException {
        return input.readChar();
    }

    @Override
    public int readInt() throws IOException {
        return input.readInt();
    }

    @Override
    public long readUnsignedInt() throws IOException {
        return input.readUnsignedInt();
    }

    @Override
    public long readLong() throws IOException {
        return input.readLong();
    }

    @Override
    public float readFloat() throws IOException {
        return input.readFloat();
    }

    @Override
    public double readDouble() throws IOException {
        return input.readDouble();
    }

    @Override
    public String readLine() throws IOException {
        return input.readLine();
    }

    @Override
    public String readUTF() throws IOException {
        return input.readUTF();
    }

    @Override
    public void readFully(byte[] b, int off, int len) throws IOException {
        input.readFully(b, off, len);
    }

    @Override
    public void readFully(byte[] b) throws IOException {
        input.readFully(b);
    }

    @Override
    public void readFully(short[] s, int off, int len) throws IOException {
        input.readFully(s, off, len);
    }

    @Override
    public void readFully(char[] c, int off, int len) throws IOException {
        input.readFully(c, off, len);
    }

    @Override
    public void readFully(int[] i, int off, int len) throws IOException {
        input.readFully(i, off, len);
    }

    @Override
    public void readFully(long[] l, int off, int len) throws IOException {
        input.readFully(l, off, len);
    }

    @Override
    public void readFully(float[] f, int off, int len) throws IOException {
        input.readFully(f, off, len);
    }

    @Override
    public void readFully(double[] d, int off, int len) throws IOException {
        input.readFully(d, off, len);
    }

    @Override
    public long getStreamPosition() throws IOException {
        return input.getStreamPosition();
    }

    @Override
    public int getBitOffset() throws IOException {
        return input.getBitOffset();
    }

    @Override
    public void setBitOffset(int bitOffset) throws IOException {
        input.setBitOffset(bitOffset);
    }

    @Override
    public int readBit() throws IOException {
        return input.readBit();
    }

    @Override
    public long readBits(int numBits) throws IOException {
        return input.readBits(numBits);
    }

    @Override
    public long length() throws IOException {
        return input.length();
    }

    @Override
    public int skipBytes(int n) throws IOException {
        return input.skipBytes(n);
    }

    @Override
    public long skipBytes(long n) throws IOException {
        return input.skipBytes(n);
    }

    @Override
    public void seek(long pos) throws IOException {
        input.seek(pos);
    }

    @Override
    public void mark() {
        input.mark();
    }

    @Override
    public void flushBefore(long pos) throws IOException {
        input.flushBefore(pos);
    }

    @Override
    public void flush() throws IOException {
        input.flush();
    }

    @Override
    public long getFlushedPosition() {
        return input.getFlushedPosition();
    }

    @Override
    public boolean isCached() {
        return input.isCached();
    }

    @Override
    public boolean isCachedMemory() {
        return input.isCachedMemory();
    }

    @Override
    public boolean isCachedFile() {
        return input.isCachedFile();
    }
}
