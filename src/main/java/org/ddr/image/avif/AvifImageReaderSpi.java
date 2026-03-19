package org.ddr.image.avif;

import com.twelvemonkeys.imageio.spi.ImageReaderSpiBase;

import javax.imageio.ImageReader;
import javax.imageio.stream.ImageInputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.Locale;
import java.util.Set;


public final class AvifImageReaderSpi extends ImageReaderSpiBase {
    private static final Set<String> TYPES;

    static {
        TYPES = new HashSet<>(4);
//        TYPES.add("mif1");
//        TYPES.add("msf1");
        TYPES.add("miaf");
        TYPES.add("avif");
        TYPES.add("avis");
        TYPES.add("avio");
    }
    public AvifImageReaderSpi() {
        super(new AvifProviderInfo());
    }

    @Override
    public boolean canDecodeInput(Object source) throws IOException {
        return source instanceof ImageInputStream && canDecode((ImageInputStream) source);
    }

    private boolean canDecode(ImageInputStream input) throws IOException {
        try {
            input.mark();
            for (int i = 0; i < 4; i++) {
                input.read();
            }
            if (input.read() == 'f' && input.read() == 't' && input.read() == 'y' && input.read() == 'p') {
                byte[] bytes = new byte[4];
                int length = input.read(bytes);
                if (length == 4) {
                    String s = new String(bytes);
                    return TYPES.contains(s);
                }
            }
        } catch (Exception ignored) {
        } finally {
            input.reset();
        }
        return false;
    }

    @Override
    public ImageReader createReaderInstance(Object extension) throws IOException {
        return new AvifImageReader(this);
    }

    @Override
    public String getDescription(Locale locale) {
        return "AV1 Image File (AVIF) format image reader";
    }
}
