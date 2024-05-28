package org.ddr.image.heif;

import com.twelvemonkeys.imageio.spi.ImageReaderSpiBase;

import javax.imageio.ImageReader;
import javax.imageio.stream.ImageInputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.Locale;
import java.util.Set;


public final class HeifImageReaderSpi extends ImageReaderSpiBase {
    private static final Set<String> TYPES;

    static {
        TYPES = new HashSet<>(8);
        TYPES.add("mif1");
        TYPES.add("msf1");
        TYPES.add("heic");
        TYPES.add("heix");
        TYPES.add("hevc");
        TYPES.add("hevx");
    }
    public HeifImageReaderSpi() {
        super(new HeifProviderInfo());
    }

    @Override
    public boolean canDecodeInput(Object source) throws IOException {
        return source instanceof ImageInputStream && canDecode((ImageInputStream) source);
    }

    private boolean canDecode(ImageInputStream input) throws IOException {
        // https://docs.oracle.com/javase/7/docs/technotes/guides/imageio/spec/extending.fm3.html
        // https://github.com/strukturag/libheif/blob/e64bb552f5d48fee5daf69c8c2fd59ec3eee0818/libheif/heif.cc#L102
        // https://devstreaming-cdn.apple.com/videos/wwdc/2017/513fzgbviu23l/513/513_high_efficiency_image_file_format.pdf?dl=1
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
        return new HeifImageReader(this);
    }

    @Override
    public String getDescription(Locale locale) {
        return "High Efficiency Image File (HEIF) format image reader";
    }
}
