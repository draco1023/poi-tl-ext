package org.ddr.image.heif;

import com.twelvemonkeys.imageio.spi.ImageReaderSpiBase;

import javax.imageio.ImageReader;
import javax.imageio.stream.ImageInputStream;
import java.io.IOException;
import java.util.Locale;


public final class HeifImageReaderSpi extends ImageReaderSpiBase {
    public HeifImageReaderSpi() {
        super(new HeifProviderInfo());
    }

    @Override
    public boolean canDecodeInput(Object source) throws IOException {
        return source instanceof ImageInputStream && canDecode((ImageInputStream) source);
    }

    private boolean canDecode(ImageInputStream input) throws IOException {
        // https://github.com/strukturag/libheif/blob/e64bb552f5d48fee5daf69c8c2fd59ec3eee0818/libheif/heif.cc#L102
        // https://devstreaming-cdn.apple.com/videos/wwdc/2017/513fzgbviu23l/513/513_high_efficiency_image_file_format.pdf?dl=1
        try {
            input.mark();
            for (int i = 0; i < 4; i++) {
                input.read();
            }
            if (input.read() == 'f' && input.read() == 't' && input.read() == 'y' && input.read() == 'p'
                    && input.read() == 'h' && input.read() == 'e' && input.read() == 'i' && input.read() == 'c') {
                return true;
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
