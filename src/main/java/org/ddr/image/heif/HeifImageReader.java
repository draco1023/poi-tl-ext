package org.ddr.image.heif;

import com.twelvemonkeys.imageio.ImageReaderBase;
import org.apache.commons.io.IOUtils;
import org.ddr.poi.util.HttpURLConnectionUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import javax.imageio.ImageReadParam;
import javax.imageio.ImageTypeSpecifier;
import javax.imageio.spi.ImageReaderSpi;
import javax.imageio.stream.ImageInputStream;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.util.Iterator;

public class HeifImageReader extends ImageReaderBase {
    private static final Logger log = LoggerFactory.getLogger(HeifImageReader.class);

    public HeifImageReader(ImageReaderSpi provider) {
        super(provider);
    }

    @Override
    protected void resetMembers() {

    }

    @Override
    public int getWidth(int imageIndex) throws IOException {
        // TODO read 'ispe' 'irot'
        return 0;
    }

    @Override
    public int getHeight(int imageIndex) throws IOException {
        // TODO read 'ispe' 'irot'
        return 0;
    }

    @Override
    public Iterator<ImageTypeSpecifier> getImageTypes(int imageIndex) throws IOException {
        return null;
    }

    @Override
    public BufferedImage read(int imageIndex, ImageReadParam param) throws IOException {
        Object input = getInput();
        if (input instanceof ImageInputStream) {
            return convert((ImageInputStream) input, param);
        }
        return null;
    }

    BufferedImage convert(ImageInputStream input, ImageReadParam param) throws IOException {
        HttpURLConnection uploadConnection = null;
        HttpURLConnection convertConnection = null;

        HeicOnlineParam heicOnlineParam = param instanceof HeicOnlineParam ? (HeicOnlineParam) param : new HeicOnlineParam();

        try {
            uploadConnection = HttpURLConnectionUtils.connect("https://s1.heic.online/heic/");
            uploadConnection.setRequestMethod("POST");
            uploadConnection.setDoOutput(true);
            uploadConnection.setRequestProperty("X-File-Name", "public.heic");
            uploadConnection.setRequestProperty("Content-Type", "application/octet-stream");
            uploadConnection.setRequestProperty("Origin", "https://heic.online");
            uploadConnection.setRequestProperty("Referer", "https://heic.online/");

            try (OutputStream outputStream = uploadConnection.getOutputStream()) {
                byte[] buffer = new byte[8192];
                int n;
                while (-1 != (n = input.read(buffer))) {
                    outputStream.write(buffer, 0, n);
                }

                outputStream.flush();
            }

            // 获取上传响应
            int uploadResponseCode = uploadConnection.getResponseCode();
            String fileId;
            try (InputStream uploadResponse = uploadConnection.getInputStream()) {
                fileId = IOUtils.toString(uploadResponse, StandardCharsets.UTF_8).trim();
            }
            if (uploadResponseCode == HttpURLConnection.HTTP_OK) {
                if (log.isDebugEnabled()) {
                    log.debug("Heic uploaded: {}", fileId);
                }
                convertConnection = HttpURLConnectionUtils.connect("https://s1.heic.online/heic/");
                convertConnection.setRequestMethod("POST");
                convertConnection.setDoOutput(true);
                convertConnection.setRequestProperty("Content-Type", "application/x-www-form-urlencoded; charset=UTF-8");
                convertConnection.setRequestProperty("Referer", "https://heic.online/");
                String data = heicOnlineParam.withId(fileId);
                convertConnection.setRequestProperty("Content-Length", String.valueOf(data.length()));
                try (OutputStream convertOutput = convertConnection.getOutputStream()) {
                    IOUtils.write(data, convertOutput, StandardCharsets.UTF_8);
                }
                int convertResponseCode = convertConnection.getResponseCode();
                String json;
                try (InputStream convertResponse = convertConnection.getInputStream()) {
                    json = IOUtils.toString(convertResponse, StandardCharsets.UTF_8);
                }
                if (convertResponseCode == HttpURLConnection.HTTP_OK && json.contains("SUCCESS")) {
                    if (log.isDebugEnabled()) {
                        log.debug("Heic converted: {}", json);
                    }
                    String url = "https://s1.heic.online/upload/" + fileId + "/";
                    return ImageIO.read(new URL(url));
                } else {
                    log.warn("Failed to convert heic image. Response code: {}, {}", convertResponseCode, json);
                }
            } else {
                log.warn("Failed to upload image. Response code: {}, {}", uploadResponseCode, fileId);
            }

        } catch (Exception e) {
            log.warn("Failed to convert heic image", e);
            IOUtils.close(uploadConnection);
            IOUtils.close(convertConnection);
        }

        return null;
    }
}
