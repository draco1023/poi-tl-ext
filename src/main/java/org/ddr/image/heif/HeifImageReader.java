package org.ddr.image.heif;

import com.drew.imaging.FileType;
import com.drew.imaging.ImageMetadataReader;
import com.drew.imaging.ImageProcessingException;
import com.drew.metadata.Metadata;
import com.twelvemonkeys.imageio.ImageReaderBase;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.ddr.image.ImageInputStreamWrapper;
import org.ddr.image.MetadataReader;
import org.ddr.poi.util.HttpURLConnectionUtils;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Element;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import javax.imageio.ImageReadParam;
import javax.imageio.ImageTypeSpecifier;
import javax.imageio.spi.ImageReaderSpi;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.nio.charset.StandardCharsets;
import java.util.Iterator;

public class HeifImageReader extends ImageReaderBase {
    private static final Logger log = LoggerFactory.getLogger(HeifImageReader.class);
    protected final MetadataReader metadataReader;

    protected ImageInputStreamWrapper wrapper;
    protected Metadata metadata;
    protected Dimension dimension;
    protected String format;

    public HeifImageReader(ImageReaderSpi provider) {
        this(provider, new HeifMetadataReader(), "heic");
    }

    protected HeifImageReader(ImageReaderSpi provider, MetadataReader metadataReader, String format) {
        super(provider);
        this.metadataReader = metadataReader;
        this.format = format;
    }

    @Override
    protected void resetMembers() {
        metadata = null;
        dimension = null;
    }

    @Override
    public void setInput(Object input, boolean seekForwardOnly, boolean ignoreMetadata) {
        super.setInput(input, seekForwardOnly, ignoreMetadata);

        if (imageInput != null) {
            wrapper = new ImageInputStreamWrapper(imageInput);
            try {
                metadata = ImageMetadataReader.readMetadata(wrapper, 0, FileType.Heif);
                dimension = metadataReader.getDimension(metadata);
            } catch (IOException | ImageProcessingException e) {
                log.warn("Failed to read metadata", e);
            }
        }
    }

    @Override
    public int getWidth(int imageIndex) throws IOException {
        return dimension == null ? 0 : dimension.width;
    }

    @Override
    public int getHeight(int imageIndex) throws IOException {
        return dimension == null ? 0 : dimension.height;
    }

    @Override
    public Iterator<ImageTypeSpecifier> getImageTypes(int imageIndex) throws IOException {
        return null;
    }

    @Override
    public BufferedImage read(int imageIndex, ImageReadParam param) throws IOException {
        return convert(param);
    }

    BufferedImage convert(ImageReadParam param) {
        HttpURLConnection uploadConnection = null;
        HttpURLConnection convertConnection = null;
        HttpURLConnection downloadConnection = null;

        try {
            uploadConnection = HttpURLConnectionUtils.connect("https://ezgif.com/" + format + "-to-jpg");
            uploadConnection.setInstanceFollowRedirects(false);
            HttpURLConnectionUtils.initUserAgent(uploadConnection);
            uploadConnection.setRequestProperty("Referer", "https://ezgif.com/" + format + "-to-jpg");
            String boundary = HttpURLConnectionUtils.initFormData(uploadConnection);

            try (OutputStream outputStream = uploadConnection.getOutputStream()) {
                byte[] boundaryBytes = ("--" + boundary).getBytes();
                wrapper.seek(0);
                HttpURLConnectionUtils.addFormData(outputStream, boundaryBytes, "new-image", "some." + format, wrapper);

                outputStream.write(boundaryBytes);
                outputStream.write("--".getBytes());
                outputStream.write(HttpURLConnectionUtils.newLineBytes);
                outputStream.flush();
            }

            // 获取上传响应
            int uploadResponseCode = uploadConnection.getResponseCode();
            if (uploadResponseCode == HttpURLConnection.HTTP_MOVED_TEMP) {
                String location = uploadConnection.getHeaderField("Location");
                String convertUrl = StringUtils.substringBeforeLast(location, ".");
                String fileId = StringUtils.substringAfterLast(convertUrl, "/");
                convertUrl += "?ajax=true";

                if (log.isDebugEnabled()) {
                    log.debug("{} uploaded: {}", format, fileId);
                }
                convertConnection = HttpURLConnectionUtils.connect(convertUrl);
                HttpURLConnectionUtils.initUserAgent(convertConnection);
                convertConnection.setRequestProperty("Referer", location);
                boundary = HttpURLConnectionUtils.initFormData(convertConnection);
                try (OutputStream convertOutput = convertConnection.getOutputStream()) {
                    byte[] boundaryBytes = ("--" + boundary).getBytes();

                    HttpURLConnectionUtils.addFormData(convertOutput, boundaryBytes, "file", fileId, null);
                    HttpURLConnectionUtils.addFormData(convertOutput, boundaryBytes, "percentage", "90", null);
                    HttpURLConnectionUtils.addFormData(convertOutput, boundaryBytes, "percentager", "90", null);
                    HttpURLConnectionUtils.addFormData(convertOutput, boundaryBytes, "background", "#ffffff", null);
                    HttpURLConnectionUtils.addFormData(convertOutput, boundaryBytes, "backgroundc", "#ffffff", null);
                    HttpURLConnectionUtils.addFormData(convertOutput, boundaryBytes, "ajax", "true", null);

                    convertOutput.write(boundaryBytes);
                    convertOutput.write("--".getBytes());
                    convertOutput.write(HttpURLConnectionUtils.newLineBytes);
                    convertOutput.flush();
                }
                int convertResponseCode = convertConnection.getResponseCode();
                if (convertResponseCode == HttpURLConnection.HTTP_OK) {
                    try (InputStream convertResponse = convertConnection.getInputStream()) {
                        Element body = Jsoup.parse(convertResponse, StandardCharsets.UTF_8.name(), "").body();
                        if (log.isDebugEnabled()) {
                            log.debug("{} converted: {}", format, body.html());
                        }
                        for (Element img : body.select("img")) {
                            String src = img.attr("src");
                            if (StringUtils.contains(src, "ezgif")) {
                                String url = "https:" + src;
                                downloadConnection = HttpURLConnectionUtils.connect(url);
                                HttpURLConnectionUtils.initUserAgent(downloadConnection);
                                try (InputStream downloadResponse = downloadConnection.getInputStream()) {
                                    return ImageIO.read(downloadResponse);
                                }
                            }
                        }
                    }
                } else {
                    log.warn("Failed to convert {} image. Response code: {}", format, convertResponseCode);
                }
            } else {
                log.warn("Failed to upload image. Response code: {}", uploadResponseCode);
            }

        } catch (Exception e) {
            log.warn("Failed to convert {} image", format, e);
            IOUtils.close(uploadConnection);
            IOUtils.close(convertConnection);
            IOUtils.close(downloadConnection);
        }

        return null;
    }
}
