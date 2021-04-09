package org.ddr.poi.html.tag;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Document;
import org.ddr.poi.html.ElementRenderer;
import org.ddr.poi.html.HtmlConstants;
import org.ddr.poi.html.HtmlRenderContext;
import org.ddr.poi.util.ByteArrayCopyStream;
import org.ddr.poi.util.HttpURLConnectionUtils;
import org.jsoup.nodes.Element;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.util.Base64;
import java.util.HashMap;
import java.util.Map;

/**
 * img标签渲染器
 *
 * @author Draco
 * @since 2021-02-09
 */
public class ImageRenderer implements ElementRenderer {
    private static final Logger log = LoggerFactory.getLogger(ImageRenderer.class);

    private static final String[] TAGS = {HtmlConstants.TAG_IMG};
    private static final String HTTP = "http";
    private static final String DOUBLE_SLASH = "//";
    private static final String BASE64_PREFIX = "data:";
    private static final Map<String, Integer> PICTURE_TYPES = new HashMap<>(12);

    static {
        PICTURE_TYPES.put("emf", Document.PICTURE_TYPE_EMF);
        PICTURE_TYPES.put("wmf", Document.PICTURE_TYPE_WMF);
        PICTURE_TYPES.put("pict", Document.PICTURE_TYPE_PICT);
        PICTURE_TYPES.put("jpeg", Document.PICTURE_TYPE_JPEG);
        PICTURE_TYPES.put("jpg", Document.PICTURE_TYPE_JPEG);
        PICTURE_TYPES.put("png", Document.PICTURE_TYPE_PNG);
        PICTURE_TYPES.put("dib", Document.PICTURE_TYPE_DIB);
        PICTURE_TYPES.put("gif", Document.PICTURE_TYPE_GIF);
        PICTURE_TYPES.put("tiff", Document.PICTURE_TYPE_TIFF);
        PICTURE_TYPES.put("eps", Document.PICTURE_TYPE_EPS);
        PICTURE_TYPES.put("bmp", Document.PICTURE_TYPE_BMP);
        PICTURE_TYPES.put("wpg", Document.PICTURE_TYPE_WPG);
    }

    /**
     * 元素渲染结束需要执行的逻辑
     *
     * @param element HTML元素
     * @param context 渲染上下文
     */
    @Override
    public boolean renderStart(Element element, HtmlRenderContext context) {
        String src = element.attr(HtmlConstants.ATTR_SRC);
        if (StringUtils.startsWithIgnoreCase(src, HTTP)) {
            handleRemoteImage(element, context, src);
        } else if (StringUtils.startsWith(src, DOUBLE_SLASH)) {
            // 某些图片链接为了跟随网站协议而隐去了协议名称
            handleRemoteImage(element, context, HTTP + HtmlConstants.COLON + src);
        } else if (StringUtils.startsWith(src, BASE64_PREFIX)) {
            handleBase64(element, context, src);
        }
        return false;
    }

    /**
     * 处理base64图片
     *
     * @param element HTML元素
     * @param context 渲染上下文
     * @param src 图片base64数据
     */
    private void handleBase64(Element element, HtmlRenderContext context, String src) {
        int index = src.indexOf(HtmlConstants.COMMA.charAt(0));
        String data = src.substring(index + 1);
        String format = StringUtils.substringBetween(src.substring(0, index), HtmlConstants.SLASH, HtmlConstants.SEMICOLON);
        // org.apache.poi.sl.usermodel.PictureData.PictureType
        // FIXME 尚不支持svg
        if (format.contains(HtmlConstants.MINUS)) {
            format = StringUtils.substringAfterLast(format, HtmlConstants.MINUS);
        } else if (format.contains(HtmlConstants.PLUS)) {
            format = StringUtils.substringBefore(format, HtmlConstants.PLUS);
        }

        byte[] bytes;
        try {
            bytes = Base64.getDecoder().decode(data);
        } catch (Exception e) {
            log.warn("Failed to load image due to illegal base64 data: {}", src);
            return;
        }
        BufferedImage image;
        try (InputStream inputStream = new ByteArrayInputStream(bytes)) {
            image = ImageIO.read(inputStream);
            inputStream.reset();

            Integer type = PICTURE_TYPES.getOrDefault(format,
                    image.getColorModel().hasAlpha() ? Document.PICTURE_TYPE_PNG : Document.PICTURE_TYPE_JPEG);

            addPicture(element, context, inputStream, type, image.getWidth(), image.getHeight());
        } catch (IOException | InvalidFormatException e) {
            log.warn("Failed to load image: {}", src, e);
        } finally {
            // 释放资源
            image = null;
        }
    }

    /**
     * 处理远程图片
     *
     * @param element HTML元素
     * @param context 渲染上下文
     * @param src 图片链接地址
     */
    private void handleRemoteImage(Element element, HtmlRenderContext context, String src) {
        String extension = FilenameUtils.getExtension(StringUtils.substringBefore(src, HtmlConstants.QUESTION)).toLowerCase();
        Integer type = PICTURE_TYPES.get(extension);

        ByteArrayCopyStream outputStream = null;
        InputStream inputStream = null;
        HttpURLConnection connect = null;
        BufferedImage image;
        try {
            connect = HttpURLConnectionUtils.connect(src);
            image = ImageIO.read(connect.getInputStream());
            if (image == null) {
                log.warn("Illegal image url: {}", src);
                return;
            }
            int size = image.getData().getDataBuffer().getSize();
            outputStream = new ByteArrayCopyStream(size);

            if (type == null) {
                if (image.getColorModel().hasAlpha()) {
                    extension = "png";
                    type = Document.PICTURE_TYPE_PNG;
                } else {
                    extension = "jpeg";
                    type = Document.PICTURE_TYPE_JPEG;
                }
            }

            ImageIO.write(image, extension, outputStream);
            inputStream = outputStream.toInput();
            addPicture(element, context, inputStream, type, image.getWidth(), image.getHeight());
        } catch (IOException | InvalidFormatException e) {
            log.warn("Failed to load image: {}", src, e);
        } finally {
            IOUtils.close(connect);
            IOUtils.closeQuietly(outputStream);
            IOUtils.closeQuietly(inputStream);
            // 释放资源
            image = null;
        }
    }

    @Override
    public String[] supportedTags() {
        return TAGS;
    }

    @Override
    public boolean renderAsBlock() {
        return false;
    }

    /**
     * 添加图片
     *
     * @param element HTML元素
     * @param context 渲染上下文
     * @param inputStream 图片数据流
     * @param type 图片类型
     * @param widthInPixels 图片宽度（像素）
     * @param heightInPixels 图片高度（像素）
     */
    private void addPicture(Element element, HtmlRenderContext context, InputStream inputStream, int type,
                            int widthInPixels, int heightInPixels) throws InvalidFormatException, IOException {
        // 容器限制宽度
        int containerWidth = context.getAvailableWidthInEMU();
//        int containerHeight = context.getAvailablePageHeight();
        // 图片原始宽高
        int widthInEMU = Units.pixelToEMU(widthInPixels);
        int heightInEMU = Units.pixelToEMU(heightInPixels);

        boolean declaredWidth = false;
        boolean declaredHeight = false;

        String width = context.getPropertyValue(HtmlConstants.CSS_WIDTH);
        if (width.length() > 0) {
            declaredWidth = true;
        } else {
            // width attribute is overridden by style, the same to height
            // https://css-tricks.com/whats-the-difference-between-width-height-in-css-and-width-height-html-attributes/
            width = element.attr(HtmlConstants.ATTR_WIDTH);
            if (NumberUtils.isParsable(width)) {
                width += HtmlConstants.PX;
                declaredWidth = true;
            }
        }

        String maxWidth = context.getPropertyValue(HtmlConstants.CSS_MAX_WIDTH);
        widthInEMU = context.computeLengthInEMU(width, maxWidth, widthInEMU, containerWidth);

        String height = context.getPropertyValue(HtmlConstants.CSS_HEIGHT);
        if (height.length() > 0) {
            declaredHeight = true;
        } else {
            height = element.attr(HtmlConstants.ATTR_HEIGHT);
            if (NumberUtils.isParsable(height)) {
                height += HtmlConstants.PX;
                declaredHeight = true;
            }
        }

        String maxHeight = context.getPropertyValue(HtmlConstants.CSS_MAX_HEIGHT);
        heightInEMU = context.computeLengthInEMU(height, maxHeight, heightInEMU, Integer.MAX_VALUE);

        // 如果只声明了宽或高，则同比计算对应尺寸
        if (declaredWidth && !declaredHeight) {
            heightInEMU = heightInPixels * widthInEMU / widthInPixels;
        } else if (!declaredWidth && declaredHeight) {
            widthInEMU = widthInPixels * heightInEMU / heightInPixels;
            if (widthInEMU > containerWidth) {
                widthInEMU = containerWidth;
                heightInEMU = heightInPixels * widthInEMU / widthInPixels;
            }
        }

        context.renderPicture(inputStream, type, HtmlConstants.TAG_IMG,
                widthInEMU, heightInEMU);
    }

}
