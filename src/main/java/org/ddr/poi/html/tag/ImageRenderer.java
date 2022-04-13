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

package org.ddr.poi.html.tag;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.common.usermodel.PictureType;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.SVGPictureData;
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
    private static final Map<String, ImageType> PICTURE_TYPES = new HashMap<>(ImageType.values().length);

    static {
        for (ImageType type : ImageType.values()) {
            PICTURE_TYPES.put(type.getExtension(), type);
        }

        SVGPictureData.initRelation();
    }

    enum ImageType {
        EMF(Document.PICTURE_TYPE_EMF),
        WMF(Document.PICTURE_TYPE_WMF),
        PICT(Document.PICTURE_TYPE_PICT),
        JPEG(Document.PICTURE_TYPE_JPEG),
        JPG(Document.PICTURE_TYPE_JPEG),
        PNG(Document.PICTURE_TYPE_PNG),
        DIB(Document.PICTURE_TYPE_DIB),
        GIF(Document.PICTURE_TYPE_GIF),
        TIF(Document.PICTURE_TYPE_TIFF),
        TIFF(Document.PICTURE_TYPE_TIFF),
        EPS(Document.PICTURE_TYPE_EPS),
        BMP(Document.PICTURE_TYPE_BMP),
        WPG(Document.PICTURE_TYPE_WPG);

        private final int type;

        ImageType(int type) {
            this.type = type;
        }

        public String getExtension() {
            return name().toLowerCase();
        }

        public int getType() {
            return type;
        }
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

            int type = PICTURE_TYPES.getOrDefault(format, typeOf(image)).getType();
            boolean svg = HtmlConstants.TAG_SVG.equals(format);
            addPicture(element, context, inputStream, type, image.getWidth(), image.getHeight(), svg ? bytes : null);
        } catch (IOException | InvalidFormatException e) {
            log.warn("Failed to load image: {}", src, e);
        } finally {
            // 释放资源
            image = null;
        }
    }

    /**
     * 根据图片反推类型
     *
     * @param image 图片
     * @return 图片类型
     */
    protected ImageType typeOf(BufferedImage image) {
        return image.getColorModel().hasAlpha() ? ImageType.PNG : ImageType.JPG;
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
        ImageType type = PICTURE_TYPES.get(extension);

        ByteArrayCopyStream outputStream = null;
        InputStream inputStream = null;
        HttpURLConnection connect = null;
        BufferedImage image;
        try {
            connect = HttpURLConnectionUtils.connect(src);
            InputStream urlStream = connect.getInputStream();
            boolean svg = connect.getHeaderField("content-type").contains(PictureType.SVG.contentType);
            byte[] svgData = null;
            if (svg) {
                outputStream = new ByteArrayCopyStream(urlStream.available());
                IOUtils.copy(urlStream, outputStream);
                svgData = outputStream.toByteArray();
                image = ImageIO.read(outputStream.toInput());
            } else {
                image = ImageIO.read(urlStream);
            }

            if (image == null) {
                log.warn("Illegal image url: {}", src);
                return;
            }
            int size = image.getData().getDataBuffer().getSize();
            outputStream = new ByteArrayCopyStream(size);

            if (type == null) {
                type = typeOf(image);
            }

            ImageIO.write(image, type.getExtension(), outputStream);
            inputStream = outputStream.toInput();
            addPicture(element, context, inputStream, type.getType(), image.getWidth(), image.getHeight(), svgData);
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
     * @param svgData SVG数据
     */
    protected void addPicture(Element element, HtmlRenderContext context, InputStream inputStream, int type,
                              int widthInPixels, int heightInPixels,
                              byte[] svgData) throws InvalidFormatException, IOException {
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
                widthInEMU, heightInEMU, svgData);
    }

}
