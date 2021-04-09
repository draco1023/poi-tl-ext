package org.ddr.poi.html;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

class HtmlRenderPolicyTest {

    @Test
    void doRender() throws IOException {
        HtmlRenderPolicy htmlRenderPolicy = new HtmlRenderPolicy();
        Configure configure = Configure.builder()
                .bind("teachContent", htmlRenderPolicy)
                .bind("plainContent", htmlRenderPolicy)
                .build();
        Map<String, Object> data = new HashMap<>(2);
        data.put("teachContent", "<p style=\"font: italic small-caps bold 16px/2 cursive, sans-serif;\"><span style=\"font-size: 36px\">你好<strong>世界<i> 你<a href=\"http://www.baidu.com\">敢信</a></i></strong>啊！</span></p><p><span style=\"font-size: 24px\">你好<a href=\"http://www.baidu.com\"><b>世界<i> 你敢信</i></b>啊！</span><br/><img src=\"https://profile.csdnimg.cn/1/2/6/1_myhope88\" width=\"32\"/></a><img src=\"https://g.csdnimg.cn/static/user-reg-year/2x/15.png\" style=\"width:50%\"/><img src=\"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAQAAABIkb+zAAABk0lEQVR4Ae3asVHcUBSF4V8jJcoghRhyVpSgFiA1dgAlQAlQApSgVQfQATTAxruZV8qUPEbXDXg8b7zPko7n/qeCL7wzF8/zPM9Lmv15/xHAAQ5wgAMcUNGyxw7YTxoumKkffGEJFvjGDFV8YYkWuGDyWizh1kzePimgY/Is8XJ1QOEABzjgH+aAV23ASKUNaEAZEDjTBjyDMmDgRBvwCMqAjiNtwD0oA7aU2oBbUAZ8UmgDrkAZ8E6mDahBGfAKyoCRShvQgDIgcKYNeAZlwMCJNuARlAEdR8RULxVwT0wZH8sEbCmJ6RpbJuCWmAo2ywR8UhDTHbZMwBUxleyWCXgnI6YHbJmAmpiO6VXugd/3hCkDThm0AS+YMuCcoA1YY8qAS0ZtwBumDKgxZUDGhzbgGlMGFGy0AXeYMqBkpw14wJQBx/TagCdMGXDKoA14wZQB5wRtwBpTBlwyagPe/OlvWoADcu33+57Ja5ICWiZvRcASLbBihr4nIgRumKmKlh47YB0tK2Yup/jL5cyb53me9wv95P/HjXVvlAAAAABJRU5ErkJggg==\"/></p>");
//        data.put("plainContent", "<table cellspacing=\"0\" class=\"Table\" style=\"border-collapse:collapse; border:none; font-family:&quot;Times New Roman&quot;; font-size:13px; margin-left:-9px; width:566px\">\n\t<tbody>\n\t\t<tr>\n\t\t\t<td colspan=\"3\" style=\"border-bottom:2px solid black; border-left:1px solid black; border-right:1px solid black; border-top:1px solid black; vertical-align:center; width:484px\">\n\t\t\t<p align=\"center\" style=\"text-align:center\"><span style=\"font-size:10.5pt\"><span style=\"font-family:等线\"><strong><span style=\"font-size:10.5000pt\"><span style=\"font-family:宋体\"><strong>（导）学案</strong></span></span></strong></span></span></p>\n\t\t\t</td>\n\t\t\t<td style=\"border-bottom:2px solid black; border-left:none; border-right:1px solid black; border-top:1px solid black; vertical-align:center; width:81px\">\n\t\t\t<p align=\"center\" style=\"text-align:center\"><span style=\"font-size:10.5pt\"><span style=\"font-family:等线\"><strong><span style=\"font-size:10.5000pt\"><span style=\"font-family:宋体\"><strong>备 注</strong></span></span></strong></span></span></p>\n\t\t\t</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:1px solid black; border-right:1px solid black; border-top:none; vertical-align:center; width:85px\">\n\t\t\t<p style=\"text-align:left\"><span style=\"font-size:10.5pt\"><span style=\"font-family:等线\"><strong><span style=\"font-size:10.5000pt\"><span style=\"font-family:宋体\"><strong>学习目标</strong></span></span></strong></span></span></p>\n\t\t\t</td>\n\t\t\t<td colspan=\"2\" style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:center; width:399px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:2px solid black; vertical-align:center; width:81px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:1px solid black; border-right:1px solid black; border-top:none; vertical-align:center; width:85px\">\n\t\t\t<p style=\"text-align:left\"><span style=\"font-size:10.5pt\"><span style=\"font-family:等线\"><strong><span style=\"font-size:10.5000pt\"><span style=\"font-family:宋体\"><strong>学习重点</strong></span></span></strong></span></span></p>\n\t\t\t</td>\n\t\t\t<td colspan=\"2\" style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:center; width:399px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:center; width:81px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:1px solid black; border-right:1px solid black; border-top:none; vertical-align:center; width:85px\">\n\t\t\t<p style=\"text-align:left\"><span style=\"font-size:10.5pt\"><span style=\"font-family:等线\"><strong><span style=\"font-size:10.5000pt\"><span style=\"font-family:宋体\"><strong>知识链接</strong></span></span></strong></span></span></p>\n\t\t\t</td>\n\t\t\t<td colspan=\"2\" style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:center; width:399px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:center; width:81px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:1px solid black; border-right:1px solid black; border-top:none; vertical-align:center; width:85px\">\n\t\t\t<p style=\"text-align:left\"><span style=\"font-size:10.5pt\"><span style=\"font-family:等线\"><strong><span style=\"font-size:10.5000pt\"><span style=\"font-family:宋体\"><strong>学法指导</strong></span></span></strong></span></span></p>\n\t\t\t</td>\n\t\t\t<td colspan=\"2\" style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:center; width:399px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:center; width:81px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:1px solid black; border-right:1px solid black; border-top:2px solid black; vertical-align:center; width:85px\">\n\t\t\t<p style=\"text-align:left\"><span style=\"font-size:10.5pt\"><span style=\"font-family:等线\"><strong><span style=\"font-size:10.5000pt\"><span style=\"font-family:宋体\"><strong>学习</strong></span></span></strong><strong><span style=\"font-size:10.5000pt\"><span style=\"font-family:宋体\"><strong>过程</strong></span></span></strong></span></span></p>\n\t\t\t</td>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:2px solid black; vertical-align:center; width:207px\">\n\t\t\t<p style=\"text-align:left\"><span style=\"font-size:10.5pt\"><span style=\"font-family:等线\"><span style=\"font-size:10.5000pt\"><span style=\"font-family:宋体\">问题与任务</span></span></span></span></p>\n\t\t\t</td>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:2px solid black; vertical-align:center; width:191px\">\n\t\t\t<p style=\"text-align:left\"><span style=\"font-size:10.5pt\"><span style=\"font-family:等线\"><span style=\"font-size:10.5000pt\"><span style=\"font-family:宋体\">方法与要求</span></span></span></span></p>\n\t\t\t</td>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:2px solid black; vertical-align:center; width:81px\">\n\t\t\t<p style=\"text-align:left\"><span style=\"font-size:10.5pt\"><span style=\"font-family:等线\"><span style=\"font-size:10.5000pt\"><span style=\"font-family:宋体\">-</span></span></span></span></p>\n\t\t\t</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:1px solid black; border-right:1px solid black; border-top:none; vertical-align:center; width:85px\">\n\t\t\t<p style=\"margin-right:1px; text-align:left\">&nbsp;</p>\n\n\t\t\t<p style=\"margin-right:1px; text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:center; width:207px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:center; width:191px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:top; width:81px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:1px solid black; border-right:1px solid black; border-top:none; vertical-align:center; width:85px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:center; width:207px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:center; width:191px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:top; width:81px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td style=\"border-bottom:2px solid black; border-left:1px solid black; border-right:1px solid black; border-top:none; vertical-align:center; width:85px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t\t<td style=\"border-bottom:2px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:center; width:207px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t\t<td style=\"border-bottom:2px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:center; width:191px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:top; width:81px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:1px solid black; border-right:1px solid black; border-top:none; vertical-align:center; width:85px\">\n\t\t\t<p style=\"text-align:left\"><span style=\"font-size:10.5pt\"><span style=\"font-family:等线\"><strong><span style=\"font-size:10.5000pt\"><span style=\"font-family:宋体\"><strong>达标检测</strong></span></span></strong></span></span></p>\n\t\t\t</td>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:center; width:207px\">\n\t\t\t<p style=\"text-align:left\"><span style=\"font-size:10.5pt\"><span style=\"font-family:等线\"><span style=\"font-size:10.5000pt\"><span style=\"font-family:宋体\">课内</span></span><span style=\"font-size:10.5000pt\"><span style=\"font-family:宋体\">作业</span></span></span></span></p>\n\t\t\t</td>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:center; width:191px\">\n\t\t\t<p style=\"text-align:left\"><span style=\"font-size:10.5pt\"><span style=\"font-family:等线\"><span style=\"font-size:10.5000pt\"><span style=\"font-family:宋体\">课外作业</span></span></span></span></p>\n\t\t\t</td>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:2px solid black; vertical-align:center; width:81px\">\n\t\t\t<p style=\"text-align:left\"><span style=\"font-size:10.5pt\"><span style=\"font-family:等线\"><span style=\"font-size:10.5000pt\"><span style=\"font-family:宋体\">-</span></span></span></span></p>\n\t\t\t</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:1px solid black; border-right:1px solid black; border-top:none; vertical-align:center; width:85px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:center; width:207px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:center; width:191px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:center; width:81px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:1px solid black; border-right:1px solid black; border-top:none; vertical-align:center; width:85px\">\n\t\t\t<p style=\"text-align:left\"><span style=\"font-size:10.5pt\"><span style=\"font-family:等线\"><strong><span style=\"font-size:10.5000pt\"><span style=\"font-family:宋体\"><strong>总结提升</strong></span></span></strong></span></span></p>\n\t\t\t</td>\n\t\t\t<td colspan=\"2\" style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:center; width:399px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:center; width:81px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t</tr>\n\t\t<tr>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:1px solid black; border-right:1px solid black; border-top:none; vertical-align:center; width:85px\">\n\t\t\t<p style=\"text-align:left\"><span style=\"font-size:10.5pt\"><span style=\"font-family:等线\"><strong><span style=\"font-size:10.5000pt\"><span style=\"font-family:宋体\"><strong>课后反思</strong></span></span></strong></span></span></p>\n\t\t\t</td>\n\t\t\t<td colspan=\"2\" style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:center; width:399px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t\t<td style=\"border-bottom:1px solid black; border-left:none; border-right:1px solid black; border-top:none; vertical-align:center; width:81px\">\n\t\t\t<p style=\"text-align:left\">&nbsp;</p>\n\t\t\t</td>\n\t\t</tr>\n\t</tbody>\n</table>\n<p>test</p>");
        data.put("plainContent", "<p><span style=\"font-size: 36px\">你好<strong>世界<i> 你敢信</i></strong>啊！</span></p><p><span style=\"font-size: 24px\">你好<b>世界<i> 你敢信</i></b>啊！</span><br/><img src=\"https://profile.csdnimg.cn/1/2/6/1_myhope88\"/><img src=\"https://g.csdnimg.cn/static/user-reg-year/2x/15.png\"/><img src=\"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAQAAABIkb+zAAABk0lEQVR4Ae3asVHcUBSF4V8jJcoghRhyVpSgFiA1dgAlQAlQApSgVQfQATTAxruZV8qUPEbXDXg8b7zPko7n/qeCL7wzF8/zPM9Lmv15/xHAAQ5wgAMcUNGyxw7YTxoumKkffGEJFvjGDFV8YYkWuGDyWizh1kzePimgY/Is8XJ1QOEABzjgH+aAV23ASKUNaEAZEDjTBjyDMmDgRBvwCMqAjiNtwD0oA7aU2oBbUAZ8UmgDrkAZ8E6mDahBGfAKyoCRShvQgDIgcKYNeAZlwMCJNuARlAEdR8RULxVwT0wZH8sEbCmJ6RpbJuCWmAo2ywR8UhDTHbZMwBUxleyWCXgnI6YHbJmAmpiO6VXugd/3hCkDThm0AS+YMuCcoA1YY8qAS0ZtwBumDKgxZUDGhzbgGlMGFGy0AXeYMqBkpw14wJQBx/TagCdMGXDKoA14wZQB5wRtwBpTBlwyagPe/OlvWoADcu33+57Ja5ICWiZvRcASLbBihr4nIgRumKmKlh47YB0tK2Yup/jL5cyb53me9wv95P/HjXVvlAAAAABJRU5ErkJggg==\"/></p><p><math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msqrt><mn>4</mn></msqrt></math>&nbsp;<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mfrac bevelled=\"true\"><mn>1</mn><mn>4</mn></mfrac></math></p>\n" +
                "\n" +
                "<p><u><em><strong><span id=\"cke_bm_1103S\" style=\"display:none\">&nbsp;</span>&nbsp;&nbsp;王府井围殴皮肤较为</strong></em></u></p>\n" +
                "\n" +
                "<h1><big><u><em><strong>威风威风文&nbsp; &nbsp;<big><u><em><strong><big><u><em><strong><big><u><em><strong><img alt=\"cool\" bizcode=\"\" fileid=\"\" height=\"23\" src=\"https://editor_baktest.bakclass.com/plugins/smiley/images/shades_smile.png\" title=\"cool\" width=\"23\" /></strong></em></u></big></strong></em></u></big></strong></em></u></big><a href=\"http://www.baidu.com\">http://www.baidu.com</a></strong></em></u></big></h1>\n" +
                "\n" +
                "<ol>\n" +
                "\t<li style=\"color: red\"><u><em><strong>违法未</strong></em></u></li>\n" +
                "\t<li><u><em><strong>违法未</strong></em></u></li>\n" +
                "</ol>\n" +
                "\n" +
                "<p style=\"color: rgb(0,255,0)\">今天（21日）0时，位于新疆南部的<span style=\"color: hsla(30, 100%, 50%, .3);\">塔里木油田</span>年油气产量达到3003.12万吨，成为我国油气上产重要战略接替区。</p>\n" +
                "\n" +
                "<p style=\"color: hsl(30, 100%, 50%);\">　　位于大北博孜气区的大北902井近日喷出高产气流，今年这里试油的6口井均获40万立方米以上高产。目前，天山南麓还有60多部钻机正在钻进，近两年，塔里木油田超过100口气井获高产，新建天然气产能90亿立方米。</p>\n" +
                "\n" +
                "<p>　　<img alt=\"\" bizcode=\"\" fileid=\"\" src=\"https://nimg.ws.126.net/\" /><br />\n" +
                "&nbsp;</p>\n" +
                "\n" +
                "<p style=\"color: rgba(0,255,0,0.6)\">　　中石油塔里木油田公司总经理 杨学文：&ldquo;十三五&rdquo;期间，塔里木油田大力提升勘探开发力度，已经落实两个万亿方大气区，形成10亿吨级大油气区，为新一轮的油气增储上产奠定了坚实的资源基础。</p>\n" +
                "\n" +
                "<p>　　56万平方公里的塔里木盆地是我国陆上最大的含油气盆地，盆地遍布大漠戈壁，油气大多蕴藏在超过6000米的地宫深处，面临超深、超高温、超高压等极限考验，是世界级勘探禁区。油田的科技工作者攻克了一系列世界级难题，探明油气储量26.7亿吨。</p>\n" +
                "\n" +
                "<hr />\n" +
                "<hr />\n" +
                "<hr />\n" +
                "<ul>\n" +
                "\t<li><u><em><strong>威风威风</strong></em></u></li>\n" +
                "\t<li><u><em><strong>违法未</strong></em></u><ol><li>测试嵌套123213</li><li>测试嵌套2222</li></ol></li>\n" +
                "\t<li><u><em><strong>违法未</strong></em></u><ol><li>测试嵌套123213</li><li>测试嵌套2222</li></ol></li>\n" +
                "</ul>\n" +
                "\n" +
                "<p>&nbsp;</p>\n" +
                "\n" +
                "<p><img alt=\"\" bizcode=\"\" fileid=\"\" height=\"431.8407960199005\" src=\"https://test-max-public.oss-cn-hangzhou.aliyuncs.com/max/site/2020-12-29/Bj21JqjsT6WgIu3BO4egbA.png\" width=\"800\" /></p>\n");

        try (InputStream inputStream = HtmlRenderPolicyTest.class.getResourceAsStream("/notes.docx")) {
            XWPFTemplate.compile(inputStream, configure).render(data).writeToFile("poi.docx");
        }
    }
}