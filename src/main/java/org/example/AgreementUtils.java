package org.example;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.outline.PDOutlineItem;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

/**
 * @author 小凡
 * @date 2023/6/2
 */
public class AgreementUtils {

    public static final String[] CN_NUMBERS = {"零", "一", "二", "三", "四", "五", "六", "七", "八", "九"};
    public static final String[] CN_UNITS = {"", "十", "百", "千", "万"};

    public static String numberToChinese(int number) {
        StringBuilder sb = new StringBuilder();
        int unit = 0;
        while (number > 0) {
            int digit = number % 10;
            if (digit == 0) {
                if (sb.length() > 0 && sb.charAt(0) != '零') {
                    sb.insert(0, "零");
                }
            } else {
                sb.insert(0, CN_NUMBERS[digit] + CN_UNITS[unit]);
            }
            number /= 10;
            unit++;
        }
        if (sb.length() == 0) {
            sb.append("零");
        }
        return sb.toString();
    }

    /**
     * 替换word文本中的字段，只替换第一个
     *
     * @param document    word文档
     * @param target      替换目标
     * @param replacement 最终替换对象
     */
    public static void replace(XWPFDocument document, String target, String replacement) {
        outerLoop:
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            List<XWPFRun> runs = paragraph.getRuns();
            for (XWPFRun run : runs) {
                String text = run.getText(0);
                if (text != null && text.contains(target)) {
                    text = text.replace(target, replacement);
                    run.setText(text, 0);
                    break outerLoop;
                }
            }
        }
    }

    /**
     * 替换word文本中的字段，替换全文
     *
     * @param document    word文档
     * @param target      替换目标
     * @param replacement 最终替换对象
     */
    public static void replaces(XWPFDocument document, String target, String replacement) {
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            List<XWPFRun> runs = paragraph.getRuns();
            for (XWPFRun run : runs) {
                String text = run.getText(0);
                if (text != null && text.contains(target)) {
                    text = text.replace(target, replacement);
                    run.setText(text, 0);
                }
            }
        }
    }

    /**
     * 在Word文档中查找目标区域所在的段落
     *
     * @param document   word文档
     * @param targetArea 目标区域
     * @return 目标区域所在的段落序号
     */
    public static int findParagraphIndex(XWPFDocument document, String targetArea) {
        int paragraphIndex = -1;
        for (int i = 0; i < document.getParagraphs().size(); i++) {
            XWPFParagraph paragraph = document.getParagraphs().get(i);
            String text = paragraph.getText();

            if (text.contains(targetArea)) {
                paragraphIndex = i;
                break;
            }
        }
        return paragraphIndex;
    }

    /**
     * 在目标区域下面放入图片
     *
     * @param document    word文档
     * @param pictureData 图片数据流
     * @param targetArea  目标区域
     * @throws IOException            输入输出异常
     * @throws InvalidFormatException 读取写入word异常
     */
    public static void setPicture(XWPFDocument document, InputStream pictureData, String targetArea) throws IOException, InvalidFormatException {
        int signParagraphIndex = findParagraphIndex(document, targetArea);
        if (signParagraphIndex >= 0) {
            XWPFParagraph signParagraph = document.getParagraphArray(signParagraphIndex);
            XWPFRun signRun = signParagraph.createRun();
            signRun.addBreak();
            signRun.addBreak();
            signRun.addBreak();
            signRun.addPicture(pictureData, XWPFDocument.PICTURE_TYPE_PNG, "signature", Units.toEMU(100), Units.toEMU(50));
        }
    }

    /**
     * 在目标区域后面写入文本
     *
     * @param document   word文档
     * @param text       需写入的文本
     * @param targetArea 目标区域
     */
    public static void setText(XWPFDocument document, String text, String targetArea) {
        int signParagraphIndex = findParagraphIndex(document, targetArea);
        if (signParagraphIndex >= 0) {
            XWPFRun run = document.getParagraphs().get(signParagraphIndex).createRun();
            run.setText(text);
        }
    }

//    public static void wordTurnToPdf(XWPFDocument document) {
//        // 创建一个 PDF 文档
//        PDDocument pdfDocument = new PDDocument();
//
//        // 将 Word 文档中的内容转换为 PDF
//        try {
//            // 遍历 Word 文档中的每个段落
//            for(XWPFParagraph paragraph: document.getParagraphs()) {
//                // 创建一个新的 PDF 段落
//                PDOutlineItem item = new PDOutlineItem();
//                item.setTitle(paragraph.getText());
////                pdfDocument.getDocumentCatalog().getDocumentOutline().addLast(item);
//                PDPage page = new PDPage();
//                pdfDocument.addPage(page);
//                PDPageContentStream contentStream = new PDPageContentStream(pdfDocument, page);
//
//                // 遍历段落中的每个 Run
//                for(XWPFRun run: paragraph.getRuns()) {
//                    // 如果 Run 中包含图片，则将其转换为 PDF 中的 Image
//                    if(run.getEmbeddedPictures() != null && run.getEmbeddedPictures().size() > 0) {
//                        for(XWPFPicture picture: run.getEmbeddedPictures()) {
//                            byte[] pictureData = picture.getPictureData().getData();
//                            PDImageXObject image = PDImageXObject.createFromByteArray(pdfDocument, pictureData, picture.getDescription());
//                            contentStream.drawImage(image, (Float) picture.getCTPicture().getSpPr().getXfrm().getOff().getX(), page.getMediaBox().getHeight() - (Float)picture.getCTPicture().getSpPr().getXfrm().getOff().getY() - image.getHeight());
//                        }
//                    }
//                    // 如果 Run 中不包含图片，则将其转换为 PDF 中的 Text
//                    else {
//                        contentStream.beginText();
//                        contentStream.setFont(PDType1Font.HELVETICA, run.getFontSizeAsDouble().floatValue());
//                        contentStream.newLineAtOffset(run.getFontSizeAsDouble().floatValue() * 0.25f, page.getMediaBox().getHeight() - paragraph.getSpacingAfter() - run.getFontSizeAsDouble().floatValue() * 1.25f);
//                        contentStream.showText(run.text());
//                        contentStream.endText();
//                    }
//                }
//
//                contentStream.close();
//            }
//
//            // 将 PDF 文档写入文件
//            FileOutputStream out = new FileOutputStream("C:\\Users\\小凡\\Desktop\\文档PDF.pdf");
//            pdfDocument.save(out);
//            out.close();
//            System.out.println("转换成功！");
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//    }
}
