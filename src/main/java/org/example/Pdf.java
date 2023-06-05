package org.example;

import java.io.*;
import java.util.List;

import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.pdfbox.pdmodel.*;

/**
 * @author 小凡
 */
public class Pdf {
    public static void main(String[] args) throws IOException {
        // 指定 Word 文件和 PDF 文件的路径
        String wordFile = "C:\\Users\\小凡\\Desktop\\家医签约协议.docx";
        String pdfFile = "C:\\Users\\小凡\\Desktop\\家医签约协议.pdf";

        // 打开 Word 文档
        FileInputStream fis = new FileInputStream(wordFile);
        XWPFDocument document = new XWPFDocument(fis);

        // 创建 PDF 文档
        PDDocument pdfDocument = new PDDocument();

        // 遍历 Word 文档的每一页，并将其转换为 PDF 格式
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for (XWPFParagraph para : paragraphs) {
            // 创建 PDF 页面
            PDPage page = new PDPage();
            pdfDocument.addPage(page);

            // 将 Word 段落内容写入 PDF 页面
            PDPageContentStream contentStream = new PDPageContentStream(pdfDocument, page);
            contentStream.beginText();
            contentStream.setFont(PDType1Font.HELVETICA, 12);
            contentStream.newLineAtOffset(50, 700);
            contentStream.showText(para.getText());
            contentStream.endText();
            contentStream.close();
        }

        // 将 PDF 文档保存到指定路径
        FileOutputStream fos = new FileOutputStream(pdfFile);
        pdfDocument.save(fos);

        // 关闭文件流
        fis.close();
        fos.close();
        pdfDocument.close();
    }

}
