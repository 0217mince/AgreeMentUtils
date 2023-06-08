package org.example;

import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;

/**
 * @author 小凡
 */
public class PdfTest {
    public static void main(String[] args) {
        //读取Word文档
        XWPFDocument document = null;
        try {
            document = new XWPFDocument(new FileInputStream("C:\\Users\\小凡\\Desktop\\家医签约协议.docx"));
//            //创建PDF选项
//            PdfOptions options = PdfOptions.create();
//
//            //创建PDF输出流
//            FileOutputStream out = new FileOutputStream(new File("C:\\Users\\小凡\\Desktop\\kzf.pdf"));
//
//            //将Word文档转换为PDF
//            PdfConverter.getInstance().convert(document, out, options);
//
//            //关闭输出流
//            out.close();

//            //检查输出的PDF文件是否有效
//            PDDocument pdfDoc = PDDocument.load(new File("output.pdf"));
//            for (PDPage page : pdfDoc.getPages()) {
//                //处理每个页面
//            }
//            pdfDoc.close();
            // 创建PDF选项
//            PdfOptions options = PdfOptions.create();
//            options.fontEncoding("UTF-8"); // 设置字体编码
//
//            // 创建PDF输出流
//            FileOutputStream out = new FileOutputStream(new File("C:\\Users\\小凡\\Desktop\\kzf.pdf"));
//
//            // 将Word文档转换为PDF
//            PdfConverter.getInstance().convert(document, out,options);
//
//            // 关闭输出流
//            out.close();
//
//            // 检查输出的PDF文件是否有效
//            PDDocument pdfDoc = PDDocument.load(new File("output.pdf"));
//            int pageCount = pdfDoc.getNumberOfPages();
//            System.out.println("PDF文件总页数：" + pageCount);
//            pdfDoc.close();

            PdfOptions options = PdfOptions.create();
            OutputStream out = new FileOutputStream("C:\\Users\\小凡\\Desktop\\kzf.pdf");
            PdfConverter.getInstance().convert(document, out, options);
            out.close();

        } catch (IOException e) {
            throw new RuntimeException(e);
        }


    }
}
