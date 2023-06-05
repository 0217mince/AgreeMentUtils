package org.example;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;

/**
 * @author 小凡
 */
public class Main {
    public static void main(String[] args) {
        try {
            // 加载Word文档
            XWPFDocument document = new XWPFDocument(Files.newInputStream(Paths.get("C:\\Users\\小凡\\Desktop\\家医签约协议.docx")));

//            //姓名
//            setTextBySignPositionInEnd(document," 柯治凡 ","尊敬的居民朋友");
//            //签约机构
//            setTextBySignPositionInEnd(document," 浙大邵逸夫柯小帅","感谢您选择朝阳区");
//
//
//            FileOutputStream out = new FileOutputStream("C:\\Users\\小凡\\Desktop\\signed_contract.docx");
//            document.write(out);
//            out.close();
//
//            System.out.println("合同签名成功，已保存至 signed_contract.docx");

            // 在Word文档中定位签名位置
            String signText = "签名时间：2023年5月30日";
            // 在合同中查找签名区域所在的段落
            int signParagraphIndex = findSignParagraph(document, "patient");

            if (signParagraphIndex >= 0) {
                // 在签名位置插入签名图片和签名时间等文字
                XWPFParagraph signParagraph = document.getParagraphArray(signParagraphIndex);
                XWPFRun signRun = signParagraph.createRun();
                signRun.addBreak();
                signRun.addBreak();
                // 插入签名图片
                signRun.addPicture(Files.newInputStream(Paths.get("C:\\Users\\小凡\\Desktop\\a.jpg")), XWPFDocument.PICTURE_TYPE_PNG, "signature", Units.toEMU(100), Units.toEMU(50));
                // 换行
                signRun.addBreak();
                //插入签名时间等文字
                signRun.setText(signText);

                // 保存Word文档
                FileOutputStream out = new FileOutputStream("C:\\Users\\小凡\\Desktop\\signed_contract.docx");
                document.write(out);
                out.close();

                System.out.println("合同签名成功，已保存至 signed_contract.docx");
            } else {
                System.out.println("找不到签名区域");
            }
        } catch (Exception ex) {
            System.out.println("合同签名失败：" + ex.getMessage());
        }
    }

    /**
     * 在Word文档中查找签名区域所在的段落
     */
    private static int findSignParagraph(XWPFDocument document, String signArea) {
        int paragraphIndex = -1;

        for (int i = 0; i < document.getParagraphs().size(); i++) {
            XWPFParagraph paragraph = document.getParagraphs().get(i);
            String text = paragraph.getText();

            if (text.contains(signArea)) {
                paragraphIndex = i;
                break;
            }
        }

        return paragraphIndex;
    }

    private static void setTextBySignPositionInEnd(XWPFDocument document,String signText,String positioningWords) {
        int signParagraphIndex = findSignParagraph(document,positioningWords);
        if (signParagraphIndex >= 0) {
            XWPFRun run = document.getParagraphs().get(signParagraphIndex).insertNewRun(1);
            run.setText(signText);
            run.setUnderline(UnderlinePatterns.SINGLE);
            //字体，范围----效果不详
            run.setFontFamily("宋体");
            //字体大小
            run.setFontSize(12);
        }
    }

    private static void setTextBySignPositionInFirst(XWPFDocument document,String signText,String positioningWords) {
        int signParagraphIndex = findSignParagraph(document,positioningWords);
        if (signParagraphIndex >= 0) {
            XWPFRun run = document.getParagraphs().get(signParagraphIndex).insertNewRun(0);
            run.setText(signText);
            run.setUnderline(UnderlinePatterns.SINGLE);
            //字体，范围----效果不详
            run.setFontFamily("宋体");
            //字体大小
            run.setFontSize(12);
        }
    }

    private static void setTextBySignPosition(XWPFParagraph signParagraph,String signText,String positioningWords) {
        String text = signParagraph.getText();
        if (text.contains(positioningWords)) {
            int index = text.indexOf(positioningWords) + positioningWords.length();
            System.out.println(signParagraph.getRuns().size());
            XWPFRun run = signParagraph.insertNewRun(1);
            run.setText(signText);
            run.setUnderline(UnderlinePatterns.DASH_LONG);
            //字体，范围----效果不详
            run.setFontFamily("宋体");
            //字体大小 picture
            run.setFontSize(12);
        }
    }

}