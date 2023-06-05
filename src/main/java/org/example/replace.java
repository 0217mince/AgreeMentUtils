package org.example;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

public class replace {
    public static void main(String[] args) {
        try {
            FileInputStream fis = new FileInputStream("C:\\Users\\小凡\\Desktop\\家医签约协议.docx");
            XWPFDocument document = new XWPFDocument(fis);

            // 替换文本字段
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                List<XWPFRun> runs = paragraph.getRuns();
                for (int i = 0; i < runs.size(); i++) {
                    XWPFRun run = runs.get(i);
                    String text = run.getText(0);
                    if (text != null && text.contains("patient")) {
                        text = text.replace("patient", "newText");
                        run.setText(text, 0);
                    }
                }
            }

            // 保存替换后的文档
            FileOutputStream fos = new FileOutputStream("C:\\Users\\小凡\\Desktop\\signed_contract.docx");
            document.write(fos);
            fos.close();
            fis.close();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
}
