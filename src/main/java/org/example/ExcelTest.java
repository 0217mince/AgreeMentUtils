package org.example;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHdrFtrRef;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.math.BigInteger;
import java.net.URL;
import java.net.URLConnection;
import java.util.List;

/**
 * @author 小凡
 */
public class ExcelTest {
    public static void main(String[] args) {
        // 打开一个文档
        XWPFDocument document = null;
        try {
//            FileInputStream fis = new FileInputStream("C:\\Users\\小凡\\Desktop\\家医签约协议.docx");
//            document = new XWPFDocument(fis);
            String url = "https://docs-import-export-1251316161.cos.ap-guangzhou.myqcloud.com/export/docx/UShQsdJsUwSv/462d8f36a297e4b542660e28e6510189.json.docx?X-Amz-Algorithm=AWS4-HMAC-SHA256&X-Amz-Credential=AKIDBAvQgh24SZPnxur0C9qfpkQp24pMCOu8%2F20230607%2Fap-guangzhou%2Fs3%2Faws4_request&X-Amz-Date=20230607T063556Z&X-Amz-Expires=1800&X-Amz-SignedHeaders=host&response-content-disposition=attachment%3Bfilename%3D%22.docx%22%3Bfilename%2A%3DUTF-8%27%27%25E5%25AE%25B6%25E5%258C%25BB%25E7%25AD%25BE%25E7%25BA%25A6%25E5%258D%258F%25E8%25AE%25AE.docx&X-Amz-Signature=40907978c253030b6d566e1fbf51c35a4c338b6400a5a6e13d99ff5e560984da";

            URL urlObject = new URL(url);
            URLConnection connection = urlObject.openConnection();
            InputStream is = connection.getInputStream();
            document = new XWPFDocument(is);
            is.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

        // 在文档中查找特定文字
        String searchText = "附表：";
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        XWPFParagraph targetParagraph = null;
        for(XWPFParagraph paragraph: paragraphs) {
            String text = paragraph.getText();
            if(text != null && text.contains(searchText)) {
                targetParagraph = paragraph;
                break;
            }
        }

        // 如果找到了目标段落，则在其下方添加表格
        if(targetParagraph != null) {
            // 在目标段落后插入一个换行符
            XWPFRun run = targetParagraph.createRun();
            run.addBreak();
            run.addBreak();
            run.setText("一、柯治凡");
            run.setFontFamily("宋体");
            run.setFontSize(10);

            // 在换行符后面创建一个新段落，并在其中添加一个表格
            XWPFTable table = document.createTable();
            table.setWidth(8844);

            // 添加表格标题行
            XWPFTableRow headerRow = table.getRow(0);
            headerRow.setHeight(400);


            XWPFParagraph paragraph = headerRow.getCell(0).getParagraphArray(0);
            // 设置段落的对齐方式为左对齐
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            // 设置段落的缩进为0
            paragraph.setIndentationLeft(-100);
            XWPFRun run1 = paragraph.createRun();
            run1.setText("柯治凡");
            run1.setFontFamily("宋体");
            run1.setFontSize(10);
            headerRow.getCell(0).setParagraph(paragraph);

            headerRow.addNewTableCell().setText("年龄");
            headerRow.addNewTableCell().setText("性别");
            headerRow.getCell(0).setWidth("2412");
            headerRow.getCell(1).setWidth("4824");
            headerRow.getCell(2).setWidth("1608");


            // 添加数据行
            XWPFTableRow dataRow1 = table.createRow();
            dataRow1.getCell(0).setText("张三");
            dataRow1.getCell(1).setText("25");
            dataRow1.getCell(2).setText("男");

            XWPFTableRow dataRow2 = table.createRow();
            dataRow2.getCell(0).setText("李四");
            dataRow2.getCell(1).setText("30");
            dataRow2.getCell(2).setText("女");
        }

        // 将文档写入文件
        try {
            FileOutputStream out = new FileOutputStream("C:\\Users\\小凡\\Desktop\\文档.docx");
            document.write(out);
            out.close();
            System.out.println("表格添加成功！");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
