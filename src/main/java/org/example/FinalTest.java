package org.example;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;

/**
 * @author 小凡
 */
public class FinalTest {

    public static void main(String[] args) {
        try {
            // 加载Word文档
            XWPFDocument document = new XWPFDocument(Files.newInputStream(Paths.get("C:\\Users\\小凡\\Desktop\\家医签约协议.docx")));

            AgreementUtils.replace(document,"doctorPhone","1313818371");

            AgreementUtils.replace(document,"mobile","33671938913");

            AgreementUtils.replaces(document,"patient","柯治凡");

            AgreementUtils.replace(document,"organ","浙大邵逸夫");

            AgreementUtils.replace(document,"team","柯小帅");

            AgreementUtils.replace(document,"协议开始年","2023");
            AgreementUtils.replace(document,"协议开始月","6");
            AgreementUtils.replace(document,"协议开始日","2");

            AgreementUtils.replace(document,"协议结束年","2024");
            AgreementUtils.replace(document,"协议结束月","6");
            AgreementUtils.replace(document,"协议结束日","2");

            AgreementUtils.replaces(document,"year","2023");
            AgreementUtils.replaces(document,"month","6");
            AgreementUtils.replaces(document,"day","2");

            AgreementUtils.replace(document,"doctor","柯大帅");

            AgreementUtils.replace(document,"idCard","331022200002173435");

            AgreementUtils.setPicture(document,Files.newInputStream(Paths.get("C:\\Users\\小凡\\Desktop\\a.jpg")),"医生签字");

            AgreementUtils.setPicture(document,Files.newInputStream(Paths.get("C:\\Users\\小凡\\Desktop\\b.jpg")),"居民签字");

            // 保存Word文档
            FileOutputStream out = new FileOutputStream("C:\\Users\\小凡\\Desktop\\signed_contract.docx");
            document.write(out);
            out.close();

            System.out.println("合同签名成功，已保存至 signed_contract.docx");
        } catch (Exception ex) {
            System.out.println("合同签名失败：" + ex.getMessage());
        }
    }
}
