//package com.test.word;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.List;

import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.*;
import org.junit.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

/**
 * 创建word文档
 */
public class WordCreate {

    /**
     * 2007word文档创建
     */

    private String[] testDesc = new String[10];
    private String[][] part1 = new String[4][11];

    public WordCreate(String describe[]){
        this.testDesc = describe;
//        testDesc = describe;
    }

//    @Test
    public void createWord2007() {
        XWPFDocument doc = new XWPFDocument();
        //add a header
        XWPFHeader header = doc.createHeader(HeaderFooterType.DEFAULT);
        XWPFParagraph hp1 = header.createParagraph();
        XWPFRun hp1r1 = hp1.createRun();
        hp1.setAlignment(ParagraphAlignment.CENTER);
        hp1r1.setText("南京诺禾");
        hp1r1.setFontSize(12);
        hp1r1.setFontFamily("宋体");

        //add title of report

        XWPFParagraph title = doc.createParagraph();
        title.setAlignment(ParagraphAlignment.CENTER);
        title.setVerticalAlignment(TextAlignment.CENTER);

        // 第一页要使用title所定义的属性
        XWPFRun r1 = title.createRun();

        // 设置字体是否加粗
        r1.setBold(false);
        r1.setFontSize(15);

        // 设置使用何种字体
        r1.setFontFamily("宋体");

        // 设置上下两行之间的间距
        r1.setTextPosition(20);
        r1.setText("氧指数试验结果记录单");


        XWPFParagraph title2 = doc.createParagraph();
        XWPFRun tr2 = title2.createRun();
        tr2.setText("");
        tr2.setFontSize(15);
        tr2.setFontFamily("Calibri");

        XWPFParagraph title3 = doc.createParagraph();
        title3.setAlignment(ParagraphAlignment.CENTER);
        title3.setVerticalAlignment(TextAlignment.CENTER);
        XWPFRun tr3 = title3.createRun();
        tr3.setText("按GB/T 2406.2测定的氧指数试验结果记录单");
        tr3.setFontSize(10);
        tr3.setFontFamily("Calibri");

        XWPFParagraph title4 = doc.createParagraph();
        XWPFRun tr4 = title4.createRun();
        tr4.setText("");
        tr4.setFontSize(10);
        tr4.setFontFamily("Calibri");

        // describe this test
        XWPFParagraph describe = doc.createParagraph();
        describe.setAlignment(ParagraphAlignment.LEFT);
        describe.setVerticalAlignment(TextAlignment.CENTER);
        XWPFRun describeRun = describe.createRun();

        String[] titles = {"材料：", "试样类别：", "点燃方法：", "状态调节方法：", "氧浓度增量(d)：",
                "氧指数[浓度,%(体积分数)]：", "σ：", "试验日期：", "实验室 No.：", "实验 No.："};

        for(int i = 0; i < 10; i++){
            describeRun.setText(titles[i] + this.testDesc[i] + '\n');
            describeRun.setFontSize(10);
            describeRun.setFontFamily("宋体");
        }
        XWPFRun describeRun2 = describe.createRun();
        describeRun.setText("\n" + "第1部分：氧浓度间隔≤1%(体积分数)的一对\"X\"和\"O\"反应的氧浓度测定(按8.5)" + '\n');
        describeRun.setFontSize(10);
        describeRun.setFontFamily("宋体");

        // create table1 for test step 1
        XWPFTable table1 = doc.createTable();
        CTTblWidth table1Width = table1.getCTTbl().addNewTblPr().addNewTblW();
        table1Width.setType(STTblWidth.PCT);
//        int[] colWidthArr = new int[] {2592, 648, 648, 648, 648, 648, 648, 648, 648, 648, 648};


        table1Width.setW(BigInteger.valueOf(9072));

//        table1.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(6000));
//        table1.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(2000));

        XWPFTableRow tableRow1 = table1.getRow(0);
        tableRow1.getCell(0).setText("氧浓度(体积分数)/%");

//        tableRow1.getCell(0).set


        for(int i = 0; i < 10; i ++){
            tableRow1.addNewTableCell().setText(part1[0][i]);
        }

        XWPFTableRow rows[] = new XWPFTableRow[3];

        String[] table1Title = {"燃烧时间/s", "燃烧长度/mm", "反应(\"X\")或\"O\""};
        for(int i = 0; i < 3; i ++){
            rows[i] = table1.createRow();
            rows[i].getCell(0).setText(table1Title[i]);
            for(int j = 0; j < 10; j ++){
                rows[i].getCell(j + 1).setText(part1[i][j]);
            }
        }

        // 设置字体对齐方式

        FileOutputStream out;
        try {
            out = new FileOutputStream("word2007.docx");
            // 以下代码可进行文件下载
            // response.reset();
            // response.setContentType("application/x-msdownloadoctet-stream;charset=utf-8");
            // response.setHeader("Content-Disposition",
            // "attachment;filename=\"" + URLEncoder.encode(fileName, "UTF-8"));
            // OutputStream out = response.getOutputStream();
            // this.doc.write(out);
            // out.flush();

            doc.write(out);
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("success");

    }




}