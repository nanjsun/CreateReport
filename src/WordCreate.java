//package com.test.word;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.List;

import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.*;
import org.junit.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import javax.swing.text.TabExpander;
import java.util.ArrayList;

public class WordCreate {

    private String[] testDesc = new String[10];
    private String[][] part1 = new String[4][11];

    public WordCreate(String describe[]){
        this.testDesc = describe;
    }

//    @Test
    public void createWord2007() {
        XWPFDocument doc = new XWPFDocument();
        //add a header
        addHeader(doc);
        addTestDetail(doc);
        addText(doc, "第1部分：氧浓度间隔≤1%(体积分数)的一对\"X\"和\"O\"反应的氧浓度测定(按8.5)");


        addPart1Table(doc);
        addText(doc, "\n此反应中的\"O\"反应的氧浓度=18.0%(体积分数)(该浓度将再次用于第二部分首次测量的浓度)。\n");

        addText(doc, "第2部分：氧指数的测定(按8.6)");


        addText(doc, "连续改变氧浓度的步长d=0.2%(体积分数)[除非另有说明，首选0.2%(体积分数)]。");

        addPart2Table(doc);

        addText(doc, "\n第三部分：氧浓度步长%d的校验\n");

        addPart3Table(doc);

        addText(doc, "\nd=0.2\tσ=sqrt{sum[(Ci-OI)^2]/(n-1)}=0.110\t2/3*σ=0.073333  3/2*σ=0.165000\n");
        addText(doc, "校验结果：合适");

        addText(doc, "使用仪器：\n" +
                "1、NH-OI-01型智能氧指数测定仪\t\t仪器编号：\n" +
                "2、1000mm钢直尺 \t\t精度：1mm\t\t仪器编号：\n" +
                "3、秒表\t\t精度：±0.25s\t\t\t仪器编号：\n" +
                "\n" +
                "检验环境：℃，%RH\t\t\t检验日期：2015917\t\t\t测试员：\n");
//        addPart2Table(doc);

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
    public void addHeader(XWPFDocument doc){
        XWPFParagraph title = doc.createParagraph();
        title.setAlignment(ParagraphAlignment.CENTER);
        title.setVerticalAlignment(TextAlignment.CENTER);
        XWPFRun r1 = title.createRun();
        r1.setBold(false);
        r1.setFontSize(20);
        r1.setFontFamily("宋体");

        r1.setText("氧指数试验原始记录\n");
        r1.addBreak();

        XWPFParagraph titlePage = doc.createParagraph();
        XWPFRun r2 = titlePage.createRun();
        titlePage.setAlignment(ParagraphAlignment.RIGHT);
        titlePage.setVerticalAlignment(TextAlignment.CENTER);
        r2.setFontSize(10);
        r2.setFontFamily("宋体");
        r2.setText("共      页 第      页");
        r2.addBreak();
    }

    public void addTestDetail(XWPFDocument doc){

        XWPFTable table = doc.createTable();
        table.getCTTbl().getTblPr().unsetTblBorders();

        CTTblWidth tableWidth = table.getCTTbl().addNewTblPr().addNewTblW();
        tableWidth.setType(STTblWidth.DXA);
        tableWidth.setW(BigInteger.valueOf(9072));

        XWPFTableRow tableRowOne = table.getRow(0);
        tableRowOne.getCell(0).setText("样品编号：");
        tableRowOne.addNewTableCell().setText("检验标准：GB/T 2406.2-2009");

        XWPFTableRow tableRowTwo = table.createRow();
        tableRowTwo.getCell(0).setText("材料：");
        tableRowTwo.getCell(1).setText("试样类别：IV(mm厚)");

        XWPFTableRow tableRowThree = table.createRow();
        tableRowThree.getCell(0).setText("点燃方法：顶部点燃法");
        tableRowThree.getCell(1).setText("氧指数[浓度,%(体积分数)]：18.1%");

        XWPFTableRow tableRowFour = table.createRow();
        tableRowFour.getCell(0).setText("氧浓度增量(d)：0.2%(体积分数)");
    }

    public void addPart1Table(XWPFDocument doc){

        XWPFTable table = doc.createTable();
        CTTblWidth table1Width = table.getCTTbl().addNewTblPr().addNewTblW();
        table1Width.setType(STTblWidth.PCT);
                table1Width.setW(BigInteger.valueOf(9072));

        XWPFTableRow tableRow1 = table.getRow(0);
        tableRow1.getCell(0).setText("氧浓度(体积分数)/%");

        for(int i = 0; i < 10; i ++){
            tableRow1.addNewTableCell().setText(part1[0][i]);
        }

        XWPFTableRow rows[] = new XWPFTableRow[3];

        String[] table1Title = {"燃烧时间/s", "燃烧长度/mm", "反应(\"X\")或\"O\""};
        for(int i = 0; i < 3; i ++){
            rows[i] = table.createRow();
            rows[i].getCell(0).setText(table1Title[i]);
            for(int j = 0; j < 10; j ++){
                rows[i].getCell(j + 1).setText(part1[i][j]);
            }
        }
    }


    public void addPart2Table(XWPFDocument doc){
        XWPFTable table = doc.createTable();
        CTTblWidth tableWidth = table.getCTTbl().addNewTblPr().addNewTblW();
        tableWidth.setType(STTblWidth.DXA);
        tableWidth.setW(BigInteger.valueOf(9072));

        table.setWidth(5*1440);
        XWPFTableRow rowOne = table.getRow(0);
        rowOne.addNewTableCell().setText("Nt系列测量");

        XWPFTableRow rowTwo = table.createRow();
        rowTwo.getCell(0).setText("");
        rowTwo.getCell(1).setText("Nt系列测量(8.6.1和8.6.2)");
        rowTwo.createCell().setText("8.6.3");
        rowTwo.createCell().setText("cf");

        XWPFTableRow rowThree = table.createRow();
        rowThree.getCell(0).setText("氧浓度(体积分数)/%");
        rowThree.getCell(1).setText("18.0");
        rowThree.createCell().setText("18.2");
        rowThree.createCell().setText("dd");
        rowThree.createCell().setText("sss");
        rowThree.createCell().setText("ddd");

        XWPFTableCell cell = rowOne.getCell(0);
    }

    public void addPart3Table(XWPFDocument doc){


        XWPFTable table = doc.createTable(5, 7);
        CTTblWidth tableWidth = table.getCTTbl().addNewTblPr().addNewTblW();
        tableWidth.setType(STTblWidth.DXA);
        tableWidth.setW(BigInteger.valueOf(9072));
    }

    public void addNewPage(XWPFDocument document,BreakType breakType){
        XWPFParagraph xp = document.createParagraph();
        xp.createRun().addBreak(breakType);
    }
    public void addText(XWPFDocument doc, String text){
        XWPFParagraph pa = doc.createParagraph();
        XWPFRun run = pa.createRun();
        run.setText(text);
        run.setFontSize(10);
        run.setFontFamily("宋体");

    }


    public void saveDocument(XWPFDocument document,String savePath) throws Exception{
        FileOutputStream fos = new FileOutputStream(savePath);
        document.write(fos);
        fos.close();
    }

}








