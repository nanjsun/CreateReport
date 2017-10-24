//package com.test.word;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.List;

import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.*;
import org.junit.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import java.util.ArrayList;

/**
 * 创建word文档
 */
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
        addPart1Table(doc);
        addParagraph2(doc);

        addNewPage(doc, BreakType.PAGE);
        addPart2Table(doc);

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
    public void addHeader(XWPFDocument doc){
        XWPFHeader header = doc.createHeader(HeaderFooterType.DEFAULT);
        XWPFParagraph hp1 = header.createParagraph();
        XWPFRun hp1r1 = hp1.createRun();
        hp1.setAlignment(ParagraphAlignment.CENTER);
        hp1r1.setText("南京诺禾");
        hp1r1.setFontSize(12);
        hp1r1.setFontFamily("宋体");
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
    }

    public void addPart1Table(XWPFDocument doc){

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
    }
    public void addParagraph2(XWPFDocument doc){
        XWPFParagraph p2 = doc.createParagraph();
        XWPFRun p2r = p2.createRun();
        p2r.setBold(false);
        p2r.setFontSize(10);
        p2r.setFontFamily("宋体");
        p2r.setText("\n此反应中\"O\"反应的氧浓度=17.0%(体积分数)(该浓度将再次用于第二部分首次测量的浓度)。\n" + "\n" +
                "第2部分：氧指数的测定(按8.6)\n" +
                "连续改变氧浓度的步长d=0.2%(体积分数)[除非另有说明，首选0.2%(体积分数)]。\n");

    }

    public void addPart2Table(XWPFDocument doc){

//        XWPFTable table2 = doc.createTable();

        List<String> columnList = new ArrayList<String>();
        columnList.add("序号");
        columnList.add("姓名信息gi|姓甚|名谁");
        columnList.add("名刺信息|籍贯|营生");
        XWPFTable table = doc.createTable(2,5);

        CTTbl ttbl = table.getCTTbl();
        CTTblPr tblPr = ttbl.getTblPr() == null ? ttbl.addNewTblPr() : ttbl.getTblPr();
        CTTblWidth tblWidth = tblPr.isSetTblW() ? tblPr.getTblW() : tblPr.addNewTblW();
        CTJc cTJc=tblPr.addNewJc();
        cTJc.setVal(STJc.Enum.forString("center"));
        tblWidth.setW(new BigInteger("8000"));
        tblWidth.setType(STTblWidth.DXA);

        XWPFTableRow firstRow=null;
        XWPFTableRow secondRow=null;
        XWPFTableCell firstCell=null;
        XWPFTableCell secondCell=null;

        for(int i=0;i<2;i++){
            firstRow=table.getRow(i);
            firstRow.setHeight(380);
            for(int j=0;j<5;j++){
                firstCell=firstRow.getCell(j);
                setCellText(firstCell, "测试", "FFFFC9", 1600);
            }
        }

        firstRow=table.insertNewTableRow(0);
        secondRow=table.insertNewTableRow(1);
        firstRow.setHeight(380);
        secondRow.setHeight(380);
        for(String str:columnList){
            if(str.indexOf("|") == -1){
                firstCell=firstRow.addNewTableCell();
                secondCell=secondRow.addNewTableCell();
                createVSpanCell(firstCell, str,"CCCCCC",1600,STMerge.RESTART);
                createVSpanCell(secondCell, "", "CCCCCC", 1600,null);
            }else{
                String[] strArr=str.split("\\|");
                firstCell=firstRow.addNewTableCell();
                createHSpanCell(firstCell, strArr[0],"CCCCCC",1600,STMerge.RESTART);
                for(int i=1;i<strArr.length-1;i++){
                    firstCell=firstRow.addNewTableCell();
                    createHSpanCell(firstCell, "","CCCCCC",1600,null);
                }
                for(int i=1;i<strArr.length;i++){
                    secondCell=secondRow.addNewTableCell();
                    setCellText(secondCell, strArr[i], "CCCCCC", 1600);
                }
            }
        }
    }

    public  void setCellText(XWPFTableCell cell,String text, String bgcolor, int width) {
        CTTc cttc = cell.getCTTc();
        CTTcPr cellPr = cttc.addNewTcPr();
        cellPr.addNewTcW().setW(BigInteger.valueOf(width));
        //cell.setColor(bgcolor);
        CTTcPr ctPr = cttc.addNewTcPr();
        CTShd ctshd = ctPr.addNewShd();
        ctshd.setFill(bgcolor);
        ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
        cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);
        cell.setText(text);
    }
    public void createHSpanCell(XWPFTableCell cell,String value, String bgcolor, int width,STMerge.Enum stMerge){
        CTTc cttc = cell.getCTTc();
        CTTcPr cellPr = cttc.addNewTcPr();
        cellPr.addNewTcW().setW(BigInteger.valueOf(width));
        cell.setColor(bgcolor);
        cellPr.addNewHMerge().setVal(stMerge);
        cellPr.addNewVAlign().setVal(STVerticalJc.CENTER);
        cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);
        cttc.getPList().get(0).addNewR().addNewT().setStringValue(value);
    }

    public void createVSpanCell(XWPFTableCell cell,String value, String bgcolor, int width,STMerge.Enum stMerge){
        CTTc cttc = cell.getCTTc();
        CTTcPr cellPr = cttc.addNewTcPr();
        cellPr.addNewTcW().setW(BigInteger.valueOf(width));
        cell.setColor(bgcolor);
        cellPr.addNewVMerge().setVal(stMerge);
        cellPr.addNewVAlign().setVal(STVerticalJc.CENTER);
        cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);
        cttc.getPList().get(0).addNewR().addNewT().setStringValue(value);
    }

    public void addNewPage(XWPFDocument document,BreakType breakType){
        XWPFParagraph xp = document.createParagraph();
        xp.createRun().addBreak(breakType);
    }

    public void saveDocument(XWPFDocument document,String savePath) throws Exception{
        FileOutputStream fos = new FileOutputStream(savePath);
        document.write(fos);
        fos.close();
    }

}








