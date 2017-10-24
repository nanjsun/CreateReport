import java.io.File;
import java.io.FileOutputStream;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class CreateDocument
{
    public static void main(String[] args)throws Exception
    {
        //Blank Document
        XWPFDocument document= new XWPFDocument();
        //Write the Document in file system
        FileOutputStream out = new FileOutputStream(
                new File("createdocument.docx"));
        document.write(out);
        out.close();
        System.out.println(
                "createdocument.docx written successully");

        String[] testDesc = {"PMMA", "IV(3mm厚)","顶面点燃法", "23 23/50", "0.2%(体积分数)", "17.3", "0.151", "20130725", "东南大学", "02"};
        WordCreate word1 = new WordCreate(testDesc);
        word1.createWord2007();
    }
}

//材料：PMMA
//        试样类别：IV(3mm厚)
//        点燃方法：顶面点燃法
//        状态调节方法：23 23/50
//        氧浓度增量(d)：0.2%(体积分数)
//        氧指数[浓度,%(体积分数)]：17.3
//        σ：0.151
//        试验日期：20130725
//        实验室 No.：东南大学
//        实验 No.：02