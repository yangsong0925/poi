package word;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileOutputStream;
import java.io.File;
/**
 * Created by ys on 2018/4/8.
 */
public class Word3_CreateParagraph {

    public static void main(String[] args) throws Exception {
        //Blank(空白) Document
        XWPFDocument document= new XWPFDocument();
        //Write the Document in file system
        FileOutputStream out = new FileOutputStream(new File("createparagraph.docx"));

        //create Paragraph（Paragraph 段落）
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("At w3ii.com, we strive hard to " +
                        "provide quality tutorials for self-learning " +
                        "purpose in the domains of Academics, Information " +
                        "Technology, Management and Computer ProgrammingLanguages.");
        document.write(out);
        out.close();
        System.out.println("createparagraph.docx written successfully");
    }
}
