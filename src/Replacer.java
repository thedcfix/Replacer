import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class Replacer {
	
	final static String FILE_NAME_IN = "test.docx";
	final static String FILE_NAME_OUT = "output.docx";

	public static void main(String[] args) throws FileNotFoundException, IOException, InvalidFormatException {
		
		System.out.println("Program started...go!");
	    
	    XWPFDocument doc = new XWPFDocument(new FileInputStream(new File(FILE_NAME_IN)));
		
		// replacing token in free text
		for (XWPFParagraph p : doc.getParagraphs()) {
			System.out.println(p.getText());
		    List<XWPFRun> runs = p.getRuns();
		    if (runs != null) {
		        for (XWPFRun r : runs) {
		            String text = r.getText(0);
		            
		            
		            if (text != null && text.contains("--token--")) {
		                text = text.replace("--token--", "INSERT YOUR VALUE HERE");
		                r.setText(text, 0);
		            }
		        }
		    }
		}
		
		// replacing token in tables
		for (XWPFTable tbl : doc.getTables()) {
		   for (XWPFTableRow row : tbl.getRows()) {
		      for (XWPFTableCell cell : row.getTableCells()) {
		         for (XWPFParagraph p : cell.getParagraphs()) {
		            for (XWPFRun r : p.getRuns()) {
		              String text = r.getText(0);
		              if (text != null && text.contains("--token--")) {
		                text = text.replace("--token--", "INSERT YOUR VALUE HERE");
		                r.setText(text,0);
		              }
		            }
		         }
		      }
		   }
		}
		
		FileOutputStream out = new FileOutputStream(FILE_NAME_OUT); 
		doc.write(out);
		
		out.close();
		doc.close();
		
		System.out.println("End =)");

	}

}