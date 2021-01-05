package controller;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class ReadInputDocument {

	public static void main(String[] args) {
		FileInputStream fs;
		FileOutputStream os;
		try {
			fs = new FileInputStream("C:\\Users\\Brenton\\Desktop\\SF-86 Continuation.doc");
			os = new FileOutputStream("C:\\Users\\Brenton\\Desktop\\Output.docx");
			XWPFDocument outDoc = new XWPFDocument();
			XWPFParagraph outDocParagraph;
			XWPFRun outDocRun;
			
			HWPFDocument doc = new HWPFDocument(fs);
			WordExtractor we = new WordExtractor(doc);
			Range range = doc.getRange();
			String[] paragraphs = we.getParagraphText();
			for (int i = 0; i < paragraphs.length; i++) {
				org.apache.poi.hwpf.usermodel.Paragraph pr = range.getParagraph(i);

				System.out.println(pr.getEndOffset());
				int j = 0;
				while (true) {
					CharacterRun run = pr.getCharacterRun(j++);
					System.out.println("-------------------------------");
					System.out.println(run.text());
					System.out.println("Color---" + run.getColor());
					System.out.println("getFontName---" + run.getFontName());
					System.out.println("getFontSize---" + run.getFontSize());
					System.out.println("highlight----" + run.getHighlightedColor());
					
					/*
					 * if(run.isHighlighted()) { outDocParagraph = outDoc.createParagraph();
					 * outDocRun = outDocParagraph.createRun();
					 * outDocRun.setTextHighlightColor(run.getHighlightedColor());
					 * outDocRun.setItalic(run.isItalic()); outDocRun.setStyle(run.getStyleIndex());
					 * outDocRun.setFontFamily(run.getFontFamily());
					 * outDocRun.setColor(run.getColor()); }
					 */

					if (run.getEndOffset() == pr.getEndOffset()) {
						break;
					}
				}

			}
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		

	}
}