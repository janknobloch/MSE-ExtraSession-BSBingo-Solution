package de.tum.jk.pptreader;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Map.Entry;
import java.util.stream.Collectors;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xslf.usermodel.DrawingParagraph;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class PowerpointReader {

	public static void main(String[] args) throws FileNotFoundException, IOException {

		XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("demo_presentation.pptx"));

		HashMap<String, Integer> wordCount = new HashMap<String, Integer>();
		createWordCountFromPPTX(ppt, wordCount);

		Map<String, Integer> result = sortByValue(wordCount);
		Workbook wb = createExcelSheetFromExtraction(result);

		writeExcelToFile(wb);
	}

	private static Map<String, Integer> sortByValue(HashMap<String, Integer> map) {

		Map<String, Integer> sortedNewMap = map.entrySet().stream()
				.sorted((e1, e2) -> e2.getValue().compareTo(e1.getValue()))
				.collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue, (e1, e2) -> e1, LinkedHashMap::new));
		return sortedNewMap;
	}

	private static void writeExcelToFile(Workbook wb) throws FileNotFoundException, IOException {
		OutputStream fileOut = new FileOutputStream("workbook.xls");
		wb.write(fileOut);
		wb.close();
	}

	private static void createWordCountFromPPTX(XMLSlideShow ppt, HashMap<String, Integer> wordCount) {
		for (XSLFSlide s : ppt.getSlides()) {

			// Interating over Text elements (drawingparagraphs)
			for (DrawingParagraph x : s.getCommonSlideData().getText()) {
				// create a lis of words for this slide by splitting whitespace " "
				String words[] = x.getText().toString().split(" ");

				// going over all words
				for (int i = 0; i < words.length; i++) { // WORD //# of occurences
					// check if word exist already in my hashmap <String, Integer>
					if (wordCount.containsKey(words[i])) {
						// if word is there we count up
						Integer count = wordCount.get(words[i]);
						count = count + 1;
						// and overwrite value
						wordCount.put(words[i], count);
					} else {
						// word not found lets add it with value #1
						wordCount.put(words[i], 1);
					}
				}
			}

		}
	}

	private static Workbook createExcelSheetFromExtraction(Map<String, Integer> sortedMapAsc) {
		Workbook wb = new HSSFWorkbook();
		Sheet s = wb.createSheet();
		int countRows = 0;
		for (Entry<String, Integer> entry : sortedMapAsc.entrySet()) {
			// only take value between 10 and 70 (min max occurence of words)
			if (entry.getValue() > 10 && entry.getValue() < 70) {

				Row r = s.createRow(countRows);
				r.createCell(0).setCellValue(entry.getKey().replaceAll("	", ""));
				r.createCell(1).setCellValue(entry.getValue());
				countRows++;

			}
		}
		return wb;
	}
}
