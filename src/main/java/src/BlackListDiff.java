package src;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

public class BlackListDiff {
    public static void main(String[] args) throws IOException {
        InputStream blackListStream = ClassLoader.class.getResourceAsStream("/BL.xls");
        findAndWriteDiffsToSheet(blackListStream);
    }

    private static void findAndWriteDiffsToSheet(InputStream blackListStream) throws IOException {
        LinkedList<String> blackListFromColumnA = new LinkedList<>();
        LinkedList<String> blackListFromColumnB = new LinkedList<>();

        HSSFWorkbook workbook = new HSSFWorkbook(blackListStream);
        HSSFSheet sheet = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.iterator();

        iterateOverRows(rowIterator, blackListFromColumnA, blackListFromColumnB);

        blackListStream.close();

        //find differential between columns lists
        List<String> diffAB = new ArrayList<>(blackListFromColumnA);
        diffAB.removeAll(blackListFromColumnB);
        List<String> diffBA = new ArrayList<>(blackListFromColumnB);
        diffBA.removeAll(blackListFromColumnA);

        //edit previous sheet and save to other file
        for (int i = 0; i < diffAB.size(); i++) {
            sheet.getRow(i).createCell(4).setCellValue(diffAB.get(i));
        }
        for (int i = 0; i < diffBA.size(); i++) {
            sheet.getRow(i).createCell(5).setCellValue(diffBA.get(i));
        }
        FileOutputStream outFile = new FileOutputStream(new File("C:\\Users\\User\\Desktop\\update.xls"));
        workbook.write(outFile);
        outFile.close();

//        System.out.println(blackListFromColumnA.size() + "---FROM COLUMN A: " + blackListFromColumnA);
//        System.out.println(blackListFromColumnB.size() + "---FROM COLUMN B: " + blackListFromColumnB);
//        System.out.println(diffAB.size() + "---DIFF BETWEEN COLUMNS: " + diffAB);
//        System.out.println(diffBA.size() + "---DIFF BETWEEN COLUMNS B A: " + diffBA);
    }

    private static void iterateOverRows(Iterator<Row> rowIterator, LinkedList<String> blackListFromColumnA, LinkedList<String> blackListFromColumnB) {
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            iterateOverCells(blackListFromColumnA, blackListFromColumnB, row, cellIterator);
        }
    }

    private static void iterateOverCells(LinkedList<String> blackListFromColumnA, LinkedList<String> blackListFromColumnB, Row row, Iterator<Cell> cellIterator) {
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            String stringCellValue = cell.getStringCellValue();
            executeTasks(blackListFromColumnA, blackListFromColumnB, row, cell, stringCellValue);
        }
    }

    private static void executeTasks(LinkedList<String> blackListFromColumnA, LinkedList<String> blackListFromColumnB, Row row, Cell cell, String stringCellValue) {
        if (cell.getColumnIndex() == 0) {
            blackListFromColumnA.add(stringCellValue);
            row.createCell(2).setCellValue(stringCellValue);
        }else if (cell.getColumnIndex() == 1) {
            blackListFromColumnB.add(stringCellValue);
            row.createCell(3).setCellValue(stringCellValue);
        }
    }
}
