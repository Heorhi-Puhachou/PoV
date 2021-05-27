package by.voiting.pov;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

public class Parser {
    public static int FIRST_COLUMN_WIDTH = 2000;
    public static int SECOND_COLUMN_WIDTH = 13000;
    public static int COLOR_CELL_WIDTH = 525;
    public static int FIRST_COLUMN_WITH_COLOR = 2;

    public static void main(String... args) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Voitings");

        int rowNumber = 1;
        createHeader(sheet, workbook);

        for (Voiting voiting : Storage.voitings) {
            addVoitingRow(sheet, voiting, rowNumber++);
        }

        FileOutputStream outputStream = new FileOutputStream("/home/heorhi/excel2/voitings.xlsx");
        workbook.write(outputStream);
        workbook.close();
    }

    public static void createHeader(XSSFSheet sheet, XSSFWorkbook workbook) {
        Row headerRow = sheet.createRow(0);
        sheet.setColumnWidth(0, FIRST_COLUMN_WIDTH);
        sheet.setColumnWidth(1, SECOND_COLUMN_WIDTH);
        ArrayList<String> delegates = Storage.getDelegates();
        XSSFCellStyle myStyle = workbook.createCellStyle();
        myStyle.setRotation((short) 90);

        for (int i = 0; i < delegates.size(); i++) {
            Cell headerCell = headerRow.createCell(i + FIRST_COLUMN_WITH_COLOR);
            sheet.setColumnWidth(i + FIRST_COLUMN_WITH_COLOR, COLOR_CELL_WIDTH);

            headerCell.setCellValue((i + 1) + ". " + delegates.get(i));
            headerCell.setCellStyle(myStyle);
        }
    }

    public static void addVoitingRow(XSSFSheet sheet, Voiting voiting, int rowNumber) {
        Row row = sheet.createRow(rowNumber);
        Cell dateCell = row.createCell(0);
        dateCell.setCellValue(voiting.getDate());
        Cell nameCell = row.createCell(1);
        nameCell.setCellValue(voiting.getName());

        String result = voiting.getResult();

        String[] variants = result.split("\n");
        HashMap<String, CellStyle> nameColor = new HashMap<>();

        for (String variant : variants) {
            handleVariant(sheet, variant, nameColor);
        }

        ArrayList<String> delegates = Storage.getDelegates();

        for (int i = 0; i < delegates.size(); i++) {
            Cell colorCell = row.createCell(i + FIRST_COLUMN_WITH_COLOR);
            CellStyle style = nameColor.get(delegates.get(i));
            colorCell.setCellStyle(style);
        }


    }

    public static void handleVariant(XSSFSheet sheet, String variant, HashMap<String, CellStyle> nameColor) {
        CellStyle style;
        if (variant.startsWith("ГОЛОСА ЗА:")) {
            style = getGreenStyle(sheet);
        } else if (variant.startsWith("ГОЛОСА ПРОТИВ:") || variant.startsWith("ГОЛОСА ПРОТИВ ВСЕХ:")) {
            style = getBlueStyle(sheet);
        } else if (variant.startsWith("ВОЗДЕРЖАЛИСЬ:")) {
            style = getYellowStyle(sheet);
        } else {
            style = getEmptyStyle(sheet);
        }

        String[] names = variant
                .replace("ГОЛОСА ЗА:", "")
                .replace("ГОЛОСА ПРОТИВ:", "")
                .replace("ГОЛОСА ПРОТИВ ВСЕХ:", "")
                .replace("ВОЗДЕРЖАЛИСЬ:", "")
                .trim()
                .split(",");
        for (String name : names) {
            nameColor.put(name.trim(), style);
        }
    }


    public static CellStyle getGreenStyle(XSSFSheet sheet) {
        CellStyle greenStyle = sheet.getWorkbook().createCellStyle();
        greenStyle.setFillForegroundColor(IndexedColors.SEA_GREEN.index);
        greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return greenStyle;
    }

    public static CellStyle getBlueStyle(XSSFSheet sheet) {
        CellStyle greenStyle = sheet.getWorkbook().createCellStyle();
        greenStyle.setFillForegroundColor(IndexedColors.BLUE_GREY.index);
        greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return greenStyle;
    }

    public static CellStyle getYellowStyle(XSSFSheet sheet) {
        CellStyle greenStyle = sheet.getWorkbook().createCellStyle();
        greenStyle.setFillForegroundColor(IndexedColors.YELLOW.index);
        greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return greenStyle;
    }

    public static CellStyle getEmptyStyle(XSSFSheet sheet) {
        return sheet.getWorkbook().createCellStyle();
    }

}
