package apple1417.randomizer_to_spreadsheet.spreadsheet_creator;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.awt.Color;
import java.util.HashMap;

public class CustomCellStyles {
    public enum CustomStyle {
        DEFAULT, HEADER, HEADER_BORDERLESS,
        STAR, GREEN, GREY, YELLOW, RED,
        HEADER_STAR, HEADER_GREEN, HEADER_GREY, HEADER_YELLOW, HEADER_RED;
    }

    private static HashMap<CustomStyle, XSSFCellStyle> styles;
    public static XSSFCellStyle getStyle(CustomStyle ccs) {
        return styles.get(ccs);
    }

    public static XSSFCellStyle getSigilStyle(String sigil, boolean header) {
        CustomStyle cStyle;
        switch (sigil.charAt(0)) {
            case '*': {
                cStyle = header ? CustomStyle.HEADER_STAR : CustomStyle.STAR;
                break;
            }
            case 'D': {
                cStyle = header ? CustomStyle.HEADER_GREEN : CustomStyle.GREEN;
                break;
            }
            case 'E': {
                cStyle = header ? CustomStyle.HEADER_GREY : CustomStyle.GREY;
                break;
            }
            case 'M': {
                cStyle = header ? CustomStyle.HEADER_YELLOW : CustomStyle.YELLOW;
                break;
            }
            case 'N': {
                cStyle = header ? CustomStyle.HEADER_RED : CustomStyle.RED;
                break;
            }
            default: {
                cStyle = CustomStyle.DEFAULT;
                break;
            }
        }
        return getStyle(cStyle);
    }
    public static XSSFCellStyle getSigilStyle(String sigil) {
        return getSigilStyle(sigil, false);
    }

    public static void applyToWorkbook(XSSFWorkbook workbook) {
        styles = new HashMap<CustomStyle, XSSFCellStyle>();
        for (CustomStyle ccs : CustomStyle.values()) {
            styles.put(ccs, workbook.createCellStyle());
        }

        XSSFCellStyle headerBorderless = styles.get(CustomStyle.HEADER_BORDERLESS);
        XSSFFont headerFont = workbook.createFont();
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setBold(true);
        headerBorderless.setFont(headerFont);
        headerBorderless.setAlignment(HorizontalAlignment.CENTER);
        headerBorderless.setVerticalAlignment(VerticalAlignment.CENTER);

        XSSFCellStyle header = styles.get(CustomStyle.HEADER);
        header.cloneStyleFrom(headerBorderless);
        header.setBorderBottom(BorderStyle.THIN);
        header.setBorderLeft(BorderStyle.THIN);
        header.setBorderRight(BorderStyle.THIN);
        header.setBorderTop(BorderStyle.THIN);

        styles.get(CustomStyle.STAR).setFillForegroundColor(new XSSFColor(new Color(253, 203, 156)));
        styles.get(CustomStyle.GREEN).setFillForegroundColor(new XSSFColor(new Color(183, 225, 205)));
        styles.get(CustomStyle.GREY).setFillForegroundColor(new XSSFColor(new Color(217, 217, 217)));
        styles.get(CustomStyle.YELLOW).setFillForegroundColor(new XSSFColor(new Color(252, 232, 178)));
        styles.get(CustomStyle.RED).setFillForegroundColor(new XSSFColor(new Color(244, 199, 195)));
        styles.get(CustomStyle.STAR).setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styles.get(CustomStyle.GREEN).setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styles.get(CustomStyle.GREY).setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styles.get(CustomStyle.YELLOW).setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styles.get(CustomStyle.RED).setFillPattern(FillPatternType.SOLID_FOREGROUND);

        styles.get(CustomStyle.HEADER_STAR).cloneStyleFrom(styles.get(CustomStyle.STAR));
        styles.get(CustomStyle.HEADER_GREEN).cloneStyleFrom(styles.get(CustomStyle.GREEN));
        styles.get(CustomStyle.HEADER_GREY).cloneStyleFrom(styles.get(CustomStyle.GREY));
        styles.get(CustomStyle.HEADER_YELLOW).cloneStyleFrom(styles.get(CustomStyle.YELLOW));
        styles.get(CustomStyle.HEADER_RED).cloneStyleFrom(styles.get(CustomStyle.RED));

        for (CustomStyle ccs : new CustomStyle[]{
            CustomStyle.HEADER_STAR, CustomStyle.HEADER_GREEN,
            CustomStyle.HEADER_GREY, CustomStyle.HEADER_YELLOW,
            CustomStyle.HEADER_RED}) {
            CellStyle style = styles.get(ccs);
            style.setFont(headerFont);
            style.setAlignment(HorizontalAlignment.CENTER);
            style.setVerticalAlignment(VerticalAlignment.CENTER);
        }
    }
}
