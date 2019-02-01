package apple1417.randomizer_to_spreadsheet.spreadsheet_creator;

import apple1417.randomizer.Enums.ScavengerEnding;
import apple1417.randomizer_to_spreadsheet.spreadsheet_creator.CustomCellStyles.CustomStyle;
import apple1417.randomizer.TalosProgress;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;

// Normally I'd just call these through MainFile but I want to know how many times I use each of these
import static apple1417.randomizer_to_spreadsheet.spreadsheet_creator.MainFile.allMarkers;
import static apple1417.randomizer_to_spreadsheet.spreadsheet_creator.MainFile.getMarkerName;

// This sheet shows all locations sorted by sigil type
public class LocationsBySigilSheet {
    public static void createSheet(XSSFSheet sheet, TalosProgress progress) {
        CellRangeAddress cellRange;
        XSSFRow row;
        XSSFCell cell;

        row = sheet.createRow(0);
        cell = row.createCell(0);
        cell.setCellValue("Sigil Type");
        cell.setCellStyle(CustomCellStyles.getStyle(CustomStyle.HEADER));
        cell = row.createCell(1);
        cell.setCellValue("Locations");
        cell.setCellStyle(CustomCellStyles.getStyle(CustomStyle.HEADER));

        // Map all locations to a sigil type
        ArrayList<String> allowedSigils = null;
        if (progress.getVar("Randomizer_Scavenger") != 0) {
            allowedSigils = ScavengerEnding.fromTalosProgress(progress).getAllowedSigils();
        }
        HashMap<String, ArrayList<String>> sigilLocations = new HashMap<String, ArrayList<String>>();
        for (String marker : allMarkers) {
            String sigil = TalosProgress.TETROS[progress.getVar(marker)];
            if (allowedSigils != null) {
                if (!allowedSigils.contains(sigil)) {
                    continue;
                }
            }
            sigil = sigil.substring(0, 2);
            if (!sigilLocations.containsKey(sigil)) {
                sigilLocations.put(sigil, new ArrayList<String>());
            }
            sigilLocations.get(sigil).add(marker);
        }

        // We want to show the sigil types in alphabetical order so colours are grouped
        ArrayList<String> sigilTypes = new ArrayList<String>(sigilLocations.keySet());
        Collections.sort(sigilTypes);
        int rowNum = 1;

        // Handle stars first, they'll span three rows instead of one
        if (progress.getVar("Randomizer_Scavenger") != 0) {
            sigilTypes.remove("**");
            ArrayList<String> stars = sigilLocations.get("**");
            Collections.sort(stars);

            row = sheet.createRow(1);
            cell = row.createCell(0);
            cell.setCellValue("**");
            cell.setCellStyle(CustomCellStyles.getStyle(CustomStyle.HEADER_STAR));

            for (int i = 0; i < stars.size(); i++) {
                if (i % 10 == 0 && i != 0) {
                    rowNum++;
                    row = sheet.createRow(rowNum);
                }
                row.createCell(i % 10 + 1).setCellValue(getMarkerName(stars.get(i)));
            }
            rowNum++;
        }

        for (String type : sigilTypes) {
            row = sheet.createRow(rowNum);
            cell = row.createCell(0);
            cell.setCellValue(type);
            cell.setCellStyle(CustomCellStyles.getSigilStyle(type, true));

            Collections.sort(sigilLocations.get(type));
            int columnNum = 1;
            for (String marker : sigilLocations.get(type)) {
                row.createCell(columnNum).setCellValue(getMarkerName(marker));
                columnNum++;
            }
            rowNum++;
        }

        sheet.autoSizeColumn(0);
        if (progress.getVar("Randomizer_Scavenger") != 0) {
            for (int i = 0; i < 13; i++) {
                sheet.setColumnWidth(i + 1, 4096);
            }
            sheet.addMergedRegion(new CellRangeAddress(1, 3, 0, 0));
            cellRange = new CellRangeAddress(0, 0, 1, 12);
            sheet.addMergedRegion(cellRange);
            RegionUtil.setBorderRight(BorderStyle.THIN, cellRange, sheet);
            // This sheet is a lot more constant than the rest so I can safely hardcode all these
            for (CellRangeAddress cr : new CellRangeAddress[]{
                new CellRangeAddress(1, 3, 1, 12), // Stars
                new CellRangeAddress(4, 8, 1, 12), // Greens
                new CellRangeAddress(9, 11, 1, 12), // Greys
                new CellRangeAddress(12, 18, 1, 12), // Yellows
                new CellRangeAddress(19, 25, 1, 12) /*Reds*/}) {
                RegionUtil.setBorderBottom(BorderStyle.THIN, cr, sheet);
                RegionUtil.setBorderLeft(BorderStyle.THIN, cr, sheet);
                RegionUtil.setBorderRight(BorderStyle.THIN, cr, sheet);
                RegionUtil.setBorderTop(BorderStyle.THIN, cr, sheet);
            }
            RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(26, 26, 0, 0), sheet);
            // Scavenger is going to have different dimentions
        } else {
            int rows = 0;
            int cols = 0;
            // Again just kind of hardcode
            switch (progress.getVar("Randomizer_ScavengerMode")) {
                // Connector Clip
                case 1: {
                    rows = 4;
                    cols = 2;
                    RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(3, 3, 1, 2), sheet);
                    break;
                }
                // F2 Clip
                case 2: {
                    rows = 7;
                    cols = 6;
                    RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(4, 4, 1, 6), sheet);
                    break;
                }
                // F3 Clip
                case 3: {
                    rows = 5;
                    cols = 4;
                    break;
                }
                // F6
                case 4: {
                    rows = 5;
                    cols = 4;
                    RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(4, 4, 1, 4), sheet);
                    break;
                }
            }

            for (int i = 0; i < cols; i++) {
                sheet.setColumnWidth(i + 1, 4096);
            }
            cellRange = new CellRangeAddress(0, 0, 1, cols);
            sheet.addMergedRegion(cellRange);
            RegionUtil.setBorderRight(BorderStyle.THIN, cellRange, sheet);
            cellRange = new CellRangeAddress(1, rows, 1, cols);
            RegionUtil.setBorderBottom(BorderStyle.THIN, cellRange, sheet);
            RegionUtil.setBorderLeft(BorderStyle.THIN, cellRange, sheet);
            RegionUtil.setBorderRight(BorderStyle.THIN, cellRange, sheet);
            RegionUtil.setBorderTop(BorderStyle.THIN, cellRange, sheet);
            RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(rows + 1, rows + 1, 0, 0), sheet);
        }
    }
}
