package apple1417.randomizer_to_spreadsheet.spreadsheet_creator;

import apple1417.randomizer.Enums.World;
import apple1417.randomizer.TalosProgress;
import apple1417.randomizer_to_spreadsheet.spreadsheet_creator.CustomCellStyles.CustomStyle;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;
import java.util.Arrays;

// Normally I'd just call these through MainFile but I want to know how many times I use each of these
import static apple1417.randomizer_to_spreadsheet.spreadsheet_creator.CustomCellStyles.getSigilStyle;
import static apple1417.randomizer_to_spreadsheet.spreadsheet_creator.CustomCellStyles.getStyle;
import static apple1417.randomizer_to_spreadsheet.spreadsheet_creator.MainFile.getWorldOrder;
import static apple1417.randomizer_to_spreadsheet.spreadsheet_creator.MainFile.getWorldMarkers;
import static apple1417.randomizer_to_spreadsheet.spreadsheet_creator.MainFile.getMarkerName;
import static apple1417.randomizer_to_spreadsheet.spreadsheet_creator.MainFile.getSigilType;
import static apple1417.randomizer_to_spreadsheet.spreadsheet_creator.MainFile.getWorldName;
import static apple1417.randomizer_to_spreadsheet.spreadsheet_creator.MainFile.removeHiddenScavengerSigils;

// This sheet shows the sigil at each location, sorted by world
public class AllSigilsSheet {
    public static void createSheet(XSSFSheet sheet, TalosProgress progress) {
        CellRangeAddress cellRange;
        XSSFRow row;
        XSSFCell cell;

        row = sheet.createRow(0);
        cell = row.createCell(0);
        cell.setCellValue("Location");
        cell.setCellStyle(getStyle(CustomStyle.HEADER));
        cell = row.createCell(1);
        cell.setCellValue("Sigil");
        cell.setCellStyle(getStyle(CustomStyle.HEADER));

        int rowNum = 1;
        int portalNum = 0;
        for (World w : getWorldOrder(progress)) {
            ArrayList<String> worldMarkers = getWorldMarkers(progress, w);

            // Don't want to show empty worlds
            if (worldMarkers.size() == 0) {
                portalNum++;
                continue;
            }

            row = sheet.createRow(rowNum);
            cell = row.createCell(0);
            cell.setCellStyle(getStyle(CustomStyle.HEADER));
            // For random portals without mobius we need to clarify portal/world
            if (progress.getVar("Randomizer_Portals") != -1
                && progress.getVar("Randomizer_Loop") == 0) {

                cell.setCellValue(String.format(
                    "%s (leads to %s)",
                    getWorldName(World.fromInt(portalNum)),
                    getWorldName(w)
                ));
            } else {
                cell.setCellValue(getWorldName(w));
            }
            portalNum++;

            cellRange = new CellRangeAddress(rowNum, rowNum, 0, 1);
            sheet.addMergedRegion(cellRange);
            // These two borders get messed up by the merge
            RegionUtil.setBorderLeft(BorderStyle.THIN, cellRange, sheet);
            RegionUtil.setBorderRight(BorderStyle.THIN, cellRange, sheet);
            rowNum++;
            for (String marker : worldMarkers) {
                row = sheet.createRow(rowNum);
                if (marker.charAt(0) == 'F') {
                    row.createCell(0).setCellValue("Mobius Extra");
                } else {
                    row.createCell(0).setCellValue(getMarkerName(marker));
                }
                cell = row.createCell(1);
                String sigil = getSigilType(progress, marker);
                cell.setCellValue(sigil);
                cell.setCellStyle(getSigilStyle(sigil));
                rowNum++;
            }
        }

        // This is a bunch of duplicated code but I need a way to deal with the tower stars
        if (progress.getVar("Randomizer_Loop") == 0) {
            ArrayList<String> markers = removeHiddenScavengerSigils(progress,
                new ArrayList<String>(Arrays.asList("F0-Star", "F3-Star"))
            );
            if (markers.size() > 0) {
                row = sheet.createRow(rowNum);
                cell = row.createCell(0);
                cell.setCellStyle(getStyle(CustomStyle.HEADER));
                cell.setCellValue("Nexus");
                cellRange = new CellRangeAddress(rowNum, rowNum, 0, 1);
                sheet.addMergedRegion(cellRange);
                RegionUtil.setBorderLeft(BorderStyle.THIN, cellRange, sheet);
                RegionUtil.setBorderRight(BorderStyle.THIN, cellRange, sheet);
                rowNum++;
                for (String marker : markers) {
                    String sigil = getSigilType(progress, marker);
                    row = sheet.createRow(rowNum);
                    row.createCell(0).setCellValue(getMarkerName(marker));
                    cell = row.createCell(1);
                    cell.setCellValue(sigil);
                    cell.setCellStyle(getSigilStyle(sigil));
                    rowNum++;
                }
            }
        }

        // Tidy up the sheet a bit
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
        // Weird stuff happens if I just try to surround the whole thing with a single range
        RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(0, 0, 0, 1), sheet);
        RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(rowNum, rowNum, 0, 1), sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(0, rowNum - 1, 0, 0), sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(0, rowNum - 1, 2, 2), sheet);
    }
}
