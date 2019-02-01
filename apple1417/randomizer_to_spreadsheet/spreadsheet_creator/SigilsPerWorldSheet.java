package apple1417.randomizer_to_spreadsheet.spreadsheet_creator;

import apple1417.randomizer.Enums.World;
import apple1417.randomizer.TalosProgress;
import apple1417.randomizer_to_spreadsheet.spreadsheet_creator.CustomCellStyles.CustomStyle;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;

// Normally I'd just call these through MainFile but I want to know how many times I use each of these
import static apple1417.randomizer_to_spreadsheet.spreadsheet_creator.CustomCellStyles.getSigilStyle;
import static apple1417.randomizer_to_spreadsheet.spreadsheet_creator.CustomCellStyles.getStyle;
import static apple1417.randomizer_to_spreadsheet.spreadsheet_creator.MainFile.getWorldOrder;
import static apple1417.randomizer_to_spreadsheet.spreadsheet_creator.MainFile.getWorldMarkers;
import static apple1417.randomizer_to_spreadsheet.spreadsheet_creator.MainFile.getWorldName;
import static apple1417.randomizer_to_spreadsheet.spreadsheet_creator.MainFile.removeHiddenScavengerSigils;

// This sheet shows what sigils are available in each world, like on the signs
public class SigilsPerWorldSheet {
    public static void createSheet(XSSFSheet sheet, TalosProgress progress) {
        CellRangeAddress cellRange;
        XSSFRow row;
        XSSFCell cell;

        row = sheet.createRow(0);
        cell = row.createCell(0);
        cell.setCellValue("World");
        cell.setCellStyle(getStyle(CustomStyle.HEADER));
        cell = row.createCell(1);
        cell.setCellValue("Sigils");
        cell.setCellStyle(getStyle(CustomStyle.HEADER));

        int rowNum = 1;
        int portalNum = 0;
        int maxWidth = 1;
        for (World w : getWorldOrder(progress)) {
            ArrayList<Integer> worldSigilIds = new ArrayList<Integer>();
            for (String marker : getWorldMarkers(progress, w)) {
                worldSigilIds.add(progress.getVar(marker));
            }
            if (worldSigilIds.size() > maxWidth) {
                maxWidth = worldSigilIds.size();
            }

            // Don't want to show an empty row
            if (worldSigilIds.size() == 0) {
                portalNum++;
                continue;
            }

            row = sheet.createRow(rowNum);
            cell = row.createCell(0);
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
            cell.setCellStyle(getStyle(CustomStyle.HEADER_BORDERLESS));

            // We want to put the stars on the end
            int starCount = worldSigilIds.size();
            worldSigilIds.removeIf(x -> x <= 30);
            starCount -= worldSigilIds.size();
            Collections.sort(worldSigilIds);
            // We don't care what star specifically so I can use the same id for them all
            worldSigilIds.addAll(Collections.nCopies(starCount, 1));

            for (int i = 0; i < worldSigilIds.size(); i++) {
                String sigil = TalosProgress.TETROS[worldSigilIds.get(i)].substring(0, 2);
                cell = row.createCell(i + 1);
                cell.setCellValue(sigil);
                cell.setCellStyle(getSigilStyle(sigil));
            }
            portalNum++;
            rowNum++;
        }
        // Bunch of duplicated code, but like before I need to deal with the nexus stars
        if (progress.getVar("Randomizer_Loop") == 0) {
            ArrayList<String> markers = removeHiddenScavengerSigils(progress,
                new ArrayList<String>(Arrays.asList("F0-Star", "F3-Star"))
            );
            if (markers.size() > 0) {
                row = sheet.createRow(rowNum);
                cell = row.createCell(0);
                cell.setCellValue("Nexus");
                cell.setCellStyle(getStyle(CustomStyle.HEADER_BORDERLESS));
                int F0 = progress.getVar("F0-Star");
                int F3 = progress.getVar("F3-Star");
                // There're only ever going to be two sigils here so I can just swap the ids
                if (F0 <= 30 || F0 > F3) {
                    int tmp = F3;
                    F3 = F0;
                    F0 = tmp;
                }
                String sigil = TalosProgress.TETROS[F0].substring(0, 2);
                cell = row.createCell(1);
                cell.setCellValue(sigil);
                cell.setCellStyle(getSigilStyle(sigil));
                sigil = TalosProgress.TETROS[F3].substring(0, 2);
                cell = row.createCell(2);
                cell.setCellValue(sigil);
                cell.setCellStyle(getSigilStyle(sigil));
                rowNum++;
            }
        }
        // Tidy the sheet up a bit
        sheet.autoSizeColumn(0);
        for (int i = 0; i < maxWidth; i++) {
            sheet.setColumnWidth(i + 1, 1024);
        }
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 1, maxWidth));
        RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(0, 0, 0, maxWidth), sheet);
        RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(rowNum, rowNum, 0, maxWidth), sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, new CellRangeAddress(0, rowNum - 1, maxWidth + 1, maxWidth + 1), sheet);
        cellRange = new CellRangeAddress(0, rowNum - 1, 0, 0);
        RegionUtil.setBorderLeft(BorderStyle.THIN, cellRange, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, cellRange, sheet);

    }
}
