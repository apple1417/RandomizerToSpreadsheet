package apple1417.randomizer_to_spreadsheet.spreadsheet_creator;

import apple1417.randomizer.Enums.RandomizerMode;
import apple1417.randomizer.Enums.ScavengerMode;
import apple1417.randomizer.Enums.MoodySigils;
import apple1417.randomizer.Enums.MobiusOptions;
import apple1417.randomizer.Enums.MobiusRandomArrangers;
import apple1417.randomizer.GeneratorGeneric;
import apple1417.randomizer.Rand;
import apple1417.randomizer.TalosProgress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

// Normally I'd just call these through MainFile but I want to know how many times I use each of these
import java.util.function.Function;
import java.util.function.Supplier;

import static apple1417.randomizer_to_spreadsheet.spreadsheet_creator.CustomCellStyles.getSigilStyle;
import static apple1417.randomizer_to_spreadsheet.spreadsheet_creator.MainFile.paint3;
import static apple1417.randomizer_to_spreadsheet.spreadsheet_creator.MainFile.paint4;
import static apple1417.randomizer_to_spreadsheet.spreadsheet_creator.MainFile.paint5a;
import static apple1417.randomizer_to_spreadsheet.spreadsheet_creator.MainFile.paint5b;

// This sheet just has basic info about the randomization
public class InfoSheet {
    public static void createSheet(XSSFSheet sheet, TalosProgress progress) {
        XSSFRow row;
        XSSFCell cell;

        row = sheet.createRow(1);
        row.createCell(0).setCellValue("Checksum:");
        row.createCell(1).setCellValue(progress.getChecksum(false));

        int paintSeed = progress.getVar("PaintItemSeed");
        row = sheet.createRow(3);
        row.createCell(0).setCellValue("Paint3:");
        row.createCell(1).setCellValue(paint3(paintSeed));
        row = sheet.createRow(4);
        row.createCell(0).setCellValue("Paint4:");
        row.createCell(1).setCellValue(paint4(paintSeed));
        row = sheet.createRow(5);
        row.createCell(0).setCellValue("Paint5a:");
        row.createCell(1).setCellValue(paint5a(paintSeed));
        row = sheet.createRow(6);
        row.createCell(0).setCellValue("Paint5b:");
        row.createCell(1).setCellValue(paint5b(paintSeed));

        row = sheet.createRow(8);
        row.createCell(0).setCellValue("Floor 4:");
        row.createCell(1).setCellValue(progress.getVar("Code_Floor4"));
        row = sheet.createRow(9);
        row.createCell(0).setCellValue("Floor 5:");
        row.createCell(1).setCellValue(progress.getVar("Code_Floor5"));
        row = sheet.createRow(10);
        row.createCell(0).setCellValue("Floor 6:");
        row.createCell(1).setCellValue(progress.getVar("Code_Floor6"));

        int rowIndex = 11;
        if (progress.getVar("Randomizer_ExtraSigils") == 1) {
            row = sheet.createRow(rowIndex);
            row.createCell(0).setCellValue("Random Extra Sigils:");
            row = sheet.createRow(rowIndex + 1);
            rowIndex += 2;

            // This is not something decided during generation so we have to manually work it out
            Rand r = new Rand(progress.getVar("Randomizer_ExtraSeed"));
            for (int i = 0; i < 20; i++) {
                r.next(0,0);
            }
            int[] sigilCounts = new int[]{
                30,
                2, 5, 3, 4, 4,
                1, 1, 4, 1, 2, 10, 4,
                6, 4, 10, 7, 4, 12, 6
            };
            String[] sigils = new String[] {
                "**",
                "DI", "DJ", "DL", "DT", "DZ",
                "MI", "MJ", "ML", "MO", "MS", "MT", "MZ",
                "NI", "NJ", "NL", "NO", "NS", "NT", "NZ"
            };

            int total = 0;
            int[] weights = new int[sigils.length];
            for (int i = 0; i < sigils.length; i++) {
                int w = sigilCounts[i] == 0 ? 0 : (int) (1f/sigilCounts[i]) * 100;
                weights[i] = w;
                total += w;
            }

            for (int col = 0; col < 15; col++){
                int selectedIndex = r.next(1, total) - 1;
                int sigilIndex = 0;

                while (selectedIndex > weights[sigilIndex]) {
                    selectedIndex -= weights[sigilIndex];
                    sigilIndex++;
                }

                String sigil = sigils[sigilIndex];
                cell = row.createCell(1);
                cell.setCellValue(sigil);
                cell.setCellStyle(getSigilStyle(sigil));
            }
        }

        int mobius = progress.getVar("Randomizer_Loop");
        if ((mobius & MobiusOptions.RANDOM_ARRANGERS.getMask()) != 0) {
            row = sheet.createRow(rowIndex);
            row.createCell(0).setCellValue("Random Arrangers:");
            row = sheet.createRow(rowIndex + 1);
            rowIndex += 2;
            int arrangers = progress.getVar("Randomizer_LoopArrangers");
            int colIndex = 0;
            for (MobiusRandomArrangers mra : MobiusRandomArrangers.values()) {
                if ((arrangers & mra.getMask()) != 0) {
                    row.createCell(colIndex).setCellValue(mra.toString());
                    colIndex++;
                }
            }
        }

        row = sheet.createRow(rowIndex + 2);
        row.createCell(0).setCellValue("Generator Info:");
        row.createCell(1).setCellValue(new GeneratorGeneric().getInfo().split("\n")[0]);
        rowIndex += 3;

        String pickedOptions = String.format(
            "%d/%s",
            progress.getVar("Randomizer_Seed"),
            RandomizerMode.fromTalosProgress(progress)
        );
        if (progress.getVar("Randomizer_NoGates") == 1) {
            pickedOptions += "/Open All Worlds";
        }
        if (progress.getVar("Randomizer_UnlockItems") == 1) {
            pickedOptions += "/Unlock All Items";
        }
        if (progress.getVar("Randomizer_ShowAll") == 1) {
            pickedOptions += "/Show Full Signs";
        }
        if (progress.getVar("Randomizer_Portals") == 1) {
            pickedOptions += "/Random Portals";
        }
        if (progress.getVar("Randomizer_Jetpack") == 1) {
            pickedOptions += "/Jetpack";
        }
        if (progress.getVar("Randomizer_ExtraSigils") == 1) {
            pickedOptions += "/Random Extra Sigils";
        }
        ScavengerMode scavenger = ScavengerMode.fromTalosProgress(progress);
        if (scavenger != ScavengerMode.OFF) {
            pickedOptions += "/" + scavenger.toString() + " Scavenger";
        }
        if (mobius != 0) {
            pickedOptions += "/Mobius ";
            boolean pastFirst = false;
            for (int i = 0; i < MobiusOptions.values().length; i++) {
                if ((mobius & MobiusOptions.fromInt(i).getMask()) != 0) {
                    pickedOptions += (pastFirst ? " + " : "") + MobiusOptions.fromInt(i).toString();
                    pastFirst = true;
                }
            }
        }
        MoodySigils moody = MoodySigils.fromTalosProgress(progress);
        if (moody != MoodySigils.OFF) {
            pickedOptions += "/Moody " + moody.toString();
        }

        // Doing this before adding pickedOptions because that's supposed to overflow
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
        sheet.createRow(0).createCell(0).setCellValue(pickedOptions);
    }
}
