package apple1417.randomizer_to_spreadsheet;

import apple1417.randomizer.Enums.RandomizerMode;
import apple1417.randomizer.Enums.ScavengerEnding;
import apple1417.randomizer.Enums.ScavengerMode;
import apple1417.randomizer.Enums.World;
import apple1417.randomizer.GeneratorGeneric;
import apple1417.randomizer.TalosProgress;
import javafx.application.Application;
import javafx.geometry.HPos;
import javafx.geometry.Insets;
import javafx.geometry.VPos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.ChoiceBox;
import javafx.scene.control.Control;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.GridPane;
import javafx.scene.text.Font;
import javafx.scene.text.Text;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.*;

import java.awt.Color;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;

import javafx.scene.text.Font;


public class MainWindow extends Application {
    public static void main(String[] args) {
        launch(args);
    }

    private FileChooser chooser = new FileChooser();
    private ArrayList<Control> leftOptions = new ArrayList<Control>();
    private ArrayList<CheckBox> topRightBoxes = new ArrayList<CheckBox>();
    private ArrayList<CheckBox> bottomRightBoxes = new ArrayList<CheckBox>();
    private ArrayList<Text> topRightLabels = new ArrayList<Text>();
    private ArrayList<Text> bottomRightLabels = new ArrayList<Text>();

    @Override
    public void start(Stage primaryStage) {
        FileChooser.ExtensionFilter filter = new FileChooser.ExtensionFilter("Excel Spreadsheet (*.xlsx)", "*.xlsx");
        chooser.getExtensionFilters().add(filter);
        chooser.setSelectedExtensionFilter(filter);

        GridPane root = new GridPane();
        root.setHgap(10);
        root.setVgap(5);
        root.setPadding(new Insets(7.5, 7.5, 7.5, 7.5));

        // The left side, drop down boxes and the seed input
        IntField seed = new IntField();

        ChoiceBox mode = new ChoiceBox();
        mode.getItems().addAll(RandomizerMode.values());
        mode.setValue(RandomizerMode.DEFAULT);

        ChoiceBox scavenger = new ChoiceBox();
        scavenger.getItems().addAll(ScavengerMode.values());
        scavenger.setValue(ScavengerMode.OFF);

        ChoiceBox moody = new ChoiceBox();
        moody.getItems().addAll(MoodySigils.values());
        moody.setValue(MoodySigils.OFF);

        leftOptions.addAll(Arrays.asList(seed, mode, scavenger, moody));

        for (Control c : leftOptions) {
            c.setMaxWidth(150);
            c.setMinWidth(150);
        }

        ArrayList<Text> leftLabels = new ArrayList<Text>();
        leftLabels.addAll(Arrays.asList(
            new Text("Seed"),
            new Text("Mode"),
            new Text("Scavenger Hunt"),
            new Text("Moody Sigils")
        ));

        // Just need something fill the last column
        BorderPane forceSize = new BorderPane();
        forceSize.setMinWidth(100);
        forceSize.setMaxWidth(100);

        root.addColumn(0,
            leftLabels.get(0),
            leftLabels.get(1),
            leftLabels.get(2),
            leftLabels.get(3),
            forceSize
        );
        root.addColumn(1,
            seed,
            mode,
            scavenger,
            moody
        );

        // Right side, standard checkbox options
        topRightLabels = new ArrayList<Text>();
        topRightLabels.addAll(Arrays.asList(
            new Text("Open All Worlds"),
            new Text("Unlock All Items"),
            new Text("Show Full Signs"),
            new Text("Random Portals"),
            new Text("")
        ));
        bottomRightLabels = new ArrayList<Text>();
        bottomRightLabels.addAll(Arrays.asList(
            new Text("All Sigils"),
            new Text("All of a Shape"),
            new Text("All of a Colour"),
            new Text("Eternalize Ending"),
            new Text("Two Tower Floors")
        ));

        for (int i = 0; i < 5; i++) {
            CheckBox top = new CheckBox();
            CheckBox bottom = new CheckBox();

            topRightBoxes.add(top);
            bottomRightBoxes.add(bottom);

            // Need to use BorderPanes to center the CheckBoxes
            BorderPane topPane = new BorderPane();
            BorderPane bottomPane = new BorderPane();
            topPane.setCenter(top);
            bottomPane.setCenter(bottom);

            topPane.setMaxWidth(100);
            topPane.setMinWidth(100);
            bottomPane.setMaxWidth(100);
            bottomPane.setMinWidth(100);

            // Might as well add to the grid while we're looping
            root.add(topRightLabels.get(i), 2 + i, 0);
            root.add(topPane, 2 + i, 1);
            root.add(bottomPane, 2 + i, 2);
            root.add(bottomRightLabels.get(i), 2 + i, 3);
        }
        topRightBoxes.get(2).fire();
        topRightBoxes.get(4).setDisable(true);

        // Setup the options locking that normal goes on
        scavenger.setOnAction((obs) -> {
            if (scavenger.getValue() != ScavengerMode.OFF) {
                moody.setValue(MoodySigils.OFF);
                topRightBoxes.get(0).setSelected(true);
                topRightBoxes.get(2).setSelected(false);
            }
        });
        moody.setOnAction((obs) -> {
            if (moody.getValue() != MoodySigils.OFF) {
                scavenger.setValue(ScavengerMode.OFF);
            }
        });
        // Open All Worlds
        topRightBoxes.get(0).setOnAction((obs) -> {
            if (!topRightBoxes.get(0).isSelected()) {
                scavenger.setValue(ScavengerMode.OFF);
            }
        });
        // Unlock All Items
        topRightBoxes.get(2).setOnAction((obs) -> {
            if (topRightBoxes.get(0).isSelected()) {
                scavenger.setValue(ScavengerMode.OFF);
            }
        });


        // The export button to actually save it to file
        Button exportButton = new Button("Export");
        exportButton.setOnAction((obs) -> saveFile(chooser.showSaveDialog(primaryStage)));
        exportButton.setFont(new Font(18));
        exportButton.setMinWidth(150);
        exportButton.setMaxWidth(150);
        root.add(exportButton, 0, 4, 7, 2);
        root.setHalignment(exportButton, HPos.CENTER);
        root.setValignment(exportButton, VPos.CENTER);

        primaryStage.setScene(new Scene(root, 815, 165));
        primaryStage.setResizable(false);
        primaryStage.setTitle("Randomizer To Spreadsheet");
        primaryStage.show();

        for (Text t : topRightLabels) {
            root.setHalignment(t, HPos.CENTER);
            root.setValignment(t, VPos.CENTER);
        }
        for (Text t : bottomRightLabels) {
            root.setHalignment(t, HPos.CENTER);
            root.setValignment(t, VPos.CENTER);
        }
    }

    private TalosProgress progress;

    private void saveFile(File saveLoc) {
        // No need to process everything if we can't save the file
        if (saveLoc == null) {
            return;
        }
        FileOutputStream outFile;
        try {
            outFile = new FileOutputStream(saveLoc);
        } catch (FileNotFoundException e) {
            return;
        }

        // Convert the options to a TalosProgress object
        long seed = ((IntField) leftOptions.get(0)).getValue();
        int mode = Arrays.asList(RandomizerMode.values()).indexOf((((ChoiceBox) leftOptions.get(1)).getValue()));
        int scavenger = Arrays.asList(ScavengerMode.values()).indexOf((((ChoiceBox) leftOptions.get(2)).getValue()));
        int moody = Arrays.asList(MoodySigils.values()).indexOf((((ChoiceBox) leftOptions.get(3)).getValue()));

        int mobius = 0;
        for (int i = bottomRightBoxes.size() - 1; i >= 0; i--) {
            mobius *= 2;
            if (bottomRightBoxes.get(i).isSelected()) {
                mobius += 1;
            }
        }

        HashMap<String, Integer> options = new HashMap<String, Integer>();
        options.put("Randomizer_Mode", mode);
        options.put("Randomizer_Scavenger", scavenger);
        options.put("Randomizer_Moody", moody);
        options.put("Randomizer_Loop", mobius);
        if (topRightBoxes.get(0).isSelected()) {
            options.put("Randomizer_NoGates", 1);
        }
        if (topRightBoxes.get(1).isSelected()) {
            options.put("Randomizer_UnlockItems", 1);
        }
        if (topRightBoxes.get(2).isSelected()) {
            options.put("Randomizer_ShowAll", 1);
        }
        if (topRightBoxes.get(3).isSelected()) {
            options.put("Randomizer_Portals", 1);
        }
        progress = new GeneratorGeneric(options).generate(seed);

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet;
        XSSFRow row;
        XSSFCell cell;

        // Setup some cell styles, mostly so sigils have different highlight colours
        HashMap<CustomCellStyles, XSSFCellStyle> styles = new HashMap<CustomCellStyles, XSSFCellStyle>();
        for (CustomCellStyles ccs : CustomCellStyles.values()) {
            styles.put(ccs, workbook.createCellStyle());
        }

        XSSFCellStyle headerBorderless = styles.get(CustomCellStyles.HEADER_BORDERLESS);
        XSSFFont headerFont = workbook.createFont();
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setBold(true);
        headerBorderless.setFont(headerFont);
        headerBorderless.setAlignment(HorizontalAlignment.CENTER);
        headerBorderless.setVerticalAlignment(VerticalAlignment.CENTER);

        XSSFCellStyle header = styles.get(CustomCellStyles.HEADER);
        header.cloneStyleFrom(headerBorderless);
        header.setBorderBottom(BorderStyle.THIN);
        header.setBorderLeft(BorderStyle.THIN);
        header.setBorderRight(BorderStyle.THIN);
        header.setBorderTop(BorderStyle.THIN);

        styles.get(CustomCellStyles.STAR).setFillForegroundColor(new XSSFColor(new Color(253, 203, 156)));
        styles.get(CustomCellStyles.GREEN).setFillForegroundColor(new XSSFColor(new Color(183, 225, 205)));
        styles.get(CustomCellStyles.GREY).setFillForegroundColor(new XSSFColor(new Color(217, 217, 217)));
        styles.get(CustomCellStyles.YELLOW).setFillForegroundColor(new XSSFColor(new Color(252, 232, 178)));
        styles.get(CustomCellStyles.RED).setFillForegroundColor(new XSSFColor(new Color(244, 199, 195)));
        styles.get(CustomCellStyles.STAR).setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styles.get(CustomCellStyles.GREEN).setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styles.get(CustomCellStyles.GREY).setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styles.get(CustomCellStyles.YELLOW).setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styles.get(CustomCellStyles.RED).setFillPattern(FillPatternType.SOLID_FOREGROUND);

        styles.get(CustomCellStyles.HEADER_STAR).cloneStyleFrom(styles.get(CustomCellStyles.STAR));
        styles.get(CustomCellStyles.HEADER_GREEN).cloneStyleFrom(styles.get(CustomCellStyles.GREEN));
        styles.get(CustomCellStyles.HEADER_GREY).cloneStyleFrom(styles.get(CustomCellStyles.GREY));
        styles.get(CustomCellStyles.HEADER_YELLOW).cloneStyleFrom(styles.get(CustomCellStyles.YELLOW));
        styles.get(CustomCellStyles.HEADER_RED).cloneStyleFrom(styles.get(CustomCellStyles.RED));

        for (CustomCellStyles ccs : new CustomCellStyles[]{
            CustomCellStyles.HEADER_STAR, CustomCellStyles.HEADER_GREEN,
            CustomCellStyles.HEADER_GREY, CustomCellStyles.HEADER_YELLOW,
            CustomCellStyles.HEADER_RED}) {
            CellStyle style = styles.get(ccs);
            style.setFont(headerFont);
            style.setAlignment(HorizontalAlignment.CENTER);
            style.setVerticalAlignment(VerticalAlignment.CENTER);
        }


        // This sheet just has basic info about the randomization
        sheet = workbook.createSheet("Info");

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

        String pickedOptions = String.format("%d/%s", seed, RandomizerMode.fromInt(mode).toString());
        for (int i = 0; i < topRightBoxes.size(); i++) {
            if (topRightBoxes.get(i).isSelected()) {
                pickedOptions += "/" + topRightLabels.get(i).getText();
            }
        }
        if (mobius != 0) {
            pickedOptions += "/Mobius ";
            boolean pastFirst = false;
            for (int i = 0; i < bottomRightBoxes.size(); i++) {
                if (bottomRightBoxes.get(i).isSelected()) {
                    pickedOptions += (pastFirst ? " + " : "") + bottomRightLabels.get(i).getText();
                    pastFirst = true;
                }
            }
        }

        ScavengerMode scavengerChoice = ScavengerMode.fromInt(scavenger);
        if (scavengerChoice != ScavengerMode.OFF) {
            pickedOptions += "/" + scavengerChoice.toString() + " Scavenger";
        }
        MoodySigils moodyChoice = MoodySigils.fromInt(moody);
        if (moodyChoice != MoodySigils.OFF) {
            pickedOptions += "/Moody " + moodyChoice.toString();
        }

        // Doing this before adding pickedOptions because that's supposed to go over
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
        sheet.createRow(0).createCell(0).setCellValue(pickedOptions);


        // This sheet shows the sigil at each location, sorted by world
        sheet = workbook.createSheet("All Sigils");
        row = sheet.createRow(0);
        cell = row.createCell(0);
        cell.setCellValue("Location");
        cell.setCellStyle(styles.get(CustomCellStyles.HEADER));
        cell = row.createCell(1);
        cell.setCellValue("Sigil");
        cell.setCellStyle(styles.get(CustomCellStyles.HEADER));

        int rowNum = 1;
        int portalNum = 0;
        for (World w : getWorldOrder()) {
            ArrayList<String> worldMarkers = getWorldMarkers(w);

            // Don't want to show empty worlds
            if (worldMarkers.size() == 0) {
                portalNum++;
                continue;
            }

            row = sheet.createRow(rowNum);
            cell = row.createCell(0);
            cell.setCellStyle(styles.get(CustomCellStyles.HEADER));
            // For random portals without mobius we need to clarify portal/world
            if (topRightBoxes.get(3).isSelected() && mobius == 0) {
                cell.setCellValue(String.format("%s (leads to %s)", worldToName(World.fromInt(portalNum)), worldToName(w)));
            } else {
                cell.setCellValue(worldToName(w));
            }
            portalNum++;

            CellRangeAddress cellRange = new CellRangeAddress(rowNum, rowNum, 0, 1);
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
                    row.createCell(0).setCellValue(markerToName.get(marker));
                }
                cell = row.createCell(1);
                String sigil = getSigil(progress, marker);
                cell.setCellValue(sigil);
                cell.setCellStyle(styles.get(CustomCellStyles.sigilToStyle(sigil)));
                rowNum++;
            }
        }

        // This is a bunch of duplicated code but I need a way to deal with the tower stars
        if (mobius == 0) {
            ArrayList<String> markers = removeHiddenScavengerSigils(
                new ArrayList<String>(Arrays.asList("F0-Star", "F3-Star"))
            );
            if (markers.size() > 0) {
                row = sheet.createRow(rowNum);
                cell = row.createCell(0);
                cell.setCellStyle(styles.get(CustomCellStyles.HEADER));
                cell.setCellValue("Nexus");
                CellRangeAddress cellRange = new CellRangeAddress(rowNum, rowNum, 0, 1);
                sheet.addMergedRegion(cellRange);
                RegionUtil.setBorderLeft(BorderStyle.THIN, cellRange, sheet);
                RegionUtil.setBorderRight(BorderStyle.THIN, cellRange, sheet);
                rowNum++;
                for (String marker : markers) {
                    String sigil = getSigil(progress, marker);
                    row = sheet.createRow(rowNum);
                    row.createCell(0).setCellValue(markerToName.get(marker));
                    cell = row.createCell(1);
                    cell.setCellValue(sigil);
                    cell.setCellStyle(styles.get(CustomCellStyles.sigilToStyle(sigil)));
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


        // This sheet shows what sigils are available in each world, like on the signs
        sheet = workbook.createSheet("Sigils Per World");
        row = sheet.createRow(0);
        cell = row.createCell(0);
        cell.setCellValue("World");
        cell.setCellStyle(styles.get(CustomCellStyles.HEADER));
        cell = row.createCell(1);
        cell.setCellValue("Sigils");
        cell.setCellStyle(styles.get(CustomCellStyles.HEADER));

        rowNum = 1;
        portalNum = 0;
        int maxWidth = 1;
        for (World w : getWorldOrder()) {
            ArrayList<Integer> worldSigilIds = new ArrayList<Integer>();
            for (String marker : getWorldMarkers(w)) {
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
            if (topRightBoxes.get(3).isSelected() && mobius == 0) {
                cell.setCellValue(String.format("%s (leads to %s)", worldToName(World.fromInt(portalNum)), worldToName(w)));
            } else {
                cell.setCellValue(worldToName(w));
            }
            cell.setCellStyle(styles.get(CustomCellStyles.HEADER_BORDERLESS));

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
                cell.setCellStyle(styles.get(CustomCellStyles.sigilToStyle(sigil)));
            }
            portalNum++;
            rowNum++;
        }
        // Bunch of duplicated code, but like before I need to deal with the nexus stars
        if (mobius == 0) {
            ArrayList<String> markers = removeHiddenScavengerSigils(
                new ArrayList<String>(Arrays.asList("F0-Star", "F3-Star"))
            );
            if (markers.size() > 0) {
                row = sheet.createRow(rowNum);
                cell = row.createCell(0);
                cell.setCellValue("Nexus");
                cell.setCellStyle(styles.get(CustomCellStyles.HEADER_BORDERLESS));
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
                cell.setCellStyle(styles.get(CustomCellStyles.sigilToStyle(sigil)));
                sigil = TalosProgress.TETROS[F3].substring(0, 2);
                cell = row.createCell(2);
                cell.setCellValue(sigil);
                cell.setCellStyle(styles.get(CustomCellStyles.sigilToStyle(sigil)));
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
        CellRangeAddress cellRange = new CellRangeAddress(0, rowNum - 1, 0, 0);
        RegionUtil.setBorderLeft(BorderStyle.THIN, cellRange, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, cellRange, sheet);


        // This sheet shows all locations sorted by sigil type
        sheet = workbook.createSheet("Locations by Sigil Type");
        row = sheet.createRow(0);
        cell = row.createCell(0);
        cell.setCellValue("Sigil Type");
        cell.setCellStyle(styles.get(CustomCellStyles.HEADER));
        cell = row.createCell(1);
        cell.setCellValue("Locations");
        cell.setCellStyle(styles.get(CustomCellStyles.HEADER));

        // Map all locations to a sigil type
        ArrayList<String> allowedSigils = null;
        if (scavengerChoice != ScavengerMode.OFF) {
            allowedSigils = ScavengerEnding.fromInt(progress.getVar("Randomizer_ScavengerMode") - 1).getAllowedSigils();
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
        rowNum = 1;

        // Handle stars first, they'll span three rows instead of one
        if (scavengerChoice == ScavengerMode.OFF) {
            sigilTypes.remove("**");
            ArrayList<String> stars = sigilLocations.get("**");
            Collections.sort(stars);

            row = sheet.createRow(1);
            cell = row.createCell(0);
            cell.setCellValue("**");
            cell.setCellStyle(styles.get(CustomCellStyles.HEADER_STAR));

            for (int i = 0; i < stars.size(); i++) {
                if (i % 10 == 0 && i != 0) {
                    rowNum++;
                    row = sheet.createRow(rowNum);
                }
                row.createCell(i % 10 + 1).setCellValue(markerToName.get(stars.get(i)));
            }
            rowNum++;
        }

        for (String type : sigilTypes) {
            row = sheet.createRow(rowNum);
            cell = row.createCell(0);
            cell.setCellValue(type);
            cell.setCellStyle(styles.get(CustomCellStyles.sigilToStyle(type, true)));

            Collections.sort(sigilLocations.get(type));
            int columnNum = 1;
            for (String marker : sigilLocations.get(type)) {
                row.createCell(columnNum).setCellValue(markerToName.get(marker));
                columnNum++;
            }
            rowNum++;
        }

        sheet.autoSizeColumn(0);
        if (scavengerChoice == ScavengerMode.OFF) {
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

        // Finally save the workbook
        try {
            workbook.write(outFile);
            outFile.close();
        } catch (IOException e) {}
    }

    private static HashMap<String, String> createMarkerToName() {
        HashMap<String, String> out = new HashMap<String, String>();
        out.put("A1-ASooR", "A Switch out of Reach");
        out.put("A1-Beaten Path", "Striding the Beaten Path");
        out.put("A1-OtToU", "Only the Two of Us");
        out.put("A1-Outnumbered", "Outnumbered");
        out.put("A1-PaSL", "Poking a Sleeping Lion");
        out.put("A1-Peephole", "Peephole");
        out.put("A1-Star", "A1 Star");
        out.put("A1-Trio", "Trio Bombasticus");
        out.put("A2-Guards", "The Guards Must be Crazy");
        out.put("A2-Hall of Windows", "Hall of Windows");
        out.put("A2-Star", "A2 Star");
        out.put("A2-Suicide Mission", "Suicide Mission");
        out.put("A3-ABTU Star", "A Bit Tied Up Star");
        out.put("A3-ABTU", "A Bit Tied Up");
        out.put("A3-AEP", "An Escalating Problem");
        out.put("A3-Clock Star", "Clock Star");
        out.put("A3-Stashed for Later", "Stashed for Later");
        out.put("A3-Swallowed the Key", "Locked Me Up, Swallowed the Key");
        out.put("A4-Above All That", "Above All That");
        out.put("A4-Branch it Out", "Branch it Out");
        out.put("A4-DCtS", "Don't Cross the Streams");
        out.put("A4-Push it Further", "Push it Further");
        out.put("A4-Star", "A4 Star");
        out.put("A5-FC Star", "Friendly Crossfire Star");
        out.put("A5-FC", "Friendly Crossfire");
        out.put("A5-OLB", "One Little Buzzer");
        out.put("A5-Over the Fence", "Going Over the Fence");
        out.put("A5-Two Boxes Star", "Things to Do With Two Boxes Star");
        out.put("A5-Two Boxes", "Things to Do With Two Boxes");
        out.put("A5-YKYMCTS", "You Know You Mustn't Cross the Streams");
        out.put("A6-Bichromatic", "Bichromatic Entanglement");
        out.put("A6-Deception", "Deception");
        out.put("A6-Door too Far", "A Door too Far");
        out.put("A6-Mobile Mindfield", "Mobile Mindfield");
        out.put("A6-Star", "A6 Star");
        out.put("A7-LFI", "Locked From Inside");
        out.put("A7-Pinhole", "Pinhole Windows");
        out.put("A7-Star", "A7 Star");
        out.put("A7-Trapped Inside", "Trapped Inside");
        out.put("A7-Two Buzzers", "Two Pesky Little Buzzers");
        out.put("A7-WiaL", "Windows into a Labyrinth");
        out.put("A*-DDM", "Dumb Dumb Mine");
        out.put("A*-JfW", "Jammed from Within");
        out.put("A*-Nervewrecker", "Nerve-Wrecker");
        out.put("B1-Over the Fence", "Over the Fence");
        out.put("B1-RoD", "Road of Death");
        out.put("B1-SaaS", "Something about a Star");
        out.put("B1-Star", "B1 Star");
        out.put("B1-Third Wheel", "Third Wheel");
        out.put("B1-WtaD", "Window through a Door");
        out.put("B2-Higher Ground", "Higher Ground");
        out.put("B2-Moonshot", "Moonshot");
        out.put("B2-MotM", "Man on the Moon");
        out.put("B2-Star", "B2 Star");
        out.put("B2-Tomb", "Tomb");
        out.put("B3-Blown Away", "Blown Away");
        out.put("B3-Eagle's Nest", "Eagle's Nest");
        out.put("B3-Star", "B3 Star");
        out.put("B3-Sunshot", "Sunshot");
        out.put("B3-Woosh", "Whoosh");
        out.put("B4-ABUH", "A Box Up High");
        out.put("B4-Double-Plate", "Double Plate");
        out.put("B4-RPS", "Redundant Power Supply");
        out.put("B4-Self Help", "Self-Help Tutorial");
        out.put("B4-Sphinx Star", "Sphinx Star");
        out.put("B4-TRA Star", "The Right Angle Star");
        out.put("B4-TRA", "The Right Angle");
        out.put("B4-WAtC", "Wrap Around the Corner");
        out.put("B5-Chambers", "The Four Chambers of Flying");
        out.put("B5-Iron Curtain", "Behind the Iron Curtain");
        out.put("B5-Obelisk Star", "Obelisk Star");
        out.put("B5-Plates", "Alley of the Pressure Plates");
        out.put("B5-SES", "Slightly Elevated Sigil");
        out.put("B5-Two Jammers", "Me, Myself and Our Two Jammers");
        out.put("B6-Crisscross", "Crisscross Conundrum");
        out.put("B6-Egyptian Arcade", "Egyptian Arcade");
        out.put("B6-JDaW", "Just Doors and Windows");
        out.put("B7-AFaF", "A Fan across Forever");
        out.put("B7-BLoM", "Big Lump of Mine");
        out.put("B7-BSbS Star", "Bouncing Side by Side Star");
        out.put("B7-BSbS", "Bouncing Side by Side");
        out.put("B7-Star", "Pyramid Star");
        out.put("B7-WLJ", "Whole Lotta Jamming");
        out.put("B*-Cat's Cradle", "Cat's Cradle");
        out.put("B*-Merry Go Round", "Merry Go 'Round");
        out.put("B*-Peekaboo", "Peekaboo!");
        out.put("C1-Blowback", "Blowback");
        out.put("C1-Conservatory", "The Conservatory");
        out.put("C1-Labyrinth", "Labyrinth");
        out.put("C1-MIA", "Multiply Impossible Ascension");
        out.put("C1-Star", "C1 Star");
        out.put("C2-ADaaF", "A Ditch and A Fence");
        out.put("C2-Cemetery", "Cemetery");
        out.put("C2-Rapunzel", "Rapunzel");
        out.put("C2-Short Wall", "The Short Wall");
        out.put("C2-Star", "C2 Star");
        out.put("C3-BSLS", "Big Stairs, Little Stairs");
        out.put("C3-Jammer Quarantine", "Jammer Quarantine");
        out.put("C3-Star", "C3 Star");
        out.put("C3-Three Connectors", "Three Little Connectors... and a Fan");
        out.put("C3-Weathertop", "Weathertop");
        out.put("C4-Armory", "Armory");
        out.put("C4-Oubliette Star", "Oubliette Star");
        out.put("C4-Oubliette", "Oubliette");
        out.put("C4-Stables", "Stables");
        out.put("C4-Throne Room Star", "Throne Room Star");
        out.put("C4-Throne Room", "Throne Room");
        out.put("C5-Dumbwaiter Star", "Dumbwaiter Star");
        out.put("C5-Dumbwaiter", "Dumbwaiter");
        out.put("C5-Time Crawls", "Time Crawls");
        out.put("C5-Time Flies Star", "Time Flies Star");
        out.put("C5-Time Flies", "Time Flies");
        out.put("C5-UCAJ Star", "Up Close and Jammed Star");
        out.put("C5-UCaJ", "Up Close and Jammed");
        out.put("C6-Circumlocution", "Circumlocution");
        out.put("C6-Seven Doors", "The Seven Doors of Recording");
        out.put("C6-Star", "C6 Star");
        out.put("C6-Two Way Street", "Two Way Street");
        out.put("C7-Carrier Pigeons", "Carrier Pigeons");
        out.put("C7-Crisscross", "Crisscross Conundrum Advanced");
        out.put("C7-DMS", "Dead Man's Switch");
        out.put("C7-Prison Break", "Prison Break");
        out.put("C7-Star", "C7 Star");
        out.put("C*-Cobweb", "Cobweb");
        out.put("C*-Nexus", "Nexus");
        out.put("C*-Unreachable Garden", "Unreachable Garden");
        out.put("CM-Star", "C Messenger Star");
        out.put("F0-Star", "Floor 0 Star");
        out.put("F3-Star", "Floor 3 Star");
        return out;
    }

    private static HashMap<String, String> markerToName = createMarkerToName();

    private static String worldToName(World w) {
        switch (w) {
            case A8:
                return "A Star";
            case ADEVISLAND:
                return "Developer Island";
            case B8:
                return "B Star";
            case C8:
                return "C Star";
            case CMESSENGER:
                return "C Messenger";
            default:
                return w.toString();
        }
    }

    private static String getSigil(TalosProgress progress, String marker) {
        return TalosProgress.TETROS[progress.getVar(marker)].substring(0, 2);
    }

    private static HashMap<World, String[]> createWorldToMarkers() {
        HashMap<World, String[]> out = new HashMap<World, String[]>();
        out.put(World.A1, new String[]{"A1-PaSL", "A1-Beaten Path", "A1-Outnumbered", "A1-OtToU", "A1-ASooR", "A1-Trio", "A1-Peephole", "A1-Star"});
        out.put(World.A2, new String[]{"A2-Guards", "A2-Hall of Windows", "A2-Suicide Mission", "A2-Star"});
        out.put(World.A3, new String[]{"A3-Stashed for Later", "A3-ABTU", "A3-ABTU Star", "A3-Swallowed the Key", "A3-AEP", "A3-Clock Star"});
        out.put(World.A4, new String[]{"A4-Branch it Out", "A4-Above All That", "A4-Push it Further", "A4-Star", "A4-DCtS"});
        out.put(World.A5, new String[]{"A5-Two Boxes", "A5-Two Boxes Star", "A5-YKYMCTS", "A5-Over the Fence", "A5-OLB", "A5-FC", "A5-FC Star"});
        out.put(World.A6, new String[]{"A6-Mobile Mindfield", "A6-Deception", "A6-Door too Far", "A6-Bichromatic", "A6-Star"});
        out.put(World.A7, new String[]{"A7-LFI", "A7-Trapped Inside", "A7-Two Buzzers", "A7-Star", "A7-WiaL", "A7-Pinhole"});
        out.put(World.A8, new String[]{"A*-JfW", "A*-Nervewrecker", "A*-DDM"});
        out.put(World.ADEVISLAND, new String[0]);
        out.put(World.B1, new String[]{"B1-WtaD", "B1-Third Wheel", "B1-Over the Fence", "B1-RoD", "B1-SaaS", "B1-Star"});
        out.put(World.B2, new String[]{"B2-Tomb", "B2-Star", "B2-MotM", "B2-Moonshot", "B2-Higher Ground"});
        out.put(World.B3, new String[]{"B3-Blown Away", "B3-Star", "B3-Sunshot", "B3-Eagle's Nest", "B3-Woosh"});
        out.put(World.B4, new String[]{"B4-Self Help", "B4-Double-Plate", "B4-TRA", "B4-TRA Star", "B4-RPS", "B4-ABUH", "B4-WAtC", "B4-Sphinx Star"});
        out.put(World.B5, new String[]{"B5-SES", "B5-Plates", "B5-Two Jammers", "B5-Iron Curtain", "B5-Chambers", "B5-Obelisk Star"});
        out.put(World.B6, new String[]{"B6-Crisscross", "B6-JDaW", "B6-Egyptian Arcade"});
        out.put(World.B7, new String[]{"B7-AFaF", "B7-WLJ", "B7-BSbS", "B7-BSbS Star", "B7-BLoM", "B7-Star"});
        out.put(World.B8, new String[]{"B*-Merry Go Round", "B*-Cat's Cradle", "B*-Peekaboo"});
        out.put(World.C1, new String[]{"C1-Conservatory", "C1-MIA", "C1-Labyrinth", "C1-Blowback", "C1-Star",});
        out.put(World.C2, new String[]{"C2-ADaaF", "C2-Star", "C2-Rapunzel", "C2-Cemetery", "C2-Short Wall",});
        out.put(World.C3, new String[]{"C3-Three Connectors", "C3-Jammer Quarantine", "C3-BSLS", "C3-Weathertop", "C3-Star",});
        out.put(World.C4, new String[]{"C4-Armory", "C4-Oubliette", "C4-Oubliette Star", "C4-Stables", "C4-Throne Room", "C4-Throne Room Star",});
        out.put(World.C5, new String[]{"C5-Time Flies", "C5-Time Flies Star", "C5-Time Crawls", "C5-Dumbwaiter", "C5-Dumbwaiter Star", "C5-UCaJ", "C5-UCAJ Star",});
        out.put(World.C6, new String[]{"C6-Seven Doors", "C6-Star", "C6-Circumlocution", "C6-Two Way Street",});
        out.put(World.C7, new String[]{"C7-Carrier Pigeons", "C7-DMS", "C7-Star", "C7-Prison Break", "C7-Crisscross",});
        out.put(World.C8, new String[]{"C*-Unreachable Garden", "C*-Nexus", "C*-Cobweb"});
        out.put(World.CMESSENGER, new String[]{"CM-Star"});
        return out;
    }

    private static HashMap<World, String[]> worldToMarkers = createWorldToMarkers();

    private ArrayList<String> getWorldMarkers(World w) {
        ArrayList<String> out = new ArrayList<String>(Arrays.asList(worldToMarkers.get(w)));

        int F0 = (progress.getVar("Randomizer_LoopF0") + 1) / 2;
        int F3 = (progress.getVar("Randomizer_LoopF3") + 1) / 2;
        int worldIndex = Arrays.asList(World.values()).indexOf(w) + 1;
        // The extra sigil numbering skips dev island
        if (worldIndex > 8) {
            worldIndex--;
        }

        if (F0 == worldIndex) {
            out.add("F0-Star");
        }
        if (F3 == worldIndex) {
            out.add("F3-Star");
        }

        return removeHiddenScavengerSigils(out);
    }

    // Removes the markers for sigils that would normally be hidden because of scavenger
    private ArrayList<String> removeHiddenScavengerSigils(ArrayList<String> markers) {
        if (progress.getVar("Randomizer_Scavenger") != 0) {
            ScavengerEnding ending = ScavengerEnding.fromInt(progress.getVar("Randomizer_ScavengerMode") - 1);
            ArrayList<String> allowedSigils = ending.getAllowedSigils();
            for (String marker : (ArrayList<String>) markers.clone()) {
                String sigil = TalosProgress.TETROS[progress.getVar(marker)];
                if (!allowedSigils.contains(sigil)) {
                    markers.remove(marker);
                }
            }
        }
        return markers;
    }

    private World[] getWorldOrder() {
        /*
          Worlds (1-Indexed):
            A1, A2, A3, A4, A5, A6, A7, A8, ADevIsland,
            B1, B2, B3, B4, B5, B6, B7, B8,
            C1, C2, C3, C4, C5, C6, C7, C8, CMessenger,
        */
        World[] out = new World[26];
        if (progress.getVar("Randomizer_Portals") == 1 && progress.getVar("Randomizer_Loop") != 0) {
            // If A1 = 23, the portal out of A1 leads to world 23 -> C6
            out[0] = World.A1;
            int index = 0;
            while (index < 25) {
                int worldIndex = progress.getVar(out[index].toString());
                index++;
                out[index] = World.fromInt(worldIndex - 1);
            }
        } else if (progress.getVar("Randomizer_Portals") == 1) {
            // If A1 = 23, A1 is in spot 23 -> C6
            for (World w : World.values()) {
                int worldIndex = progress.getVar(w.toString());
                out[worldIndex - 1] = w;
            }
        } else {
            return World.values();
        }
        return out;
    }

    private static String[] paint3Order = new String[]{
        "A", "E", "F", "D", "E", "F", "D", "E",
        "I", "A", "B", "C", "A", "B", "C", "A",
        "E", "I", "G", "H", "I", "G", "H", "I"
    };

    private static String paint3(int seed) {
        return paint3Order[seed % 24];
    }

    private static String paint4(int seed) {
        return Integer.toString((seed % 8 > 4) ? (seed - 4) % 8 : seed % 8);
    }

    private static int[] paint5aOrder = new int[]{
        3, 4, 3, 2, 1, 0, 4, 3,
        1, 2, 1, 0, 4, 3, 2, 1,
        4, 0, 4, 3, 2, 1, 0, 4,
        2, 3, 2, 1, 0, 4, 3, 2,
        0, 1, 0, 4, 3, 2, 1, 0
    };
    private static String[] paint5aNames = new String[]{
        "Behind Amphitheatre / Far C",
        "Suicide Mission / 20%",
        "Amphitheatre / Tower",
        "A2 Star / B",
        "Spawn / A/C"
    };

    private static String paint5a(int seed) {
        return paint5aNames[paint5aOrder[seed % 40]];
    }

    private static int[] paint5bOrder = new int[]{
        4, 0, 0, 0, 0, 0, 0, 0,
        2, 3, 3, 3, 3, 3, 3, 3,
        0, 1, 1, 1, 1, 1, 1, 1,
        3, 4, 4, 4, 4, 4, 4, 4,
        1, 2, 2, 2, 2, 2, 2, 2,
    };
    private static String[] paint5bNames = new String[]{
        "LMUStK / JfW/Nervewrecker",
        "ABTU/SfL / Right of Spawn",
        "AEP / JfW/Middle",
        "LMUStK/SfL / Middle",
        "SfL / Left of Spawn"
    };

    private static String paint5b(int seed) {
        return paint5bNames[paint5bOrder[seed % 40]];
    }

    private static String[] allMarkers = new String[]{
        "A1-ASooR", "A1-Beaten Path", "A1-OtToU", "A1-Outnumbered",
        "A1-PaSL", "A1-Peephole", "A1-Star", "A1-Trio",
        "A2-Guards", "A2-Hall of Windows", "A2-Star", "A2-Suicide Mission",
        "A3-ABTU Star", "A3-ABTU", "A3-AEP", "A3-Clock Star",
        "A3-Stashed for Later", "A3-Swallowed the Key", "A4-Above All That", "A4-Branch it Out",
        "A4-DCtS", "A4-Push it Further", "A4-Star", "A5-FC Star",
        "A5-FC", "A5-OLB", "A5-Over the Fence", "A5-Two Boxes Star",
        "A5-Two Boxes", "A5-YKYMCTS", "A6-Bichromatic", "A6-Deception",
        "A6-Door too Far", "A6-Mobile Mindfield", "A6-Star", "A7-LFI",
        "A7-Pinhole", "A7-Star", "A7-Trapped Inside", "A7-Two Buzzers",
        "A7-WiaL", "A*-DDM", "A*-JfW", "A*-Nervewrecker",
        "B1-Over the Fence", "B1-RoD", "B1-SaaS", "B1-Star",
        "B1-Third Wheel", "B1-WtaD", "B2-Higher Ground", "B2-Moonshot",
        "B2-MotM", "B2-Star", "B2-Tomb", "B3-Blown Away",
        "B3-Eagle's Nest", "B3-Star", "B3-Sunshot", "B3-Woosh",
        "B4-ABUH", "B4-Double-Plate", "B4-RPS", "B4-Self Help",
        "B4-Sphinx Star", "B4-TRA Star", "B4-TRA", "B4-WAtC",
        "B5-Chambers", "B5-Iron Curtain", "B5-Obelisk Star", "B5-Plates",
        "B5-SES", "B5-Two Jammers", "B6-Crisscross", "B6-Egyptian Arcade",
        "B6-JDaW", "B7-AFaF", "B7-BLoM", "B7-BSbS Star",
        "B7-BSbS", "B7-Star", "B7-WLJ", "B*-Cat's Cradle",
        "B*-Merry Go Round", "B*-Peekaboo", "C1-Blowback", "C1-Conservatory",
        "C1-Labyrinth", "C1-MIA", "C1-Star", "C2-ADaaF",
        "C2-Cemetery", "C2-Rapunzel", "C2-Short Wall", "C2-Star",
        "C3-BSLS", "C3-Jammer Quarantine", "C3-Star", "C3-Three Connectors",
        "C3-Weathertop", "C4-Armory", "C4-Oubliette Star", "C4-Oubliette",
        "C4-Stables", "C4-Throne Room Star", "C4-Throne Room", "C5-Dumbwaiter Star",
        "C5-Dumbwaiter", "C5-Time Crawls", "C5-Time Flies Star", "C5-Time Flies",
        "C5-UCAJ Star", "C5-UCaJ", "C6-Circumlocution", "C6-Seven Doors",
        "C6-Star", "C6-Two Way Street", "C7-Carrier Pigeons", "C7-Crisscross",
        "C7-DMS", "C7-Prison Break", "C7-Star", "C*-Cobweb",
        "C*-Nexus", "C*-Unreachable Garden", "CM-Star", "F0-Star",
        "F3-Star"
    };
}
