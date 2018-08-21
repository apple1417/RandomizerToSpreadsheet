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
import randomizer_bruteforce.Enums.RandomizerMode;
import randomizer_bruteforce.Enums.ScavengerMode;
import randomizer_bruteforce.GeneratorGeneric;
import randomizer_bruteforce.TalosProgress;

import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;


public class MainWindow extends Application {
    public static void main(String[] args) {
        launch(args);
    }

    private FileChooser chooser = new FileChooser();
    private ArrayList<Control> leftOptions = new ArrayList<Control>();
    private ArrayList<CheckBox> topRightBoxes = new ArrayList<CheckBox>();
    private ArrayList<CheckBox> bottomRightBoxes = new ArrayList<CheckBox>();

    @Override
    public void start(Stage primaryStage) {
        FileChooser.ExtensionFilter filter = new FileChooser.ExtensionFilter("Excel Spreadsheet (*.xlsx)", "*.xlsx");
        chooser.getExtensionFilters().add(filter);
        chooser.setSelectedExtensionFilter(filter);

        GridPane root = new GridPane();
        root.setHgap(10);
        root.setVgap(5);
        root.setPadding(new Insets(7.5, 7.5, 7.5, 7.5));

        /*================================*\
        |            NODE SETUP            |
        \*================================*/

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
        ArrayList<Text> topRightLabels = new ArrayList<>();
        topRightLabels.addAll(Arrays.asList(
                new Text("Open All Worlds"),
                new Text("Unlock All Items"),
                new Text("Show Full Signs"),
                new Text("Random Portals"),
                new Text("")
        ));
        ArrayList<Text> bottomRightLabels = new ArrayList<>();
        bottomRightLabels.addAll(Arrays.asList(
                new Text("All Sigils"),
                new Text("All of a Shape"),
                new Text("All of a Colour"),
                new Text("Eternalize Ending"),
                new Text("Two Tower Floors")
        ));
        ArrayList<Text> rightLabels = new ArrayList<Text>();
        rightLabels.addAll(topRightLabels);
        rightLabels.addAll(bottomRightLabels);

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

        scavenger.setOnAction((obs) -> {
            if (scavenger.getValue() == ScavengerMode.OFF) {
                topRightBoxes.get(0).setDisable(false);
                topRightBoxes.get(2).setDisable(false);
            } else {
                topRightBoxes.get(0).setDisable(true);
                topRightBoxes.get(2).setDisable(true);
                topRightBoxes.get(0).setSelected(true);
                topRightBoxes.get(2).setSelected(false);
            }
        });

        // The export button to actually save it to file
        Button exportButton = new Button("Export");
        exportButton.setOnAction((obs) -> saveFile(/*chooser.showSaveDialog(primaryStage)*/ null));
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

        for (Text t : rightLabels) {
            root.setHalignment(t, HPos.CENTER);
            root.setValignment(t, VPos.CENTER);
        }
    }

    private void saveFile(File saveLoc) {
        /*if (saveLoc == null) {
            return;
        }
        FileOutputStream out;
        try {
            out = new FileOutputStream(saveLoc);
        } catch (FileNotFoundException e) {
            return;
        }*/

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
        TalosProgress progress = new GeneratorGeneric(options).generate(seed);

        System.out.println(progress.getChecksum());


        /*
          TODO: Three sheets:
          TODO:  Sigil at each location
          TODO:  Sigils in each world
          TODO:  Locations for each sigil type
          TODO: Sort worlds by portals, A1 location gets sigils of whatever world's inside
        */


        /*
        XSSFWorkbook workbook = new XSSFWorkbook();
        try {
            workbook.write(out);
            out.close();
        } catch (IOException e) {}
        */
    }
}
