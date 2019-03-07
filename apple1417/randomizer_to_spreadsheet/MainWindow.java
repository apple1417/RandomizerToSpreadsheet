package apple1417.randomizer_to_spreadsheet;

import apple1417.randomizer.Enums.MobiusOptions;
import apple1417.randomizer.Enums.MoodySigils;
import apple1417.randomizer.Enums.RandomizerMode;
import apple1417.randomizer.Enums.ScavengerMode;
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

import java.util.ArrayList;
import java.util.Arrays;


public class MainWindow extends Application {
    public static void start() {
        launch();
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
            new Text("Jetpack"),
            new Text("Random Extra Sigils")
        ));
        bottomRightLabels = new ArrayList<Text>();

        for (int i = 0; i < 6; i++) {
            CheckBox top = new CheckBox();
            CheckBox bottom = new CheckBox();

            topRightBoxes.add(top);
            bottomRightBoxes.add(bottom);

            bottomRightLabels.add(new Text(MobiusOptions.fromInt(i).toString()));

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

        // Setup the options locking that normally goes on
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
        exportButton.setOnAction((obs) -> {
            TalosProgress options = new TalosProgress();
            options.setVar("Randomizer_Seed", ((IntField) leftOptions.get(0)).getValue());
            options.setVar("Randomizer_Mode", Arrays.asList(RandomizerMode.values()).indexOf((((ChoiceBox) leftOptions.get(1)).getValue())));
            options.setVar("Randomizer_Scavenger", Arrays.asList(ScavengerMode.values()).indexOf((((ChoiceBox) leftOptions.get(2)).getValue())));
            options.setVar("Randomizer_Moody", Arrays.asList(MoodySigils.values()).indexOf((((ChoiceBox) leftOptions.get(3)).getValue())));

            // This one needs to be explicitly set to 0 if turned off
            options.setVar("Randomizer_ShowAll", topRightBoxes.get(2).isSelected() ? 1 : 0);

            int mobius = 0;
            for (int i = 0; i < bottomRightBoxes.size(); i++) {
                if (bottomRightBoxes.get(i).isSelected()) {
                    mobius |= MobiusOptions.fromInt(i).getMask();
                }
            }
            options.setVar("Randomizer_Loop", mobius);

            if (topRightBoxes.get(0).isSelected()) {
                options.setVar("Randomizer_NoGates", 1);
            }
            if (topRightBoxes.get(1).isSelected()) {
                options.setVar("Randomizer_UnlockItems", 1);
            }
            if (topRightBoxes.get(3).isSelected()) {
                options.setVar("Randomizer_Portals", 1);
            }
            if (topRightBoxes.get(4).isSelected()) {
                options.setVar("Randomizer_Jetpack", 1);
            }
            if (topRightBoxes.get(5).isSelected()) {
                options.setVar("Randomizer_ExtraSigils", 1);
            }

            SpreadsheetCreator.createSpreadsheet(chooser.showSaveDialog(primaryStage), options);
        });
        exportButton.setFont(new Font(18));
        exportButton.setMinWidth(150);
        exportButton.setMaxWidth(150);
        root.add(exportButton, 0, 4, 8, 2);
        root.setHalignment(exportButton, HPos.CENTER);
        root.setValignment(exportButton, VPos.CENTER);

        primaryStage.setScene(new Scene(root, 925, 165));
        primaryStage.setResizable(false);
        primaryStage.setTitle("Randomizer To Spreadsheet " + CommandParser.VERSION);
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
}
