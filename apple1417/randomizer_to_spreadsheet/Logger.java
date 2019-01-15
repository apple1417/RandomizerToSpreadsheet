package apple1417.randomizer_to_spreadsheet;

import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.ButtonType;

public class Logger {
    private static boolean usesGUI = false;

    public static void setGUI() {
        usesGUI = true;
    }

    public static void error(String str) {
        if (usesGUI) {
            Alert alert = new Alert(AlertType.NONE, str, ButtonType.OK);
            alert.setTitle("Error");
            alert.showAndWait();
        } else {
            System.err.println(str);
        }
    }

    public static void error(Exception e) {
        error(e.toString());
    }
}
