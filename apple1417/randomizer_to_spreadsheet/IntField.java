package apple1417.randomizer_to_spreadsheet;

import javafx.geometry.Pos;
import javafx.scene.control.TextField;
import javafx.scene.control.TextFormatter;
import javafx.util.converter.IntegerStringConverter;

import java.util.function.UnaryOperator;

public class IntField extends TextField {
    private boolean initialSelect = false;

    public IntField(int value) {
        super();
        // Defines the allowed formatting
        UnaryOperator<TextFormatter.Change> integerFilter = (change) -> {
            // We only care about text changes
            String newText = change.getControlNewText();
            if (change.getControlText() == newText) {
                return change;
            }

            // Quick fix so that the value isn't initially selected
            if (!initialSelect) {
                initialSelect = true;
                change.setCaretPosition(change.getAnchor());
                return change;
            }

            // Remove non-numeric chars
            newText = newText.replaceAll("\\D", "");
            int oldLen = change.getText().length();
            change.setText(change.getText().replaceAll("\\D", ""));
            int newLen = change.getText().length();

            // Valid number
            if (newText.matches("(0|[1-9]\\d*)")) {
                // Make sure it fits an int
                try {
                    Integer.parseInt(newText);
                } catch (NumberFormatException e) {
                    return null;
                }
            }
            // Remove leading 0s
            if (newText.matches("0\\d+")) {
                newText = newText.substring(1);
                change.setRange(0, change.getControlText().length());
                change.setText(newText);
            }
            // Empty string
            if (newText.length() == 0) {
                change.setText("0");
            }

            return change;
        };

        setTextFormatter(new TextFormatter<Integer>(new IntegerStringConverter(), value, integerFilter));
        setAlignment(Pos.BOTTOM_RIGHT);
    }

    public IntField() {
        this(0);
    }

    public int getValue() {
        return Integer.parseInt(getText());
    }

    public void setValue(int val) {
        setText(Integer.toString(val));
    }
}
