package apple1417.randomizer_to_spreadsheet;

public enum CustomCellStyles {
    DEFAULT, HEADER, HEADER_BORDERLESS,
    STAR, GREEN, GREY, YELLOW, RED,
    HEADER_STAR, HEADER_GREEN, HEADER_GREY, HEADER_YELLOW, HEADER_RED;

    public static CustomCellStyles sigilToStyle(String sigil, boolean header) {
        switch (sigil.charAt(0)) {
            case '*':
                return header ? HEADER_STAR : STAR;
            case 'D':
                return header ? HEADER_GREEN : GREEN;
            case 'E':
                return header ? HEADER_GREY : GREY;
            case 'M':
                return header ? HEADER_YELLOW : YELLOW;
            case 'N':
                return header ? HEADER_RED : RED;
            default:
                return DEFAULT;
        }
    }

    public static CustomCellStyles sigilToStyle(String sigil) {
        return sigilToStyle(sigil, false);
    }
}