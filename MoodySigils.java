public enum MoodySigils {
    OFF("Off"),
    COLOUR("Colour"),
    SHAPE("Shape"),
    BOTH("Colour + Shape");

    private String label;

    private MoodySigils(String label) {
        this.label = label;
    }

    public static MoodySigils fromInt(int i) {
        return MoodySigils.values()[i];
    }

    public String toString() {
        return label;
    }
}