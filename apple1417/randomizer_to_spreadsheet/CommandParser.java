package apple1417.randomizer_to_spreadsheet;

import apple1417.randomizer.Enums.MoodySigils;
import apple1417.randomizer.Enums.RandomizerMode;
import apple1417.randomizer.Enums.ScavengerMode;
import apple1417.randomizer.TalosProgress;
import picocli.CommandLine;
import picocli.CommandLine.*;

import java.io.File;
import java.util.Arrays;

@Command(name = "RandomizerToSpreadsheet", mixinStandardHelpOptions = true, version = "1.2.0-v11.2.0")
public class CommandParser {
    @Option(names = {"-s", "--seed"}, description = "Between 0 and 2147483647 inclusive")
    private static int seed = 0;
    @Option(names = {"-m", "--mode"}, description = "One of: [NONE, DEFAULT, SIXTY, FULLY_RANDOM, INTENDED, HARDMODE, SIXTY_HARDMODE]")
    private static RandomizerMode mode = RandomizerMode.DEFAULT;
    @Option(names = {"-c", "--scavenger"}, description = "One of: [OFF, SHORT, FULL]")
    private static ScavengerMode scav = ScavengerMode.OFF;
    @Option(names = {"-b", "--mobius"}, description = "Between 0 and 63 inclusive")
    private static int mobius = 0;
    @Option(names = {"-d", "--moody"}, description = "One of: [OFF, COLOUR, SHAPE, BOTH]")
    private static MoodySigils moody = MoodySigils.OFF;
    @Option(names = {"-o", "--open"})
    private static boolean open = false;
    @Option(names = {"-i", "--items"})
    private static boolean items = false;
    @Option(names = {"-g", "--signs"})
    private static boolean signs = true;
    @Option(names = {"-p", "--portals"})
    private static boolean portals = false;

    // Use a File[] here so that this can continue without specifying a file
    @Parameters(paramLabel = "FILE", index = "0", description = "Save location")
    private static File[] file;

    public static void main(String[] args) {
        /*
          There's a reason I don't make this implement runnable to remove all the error checking
          While the input parsing for enums seems to use Enum.name(), if you input an invalid value
           if will print the error using Enum.toString()
          I have overwritten toString on all of the ones I'm using here, so I do all this to be able
           to intercept the error message and replace it with the correct one
          This also lets me add a new error for when seed/mobius values are out of bounds
        */
        CommandParser cp = new CommandParser();
        CommandLine cmd = new CommandLine(cp);
        try {
            cmd.parse(args);
            if (cmd.isUsageHelpRequested()) {
                cmd.usage(System.out);
            } else if (cmd.isVersionHelpRequested()) {
                cmd.printVersionHelp(System.out);
            } else if (!(0 <= seed && seed <= 2147483647)) {
                System.err.println("Invalid value for option '--seed': expected value between 0 and 2147483647 inclusive");
                cmd.usage(System.err);
            } else if (!(0 <= mobius && mobius <= 63)) {
                System.err.println("Invalid value for option '--mobius': expected value between 0 and 63 inclusive");
            } else {
                cp.start();
            }
        } catch (ParameterException ex) {

            String message = ex.getMessage();
            message = message.replaceAll("\\[No Randomization, Default, 60fps Friendly, Fully Random, Intended, Hardmode, 60fps Hardmode\\]",
                "[NONE, DEFAULT, SIXTY, FULLY_RANDOM, INTENDED, HARDMODE, SIXTY_HARDMODE]");
            message = message.replaceAll("\\[Off, Short, Full\\]",
                "[OFF, SHORT, FULL]");
            message = message.replaceAll("\\[Off, Colour, Shape, Colour + Shape\\]",
                "[OFF, COLOUR, SHAPE, BOTH]");

            System.err.println(message);
            if (!UnmatchedArgumentException.printSuggestions(ex, System.err)) {
                cmd.usage(System.err);
            }
        } catch (Exception ex) {
            throw new ExecutionException(cmd, "Error while calling " + cp, ex);
        }
    }

    public void start() {
        if (file == null) {
            MainWindow.start();
            return;
        }

        TalosProgress options = new TalosProgress();
        options.setVar("Randomizer_Seed", seed);
        options.setVar("Randomizer_Mode", Arrays.asList(RandomizerMode.values()).indexOf(mode));
        options.setVar("Randomizer_Scavenger", Arrays.asList(ScavengerMode.values()).indexOf(scav));
        options.setVar("Randomizer_Moody", Arrays.asList(MoodySigils.values()).indexOf(moody));
        options.setVar("Randomizer_Loop", mobius);

        if (open) {
            options.setVar("Randomizer_NoGates", 1);
        }
        if (items) {
            options.setVar("Randomizer_UnlockItems", 1);
        }
        if (signs) {
            options.setVar("Randomizer_ShowAll", 1);
        }
        if (portals) {
            options.setVar("Randomizer_Portals", 1);
        }

        SpreadsheetCreator.createSpreadsheet(file[0], options);
    }
}
