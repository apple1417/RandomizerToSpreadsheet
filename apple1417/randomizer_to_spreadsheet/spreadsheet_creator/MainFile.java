package apple1417.randomizer_to_spreadsheet.spreadsheet_creator;

import apple1417.randomizer.Enums.*;
import apple1417.randomizer.GeneratorGeneric;
import apple1417.randomizer.TalosProgress;
import apple1417.randomizer_to_spreadsheet.Logger;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;

/*
  This class just starts creating the workbook and holds a few helper functions
  Other classes deal with creating individual sheets
*/
public class MainFile {
    public static void createSpreadsheet(File saveLoc, TalosProgress options) {
        TalosProgress progress;
        try {
            progress = new GeneratorGeneric(options).generate(options.getVar("Randomizer_Seed"));
        } catch (Exception e) {
            Logger.error("The selected options don't fully generate.\nThis is an issue with the Randomizer, not this program.");
            return;
        }

        if (saveLoc == null) {
            Logger.error("No file selected");
            return;
        }
        // Don't want to create the file until all other possible errors have been checked
        FileOutputStream outFile;
        try {
            outFile = new FileOutputStream(saveLoc);
        } catch (FileNotFoundException e) {
            Logger.error(e);
            return;
        }

        // Create the workbook
        XSSFWorkbook workbook = new XSSFWorkbook();
        CustomCellStyles.applyToWorkbook(workbook);
        // Add all of our sheets
        InfoSheet.createSheet(workbook.createSheet("Info"), progress);
        AllSigilsSheet.createSheet(workbook.createSheet("All Sigils"), progress);
        SigilsPerWorldSheet.createSheet(workbook.createSheet("Sigils Per World"), progress);
        LocationsBySigilSheet.createSheet(workbook.createSheet("Locations by Sigil Type"), progress);

        // Save the workbook
        try {
            workbook.write(outFile);
            outFile.close();
        } catch (IOException e) {
            Logger.error(e);
        }
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
    public static String getMarkerName(String marker) {
        return markerToName.get(marker);
    }

    public static String getWorldName(World w) {
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

    public static String getSigilType(TalosProgress progress, String marker) {
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
    public static ArrayList<String> getWorldMarkers(TalosProgress progress, World w) {
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

        return removeHiddenScavengerSigils(progress, out);
    }

    // Removes the markers for sigils that would normally be hidden because of scavenger
    public static ArrayList<String> removeHiddenScavengerSigils(TalosProgress progress, ArrayList<String> markers) {
        if (progress.getVar("Randomizer_Scavenger") != 0) {
            ScavengerEnding ending = ScavengerEnding.fromInt(progress.getVar("Randomizer_ScavengerMode") - 1);
            ArrayList<String> allowedSigils = ending.getAllowedSigils();
            for (int i = 0; i < markers.size(); i++) {
                String sigil = TalosProgress.TETROS[progress.getVar(markers.get(i))];
                if (!allowedSigils.contains(sigil)) {
                    markers.remove(i);
                    i--;
                }
            }
        }
        return markers;
    }

    public static World[] getWorldOrder(TalosProgress progress) {
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
    static String paint3(int seed) {
        return paint3Order[seed % 24];
    }

    static String paint4(int seed) {
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
    static String paint5a(int seed) {
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
    static String paint5b(int seed) {
        return paint5bNames[paint5bOrder[seed % 40]];
    }

    public static final String[] allMarkers = new String[]{
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
