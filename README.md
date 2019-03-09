# RandomizerToSpreadsheet
Makes spreadsheets (for routing purposes) from the settings for the [Talos Principle Randomzier](https://github.com/apple1417/Talos-Sigil-Randomizer).

[Example](https://drive.google.com/open?id=1h5kj1amMigCUU4DEzXofW4UxEiZiGDla)

### Usage
Running the file as any normal executable will give you a simple gui to input various options and select a file to output to.

When running through the command line you have the following options:
```
java -jar RandomizerToSpreadsheet.jar [-ghijopV] [-b=<mobius>] [-c=<scav>]
                                      [-d=<moody>] [-m=<mode>] [-s=<seed>] [FILE]
      [FILE]               Save location
  -b, --mobius=<mobius>    Between 0 and 63 inclusive
  -c, --scavenger=<scav>   One of: [OFF, SHORT, FULL]
  -d, --moody=<moody>      One of: [OFF, COLOUR, SHAPE, BOTH]
  -g, --signs
  -h, --help               Show this help message and exit.
  -i, --items
  -j, --jetpack
  -m, --mode=<mode>        One of: [NONE, DEFAULT, SIXTY, FULLY_RANDOM, INTENDED,
                             HARDMODE, SIXTY_HARDMODE]
  -o, --open
  -p, --portals
  -s, --seed=<seed>        Between 0 and 2147483647 inclusive
  -V, --version            Print version information and exit.
  -x, --extra
```
By not specifying a file you will be launched into the gui version instead.

### Compiling manually

This project makes use of three libraries you'll need to include if compiling manually:

 * [Apache POI](https://poi.apache.org/) (3.17)
 * [Picocli](https://github.com/remkop/picocli) (3.6.1)
 * [Talos Randomizer](https://github.com/apple1417/random_stuff/tree/master/apple1417/randomizer) (Latest)
 
Note that the Talos Randomizer "library" isn't a standard library in a jar, you want to download all the files in the
 linked directory and put them in the same overall directory structure (src/apple1417/randomizer), then either make your
 own jar and add it to your classpath or just add each of the compiled files separately.