# RandomizerToSpreadsheet
Makes spreadsheets (for routing purposes) from the settings for the [Talos Principle Randomzier](https://github.com/apple1417/Talos-Sigil-Randomizer).

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
```
By not specifying a file you will be launched into the gui version instead.