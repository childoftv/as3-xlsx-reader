#AS3 (Actionscript 3) XLSX READER

A reader for excel files in Flash, Flex and Air

<i>Copyright (c) 2011 Ben Morrow</i>



##Usage

###Easy

Just add the [swc](bin/as3-xlsx-reader.swc) to your library path

###Advanced (for modifying source/building your own)

1. Clone master e.g. `git clone git@github.com:childoftv/as3-xlsx-reader.git`
2. `cd master`
3. Fetch Fzip using `git submodule foreach git pull origin master`
4. Now open the projects in the [projects](projects) directories.

[as3-xlsx-reader](projects/ass-xlsx-reader) is a library project which can be used to build the core (outputs to [bin](bin) folder)

[as3-xlsx-reader-example](projects/ass-xlsx-reader-example) is an adobe air project which can build a quick example


##Example

Look at [LoadXLSXExample.as](projects/as3-xlsx-reader-example/src/LoadXLSXExample.as) for an example that can be used from flash, flex or with adobe air.


##LICENSE:

Released under [MIT LICENSE](LICENSE)
