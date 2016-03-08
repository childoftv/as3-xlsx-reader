#AS3 (Actionscript 3) XLSX READER

A reader for excel files in Flash, Flex and Air

<i>Copyright (c) 2011 Ben Morrow</i>



##Usage

###Easy

Just add the [swc](bin/as3-xlsx-reader.swc) to your library path

###Advanced (for modifying source/building your own)

1. Clone master e.g. `git clone git@github.com:childoftv/as3-xlsx-reader.git`
2. `cd as3-xlsx-reader`
3. Fetch Fzip using `git submodule foreach git pull origin master`
4. Now open the projects in the [projects](projects) directories.

[as3-xlsx-reader](projects/ass-xlsx-reader) is a library project which can be used to build the core (outputs to [bin](bin) folder)

[as3-xlsx-reader-example](projects/ass-xlsx-reader-example) is an adobe air project which can build a quick example

####command line compilation using [Flex SDK](http://www.adobe.com/devnet/flex/flex-sdk-download.html)

1. `cd as3-xlsx-reader`
2. Build the library: `compc -include-sources=src -library-path=libs/fzip/bin/fzip.swc  -output bin/as3-xlsx-reader.swc``
3. Build the Example: `mxmlc -source-path=projects/as3-xlsx-reader-example/src/ -library-path=bin/as3-xlsx-reader.swc -static-link-runtime-shared-libraries=true -use-network=false -debug=true -output=projects/as3-xlsx-reader-example/bin-debug/LoadXLSXExample.swf  projects/as3-xlsx-reader-example/src/LoadXLSXExample.as`
4. Copy the Spreadsheet to the debug folder: `cp projects/as3-xlsx-reader-example/assets/*.xlsx projects/as3-xlsx-reader-example/bin-debug/`
5. run with fdb: `fdb projects/as3-xlsx-reader-example/bin-debug/LoadXLSXExample.swf`
6. type `continue` for output


##Example

Look at [LoadXLSXExample.as](projects/as3-xlsx-reader-example/src/LoadXLSXExample.as) for an example that can be used from flash, flex or with adobe air.


##LICENSE:

Released under [MIT LICENSE](LICENSE)
