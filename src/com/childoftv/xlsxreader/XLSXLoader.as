/*
MIT LICENSE:
http://www.opensource.org/licenses/mit-license.php
Copyright (c) 2011 Ben Morrow
Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
*/
package com.childoftv.xlsxreader
{
	
	
	
	
	import flash.events.Event;
	import flash.events.EventDispatcher;
	import flash.events.IOErrorEvent;
	import flash.net.URLRequest;
	import flash.utils.ByteArray;
	import flash.utils.Dictionary;
	
	import deng.fzip.FZip;
	import deng.fzip.FZipFile;
	
	
	[Event(name="complete",type="flash.events.Event")]
	
	/**
	 * A class to load a Microsoft Excel 2007+ .XLSX Spreadsheet (described here: http://en.wikipedia.org/wiki/Office_Open_XML)
	 * 
	 * 
	 * 
	 * @example Loading an Excel file and reading a cell:
	 * <listing version="3.0">
	 * 
	 * //Create the Excel Loader
	 * var excel_loader:XLSXLoader=new XLSXLoader();
	 * 
	 * //Listen for when the file is loaded
	 * excel_loader.addEventListener(Event.COMPLETE,function (e:Event) {
	 * 
	 *   //Access a worksheet by name ('Sheet1')
	 *   var sheet_1:Worksheet=excel_loader.worksheet("Sheet1");    
	 *   
	 *   //Access a cell in sheet 1 and output to trace
	 *   trace("Cell A3="+sheet_1.getCellValue("A3")) //outputs: Cell A3=Hello World;
	 * 
	 * 
	 * });
	 * 
	 * //Load the file
	 *excel_loader.load("Example Spreadsheet.xlsx");
	 *
	 * </listing>
	 *
	 */
	public class XLSXLoader extends EventDispatcher
	{
		private var zipProcessor:FZip =new FZip();
		
		public static var openXMLNS:Namespace=new Namespace("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		
		private var file:String="none";
		private var sharedStringsCache:XML;
		private var manifestCache:XML;
		private var worksheetCache:Dictionary=new Dictionary();
		
		default xml namespace=openXMLNS;
		/**
		 * Creates an XLSXLoader which can be used to load a spreadsheet
		 *
		 */
		public function XLSXLoader()
		{
			
			zipProcessor.addEventListener("complete",completed);
			zipProcessor.addEventListener("ioError",function(e:IOErrorEvent):void{trace("ZIP Processor IO error:");trace(e)});
			
		}
		/**
		 * Returns the openxml namespace as a Namespace object
		 * 
		 * @return the openxml namespace as a Namespace object
		 * 
		 */     
		public function getNamespace():Namespace
		{
			return openXMLNS;
		}
		
		/**
		 * Load a spreadsheet from a valid file path or URI
		 *
		 * @param sFile A String specifying a valid file or URI.
		 *
		 */
		public function load(sFile:String):void
		{
			file=sFile;
			addEventHandlers();
			
			trace("Attempting to load "+sFile );
			zipProcessor.load(new URLRequest(sFile));
			
		}
		
		/**
		 * Load a spreadsheet from a valid file path or URI
		 *
		 * @param sFile A String specifying a valid file or URI.
		 *
		 */
		public function loadFromByteArray(bytes:ByteArray, fileSrc:String = "From Byte Array"):void
		{
			file=fileSrc;
			addEventHandlers();
			
			trace("Attempting to load bytes ("+fileSrc+") size:"+bytes.length );
			zipProcessor.loadBytes(bytes);
			
		}
		
		
		internal function completed(e:Event):void
		{
			trace("'"+file+"' unzipped and loaded: " +zipProcessor.getFileCount()+ ' files are inside the xlsx');
		}
		
		/**
		 * Gets the named worksheet from the loaded spreadsheet as a new com.childoftv.xlsxreader.Worksheet Object
		 * 
		 * @param wName a valid worksheet name within the loaded spreadsheet
		 * @return a Worksheet object
		 * 
		 */     
		public function worksheet(wName:String):Worksheet
		{
			
			default xml namespace=openXMLNS;
			if (! manifestCache)
			{
				manifestCache=retrieveXML("xl/workbook.xml");
			}
			var ret:Worksheet;
			var val:*=manifestCache..sheet.(@name==wName);
			var index:Number=val.childIndex();
			
			ret= worksheetbyId(index+1,wName);
			return ret;
		}
		
		/**
		 * Tests whether the provided name is the name of a worksheet in the loaded spreadsheet.
		 * 
		 * @param wName String with the name of a worksheet
		 * @return Returns true
		 * 
		 */     
		public function isSheetName(wName:String):Boolean
		{
			default xml namespace=openXMLNS;
			if (! manifestCache)
			{
				manifestCache=retrieveXML("xl/workbook.xml");
			}
			
			return Boolean(manifestCache.sheets.sheet.(@name==wName).length() > 0);
		}
		
		/**
		 * returns names of sheets in xlsx
		 *
		 * @return Returns Vector.<String> sheet names
		 */
		public function getSheetNames():Vector.<String>
		{
			default xml namespace=openXMLNS;
			if (! manifestCache)
			{
				manifestCache=retrieveXML("xl/workbook.xml");
			}
			
			var sheetNames:Vector.<String> = new Vector.<String>();
			for each(var sheetName:String in manifestCache.sheets.sheet.@name)
			sheetNames.push(sheetName);
			
			return sheetNames;
		}
		
		private function worksheetbyId(id:Number,wName:String="name not available"):Worksheet
		{
			
			if (! worksheetCache[id])
			{
				worksheetCache[id]=new Worksheet(wName,this,retrieveXML("xl/worksheets/sheet"+id+".xml"));
			}
			return worksheetCache[id];
			
		}
		/**
		 * @private  
		 * 
		 * Looks up the internal shared string database XML
		 * 
		 */     
		internal function sharedStrings():XML
		{
			if (! sharedStringsCache)
			{
				sharedStringsCache=retrieveXML("xl/sharedStrings.xml");
			}
			return sharedStringsCache;
		}
		
		/**
		 * @private  
		 * 
		 *Retrieves a specific shared string
		 * 
		 */ 
		internal function sharedString(index:String):String
		{
			default xml namespace=openXMLNS;
			if (index==""||! index)
			{
				return "";
			}
			else
			{
				
				
				return sharedStrings().child(index).t.toString();;
			}
		}
		private function retrieveXML(path:String):XML
		{
			
			var file:FZipFile=zipProcessor.getFileByName(path);
			return convertToOpenXMLNS(file.getContentAsString(false));
			
		}
		
		
		private function convertToOpenXMLNS(s:String):XML
		{
			
			XML.ignoreProcessingInstructions=true;
			var XMLDoc:XML=XML(s);
			XMLDoc.normalize();
			return XMLDoc;
		}
		
		/**
		 * @private  
		 */ 
		protected function defaultHandler(evt:Event):void
		{
			dispatchEvent(evt.clone());
		}
		/**
		 * @private  
		 */ 
		protected function defaultErrorHandler(evt:Event):void
		{
			trace(evt);
			close();
			dispatchEvent(evt.clone());
		}
		/**
		 * @private  
		 */ 
		protected function addEventHandlers():void
		{
			
			zipProcessor.addEventListener(Event.COMPLETE, defaultHandler);
		}
		
		/**
		 * @private  
		 */ 
		protected function removeEventHandlers():void
		{
			zipProcessor.removeEventListener(Event.COMPLETE, defaultHandler);
			
		}
		/**
		 * Closes the open xlsx file and frees the available memory 
		 * 
		 */         
		public function close():void
		{
			if (zipProcessor)
			{
				
				removeEventHandlers();
				zipProcessor.close();
				zipProcessor=null();
				manifestCache=null;
				worksheetCache=null;
			}
		}
		
		
	}
}