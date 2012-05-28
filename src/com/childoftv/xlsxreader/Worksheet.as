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
	/**
	 * A wrapper class for an individual XLSX Worksheet loaded through an XLSXLoader object
	 * Instances of this class are created by the XLSX.worksheet function.
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
	public class Worksheet
	{
		
		private var xml:XML;
		private var fileLink:XLSXLoader;
		private var ns:Namespace;
		private var sheetName:String;
		
		/**
		 * Creates a new Worksheet Object from a file loader. This consturctor is designed to be called from the XLSXLoader.worksheet() function only
		 * 
		 * @param sSheetName worksheet name
		 * @param FileLink link to the XLSX Loader
		 * @param input the worksheet as XML
		 * 
		 */     
		public function Worksheet(sSheetName:String,FileLink:XLSXLoader,input:XML)
		{
			sheetName=sSheetName;
			xml=input;
			fileLink=FileLink;
			ns=fileLink.getNamespace();
			default xml namespace=ns;
		}
		
		/**
		 * Returns an XML representation of the worksheet
		 * 
		 * @return the worksheet as XML
		 * 
		 */     
		public function toXML():XML
		{
			return xml;
		}
		
		/**
		 * Gets the XML representation of a single cell
		 * 
		 * @param cellRef a standard spreadsheet single cell reference (e.g. "A:3")
		 * @return the cell value as XML
		 * 
		 */ 
		public function getCell(cellRef:String):XMLList
		{
			cellRef=cellRef.toUpperCase();
			var row:Number=Number(cellRef.match(/[0-9]+/)[0]);
			var column:String=cellRef.match(/[A-Z]+/)[0];
			trace("getCell:"+cellRef, row, column);
			
			return getRows(column,row,row);
		}
		/**
		 * Gets the String value of a single cell
		 * 
		 * @param cellRef a standard spreadsheet single cell reference (e.g. "A:3")
		 * @return the cell value as a string
		 * 
		 */ 
		public function getCellValue(cellRef:String):String
		{
			var xml:XMLList=getCell(cellRef);
			if(xml.v.valueOf())
			{
				return xml.v.valueOf();
			}else{
				return null;
			}
		}
			
		private function getRawRows(column:String="A",from:Number=1,to:Number=-1):XMLList
		{
			// returns the  raw (ie shared strings not converted) 
			//rows in a given column within a certain range
                        if(to==-1)
				to=xml.sheetData.row.length();
			return xml.sheetData.row.(@r>=from && @r<= to).c.(@r.match(/^[A-Z]+/)[0]==column);
		}
		/**
		 * Provides an XML list representation of a range of rows in a given column as a list of xml v tags
		 * 
		 * @param column the column name e.g. "A"
		 * @param from the row number to start at e.g. 1
		 * @param to the row number to end at e.g. 10
		 * @return an XMLList of the requested rows in a single column as a list of xml v tags
		 * 
		 */     
		public function getRows(column:String="A",from:Number=1,to:Number=1000):XMLList
		{
			// returns the  converted (ie shared strings are converted) 
			//rows in a given column within a certain range
			return fillRowsWithValues(getRawRows(column,from,to));
		}
		/**
		 * Provides an XML list representation of a range of rows in a given column as a list of xml values
		 * 
		 * @param column the column name e.g. "A"
		 * @param from the row number to start at e.g. 1
		 * @param to the row number to end at e.g. 10
		 * @return an XMLList of the requested rows in a single column as a list of xml values
		 * 
		 */ 
		public function getRowsAsValues(column:String="A",from:Number=1,to:Number=1000):XMLList
		{
			// returns the  converted (ie shared strings are converted) 
			//values in a given column within a certain range
			
			return getRows(column,from,to).v;
		}
		private function rowsToValues(rows:XMLList):XMLList
		{
			//converts a set of rows to values
			return fillRowsWithValues(rows).v;
		}
		private function fillRowsWithValues(rows:XMLList):XMLList
		{
			// takes a set of rows and inserts the correct values
			var copy:XMLList=rows.copy();
			for each (var item:Object in copy)
			{
				//trace(sharedString(item.v.toString()));
				if(item.f.(children().length()!=0)+""=="") // If it's the result of a formula, no need to replace
				{
					if(item.@t=="str")
						item.v=fileLink.sharedString(item.v.toString());
					if(item.@t=="s")
						item.v=fileLink.sharedString(item.v.toString());
				}
			}
			return copy;
		}
		public function toString():String
		{
			return xml.toString();
		}
		public function toXMLString():String
		{
			return xml.toXMLString();
		}
		
	}
}
