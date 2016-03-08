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
package
{
	import com.childoftv.xlsxreader.Worksheet;
	import com.childoftv.xlsxreader.XLSXLoader;
	
	import flash.display.Sprite;
	import flash.events.Event;
	import flash.text.TextField;
	
	public class LoadXLSXExample extends Sprite
	{
		//Create the Excel Loader
		protected var excel_loader:XLSXLoader = new XLSXLoader();
		private var textField:TextField = new TextField();
		
		public function LoadXLSXExample()
		{
			
			setupScreenText();
			
			//Listen for when the file is loaded
			excel_loader.addEventListener(Event.COMPLETE, loadingComplete);
			
			//Load the file
			excel_loader.load("Example Spreadsheet.xlsx");
			
			
		}
		
		private function setupScreenText():void
		{
			textField.multiline = true;
			textField.height = 100;
			textField.width = 400;
			addChild(textField);
		}
		private var log:String = "";
		private function logline(s:String):void
		{
			trace(s);
			
			var line:String = s + "\n";
			textField.appendText(line);
			
		}
		
		//Handler for loading complete
		private function loadingComplete(e:Event):void
		{
			//Access a worksheet by name ('Sheet1')
			var sheet_1:Worksheet=excel_loader.worksheet("Sheet1");
			
			//Access a cell in sheet 1 and output to trace
			logline("Cell A3=" + sheet_1.getCellValue("A3")) //outputs: Cell A3=Hello World;
			logline("Cell A4=" + sheet_1.getCellValue("A4")) //outputs: Cell A4=Hello Excel with colors;
			logline("Cell A5=" + sheet_1.getCellValue("A5")) //outputs: Cell A3=Hello Excel with the same color;
		}
	}
}