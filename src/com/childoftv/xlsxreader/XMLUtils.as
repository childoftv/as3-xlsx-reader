package com.childoftv.xlsxreader 
{
	import deng.fzip.FZipFile;
	import flash.xml.XMLDocument;
	/**
	 * ...
	 * @author Vector.Lee
	 */
	public class XMLUtils 
	{
		
		public function XMLUtils() 
		{
			
		}
		
		static public function retrieveXML(path:String):XML
		{
			
			var file:FZipFile=zipProcessor.getFileByName(path);
			return convertToOpenXMLNS(file.getContentAsString(false));
			
		}
		
		
		static public function convertToOpenXMLNS(s:String):XML
		{
			
			XML.ignoreProcessingInstructions = true;
			XML.ignoreWhitespace = false;
			var XMLDocument:XML=XML(s);
			XMLDoc.normalize();
			XML.ignoreWhitespace = true;
			return XMLDoc;
		}
		
	}

}