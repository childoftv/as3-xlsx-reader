package com.childoftv.xlsxreader 
{
	/**
	 * ...
	 * @author Vector.Lee
	 */
	public class SharedStrings 
	{
		private var sharedStringsCache:XML;
		
		public function SharedStrings() 
		{
			
		}
		
		/**
		 * @private  
		 * 
		 * Looks up the internal shared string database XML
		 * 
		 */     
		private function sharedStrings():XML
		{
			if (! sharedStringsCache)
			{
				sharedStringsCache=XMLUtils.retrieveXML("xl/sharedStrings.xml");
			}
			return sharedStringsCache;
		}
		
		/**
		 * get shared strings by sharedId
		 * @param	sharedId
		 * @return
		 */
		public function getSharedStrings(sharedId:uint):XMLList
		{
			sharedStrings();
			return sharedStringsCache.child(sharedId);
		}
		
	}

}