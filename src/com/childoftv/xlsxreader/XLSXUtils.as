package main {
	import flash.globalization.StringTools;
	public class XLSXUtils
	{
		static private const A2Z:Array = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y","Z"];
		static private const AZ_COUNT:uint = 26;
		public function XLSXUtils()
		{
		}
		
		static public function num2AZ(value:uint):String
		{
			var az:String = "";
			var remain:uint = value % AZ_COUNT;
			var count:uint = value / AZ_COUNT;
			
			
			if (count > 0)
			{
				az = num2AZ(count)
			}
			az = A2Z[remain]+az;
			return az;
		}
		
		static public function AZ2Num(value:String):uint
		{
			value ||= "";
			var num:uint = 0;
			
			var char:String = value.charAt(0);
			var index:int = A2Z.indexOf(char);
			if (index== -1)
				throw "value must be 'A'-'Z'";
			var length:uint = value.length;
			if (length > 1)
			{
				num += AZ2Num( value.slice(1) );
			}
			num = index * Math.pow(AZ_COUNT, length-1);
			return num+1;
		}
	}
}