package  {
	import flash.globalization.StringTools;
	public class XLSXUtils
	{
		static private const A2Z:Array = ["0","A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y","Z"];
		static private const AZ_COUNT:uint = 26;
		public function XLSXUtils()
		{
		}
		
		
		static public function num2AZ(value:Number):String
		{
			var az:String = "";
			if (value <= AZ_COUNT)
			{
				az = A2Z[value];
			}else {
				var remain:uint = value % AZ_COUNT;
				var count:Number = value / AZ_COUNT;
				count = Math.floor(count);
				
				if (count > 0)
				{
					az = num2AZ(count)
				}
				az += A2Z[remain];
			}
			
			trace(az);
			return az;
		}
		
		static public function AZ2Num(value:String):Number
		{
			value ||= "";
			var num:Number = 0;
			
			var char:String = value.charAt(0);
			var index:Number = A2Z.indexOf(char);
			if (index== -1)
				throw "value must be 'A'-'Z'";
			var length:uint = value.length;
			num = index * Math.pow(AZ_COUNT, length-1);
			if (length > 1)
			{
				num += AZ2Num( value.slice(1) );
			}
			
			return num;
		}
	}
}