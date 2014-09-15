package org.as3.ms.events
{
	import flash.events.Event;
	
	public class XLSEvent extends Event
	{
		public static const SIZING_SPREADSHEET 			: String 	= "sizingSpreadsheet";
		public static const SIZING_SPREADSHEET_OVER 	: String 	= "sizingSpreadsheetOver";
		
		public var numRows 		: uint;
		public var numColumns 	: uint; 
		public var bytesLoaded 	: uint;
		public var bytesTotal	: uint;
		
		public function XLSEvent(type:String,numRows:uint,numColumns:uint, bubbles:Boolean=false, cancelable:Boolean=false)
		{
			this.numRows 		= numRows;
			this.numColumns 	= numColumns;
			this.bytesLoaded 	= bytesLoaded;
			this.bytesTotal		= bytesTotal;
			super(type, bubbles, cancelable);
		}
		
		
		public override function clone():Event{
			return new XLSEvent(type,numRows,numColumns,bubbles,cancelable);
		}
	}
}