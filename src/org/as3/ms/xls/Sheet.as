package org.as3.ms.xls {
	import flash.events.EventDispatcher;
	
	import org.as3.ms.events.XLSEvent;
	import org.as3.ms.xls.formula.Formula;
	
	
	
	
	
	[Event(name="sizingSpreadsheet", type="org.as3.ms.events.XLSEvent")]
	[Event(name="sizingSpreadsheetOver", type="org.as3.ms.events.XLSEvent")]
	
	/**
	 * Used to represent a single sheet as a series of Cells. It also handles things like the date mode and format/extended
	 * format tables.
	 */
	public class Sheet extends EventDispatcher {
		private var _formats:Array;
		private var _xformats:Array;
		public var _values:Array;
		private var _header:String;
		private var _footer:String;
		private var _dateMode:uint;
		
		private var _name:String;
		
		public var _rows:uint;
		public var _cols:uint;
		
		public function Sheet() {
			_formats = new Array();
			_xformats = new Array();
			_values = new Array();
			_header = "";
			_footer = "";
		}
		
		/**
		 * Resizes the given sheet creating blank Cells in any new rows or columns created. This has no effect if the given
		 * dimensions are smaller than the current dimensions.
		 *
		 * @param rows The new number of rows
		 * @param cols The new number of columns
		 *
		 */
		public function resize(rows:uint, cols:uint):void {
			
			_rows = Math.max(rows, _rows);
			_cols = Math.max(cols, _cols);
			var sizingEvent:XLSEvent = new XLSEvent("sizingSpreadsheet",_rows,_cols,true,false);
			dispatchEvent(sizingEvent);
			// Add needed rows
			while(_values.length <= _rows) {
				_values.push(new Array());
			}
			
			// Add needed columns
			for(var row:uint = 0; row < _values.length; row++) {
				for(var col:uint = 0; col < _cols; col++) {
					if(!(_values[row][col] is Cell)) {
						_values[row][col] = new Cell();
						_values[row][col].dateMode = _dateMode;
					}
				}
			}
			sizingEvent = new XLSEvent("sizingSpreadsheetOver",_rows,_cols,true,false);
			dispatchEvent(sizingEvent);
		}
		
		public function mergeWith(sheetToMerge:Sheet):void{
			values = values.concat(sheetToMerge.values);
			_rows 	+= sheetToMerge.rows;
			_cols	= Math.max(cols,sheetToMerge.cols);
		}
		
		/**
		 * Gets the cell object at the given location
		 * @param row
		 * @param col
		 * @return The Cell at the given location
		 *
		 */
		public function getCell(row:uint, col:uint):Cell {
			// Create a cell if one doesn't exist yet
			if(!(_values[row][col] is Cell)) {
				_values[row][col] = new Cell();
				_values[row][col].dateMode = _dateMode;
			}
			
			return _values[row][col];
		}
		
		/**
		 * Sets the value of the given cell. If a formula is assigned then the Cell's formula property is updated.
		 * @param row
		 * @param col
		 * @param value
		 *
		 */
		public function setCell(row:uint, col:uint, value:*):Cell {
			var cell:Cell;
			if ((row+1) > _rows || (col+1) > _cols) {
				resize(row+1, col+1);
			}
			cell = (_values[row][col] as Cell);
			if(value is Formula) {
				cell.formula = value;
			} else {
				cell.value = value;
			}
			return cell;
		}
		
		
		/**
		 *
		 * @return The values contained in the array. This can be used directly as the dataProvider of a DataGrid
		 * to display the contents of this sheet.
		 *
		 */
		public function get values():Array { return _values; }
		public function set values(valueArray : Array):void { _values = valueArray; }
		
		public function get formats():Array { return _formats; }
		public function set formats(value:Array):void { _formats = value; }
		
		public function get xformats():Array { return _xformats; }
		public function set xformats(value:Array):void { _xformats = value; }
		
		public function get rows():uint { return _rows; }
		public function get cols():uint { return _cols; }
		
		public function get header():String { return _header; }
		public function set header(value:String):void { _header = value; }
		
		public function get footer():String { return _footer; }
		public function set footer(value:String):void { _footer = value; }
		
		public function get dateMode():uint { return _dateMode; }
		public function set dateMode(value:uint):void {
			_dateMode = value;
			for(var r:uint = 0; r < _values.length; r++) {
				for(var c:uint = 0; c < _values[r].length; c++) {
					_values[r][c].dateMode = _dateMode;
				}
			}
		}
		
		public function get name():String { return _name; }
		public function set name(value:String):void { _name = value; }
	}
}