package org.as3.ms.xls {
	import flash.errors.ScriptTimeoutError;
	import flash.events.Event;
	import flash.events.EventDispatcher;
	import flash.events.ProgressEvent;
	import flash.events.TimerEvent;
	import flash.system.MessageChannel;
	import flash.system.MessageChannelState;
	import flash.system.Worker;
	import flash.utils.ByteArray;
	import flash.utils.Endian;
	import flash.utils.Timer;
	
	import mx.messaging.events.ChannelEvent;
	
	import org.as3.ms.biff.BIFFReader;
	import org.as3.ms.biff.BIFFVersion;
	import org.as3.ms.biff.BIFFWriter;
	import org.as3.ms.biff.Record;
	import org.as3.ms.cdf.CompoundDocument;
	import org.as3.ms.events.XLSEvent;
	import org.as3.ms.ooxml.OfficeOpenXML;
	import org.as3.ms.xls.formula.Formula;
	import org.as3.ms.xls.style.XFormat;
	
	
	
	[Event(name="sizingSpreadsheet", type="org.as3.ms.events.XLSEvent")]
	[Event(name="sizingSpreadsheetOver", type="org.as3.ms.events.XLSEvent")]
	[Event(name="excelProgess", type="flash.events.ProgressEvent")]
	[Event(name="excelComplete", type="flash.events.ProgressEvent")]
	[Event(name="unpackingOOXMLFile", type="flash.events.Event")]
	[Event(name="unpackingOOXMLFileOver", type="flash.events.Event")]
	[Event(name="unpackingOOXMLFileFailed", type="flash.events.Event")]
	/**
	 *
	 * ExcelFile has been modified to handle marshalling and multi-processing with workers
	 * 
	 */
	public class ExcelFile extends EventDispatcher {
		public static const BASE1899:uint = 0;
		public static const BASE1904:uint = 1;
		
		public static const GLOBALS:uint = 0x05;
		public static const SHEET:uint = 0x10;
		
		private var biffReader		: BIFFReader;
		private var version			: uint;
		private var dateMode		: uint;
		
		private var globalFormats	: Array = new Array();
		private var globalXFormats	: Array = new Array();
		private var notes			: Array;
		
		private var currentFileType	: uint;
		private var currentSheet	: Sheet;
		private var currentSheetIdx	: uint = 0;
		public var _sheets			: Array = new Array();
		
		private var lastRecordType	: uint;
		
		private var _sst			: Array;
		
		
		
		
		private const handlers		: Array = initHandlers();
		private static const ignore	: Array = [
			0x4D, 0xE1, 0xC0, 0xC1, 0xE1,
			0xE2, 0x5D, 0x9C, 0xBF, 0xEB,
			0xEE, 0xF1, 0x13D, 0x1AF, 0x1B6,
			0x1B7, 0x1BC, 0x1C0, 0x1C1, 0x1C2,
			0x863, 0x8C8, 0x105C];
		
		private function initHandlers():Array {
			var handlers:Array = [
				dimensions,
				blank, integer, number, label, boolerr,
				formula, string, row, bof, eof, // 10
				index, calccount, calcmode, precision, refmode,
				delta, iteration, protect, password, header, // 20
				footer, externcount, externsheet, name, windowprotect,
				vertpagebreaks, horizpagebreaks, note, selection, format, // 30
				builtinfmtcount, columndefault, array, datemode, externname,
				colwidth, defrowheight, leftmargin, rightmargin, topmargin, // 40
				bottommargin, printheaders, printgridlines, null, null,
				null, filepass, null, font, font2, // 50
				null, null, null, tableop, tableop2,
				null, null, null, null, cont, // 60 // Continue
				window1, window2, null, backup, pane,
				codepage, xf, ixfe, efont, null, // 70
				null, null, null, null, null,
				null, null, null, null, null, // 80
				dconref, null, null, null, defcolwidth,
				builtinfmtcount, null, null, xct, crn, // 90
				filesharing, writeaccess, null, uncalced, saverecalc,
				null, null, null, objectprotect, null, // 100
				null, null, null, null, null,
				null, null, null, null, null, // 110
				null, null, null, null, null,
				null, null, null, null, null, // 120
				null, null, null, null, colinfo,
				null, null, guts, wsbool, gridset, // 130
				hcenter, vcenter, boundsheet, writeprot, null,
				null, null, null, null, country, // 140
				hideobj, null, null, null, null,
				palette, null, null, null, null, // 150
				null, null, null, null, null,
				null, null, null, null, null, // 160
				setup, null, null, null, null,
				null, null, null, null, null, // 170
				gcw, null, null, null, null,
				null, null, null, null, null, // 180
				null, null, null, null, null,
				null, null, null, mulrk, mulblank, // 190
				null, null, null, null, null,
				null, null, null, null, null, // 200
				null, null, null, null, null,
				null, null, null, null, null, // 210
				null, null, null, null, dbcell,
				null, null, bookbool, null, null, // 220
				null, null, null, xf, null,
				null, null, null, mergedcells, null, // 230
				null, null, null, null, null,
				null, null, null, phonetic, null, // 240
				null, null, null, null, null,
				null, null, null, null, null, // 250
				null, sst, labelsst, null, extsst,
				null, null, null, null, null, // 260
				null, null, null, null, null,
				null, null, null, null, null, // 270
				null, null, null, null, null,
				null, null, null, null, null, // 280
				null, null, null, null, null,
				null, null, null, null, null, // 290
				null, null, null, null, null,
				null, null, null, null, null, // 300
				null, null, null, null, null,
				null, null, null, null, null, // 310
				null, null, null, null, null,
				null, null, null, null, null, // 320
				null, null, null, null, null,
				null, null, null, null, null, // 330
				null, null, null, null, null,
				null, null, null, null, null, // 340
				null, null, null, null, null,
				null, null, null, null, null, // 350
				null, uselfs, dsf, null, null,
				null, null, null, null, null, // 360
				null, null, null, null, null,
				null, null, null, null, null, // 370
				null, null, null, null, null,
				null, null, null, null, null, // 380
				null, null, null, null, null,
				null, null, null, null, null, // 390
				null, null, null, null, null,
				null, null, null, null, null, // 400
				null, null, null, null, null,
				null, null, null, null, null, // 410
				null, null, null, null, null,
				null, null, null, null, null, // 420
				null, null, null, null, null,
				null, null, null, null, null, // 430
				null, null, null, null, null,
				null, null, null, null, dv, // 440
				null, null, null, null, null,
				null, null, null, null, null, // 450
				null, null, null, null, null,
				null, null, null, null, null, // 460
				null, null, null, null, null,
				null, null, null, null, null, // 470
				null, null, null, null, null,
				null, null, null, null, null, // 480
				null, null, null, null, null,
				null, null, null, null, null, // 490
				null, null, null, null, null,
				null, null, null, null, null, // 500
				null, null, null, null, null,
				null, null, null, null, null, // 510
				null, dimensions, blank, null, number,
				label, null, formula, null, row, // 520
				bof, null, index, null, null  					//handlers[Type.BOF_3] = bof;
			];
			handlers[Type.RK] = rk;
			handlers[Type.SHAREDFMLA] = shrfmla;
			handlers[Type.SHEETPROTECTION] = sheetprotection;
			handlers[Type.RANGEPROTECTION] = rangeprotection;
			handlers[Type.DEFAULTROWHEIGHT_3458] = defrowheight;
			handlers[Type.FONT_34] = font;
			handlers[Type.WINDOW2_3458] = window2;
			handlers[Type.XF_3] = xf;
			handlers[Type.STYLE] = style;
			handlers[Type.FORMULA_4] = formula; // handlers[1030] = formula;
			handlers[Type.BOF_4] = bof;//handlers[1033] = bof;
			handlers[Type.FORMAT_458] = format; //handlers[1054] = format;
			handlers[Type.XF_4] = xf;//handlers[1091] = xf;
			handlers[Type.BOF_58] = bof; //handlers[2057] = bof;
			
			return handlers;
		}
		
		/**
		 *
		 * @return An ArrayCollection with the sheets in the Excel file.
		 *
		 */
		public function get sheets():Array { return _sheets; }
		
		/**
		 * Saves the first sheet in the sheets array as a BIFF2 document. Saving formulas
		 * is not currently supported.
		 * @return A ByteArray containing the saved sheet in BIFF2 form
		 *
		 */
		public function saveToByteArray(charset:String="UTF8"):ByteArray {
			var s:Sheet = _sheets[0] as Sheet;
			
			var br:BIFFWriter = new BIFFWriter();
			
			// Write the BOF and header records
			var bof:Record = new Record(Type.BOF);
			bof.data.writeShort(BIFFVersion.BIFF2);
			bof.data.writeByte(0);
			bof.data.writeByte(0x10);
			br.writeTag(bof);
			
			// Date mode
			var dateMode:Record = new Record(Type.DATEMODE);
			dateMode.data.writeShort(1);
			br.writeTag(dateMode);
			
			// Store built in formats
			var formats:Array = ["General",
				"0", "0.00", "#,##0", "#,##0.00",
				"", "", "", "",
				"0%", "0.00%", "0.00E+00",
				"#?/?", "#??/??",
				"M/D/YY", "D-MMM-YY", "D-MMM", "MMM-YY"];
			
			var numfmt:Record = new Record(Type.BUILTINFMTCOUNT);
			numfmt.data.writeShort(formats.length);
			br.writeTag(numfmt);
			
			for(var n:uint = 0; n < formats.length; n++) {
				var fmt:Record = new Record(Type.FORMAT);
				fmt.data.writeByte(formats[n].length);
				fmt.data.writeUTFBytes(formats[n]);
				br.writeTag(fmt);
			}
			
			var dimensions:Record = new Record(Type.DIMENSION);
			dimensions.data.writeShort(0);
			dimensions.data.writeShort(s.rows+1);
			dimensions.data.writeShort(0);
			dimensions.data.writeShort(s.cols+1);
			br.writeTag(dimensions);
			
			for(var r:uint = 0; r < s.rows; r++) {
				for(var c:uint = 0; c < s.cols; c++) {
					var value:* = s.getCell(r, c).value;
					var cell:Record = new Record(1);
					cell.data.writeShort(r);
					cell.data.writeShort(c);
					
					if(value is Date) {
						var dateNum:Number = (value.time / 86400000) + 24106.667;
						cell.type = Type.NUMBER;
						cell.data.writeByte(0);
						cell.data.writeByte(15);
						cell.data.writeByte(0);
						cell.data.writeDouble(dateNum);
					} else if(isNaN(Number(value)) == false && String(value) != "") {
						cell.type = Type.NUMBER;
						cell.data.writeByte(0);
						cell.data.writeByte(0);
						cell.data.writeByte(0);
						cell.data.writeDouble(value);
					} else if(String(value).length > 0) {
						cell.type = Type.LABEL;
						cell.data.writeByte(0);
						cell.data.writeByte(0);
						cell.data.writeByte(0);
						var len:uint = String(value).length;
						cell.data.writeByte(len);
						cell.data.writeMultiByte(value, charset);
					} else {
						cell.type = Type.BLANK;
						cell.data.writeByte(0);
						cell.data.writeByte(0);
						cell.data.writeByte(0);
					}
					
					br.writeTag(cell);
				}
			}
			
			
			// Finally, the closing EOF record
			var eof:Record = new Record(Type.EOF);
			br.writeTag(eof);
			
			br.stream.position = 0;
			return br.stream;
		}
		
		
		
		
		/**
		 * Loads the sheets from a ByteArray containing an Excel file. If the ByteArray contains a CDF file the Workbook stream
		 * will be extracted and loaded.
		 * The original method
		 *
		 * @see com.as3xls.cdf.CDFReader
		 */
		public function loadFromByteArrayNoThreading(xls:ByteArray):void {
			// Newer workbooks are actually cdf files which must be extracted
			if(CompoundDocument.isCompoundDocumentFormatFile(xls)) {
				var cdf:CompoundDocument = new CompoundDocument(xls);
				xls = cdf.loadDirectoryEntry(1);
			}
			
			biffReader = new BIFFReader(xls);
			
			var unknown:Array = [];
			var r:Record;
			
			
			while((r = biffReader.readTag()) != null) {
				if(ignore.indexOf(r.type) != -1) {
					continue;
				}
				
				if(r.type != Type.CONTINUE) {
					lastRecordType = r.type;
				}
				
				if(handlers[r.type] is Function) {
					if (version==BIFFVersion.BIFF2) {
						if (!currentSheet){
							currentSheet = new Sheet();
							currentSheet.name = 'test';
						}
					}
					(handlers[r.type] as Function).call(this, r, currentSheet);
				} else {
					unknown.push(r.type);
				}
				
				var progressEvent:ProgressEvent = new ProgressEvent("excelProgess",false,true,biffReader.stream.position,biffReader.stream.length);
				dispatchEvent(progressEvent);
			}
			if (version==BIFFVersion.BIFF2) {
				_sheets.push(currentSheet);
			}
			if(unknown.length > 0) {
				//throw new Error("Unsupported BIFF records: " + unknown.join(", "));
			}
			
			
		}
		
		
		
		/**
		 * faster
		 * replace ignore.indexOf(r.type) != -1
		 * 
		 *  
		 * @param type
		 * @return 
		 * 
		 */
		private function ignoreType(type:uint):Boolean{
			switch(type){
				case 0x4D:
				case 0xE1:
				case 0xC0:
				case 0xC1:
				case 0xE1:
				case 0xE2:
				case 0x5D:
				case 0x9C:
				case 0xBF:
				case 0xEB:
				case 0xEE:
				case 0xF1:
				case 0x13D:
				case 0x1AF:
				case 0x1B6:
				case 0x1B7:
				case 0x1BC:
				case 0x1C0:
				case 0x1C1:
				case 0x1C2:
				case 0x863:
				case 0x8C8:
				case 0x105C:
					return true;
					
			}
			return false;
		}
		
		private function onSizingSpreadsheetEvent(event:XLSEvent):void{
			dispatchEvent(event);
		}
		
		/**
		 * used in worker multi threading
		 *
		 * @param replyChannel
		 * @param eventChannel
		 * @param xls
		 * @param fileExtension xls or xlsx
		 * 
		 */
		public function loadFromByteArrayInWorker(replyChannel:MessageChannel,eventChannel:MessageChannel,xls:ByteArray,fileExtension:String):void{
			var onSizingSpreadsheetEvent : Function = function(event:XLSEvent):void{
				eventChannel.send({type:event.type,numRows:event.numRows,numColumns:event.numColumns});
			}
			var onProgressEvent		: Function = function(event:Event):void{
				if(event is ProgressEvent){
					var pe : ProgressEvent = event as ProgressEvent;
					eventChannel.send({type:event.type,bytesLoaded:pe.bytesLoaded,bytesTotal:pe.bytesTotal});
				} else {
					eventChannel.send({type:event.type});
				}
			}
			this.addEventListener(XLSEvent.SIZING_SPREADSHEET,onSizingSpreadsheetEvent);
			this.addEventListener(XLSEvent.SIZING_SPREADSHEET_OVER,onSizingSpreadsheetEvent);
			this.addEventListener("excelFormatError",onProgressEvent);
			this.addEventListener("excelProgess",onProgressEvent);
			this.addEventListener("excelComplete",onProgressEvent);
			this.addEventListener("unpackingOOXMLFile",onProgressEvent);
			this.addEventListener("unpackingOOXMLFileOver",onProgressEvent);
			this.addEventListener("unpackingOOXMLFileFailed",onProgressEvent);
			
			var result : ExcelFile = loadFromByteArray(xls,fileExtension);
			replyChannel.send(result);
			this.removeEventListener(XLSEvent.SIZING_SPREADSHEET,onSizingSpreadsheetEvent);
			this.removeEventListener(XLSEvent.SIZING_SPREADSHEET_OVER,onSizingSpreadsheetEvent);
			this.removeEventListener("excelFormatError",onProgressEvent);
			this.removeEventListener("excelProgess",onProgressEvent);
			this.removeEventListener("excelComplete",onProgressEvent);
			this.removeEventListener("unpackingOOXMLFile",onProgressEvent);
			this.removeEventListener("unpackingOOXMLFileOver",onProgressEvent);
			this.removeEventListener("unpackingOOXMLFileFailed",onProgressEvent);
		}
		
		/**
		 * loads transforms .xls and .xlsx files into an ExcelFile object.
		 * used either in a worker or directly inside the main process.
		 * 
		 * 
		 *  
		 * @param xls
		 * @param fileExtension used to differentiate .xls and .xlsx files
		 * @return 
		 * 
		 */
		public function loadFromByteArray(xls:ByteArray,fileExtension:String):ExcelFile{
			var result : ExcelFile;
			if(fileExtension == "xls"){
				result = loadFromXlsByteArray(xls);
			} else if(fileExtension == "xlsx"){
				result = loadFromXlsxByteArray(xls);
			}
			return result;
		}
		
		
		public function loadFromXlsByteArray(xls:ByteArray):ExcelFile{
			// Newer workbooks are actually cdf files which must be extracted
			if(CompoundDocument.isCompoundDocumentFormatFile(xls)) {
				var cdf:CompoundDocument = new CompoundDocument(xls);
				xls = cdf.loadDirectoryEntry(1);
				xls.shareable = true;
			}
			
			biffReader = new BIFFReader(xls);
			var unknown:Array = [];
			var r:Record;
			var progressEvent:*;
			var operationEvent:*;
			var threadSimulator:Timer = new Timer(100,1);
			var previousType : uint = uint.MAX_VALUE;
			var counter: uint = 0;
			var complete:int = 0;
			
			while((r = biffReader.readTag()) != null){
				if(ignoreType(r.type)){
					continue;
				} else {
					
					if(r.type != Type.CONTINUE) {
						lastRecordType = r.type;
					}
					
					if(handlers[r.type] is Function) {
						if (version==BIFFVersion.BIFF2) {
							if (!currentSheet){
								currentSheet = new Sheet();
								currentSheet.name = 'test';
								currentSheet.addEventListener(XLSEvent.SIZING_SPREADSHEET,onSizingSpreadsheetEvent);
								currentSheet.addEventListener(XLSEvent.SIZING_SPREADSHEET_OVER,onSizingSpreadsheetEvent);
							}
						}
						
						try{
							++counter;
							(handlers[r.type] as Function).call(this, r, currentSheet);
							previousType = r.type;
						} catch (e:Error){
							
							dispatchEvent(new ProgressEvent("excelFormatError",false,false,biffReader.stream.position,biffReader.stream.length));
							return null;
							
						}
					} else {
						unknown.push(r.type);
					}							
				}
				if(counter%10 == 0){
					dispatchEvent(new ProgressEvent("excelProgess",false,false,biffReader.stream.position,biffReader.stream.length));
				}
			}
			
			if (version==BIFFVersion.BIFF2) {
				_sheets.push(currentSheet);
			}
			if(unknown.length > 0) {
				//throw new Error("Unsupported BIFF records: " + unknown.join(", "));
			}
			dispatchEvent(new ProgressEvent("excelComplete",false,false,biffReader.stream.position,biffReader.stream.length));
			return this;
		}
		
		public function loadFromXlsxByteArray(xls:ByteArray):ExcelFile{
			var ooxml :OfficeOpenXML = OfficeOpenXML.open(xls,this);
			ooxml.ooxmlBuildExcelFile(this);
			
			return this;
		}
		
		
		public function mergeSheetsInWorker(replyChannel:MessageChannel,eventChannel:MessageChannel):void{
			mergeSheets();
		}
		
		/**
		 * merge all sheets into one, the first sheet
		 * used on Excel 2003 when a single sheet cannot handle all the data
		 * 
		 */
		public function mergeSheets():void{
			var numSheets: int = sheets.length;
			if(numSheets>1){
				var index:int = 0;
				var refSheet:Sheet = sheets[0];
				var sheetToMerge:Sheet;
				while(++index < numSheets){
					sheetToMerge = sheets[index];
					refSheet.mergeWith(sheetToMerge);
				}
			}
		}
		
		/**
		 *
		 * simulates a multi threaded loadFromByteArray(...
		 * should be replaced by a worker
		 *  
		 * @param xls
		 * @param onLoaded
		 * @param loop
		 * 
		 */
		public function loadFromByteArraySimulatedThreading(xls:ByteArray,onLoaded:Function,loop:uint = 2000):void {
			// Newer workbooks are actually cdf files which must be extracted
			if(CompoundDocument.isCompoundDocumentFormatFile(xls)) {
				var cdf:CompoundDocument = new CompoundDocument(xls);
				xls = cdf.loadDirectoryEntry(1);
			}
			
			biffReader = new BIFFReader(xls);
			
			var unknown:Array = [];
			var r:Record;
			var progressEvent:ProgressEvent;
			var operationEvent:Event;
			var threadSimulator:Timer = new Timer(100,1);
			var previousType : uint = uint.MAX_VALUE;
			var counter: uint = 0;				
			
			var onTimerHandler:Function = function(event:TimerEvent):void{
				var loopIndex:uint = loop;
				var complete:int = 0;
				
				while(loopIndex--){
					if((r = biffReader.readTag()) != null){
						//if(ignore.indexOf(r.type) != -1) {
						if(ignoreType(r.type)){
							continue;
						} else {
							
							if(r.type != Type.CONTINUE) {
								lastRecordType = r.type;
							}
							
							if(handlers[r.type] is Function) {
								if (version==BIFFVersion.BIFF2) {
									if (!currentSheet){
										currentSheet = new Sheet();
										currentSheet.name = 'test';
										currentSheet.addEventListener("sizingSpreadsheet",onSizingSpreadsheetEvent);
										currentSheet.addEventListener("sizingSpreadsheetOver",onSizingSpreadsheetEvent);
									}
								}
								try{
									++counter;
									(handlers[r.type] as Function).call(this, r, currentSheet);
									previousType = r.type;
								} catch (e:Error){
									trace(counter,"error",r.type, Type.getType(r.type), "previous",Type.getType(previousType));
									previousType = r.type;
									complete = -1;
									break;
								}
							} else {
								unknown.push(r.type);
							}							
						}
					} else {
						
						if (version==BIFFVersion.BIFF2) {
							_sheets.push(currentSheet);
						}
						if(unknown.length > 0) {
							//throw new Error("Unsupported BIFF records: " + unknown.join(", "));
						}
						complete = 1;
						break;
					}
				}
				progressEvent = new ProgressEvent("excelProgess",true,false,biffReader.stream.position,biffReader.stream.length);
				dispatchEvent(progressEvent);
				
				if(complete == 1){
					progressEvent = new ProgressEvent("excelComplete",true,false,biffReader.stream.length,biffReader.stream.length);
					dispatchEvent(progressEvent);
					onLoaded();
				} else if (complete == 0) {
					threadSimulator.reset();
					threadSimulator.start();
				} else {
					threadSimulator.stop();
					progressEvent = new ProgressEvent("excelFormatError",true,false,biffReader.stream.length,biffReader.stream.length);
					dispatchEvent(progressEvent);
				}
			}
			threadSimulator.addEventListener(TimerEvent.TIMER,onTimerHandler);
			threadSimulator.start();
		}
		
		
		// Values
		private function integer(r:Record, s:Sheet):void {
			var row:uint = r.data.readUnsignedShort();
			var col:uint = r.data.readUnsignedShort();
			
			// Cell attributes
			var attr1:uint = r.data.readUnsignedByte();
			var attr2:uint = r.data.readUnsignedByte();
			var attr3:uint = r.data.readUnsignedByte();
			
			// Figure out the format
			var format:uint = attr2 & 0x3F;
			
			// Integer values can only be unsigned
			var value:Number = r.data.readUnsignedShort();
			
			// Figure out the format
			var formatString:String = s.formats[format];
			s.setCell(row, col, value);
			s.getCell(row, col).format = formatString;
		}
		
		private function number(r:Record, s:Sheet):void {
			var row:uint = r.data.readUnsignedShort();
			var col:uint = r.data.readUnsignedShort();
			
			if(r.type < 0x200) {
				// BIFF2
				// Cell attributes
				r.data.readUnsignedByte();
				r.data.readUnsignedByte();
				r.data.readUnsignedByte();
			} else {
				// BIFF>2
				var indexToXF:uint = r.data.readUnsignedShort();
			}
			
			
			var value:Number = r.data.readDouble();
			
			if(version == BIFFVersion.BIFF2) {
				s.setCell(row, col, value);
			} else {
				s.setCell(row, col, value);
				// indexToXF can be an issue
				var fmt:String = s.formats[s.xformats[indexToXF].format];
				if(fmt == null || fmt.length == 0) {
					fmt = Formatter.builtInFormats[s.xformats[indexToXF].format];
				}
				s.getCell(row, col).format = fmt;
			}
			
		}
		
		private function label(r:Record, s:Sheet):void {
			var row:uint = r.data.readUnsignedShort();
			var col:uint = r.data.readUnsignedShort();
			
			var len:uint;
			
			if(r.type == Type.LABEL) {
				// BIFF2
				// Cell attributes
				r.data.readUnsignedByte();
				r.data.readUnsignedByte();
				r.data.readUnsignedByte();
				
				len = r.data.readUnsignedByte();
			} else {
				// BIFF3+
				var indexToXF:uint = r.data.readUnsignedShort();
				len = r.data.readUnsignedShort();
			}
			
			var value:String = r.data.readUTFBytes(len);
			
			s.setCell(row, col, value);
			
			if(version != BIFFVersion.BIFF2) {
				var fmt:String = s.formats[s.xformats[indexToXF].format];
				if(fmt == null || fmt.length == 0) {
					fmt = Formatter.builtInFormats[s.xformats[indexToXF].format];
				}
				s.getCell(row, col).format = fmt;
			}
		}
		
		private function labelsst(r:Record, s:Sheet):void {
			var row:uint = r.data.readUnsignedShort();
			var col:uint = r.data.readUnsignedShort();
			
			if(row == 65533){
				
			}
			
			var xfIndex:uint = r.data.readUnsignedShort();
			var sstIndex:uint = r.data.readUnsignedInt();
			
			var value:String = _sst[sstIndex];
			
			
			s.setCell(row, col, value);
		}
		
		private function rk(r:Record, s:Sheet):void {
			var row:uint = r.data.readUnsignedShort();
			var col:uint = r.data.readUnsignedShort();
			var indexToXF:uint = r.data.readUnsignedShort();
			
			
			var value:Number = readRK(r, s);
			
			s.setCell(row, col, value);
			
			var fmt:String = s.formats[s.xformats[indexToXF].format];
			if(fmt == null || fmt.length == 0) {
				fmt = Formatter.builtInFormats[s.xformats[indexToXF].format];
			}
			
			s.getCell(row, col).format = fmt;
			
		}
		
		private function note(r:Record, s:Sheet):void {
			var row:uint = r.data.readUnsignedShort();
			var col:uint = r.data.readUnsignedShort();
			
			var note:String;
			if(version <= BIFFVersion.BIFF5) {
				var totalLength:uint = r.data.readUnsignedShort();
				note = r.data.readUTFBytes(r.data.bytesAvailable);
				
			} else {
				var flags:uint = r.data.readUnsignedShort();
				var idx:uint = (r.data.readUnsignedShort() - 1)*2;
				var author:String = r.readUnicodeStr16();
				note = notes[idx];
			}
			
			s.getCell(row, col).note = note;
		}
		
		private function mulblank(r:Record, s:Sheet):void {
			var row:uint = r.data.readUnsignedShort();
			var col:uint = r.data.readUnsignedShort();
			
			while(r.data.bytesAvailable > 2) {
				var indexToXF:uint = r.data.readUnsignedShort();
				s.setCell(row, col, "");
				col++;
			}
		}
		
		/**
		 * MULRK – Multiple RK
		 * BIFF2			BIFF3			BIFF4			BIFF5			BIFF8
		 * —				—				—				00BD			00BD
		 * This record represents a cell range containing RK value cells. All cells are located in the same row.
		 * Record MULRK, BIFF5-BIFF8:
		 * Offset			Size			Contents
		 * 0				2				Index to row
		 * 2				2				Index to first column (fc)
		 * 4				6∙nc			List of	nc = lc–fc +1 XF/RK structures. Each XF/RK contains:
		 * 									Offset		Size		Contents
		 * 									0			2			Index to XF record
		 * 									2			4			RK value
		 * 4+6∙nc			2				Index to last column (lc)
		 *  
		 * @param r
		 * @param s
		 * 
		 */
		private function mulrk(r:Record, s:Sheet):void {
			var row:uint = r.data.readUnsignedShort();
			var col:uint = r.data.readUnsignedShort();
			while(r.data.bytesAvailable > 2) {
				var indexToXF:uint = r.data.readUnsignedShort();
				var value:Number = readRK(r, s);
				var fmt:String = s.formats[s.xformats[indexToXF].format];
				if(fmt == null || fmt.length == 0) {
					fmt = Formatter.builtInFormats[s.xformats[indexToXF].format];
				}
				s.setCell(row, col, value);
				s.getCell(row, col).format = fmt;
				
				col++;
			}
		}
		
		
		/**
		 * An RK value is an encoded integer or floating-point value.
		 * RK values have a size of 4 bytes and are used to decrease file size for floating-point values.
		 * Structure of an RK value (32-bit value), BIFF3-BIFF8:
		 * Bit				Mask				Contents
		 * 0				00000001			0 = Value not changed		1 = Encoded value is multiplied by 100
		 * 1				00000002			0 = Floating-point value	1 = Signed integer value
		 * 31-2				FFFFFFFC			Encoded value
		 * 
		 * If bit 1 is cleared, the encoded value represents the 30 most significant bits of an IEEE 754 floating-point value (6-bit double precision).
		 * The 34 least significant bits must be set to zero. If bit 1 is set, the encoded value represents a signed 30- bit integer value.
		 * To get the correct integer, the encoded value has to be shifted right arithmetically by 2 bits.
		 * If bit 0 is set, the decoded value (both integer and floating-point) must be divided by 100 to get the final result.
		 * 
		 * Examples:
		 * RK value				Type				Div 100				Encoded value				Decoded value				Result
		 * 3FF00000				float				no					3FF00000					3FF0000000000000= 1.0		1.0
		 * 3FF00001				float				yes					3FF00000					3FF0000000000000= 1.0		0.01
		 * 004B5646				integer				no					004B5644					0012D591= 1234321			1234321
		 * 004B5647				integer				yes					004B5644					0012D591  = 1234321			12343.21
		 * 
		 * @param r
		 * @param s
		 * @return 
		 * 
		 */		
		private function readRK(r:Record, s:Sheet):Number {
			var raw:uint = r.data.readUnsignedInt();
			var div100:Boolean = (raw & 0x00000001) == 1;
			var intVal:Boolean = (raw & 0x00000002) == 2;
			
			r.data.position -= 4;
			
			var value:Number;
			if(intVal) {
				value = r.data.readInt() >> 2;
			} else {
				var b:ByteArray = new ByteArray();
				b[7] = 0;
				b[6] = 0;
				b[5] = 0;
				b[4] = 0;
				b[3] = r.data.readUnsignedByte();
				b[2] = r.data.readUnsignedByte();
				b[1] = r.data.readUnsignedByte();
				b[0] = r.data.readUnsignedByte();
				value = b.readDouble();
			}
			
			if(div100) {
				value = Math.round(value) / 100;
			}
			
			return value;
		}
		
		private function blank(r:Record, s:Sheet):void {
			var row:uint = r.data.readUnsignedShort();
			var col:uint = r.data.readUnsignedShort();
			
			s.setCell(row, col, "");
		}
		
		private function formula(r:Record, s:Sheet):void {
			var row:uint = r.data.readUnsignedShort();
			var col:uint = r.data.readUnsignedShort();
			
			// Cell attributes
			if(version == BIFFVersion.BIFF2) {
				r.data.readUnsignedByte();
				r.data.readUnsignedByte();
				r.data.readUnsignedByte();
			} else {
				var indexToXF:uint = r.data.readUnsignedShort();
			}
			
			var result:Number = r.data.readDouble();
			
			var alwaysRecalculate:Boolean;
			if(version == BIFFVersion.BIFF2) {
				alwaysRecalculate = r.data.readUnsignedByte() == 1;
			} else {
				alwaysRecalculate = r.data.readUnsignedShort() == 0;
			}
			
			if(version >= BIFFVersion.BIFF5) {
				// For some reason in BIFF5-8 there are 4 unused bytes before the token array
				r.data.position += 4;
			}
			
			var tokenArrSize:uint;
			if(version == BIFFVersion.BIFF2) {
				tokenArrSize = r.data.readUnsignedByte();
			} else {
				tokenArrSize = r.data.readUnsignedShort();
			}
			var tokens:ByteArray = new ByteArray();
			r.data.readBytes(tokens, 0, tokenArrSize);
			
			var f:Formula = new Formula(row, col, tokens, version, s);
			s.setCell(row, col, f);
			
			var fmt:String = s.formats[s.xformats[indexToXF].format];
			if(fmt == null || fmt.length == 0) {
				fmt = Formatter.builtInFormats[s.xformats[indexToXF].format];
			}
			
			s.getCell(row, col).format = fmt;
		}
		
		private function shrfmla(r:Record, s:Sheet):void {
			var firstRow:uint = r.data.readUnsignedShort();
			var lastRow:uint = r.data.readUnsignedShort();
			var firstCol:uint = r.data.readUnsignedByte();
			var lastCol:uint = r.data.readUnsignedByte();
			
			
			// Not used
			r.data.position++;
			
			var numExistingFormulaRecords:uint = r.data.readUnsignedByte();
			
			var tokLen:uint = version == BIFFVersion.BIFF2 ? r.data.readUnsignedByte() : r.data.readUnsignedShort();
			
			// Next comes the formula
			var tokens:ByteArray;
			
			for(var rw:uint = firstRow; rw < lastRow; rw++) {
				for(var c:uint = firstCol; c <= lastCol; c++) {
					tokens = new ByteArray();
					tokens.endian = Endian.LITTLE_ENDIAN;
					r.data.readBytes(tokens, 0, tokLen);
					r.data.position -= tokLen;
					s.getCell(rw, c).sharedTokens = tokens;
				}
			}
		}
		
		
		
		/**
		 * Shared String Table setup
		 * 
		 * A BIFF8 workbook collects the strings of all text cells in a global list, the Shared String Table.
		 * This table is located in the record SST in the Workbook Globals Substream
		 * 
		 * SST – Shared String Table
		 * BIFF2	BIFF3	BIFF4	BIFF5	BIFF8
		 * 	—		  —		  —		  —		 0x00FC
		 * This record contains a list of all strings used anywhere in the workbook. Each string occurs only once.
		 * The workbook uses indexes into the list to reference the strings.
		 * Record SST, BIFF8:
		 * Offset		Size		Contents
		 * 	 0			 4			Total number of strings	in the workbook (see below)
		 *   4			 4			Number of following strings (nm)
		 *   8 			var. 		List of nm Unicode strings, 16 - bit string length
		 * 
		 * The first field of the SST record counts the total occurrence of strings in the workbook. For instance, the string “AAA”
		 * is used 3 times and the string “BBB” is used 2 times. The first field contains 5 (3+2) and the second field contains 2 (field "AAA" and field "BBB", followed by the two strings.
		 * 
		 * @param r
		 * @param s
		 * 
		 */
		private function sst(r:Record, s:Sheet):void {
			//Total number of strings	in the workbook
			var numWorkbookStrings:uint = r.data.readUnsignedInt();
			// offset 4
			// Number of following strings
			var sstSize:uint = r.data.readUnsignedInt();
			// offset 8
			// List of nm Unicode strings, 16 - bit string length
			var dataArr:Array = [r.data];
			while(true){
				var d:ByteArray = get_continuation_data();
				if (!d) {
					break;
				} else {
					dataArr.push(d);
				}
			}
			_sst = unpack_sst(dataArr, sstSize);
		}
		
		/**
		 * builds the SST Shared String Table
		 * 
		 * @param dataArr
		 * @param sstSize
		 * @return 
		 * 
		 */
		private function unpack_sst(dataArr:Array, sstSize:uint):Array {
			var _currIndex:uint = 0;
			var _currData:ByteArray = dataArr[_currIndex];
			
			// Now unpack
			var _strings:Array = new Array();
			for(var n:uint = 0; n < sstSize; n++) {
				// this is very much like readUnicodeStr16(), except it deals with continuations
				var len:uint = _currData.readUnsignedShort();
				var opts:uint = _currData.readByte();
				var compressed:Boolean = (opts & 0x01) == 0;
				var asianPhonetic:Boolean = (opts & 0x04) == 0x04;
				var richtext:Boolean = (opts & 0x08) == 0x08;
				var toSkip:uint = 0;
				// We need to skip past these if they're present
				if (richtext) {
					toSkip += 4 * _currData.readShort();
				}
				if (asianPhonetic) {
					toSkip += _currData.readUnsignedInt();
				}
				
				var fullString:String = "";
				var charsGot:uint = 0;
				while (true) {
					var charsNeed:uint = len - charsGot;
					var charsAvail:uint;
					var _strArray:Array = [];
					var i:uint;
					if (compressed) {
						// This is compressed UTF-16, not UTF-8, so we don't use readUTFBytes()
						charsAvail = charsNeed > _currData.bytesAvailable ? _currData.bytesAvailable : charsNeed;
						for (i = 0; i < charsAvail; i++){
							_strArray.push(_currData.readUnsignedByte());
						}
					} else {
						// Treating string as UCS-2, rather than UTF-16 (i.e. ignoring surrogate pairs)
						// readMultiByte() claims to do this, but doesn't seem to work...
						charsAvail = charsNeed > (_currData.bytesAvailable/2) ? (_currData.bytesAvailable/2) : charsNeed;
						for (i = 0; i < charsAvail; i++){
							_strArray.push(_currData.readUnsignedShort());
						}
					}
					var partialString:String = String.fromCharCode.apply(null, _strArray);
					fullString += partialString;
					charsGot += charsAvail;
					if (charsGot == len) {
						break;
					}
					_currIndex += 1;
					_currData = dataArr[_currIndex];
					var new_opts:uint = _currData.readByte();
					compressed = (new_opts & 0x01) == 0;
				}
				_currData.position += toSkip;
				if (_currData.position >= _currData.length) {
					_currIndex += 1;
					if (_currIndex < dataArr.length){
						_currData = dataArr[_currIndex];
					}
				}
				_strings.push(fullString);
			}
			return _strings;
		}
		
		/**
		 * 
		 * @return 
		 * 
		 */
		private function get_continuation_data():ByteArray {
			var ba:ByteArray;
			var pos:uint = biffReader.stream.position;
			var r:Record = biffReader.readTag();
			if (r.type != Type.CONTINUE){
				biffReader.stream.position = pos;
				return null;
			} else {
				return r.data;
			}
		}
		
		private function dimensions(r:Record, s:Sheet):void {
			// For some reason sometimes the dimension record is blank. Reading it fails in this case
			if(r.data.length == 0) {
				return;
			}
			
			// Using the biff version doesn't seem to work; instead use the record size to figure out
			// whether the row indeces are ints or shorts;
			var firstRow	: uint = r.length == 14 ? r.data.readUnsignedInt() : r.data.readUnsignedShort();
			var lastRow		: uint = r.length == 14 ? r.data.readUnsignedInt() : r.data.readUnsignedShort();
			var firstCol	: uint = r.data.readUnsignedShort();
			var lastCol		: uint = r.data.readUnsignedShort();
			if (s) {
				s.resize(lastRow, lastCol);
			}
		}
		
		private function boundsheet(r:Record, s:Sheet):void {
			var sheetOffset	: uint = r.data.readUnsignedInt();
			var visibility	: uint = r.data.readUnsignedByte();
			var sheetType	: uint = r.data.readUnsignedByte();
			
			var name:String;
			if(version == BIFFVersion.BIFF8){
				// Stored as 16-bit unicode string
				name = r.readUnicodeStr16(true);
			} else {
				// Stored as 8-bit ascii string
				var len:uint = r.data.readUnsignedByte();
				name = r.data.readUTFBytes(len);
			}
			var l_currentSheet:Sheet;
			l_currentSheet = new Sheet();
			l_currentSheet.addEventListener("sizingSpreadsheet",onSizingSpreadsheetEvent);
			l_currentSheet.addEventListener("sizingSpreadsheetOver",onSizingSpreadsheetEvent);
			
			l_currentSheet.dateMode = dateMode;
			l_currentSheet.name = name;
			l_currentSheet.formats = currentSheet.formats.concat(globalFormats);
			l_currentSheet.xformats = currentSheet.xformats.concat(globalXFormats);
			_sheets.push(l_currentSheet);
		}
		
		// Formatting
		
		private function builtinfmtcount(r:Record, s:Sheet):void {
			var numBuildInFormats:uint = r.data.readUnsignedShort();
		}
		
		private function xf(r:Record, s:Sheet):void {
			var font:uint;
			var format:uint;
			switch(version) {
				case BIFFVersion.BIFF2:
					font = r.data.readUnsignedByte();
					r.data.position++;
					format = r.data.readUnsignedByte() & 0x3F;
					break;
				case BIFFVersion.BIFF3:
				case BIFFVersion.BIFF4:
					font = r.data.readUnsignedByte();
					format = r.data.readUnsignedByte();
					break;
				case BIFFVersion.BIFF5:
					font = r.data.readUnsignedShort();
					format = r.data.readUnsignedShort();
					break;
				case BIFFVersion.BIFF8:
					font = r.data.readUnsignedShort();
					format = r.data.readUnsignedShort();
					break;
			}
			
			if(s is Sheet) {
				s.xformats.push(new XFormat(r.type, format));
			} else {
				globalXFormats.push(new XFormat(r.type, format));
			}
		}
		
		private function efont(r:Record, s:Sheet):void {
			var color:uint = r.data.readUnsignedShort();
		}
		
		private function font(r:Record, s:Sheet):void {
			var height:Number = r.data.readUnsignedShort();
			var attributes:uint = r.data.readUnsignedShort();
			
			if(r.type == 0x231 || version >= BIFFVersion.BIFF5) {
				var colorIndex:uint = r.data.readUnsignedShort();
			}
			
			var len:uint;
			var name:String;
			if(r.type == 0x231 && version <= BIFFVersion.BIFF4) {
				len = r.data.readUnsignedByte();
				name = r.data.readUTFBytes(len);
			}
			
			
		}
		
		private function font2(r:Record, s:Sheet):void { }
		
		private function format(r:Record, s:Sheet):void {
			var index:uint = NaN;
			if(version == BIFFVersion.BIFF4) {
				r.data.position += 2;
			} else if (version == BIFFVersion.BIFF8 || version == BIFFVersion.BIFF5){
				index = r.data.readUnsignedShort();
			}
			
			var string:String;
			if(version == BIFFVersion.BIFF8){
				// Stored as 16-bit unicode string
				string = r.readUnicodeStr16();
			} else {
				// Stored as 8-bit ascii string
				var len:uint = r.data.readUnsignedByte();
				string = r.data.readUTFBytes(len);
			}
			
			if(s is Sheet) {
				isNaN(index) ? s.formats.push(string) : s.formats[index] = string;
			} else {
				isNaN(index) ? globalFormats.push(string) : globalFormats[index] = string;
			}
		}
		
		private function mergedcells(r:Record, s:Sheet):void { }
		
		private function palette(r:Record, s:Sheet):void { }
		
		private function style(r:Record, s:Sheet):void { }
		
		private function columndefault(r:Record, s:Sheet):void { }
		
		private function ixfe(r:Record, s:Sheet):void { }
		
		private function colinfo(r:Record, s:Sheet):void { }
		
		// Options
		
		private function codepage(r:Record, s:Sheet):void {
			var codepage					: uint = r.data.readUnsignedShort();
		}
		
		private function defcolwidth(r:Record, s:Sheet):void {
			var defColWidth					: uint = r.data.readUnsignedShort();
		}
		
		private function window1(r:Record, s:Sheet):void {
			var x							: uint 		= r.data.readUnsignedShort();
			var y							: uint 		= r.data.readUnsignedShort();
			var width						: uint 		= r.data.readUnsignedShort();
			var height						: uint 		= r.data.readUnsignedShort();
			var hidden						: Boolean 	= r.data.readUnsignedByte() == 1;
		}
		
		private function window2(r:Record, s:Sheet):void {
			if(r.type == 0x003E) {
				// BIFF2
				var showFormulas			: Boolean 	= r.data.readUnsignedByte() == 1;
				var showGridLines			: Boolean	= r.data.readUnsignedByte() == 1;
				var showRowColHeaders		: Boolean 	= r.data.readUnsignedByte() == 1;
				var frozen					: Boolean 	= r.data.readUnsignedByte() == 1;
				var showZeros				: Boolean 	= r.data.readUnsignedByte() == 1;
				var topRowVisible			: uint 		= r.data.readUnsignedShort();
				var leftColVisible			: uint 		= r.data.readUnsignedShort();
				var showHeadersDefaultColor	: Boolean 	= r.data.readUnsignedByte() == 1;
				var headerColor				: uint 		= r.data.readUnsignedInt();
			} else {
				// BIFF>2
				var options					: uint 		= r.data.readUnsignedShort();
				leftColVisible 							= r.data.readUnsignedShort();
				showHeadersDefaultColor 				= r.data.readUnsignedByte() == 1;
				headerColor 							= r.data.readUnsignedInt();
			}
		}
		
		private function backup(r:Record, s:Sheet):void {
			var saveBackup:Boolean = r.data.readUnsignedShort() == 1;
		}
		
		private function selection(r:Record, s:Sheet):void {
			// Meh. Who cares?
		}
		
		private function bof(r:Record, s:Sheet):void {
			var version			: uint 		= r.data.readUnsignedShort();
			var fileType		: uint 		= r.data.readUnsignedShort();
			var rupBuild		: uint 		= r.data.readUnsignedShort();
			var rupYear			: uint 		= r.data.readUnsignedShort();
			
			var tmp				: uint 		= r.data.readUnsignedInt();
			var fWin			: uint 		= tmp & 0x00001;
			tmp = tmp>>1;
			var fRisc			: uint 		= tmp & 0x00001;
			tmp = tmp>>1;
			var fBeta			: uint 		= tmp & 0x00001;
			tmp = tmp>>1;
			var fWinAny			: uint 		= tmp & 0x00001;
			tmp = tmp>>1;
			var fMacAny			: uint 		= tmp & 0x00001;
			tmp = tmp>>1;
			var fBetaAny		: uint 		= tmp & 0x00001;
			tmp = tmp>>1;
			var unused1			: uint 		= tmp & 0x00003;
			tmp = tmp>>2;
			var fRiscAny		: uint 		= tmp & 0x00001;
			tmp = tmp>>1;
			var fOOM			: uint 		= tmp & 0x00001;
			tmp = tmp>>1;
			var fGImp			: uint 		= tmp & 0x00001;
			tmp = tmp>>1;
			var unused2			: uint 		= tmp & 0x00003;
			tmp = tmp>>2;
			var fFontLimit		: uint 		= tmp & 0x00001;
			tmp = tmp>>1;
			var verXLHigh		: uint 		= tmp & 0x0000F;
			tmp = tmp>>4;
			var unused3			: uint 		= tmp & 0x00001;
			tmp = tmp>>0;
			var reserved1		: uint 		= tmp & 0x01FFF;
			
			tmp						   		= r.data.readUnsignedInt();
			var verLowestBiff	: uint 		= tmp & 0x000FF;
			tmp = tmp>>8;
			var verXLSaved		: uint 		= tmp & 0x0000F;
			tmp = tmp>>4;
			var reserved2		: uint 		= tmp & 0x0FFFFF;
			
			currentFileType = fileType;
			if(fileType == SHEET) {
				if(_sheets.length == 0) {
					var newSheet: Sheet 	= new Sheet();
					newSheet.dateMode 		= dateMode;
					_sheets.push(newSheet);
					newSheet.formats 		= newSheet.formats.concat(globalFormats);
					newSheet.xformats 		= newSheet.xformats.concat(globalXFormats);
					newSheet.addEventListener("sizingSpreadsheet",onSizingSpreadsheetEvent);
					newSheet.addEventListener("sizingSpreadsheetOver",onSizingSpreadsheetEvent);
				}
				notes 						= new Array();
				currentSheet 				= _sheets[currentSheetIdx] as Sheet;
				currentSheetIdx++;
			}
			if(r.type == 0x9) {
				this.version 				= BIFFVersion.BIFF2;
			} else if(r.type == 0x209) {
				this.version 				= BIFFVersion.BIFF3;
			} else if(r.type == 0x409) {
				this.version 				= BIFFVersion.BIFF4;
			} else if(r.type == 0x809 && version != 0x600) {
				this.version = BIFFVersion.BIFF5;
			} else if(fileType == 0x05) {
				this.version 				= BIFFVersion.BIFF8;
			}
		}
		
		private function calcmode(r:Record, s:Sheet):void {
			var mode		: uint 			= r.data.readUnsignedShort();
		}
		
		private function calccount(r:Record, s:Sheet):void {
			var iterations	: uint 			= r.data.readUnsignedShort();
		}
		
		private function refmode(r:Record, s:Sheet):void {
			var mode		: uint 			= r.data.readUnsignedShort();
		}
		
		private function iteration(r:Record, s:Sheet):void {
			var iteration	: Boolean 		= r.data.readUnsignedShort() == 1;
		}
		
		private function delta(r:Record, s:Sheet):void {
			var delta		: Number 		= r.data.readDouble();
		}
		
		private function precision(r:Record, s:Sheet):void {
			var fullPrecision: Boolean 		= r.data.readUnsignedShort() == 1;
		}
		
		private function printgridlines(r:Record, s:Sheet):void {
			var printGridlines: Boolean 	= r.data.readUnsignedShort() == 1
		}
		
		private function defrowheight(r:Record, s:Sheet):void {
			var defRowHeight	: Number 	= r.data.readUnsignedShort();
		}
		
		private function datemode(r:Record, s:Sheet):void {
			var baseDate		: uint 		= r.data.readUnsignedShort();
			dateMode 						= baseDate;
			for(var n:uint = 0; n < _sheets.length; n++) {
				_sheets[n].dateMode = dateMode;
			}
		}
		
		private function country(r:Record, s:Sheet):void {
			var excelCountry	: uint 		= r.data.readUnsignedShort();
			var systemCountry	: uint 		= r.data.readUnsignedShort();
		}
		
		private function saverecalc(r:Record, s:Sheet):void {
			var recalculateBeforeSave : Boolean = r.data.readUnsignedByte() == 1;
		}
		
		private function wsbool(r:Record, s:Sheet):void {
			var options			: uint 		= r.data.readUnsignedShort();
		}
		
		private function bookbool(r:Record, s:Sheet):void { }
		
		private function dv(r:Record, s:Sheet):void { }
		
		private function pane(r:Record, s:Sheet):void { }
		
		private function uncalced(r:Record, s:Sheet):void {
			// Indicates that formulas were not recalculated before the sheet was saved
		}
		
		// Page layout
		private function header(r:Record, s:Sheet):void {
			var string:String;
			if(r.data.length == 0) {
				string = "";
			} else if(version == BIFFVersion.BIFF8) {
				string = r.readUnicodeStr16();
			} else {
				var len:uint = r.data.readUnsignedByte();
				string = r.data.readUTFBytes(len);
			}
			s.header = string;
		}
		
		private function footer(r:Record, s:Sheet):void {
			if(r.data.bytesAvailable == 0 ) {
				return;
			}
			var len:uint = r.data.readUnsignedByte();
			var string:String = r.data.readUTFBytes(len).substr(2); // Skip two bytes b/c of commands (left, center, etc)
			s.footer = string;
		}
		
		private function vertpagebreaks(r:Record, s:Sheet):void { }
		
		private function horizpagebreaks(r:Record, s:Sheet):void { }
		
		private function leftmargin(r:Record, s:Sheet):void {
			var size:Number = r.data.readDouble();
		}
		
		private function rightmargin(r:Record, s:Sheet):void {
			var size:Number = r.data.readDouble();
		}
		
		private function topmargin(r:Record, s:Sheet):void {
			var size:Number = r.data.readDouble();
		}
		
		private function bottommargin(r:Record, s:Sheet):void {
			var size:Number = r.data.readDouble();
		}
		
		private function printheaders(r:Record, s:Sheet):void {
			var printHeaders:Boolean = r.data.readUnsignedShort() == 1;
		}
		
		private function setup(r:Record, s:Sheet):void {
			var paperSize:uint = r.data.readUnsignedShort();
			var scaleFactor:uint = r.data.readUnsignedShort();
			var startPageNumber:uint = r.data.readUnsignedShort();
			var maxWidthInPages:uint = r.data.readUnsignedShort();
			var maxHeightInPages:uint = r.data.readUnsignedShort();
			
			var optionFlags:uint;
			if(version == BIFFVersion.BIFF4) {
				optionFlags = r.data.readUnsignedShort();
			} else {
				optionFlags = r.data.readUnsignedShort();
				var printDPI:uint = r.data.readUnsignedShort();
				var verticalPrintDPI:uint = r.data.readUnsignedShort();
				var headerMargin:Number = r.data.readDouble();
				var footerMargin:Number = r.data.readDouble();
				var copies:uint = r.data.readUnsignedShort();
			}
		}
		
		private function gcw(r:Record, s:Sheet):void {
			// Global column width
			var bitfieldSize:uint = r.data.readUnsignedShort();
		}
		
		private function guts(r:Record, s:Sheet):void {
			var rowOutlineWidth:uint = r.data.readUnsignedShort();
			var colOutlineHeight:uint = r.data.readUnsignedShort();
			var visibleRowLevels:uint = r.data.readUnsignedShort();
			var visibleColLevels:uint = r.data.readUnsignedShort();
		}
		
		private function hcenter(r:Record, s:Sheet):void {
			// 0 = left align, 1 = centered
			var center:uint = r.data.readUnsignedShort();
		}
		
		private function vcenter(r:Record, s:Sheet):void {
			// 0 = top align, 1 = centered
			var center:uint = r.data.readUnsignedShort();
		}
		
		private function hideobj(r:Record, s:Sheet):void {
			/*
			0 = Show objects
			1 = Show placeholders
			2 = Hide objects
			*/
			var viewingMode:uint = r.data.readUnsignedShort();
		}
		
		private function gridset(r:Record, s:Sheet):void {
			var printGridLinesOptionEverChanged:Boolean = r.data.readUnsignedByte() == 1;
		}
		
		private function colwidth(r:Record, s:Sheet):void { }
		
		
		
		// Security
		private function protect(r:Record, s:Sheet):void { }
		
		private function password(r:Record, s:Sheet):void { }
		
		private function rangeprotection(r:Record, s:Sheet):void { }
		
		private function sheetprotection(r:Record, s:Sheet):void { }
		
		private function windowprotect(r:Record, s:Sheet):void {
			var windowsProtected:Boolean = r.data.readUnsignedShort() == 1;
		}
		
		private function filesharing(r:Record, s:Sheet):void { }
		
		private function writeaccess(r:Record, s:Sheet):void {
			var len:uint = r.data.readUnsignedByte();
			var username:String = r.data.readUTFBytes(len);
		}
		
		private function writeprot(r:Record, s:Sheet):void { }
		
		private function filepass(r:Record, s:Sheet):void { }
		
		private function objectprotect(r:Record, s:Sheet):void { }
		
		
		// Junk that might be usedful that I don't really need
		
		private function row(r:Record, s:Sheet):void {
			var rowNum:uint = r.data.readUnsignedShort();
			var firstCol:uint = r.data.readUnsignedShort();
			var lastCol:uint = r.data.readUnsignedShort()-1;
			var rowHeight:uint = r.data.readUnsignedShort();
			var microsoftUse:uint = r.data.readUnsignedShort();
			var defaultCellAttrs:Boolean = r.data.readUnsignedByte() == 1;
			var cellRecordsOffset:uint = r.data.readUnsignedShort();
			
			// Last three bytes are default cell attributes
		}
		
		private function index(r:Record, s:Sheet):void {
			// We don't need no stinkin' index
		}
		
		private function string(r:Record, s:Sheet):void { }
		
		private function eof(r:Record, s:Sheet):void {
			// Yay! Done!
		}
		
		private function extsst(r:Record, s:Sheet):void { }
		
		private function uselfs(r:Record, s:Sheet):void { }
		
		private function phonetic(r:Record, s:Sheet):void { }
		
		private function dsf(r:Record, s:Sheet):void { }
		
		private function dbcell(r:Record, s:Sheet):void { }
		
		private function array(r:Record, s:Sheet):void {
			var firstRow:uint = r.data.readUnsignedShort();
			var lastRow:uint = r.data.readUnsignedShort();
			var firstCol:uint = r.data.readUnsignedByte();
			var lastCol:uint = r.data.readUnsignedByte();
			
			var alwaysRecalculate:Boolean;
			if(version == BIFFVersion.BIFF2) {
				alwaysRecalculate = r.data.readUnsignedByte() == 0x01;
			} else {
				alwaysRecalculate = (r.data.readUnsignedShort() & 0x0001) == 0x0001;
			}
			
			if(version == BIFFVersion.BIFF8) {
				r.data.position += 4;
			}
			
			var tokens:ByteArray = new ByteArray();
			r.data.readBytes(tokens, 0, r.data.bytesAvailable);
		}
		
		
		
		
		// Haven't the slightest
		private function externcount(r:Record, s:Sheet):void { }
		
		private function externsheet(r:Record, s:Sheet):void { }
		
		private function name(r:Record, s:Sheet):void { }
		
		private function boolerr(r:Record, s:Sheet):void { }
		
		private function externname(r:Record, s:Sheet):void { }
		
		private function xct(r:Record, s:Sheet):void { }
		
		private function crn(r:Record, s:Sheet):void { }
		
		private function dconref(r:Record, s:Sheet):void { }
		
		private function tableop(r:Record, s:Sheet):void { }
		
		private function tableop2(r:Record, s:Sheet):void { }
		
		private function cont(r:Record, s:Sheet):void {
			switch(lastRecordType) {
				case 0xEC:
					var flags:uint = r.data.readUnsignedByte();
					var note:String = r.data.readUTFBytes(r.data.bytesAvailable);
					notes.push(note);
					break;
				default:
					break;
			}
		}
	}
}