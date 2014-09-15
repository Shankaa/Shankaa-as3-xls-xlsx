package org.as3.ms.ooxml
{
	import flash.events.Event;
	import flash.events.EventDispatcher;
	import flash.events.ProgressEvent;
	import flash.utils.ByteArray;
	
	import org.as3.ms.events.XLSEvent;
	import org.as3.ms.xls.ExcelFile;
	import org.as3.ms.xls.Sheet;
	import org.as3commons.zip.Zip;
	import org.as3commons.zip.ZipFile;
	

	/**
	 * 
	 * @author Hugues Sansen, Shankaa
	 * 
	 * 
	 * This is a version limited to xlsx of the implementation of Office Open XML 
	 * 
	 * 
	 * http://en.wikipedia.org/wiki/Office_Open_XML
	 * Office Open XML (also informally known as OOXML or OpenXML) is a zipped, XML-based file format developed by Microsoft
	 * for representing spreadsheets, charts, presentations and word processing documents. The format was initially standardised
	 * by Ecma (as ECMA-376) and, in later versions, by ISO and IEC (as ISO/IEC 29500).
	 * Starting with Microsoft Office 2007, the Office Open XML file formats have become the default target file format of Microsoft Office.
	 * Microsoft Office 2010 provides read support for ECMA-376, read/write support for ISO/IEC 29500 Transitional, and read support
	 * for ISO/IEC 29500 Strict. Microsoft Office 2013 additionally supports both reading and writing of ISO/IEC 29500 Strict.
	 * 
	 * 
	 * requires http://www.as3commons.org/as3-commons-zip library
	 * caution :
	 * Limitations:
	 * 		* ZIP feature versions > 2.0 are not supported
	 * 		* ZIP archives containing data descriptor records are not supported.
	 * 		* If running in the Flash Player browser plugin, Zip requires ZIPs to be patched (Adler32 checksums need to be added).
	 * 		This is not required if Zip runs in the Adobe AIR runtime or if files contained in the ZIP are not compressed.
	 * 
	 */
	public class OfficeOpenXML
	{
		private static const DATE_0				: Date = new Date(1899,11,30);
		
		/**
		 * ECMA-376, 4th Edition
		 * Office Open XML File Formats — Fundamentals and Markup Language Reference
		 * 
		 * 18.8.30 numFmt (Number Format)
		 * 
		 * ID formatCode
		 * 0 General
		 * 1 0
		 * 2 0.00
		 * 3 #,##0
		 * 4 #,##0.00
		 * 9 0%
		 * 10 0.00%
		 * 11 0.00E+00
		 * 12 # ?/?
		 * 13 # ??/??
		 * 14 mm-dd-yy
		 * 15 d-mmm-yy
		 * 16 d-mmm
		 * 17 mmm-yy
		 * 18 h:mm AM/PM
		 * 19 h:mm:ss AM/PM
		 * 20 h:mm
		 * 21 h:mm:ss
		 * 22 m/d/yy h:mm
		 * 37 #,##0 ;(#,##0)
		 * 38 #,##0 ;[Red](#,##0)
		 * 39 #,##0.00;(#,##0.00)
		 * 40 #,##0.00;[Red](#,##0.00)
		 * 45 mm:ss
		 * 46 [h]:mm:ss
		 * 47 mmss.0
		 * 48 ##0.0E+0
		 * 49 @ 
		 */
		private static const DEFAULT_NUMBER_FORMAT_ALL_LANGUAGES	: Array = ["General","0","0.00","#,##0","#,##0.00",
			null,null,null,null,"0%","# ?/?","# ??/??","mm-dd-yy",
			"d-mmm-yy","d-mmm","mmm-yy","h:mm AM/PM","h:mm:ss AM/PM",
			"h:mm","h:mm:ss","m/d/yy h:mm",null,null,
			null,null,"#,##0 ;(#,##0)","#,##0 ;[Red](#,##0)","#,##0.00;(#,##0.00)",
			"#,##0.00;[Red](#,##0.00)",null,null,null,null,
			"mm:ss","[h]:mm:ss","mmss.0","##0.0E+0","@" ];
		
		private static const DEFAULT_NUMERICAL_FORMATS				: Array = [0,1,2,3,4,9,10,11,12,13,37,38,39,40,48,49];
		private static const DEFAULT_DATE_FORMATS					: Array	= [14,15,16,17,22];
		private static const DEFAULT_TIME_FORMATS					: Array = [18,1,20,21,45,46,47];
		
		private var zip												: Zip	= new Zip()
		
		/**
		 * [Content_Types].xml
		 * This file provided MIME type information for parts of the package,
		 * using defaults for certain file extensions and overrides for parts specificied by IRI.
		 */
		private var content_Types_xml								: XML;
		
		/**
		 * built with [Content_Types].xml
		 * gives the file structure inside the zip document
		 */
		private var documentStructure 								: XML;
		
		public var isOfficeOpenXML									: Boolean;
		
		
		
		/**
		 * 
		 * unzips an OOXML document,
		 * if a requester is not null sends Events when unzip starts("unpackingOOXMLFile"), is over("unpackingOOXMLFileOver")
		 * or fails ("unpackingOOXMLFileFailed").
		 * 
		 * @param xlsxStream
		 * @param requester
		 * @return 
		 * 
		 */
		public static function open(xlsxStream:ByteArray,requester:EventDispatcher= null):OfficeOpenXML{
			var instance : OfficeOpenXML =   new OfficeOpenXML();
			if(requester){
				requester.dispatchEvent(new Event("unpackingOOXMLFile"));
			}
			try{
				instance.zip.loadBytes(xlsxStream);
				instance.isOfficeOpenXML = true;
				if(requester){
					requester.dispatchEvent(new Event("unpackingOOXMLFileOver"));
				}
			} catch (error:org.as3commons.zip.ZipErrorEvent){
				if(requester){
					requester.dispatchEvent(new Event("unpackingOOXMLFileFailed"));
				}
				instance.isOfficeOpenXML = false;
			}
			return instance;
		}
		
		private function analyze():void{
			var zipFile 			: ZipFile 	= zip.getFileByName("[Content_Types].xml");
			var content_Types_xmlba : ByteArray = zipFile.content;
			content_Types_xml = new XML(content_Types_xmlba);
			getDocumentStructure(content_Types_xml);
		}
		
		
		private function getDocumentStructure(contentTypes:XML):XML{
			var documentStructure 	: XML = <root/>;
			var partName			: String;
			var struct				: Array;
			var branch				: XML;
			var fileNameA			: Array;
			var fileName			: String;
			var fileSuffix			: String;
			for each(var override : XML in contentTypes.Override){
				branch = documentStructure;
				partName = override.@PartName;
				struct = partName.split("/");
				for each (var n : String in struct){
					fileNameA = n.split(".");
					fileName = fileNameA[0];
					if(fileNameA.length == 1){
						if(branch.hasOwnProperty(fileName)){
							branch = branch[fileName][0];
						} else {
							branch.appendChild(<{fileName}/>);
						}
					} else {
						fileSuffix = fileNameA[1];
						branch.appendChild(<{fileName} suffix={fileSuffix}/>);
					}
				}
			}
			
			return documentStructure;
		}
		
		public function getWorkBook():XML{
			return new XML(zip.getFileByName("xl/workbook.xml").content);
		}
		
		public function getWorkBookRels():XML{
			return new XML(zip.getFileByName("xl/_rels/workbook.xml.rels").content);
		}
		
		public function getSharedStrings():XML{
			return new XML(zip.getFileByName("xl/sharedStrings.xml").content);
		}
		
		public function getSheetNames():Array{
			var names		: Array = [];
			var workbook : XML = getWorkBook();
			for each (var sheet : XML in workbook.xl.worksheets.sheets.sheet){
				names.push(sheet.@name.toString());
			}
			return names;
		}
		
		public function getXLStyles():XML{
			return XML(zip.getFileByName("xl/styles.xml").content);
		}
		
		public function getSheet(sheetDescription:XML,workbook:XML):XML{
			var ns	: String = workbook.namespace("r");
			default xml namespace = ns;
			var rId : String = sheetDescription.@id;
			var rels : XML = getWorkBookRels();
			ns = rels.namespace();
			default xml namespace = ns;
			var relationship : XML = rels.Relationship.(@Id == rId)[0];
			var target	: String = relationship.@Target;
			return XML(zip.getFileByName("xl/"+target).content);
		}
		
		public function getSheetNamed(name:String):XML{
			var workbook 	: XML 	= getWorkBook();
			try{
				var sheet		: XML	= workbook.xl.worksheets.sheets.sheet.(@name == name)[0];
				var index		: int	= sheet.childIndex();
				return XML(zip.getFileByName("xl/worksheets/sheet"+index+".xml").content);
			} catch (error:Error){
				
			}
			return null;
		}
		
		public function ooxmlBuildExcelFile(excelFile	: ExcelFile):void {
			var workbook				: XML 		= getWorkBook();
			var workbookNs				: String	= workbook.namespace().toString();
			var workbookRelationshipNs	: String	= workbook.namespace("r").toString();
			
			var sheetIndex				: int 		= 1;
			var xSheet					: XML;
			var sharedStrings			: XML		= getSharedStrings();
			var sharedStringTable 		: Array 	= getSharedStringTable(sharedStrings);
			var xlStyles				: XML		= getXLStyles();
			var sheet					: Sheet;
			var dimensions				: Array		= [];
			var dimension 				: *	;
			var workBookNumberOfCells	: uint		= 0;
			var allocatedCells			: Number	= 0;
			var sheetDescription 		: XML;
			
			
			default xml namespace	= workbookNs;
			
			for each (sheetDescription in workbook.sheets.sheet){
				xSheet							= getSheet(sheetDescription,workbook);
				dimension						= getSheetDimension(xSheet,workbookNs);
				dimensions.push(dimension);
				workBookNumberOfCells 					+= dimension.numberOfCells;
			}
			var i					: int		= 0;
			sheetIndex							= 1;
			for each (sheetDescription in workbook.sheets.sheet){
				xSheet							= getSheet(sheetDescription,workbook);
				sheet							= new Sheet();
				excelFile.sheets.push(sheet);
				allocatedCells = ooxmlBuildSheet(excelFile,xlStyles,sheet,xSheet,sharedStringTable,dimensions[i++],workBookNumberOfCells,allocatedCells);
			}
			excelFile.dispatchEvent(new ProgressEvent("excelComplete",false,false,allocatedCells,workBookNumberOfCells));
		}
		
		
		/**
		 * does not include neither formulas nor styles
		 * 
		 * builds a sheet from the XML definitions
		 * 
		 * see http://officeopenxml.com/
		 *  
		 * @param sheet
		 * @param xSheet
		 * @param sharedStringTable
		 * @param dimension
		 * @param workBookNumberOfCells
		 * @param allocatedCells
		 * 
		 */
		private function ooxmlBuildSheet(excelFile:ExcelFile,xlStyles:XML,sheet : Sheet, xSheet : XML,sharedStringTable : Array,dimension : *,workBookNumberOfCells: uint,allocatedCells: Number):Number{
			var dimension				: *;
			var rowIndex				: int;
			var columnIndex				: int;
			var position				: *;
			
			var cellRef					: String;
			var cellType				: String;
			var cellStyle				: Number;
			var cellVm					: Number;
			
			var value					: String;
			var formula					: String;
			var isRichTextInLine		: String;
			var extLst					: String;
			var date					: Date;
			
			var s						: String;
			
			var progress				: int;
			
			var sheetNs					: String 	= xSheet.namespace().toString();
			
			
			
			excelFile.dispatchEvent(new XLSEvent("sizingSpreadsheet",dimension.rows,dimension.columns));
			sheet.resize(dimension.rows,dimension.columns);
			excelFile.dispatchEvent(new XLSEvent("sizingSpreadsheetOver",dimension.rows,dimension.columns));
			rowIndex = 0;
			default xml namespace					= sheetNs;
			for each (var xRow : XML in xSheet.sheetData.row){
				columnIndex = 0;
				for each (var xCell : XML in xRow.c){
					++allocatedCells;
					/*
					 An A1 style reference to the location of this cell
					 The possible values for this attribute are defined by the ST_CellRef simple type
					 (§18.18.7).
					*/
					cellRef				= xCell.@r;
					/*
					 The index of this cell's style. Style records are stored in the Styles Part.
					 The possible values for this attribute are defined by the W3C XML Schema unsignedInt
					 datatype.
					*/
					cellStyle			= Number(xCell.@s);
					/* 
					 An enumeration representing the cell's data type.
					 The possible values for this attribute are defined by the ST_CellType simple type
					 (§18.18.11).
					*/
					cellType 			= xCell.@t;
					/*
					 The zero-based index of the value metadata record associated with this cell's value.
					 Metadata records are stored in the Metadata Part. Value metadata is extra information
					 stored at the cell level, but associated with the value rather than the cell itself. Value
					 metadata is accessible via formula reference.
					 The possible values for this attribute are defined by the W3C XML Schema unsignedInt
					 datatype.
					*/
					cellVm				= Number(xCell.@vm);
					/*
					Child Elements
					extLst (Future Feature Data Storage Area)
					f (Formula)
					is (Rich Text Inline)
					v (Cell Value)
					*/			
					
					extLst				= xCell.extLst.toString();
					position 			= getCellPosition(cellRef);
					
					/*
					The make-up of a cell is important in understanding the overall architecture of the spreadsheet content.
					Each cell specifies its type with the t attribute.
					Possible values include:
					b for boolean
					d for date
					e for error
					inlineStr for an inline string (i.e., not stored in the shared strings part, but directly in the cell)
					n for number
					s for shared string (so stored in the shared strings part and not in the cell)
					str for a formula (a string representing the formula)
					*/
					switch(cellType){
						case "b":
							value 				= xCell.v.toString();
							sheet.setCell(position.row,position.column,value == "true");
							break;
						case "d":
							value 				= xCell.v.toString();
							date = computeDate(Number(value),cellStyle,xlStyles);
							sheet.setCell(position.row,position.column,date);
							break;
						case "e":
							value 				= xCell.v.toString();
							sheet.setCell(position.row,position.column,value);
							break;
						case "inlineStr":
							value 				= xCell.elements("is").t.toString();
							//For inline strings, the value is within an <is> element. But of course the actual text is further nested within a t since the text can be formatted.
							sheet.setCell(position.row,position.column,value);
							break;
						case "n":
							value 				= xCell.v.toString();
							sheet.setCell(position.row,position.column,Number(value));
							break;
						case "s":
							value 				= xCell.v.toString();
							s = sharedStringTable[value];
							sheet.setCell(position.row,position.column,s);
							break;
						case "str":
							//formula
							value 				= xCell.v.toString();
							formula				= xCell.f.toString();
							sheet.setCell(position.row,position.column,Number(value));
							//new Formula(position.row,position.column,
							break;
						default :
							value 				= xCell.v.toString();
							var num : Number = Number(value);
							//computeValueFormat(value,cellStyle,xlStyles);
							sheet.setCell(position.row,position.column,num);
							break;
					}
					++columnIndex;
					progress = allocatedCells/workBookNumberOfCells *1000;
					if(progress%10 == 0){
						excelFile.dispatchEvent(new ProgressEvent("excelProgess",false,false,allocatedCells,workBookNumberOfCells));
					}
				}
				++rowIndex;
				
			}
			return allocatedCells;
		}
		
		
		private function computeDate(value:Number,styleValue:Number,xlStyles:XML):Date{
			var date : Date = new Date(DATE_0);
			date.date += value;
			return date;
		}
		
		private function computeValueFormat(value:Number,styleValue:Number,xlStyles:XML=null):*{
			var result 	: *;
			var xf		: XML;
			var numFmt	: XML;
			var date	: Date;
			if(DEFAULT_NUMERICAL_FORMATS.indexOf(styleValue)>-1){
				return value;
			} else if(DEFAULT_DATE_FORMATS.indexOf(styleValue)>-1){
				date = new Date(DATE_0);
				date.date += value;
				result = date;
			} else if(DEFAULT_TIME_FORMATS.indexOf(styleValue)>-1){
				
				result = value;
			} else {
				//xf 		= xlStyles.cellXfs.xf[styleValue];
				//numFmt 	= xlStyles.numFmts.numFmt.(@numFmtId.toString() == xf.@numFmtId.toString())[0];
				//
				result = value;
			}
			return result;
		}
		
		
		
		private function getSharedString(sharedStrings	: XML,stringIndex : int):String{
			return sharedStrings.si[stringIndex].t.toString();
		}
		
		private function getSharedStringTable(sharedStrings: XML):Array{
			var sharedStringTable 	: Array = [];
			var index				: uint = 0;
			var ns					: String = sharedStrings.namespace().toString();
			default xml namespace	= ns;
			for each (var sharedString : XML in sharedStrings.si){
				sharedStringTable[index++] = sharedString.t.toString();
			}
			return sharedStringTable;
		}
		
		/**
		 * we assume that the dimension mentionned in the <dimension ref=""/> field of the sheet is correct.
		 * It may not. In some cases (files generated by none MS software?), all the sheets have the biggest sheet dimmension.
		 * 
		 * 
		 * @param xSheet
		 * @param ns the namespace of the xSheet
		 * @return an object with rows, columns and numberOfCells attributes
		 * 
		 */
		private function getSheetDimension(xSheet : XML, ns:String):*{
			default xml namespace = ns;
			var dimension	: String	= xSheet.dimension.@ref.toString();
			
			var dimensionObject : * 	= {rows:0,columns:0};
			var regExp		: RegExp	= /\d+/g;
			var row0		: String 	= regExp.exec(dimension)[0];
			var rown		: String	= regExp.exec(dimension)[0];
			dimensionObject.rows		= Number(rown)-Number(row0)+1
			regExp						= /[a-z]+/gi;
			var col0		: String 	= regExp.exec(dimension)[0];
			var coln		: String	= regExp.exec(dimension)[0];
			col0						= col0.toLowerCase();
			coln						= coln.toLowerCase();
			
			regExp						= /[a-z]/gi;
			var a			: int		= "a".charCodeAt(0);
			var col0Length	: int		= col0.length;
			var colnLength	: int		= coln.length;
			var index		: int		= 0;
			var cols0		: int		= 0;
			var colsn		: int		= 0;
			while(index<col0Length){
				cols0 = (cols0*26)+(col0.charCodeAt(index++)-a);
			}
			index = 0;
			while(index<colnLength){
				colsn = (colsn*26)+(coln.charCodeAt(index++)-a);
			}
			dimensionObject.columns = colsn - cols0 +1;
			dimensionObject.numberOfCells = dimensionObject.rows*dimensionObject.columns;
			return dimensionObject;
		}
		
		private function transformColumnLettersToNumber(Cc:String):int{
			var index 		: int		= 0;
			var cc			: String	= Cc.toLowerCase();
			var ccLength	: int		= cc.length;
			var a			: int		= "a".charCodeAt(0);
			var col			: int;
			while(index<ccLength){
				col = (col*26)+(cc.charCodeAt(index++)-a);
			}
			return col;
		}
		
		private function getCellPosition(r:String):*{
			var position	: * = {};
			var regExp		: RegExp	= /\d+/g;
			var row			: String 	= regExp.exec(r)[0];
			regExp						= /[a-z]+/gi;
			var col			: String 	= regExp.exec(r)[0];
			//in Excel rows are counted from 1
			position.row				= int(row)-1;
			//columns are counted from A (we put A = 0)
			position.column				= int(transformColumnLettersToNumber(col));
			
			return position;
		}
		
		private function getNumRows(dimension : String):uint{
			var dim				: Array 	= dimension.split(":");
			var topLeft 		: String 	= dim[0];
			var bottomRight 	: String 	= dim[1];
			var topLeftPosition : *			= getCellPosition(topLeft);
			var bottomRightPosition : *		= getCellPosition(bottomRight);
			
			var numRows			: int		= bottomRightPosition.row - topLeftPosition.row +1;
			
			return numRows;
		}
		
		private function getNumCols(dimension : String):uint{
			var dim				: Array 	= dimension.split(":");
			var topLeft 		: String 	= dim[0];
			var bottomRight 	: String 	= dim[1];
			var topLeftPosition : *			= getCellPosition(topLeft);
			var bottomRightPosition : *		= getCellPosition(bottomRight);
			
			var numCols			: int		=  bottomRightPosition.column - topLeftPosition.column +1;
			
			return numCols;
		}
	}
}