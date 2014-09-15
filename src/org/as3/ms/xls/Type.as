package org.as3.ms.xls {
	/**
	 * Used to represent BIFF Records
	 */
	public class Type {
		//All Records in BIFF2
		public static const DIMENSION		: uint = 0x0000;
		public static const DIMENSION_3458	: uint = 0x0200;
		public static const BLANK			: uint = 0x0001;
		public static const BLANK_3458		: uint = 0x0201;
		public static const INTEGER			: uint = 0x0002;
		public static const NUMBER			: uint = 0x0003;
		public static const NUMBER_3458		: uint = 0x0203;
		public static const LABEL			: uint = 0x0004;
		public static const LABEL_3458		: uint = 0x0204;
		public static const BOOLERR			: uint = 0x0005;
		public static const BOOLERR_3458	: uint = 0x0205;
		public static const FORMULA			: uint = 0x0006;
		public static const FORMULA_3		: uint = 0x0206;
		public static const FORMULA_4		: uint = 0x0406;
		public static const STRING			: uint = 0x0007;
		public static const STRING_3458		: uint = 0x0207;
		public static const ROW				: uint = 0x0008;
		public static const ROW_3458		: uint = 0x0208;
		public static const BOF				: uint = 0x0009;
		public static const BOF_3			: uint = 0x0209;
		public static const BOF_4			: uint = 0x0409;
		public static const BOF_58			: uint = 0x0809;
		public static const EOF				: uint = 0x000A;
		public static const INDEX			: uint = 0x000B;
		public static const INDEX_3458		: uint = 0x020B;
		public static const CALCCOUNT		: uint = 0x000C;
		public static const CALCMODE		: uint = 0x000D;
		public static const PRECISION		: uint = 0x000E;
		public static const REFMODE			: uint = 0x000F;
		public static const DELTA			: uint = 0x0010;
		public static const ITERATION		: uint = 0x0011;
		public static const PROTECT			: uint = 0x0012;
		public static const PASSWORD		: uint = 0x0013;
		public static const HEADER			: uint = 0x0014;
		public static const FOOTER			: uint = 0x0015;
		public static const EXTERNCOUNT		: uint = 0x0016;
		public static const EXTERNSHEET		: uint = 0x0017;
		public static const DEFINEDNAME		: uint = 0x0018;
		public static const DEFINEDNAME_34	: uint = 0x0218;
		public static const WINDOWPROTECT	: uint = 0x0019;
		public static const VERTICALPAGEBREAKS: uint = 0x001A;
		public static const HORIZONTALPAGEBREAKS: uint = 0x001B;
		public static const NOTE			: uint = 0x001C;
		public static const SELECTION		: uint = 0x001D;
		public static const FORMAT			: uint = 0x001E;
		public static const FORMAT_458		: uint = 0x041E;
		public static const BUILTINFMTCOUNT	: uint = 0x001F;
		public static const BUILTINFMTCOUNT_34: uint = 0x0056;
		public static const COLUMNDEFAULT	: uint = 0x0020;
		public static const ARRAY			: uint = 0x0021;
		public static const ARRAY_3458		: uint = 0x0221;
		public static const DATEMODE		: uint = 0x0022;
		public static const EXTERNALNAME	: uint = 0x0023;
		public static const EXTERNALNAME_34	: uint = 0x0223;
		public static const COLWIDTH		: uint = 0x0024;
		public static const DEFAULTROWHEIGHT: uint = 0x0025;
		public static const DEFAULTROWHEIGHT_3458: uint = 0x0225;
		public static const LEFTMARGIN		: uint = 0x0026;
		public static const RIGHTMARGIN		: uint = 0x0027;
		public static const TOPMARGIN		: uint = 0x0028;
		public static const BOTTOMMARGIN	: uint = 0x0029;
		public static const PRINTHEADERS	: uint = 0x002A;
		public static const PRINTGRIDLINES	: uint = 0x002B;
		public static const FILEPASS		: uint = 0x002F;
		public static const FONT			: uint = 0x0031;
		public static const FONT_34			: uint = 0x0231;
		public static const FONT2			: uint = 0x0032;
		public static const DATATABLE		: uint = 0x0036;
		public static const DATATABLE_3458	: uint = 0x0236;
		public static const DATATABLE2		: uint = 0x0037;
		public static const CONTINUE		: uint = 0x003C;  //Whenever the content of a record exceeds the given limits (see table), the record must be split. Several CONTINUE records containing the additional data are added after the parent recor
		public static const WINDOW1			: uint = 0x003D;
		public static const WINDOW2			: uint = 0x003E;
		public static const WINDOW2_3458	: uint = 0x023E;
		public static const BACKUP			: uint = 0x0040;
		public static const PANE			: uint = 0x0041;
		public static const CODEPAGE		: uint = 0x0042;
		public static const XF				: uint = 0x0043;
		public static const XF_3			: uint = 0x0243;
		public static const XF_4			: uint = 0x0443;
		public static const XF_58			: uint = 0x00E0;
		public static const IXFE			: uint = 0x0044;
		public static const FONTCOLOR		: uint = 0x0045;
		public static const PLS				: uint = 0x004D;
		public static const DCONREF			: uint = 0x0051;
		public static const DEFCOLWIDTH		: uint = 0x0055;
		
		//New Records in BIFF3
		public static const XCT				: uint = 0x0059;
		public static const CRN				: uint = 0x005A;
		public static const FILESHARING		: uint = 0x005B;
		public static const WRITEACCESS		: uint = 0x005C;
		public static const UNCALCED		: uint = 0x005E;
		public static const SAVERECALC		: uint = 0x005F;
		public static const OBJECTPROTECT	: uint = 0x0063;
		public static const COLINFO			: uint = 0x007D;
		public static const RK				: uint = 0x027E;
		public static const GUTS			: uint = 0x0080;
		public static const SHEETPR			: uint = 0x0081;
		public static const GRIDSET			: uint = 0x0082;
		public static const HCENTER			: uint = 0x0083;
		public static const VCENTER			: uint = 0x0084;
		public static const WRITEPROT		: uint = 0x0086;
		public static const COUNTRY			: uint = 0x008C;
		public static const HIDEOBJ			: uint = 0x008D;
		public static const PALETTE			: uint = 0x0092;
		public static const STYLE			: uint = 0x0293;
		
		
		//New Records in BIFF4
		public static const STANDARDWIDTH	: uint = 0x0099;
		public static const SCL				: uint = 0x00A0;
		public static const PAGESETUP		: uint = 0x00A1;
		public static const GCW				: uint = 0x00AB;
		
		
		//New Records in BIFF5
		public static const SHEET			: uint = 0x0085;
		public static const SORT			: uint = 0x0090;
		public static const MULRK			: uint = 0x00BD;
		public static const MULBLANK		: uint = 0x00BE;
		public static const RSTRING			: uint = 0x00D6;
		public static const DBCELL			: uint = 0x00D7;
		public static const BOOKBOOL		: uint = 0x00DA;
		public static const SCENPROTECT		: uint = 0x00DD;
		public static const SHAREDFMLA		: uint = 0x04BC;
		
		//New Records in BIFF8
		public static const MERGEDCELLS		: uint = 0x00E5;
		public static const BITMAP 			: uint = 0x00E9;
		public static const PHONETICPR		: uint = 0x00EF;
		public static const SST				: uint = 0x00FC;
		public static const LABELSST		: uint = 0x00FD;
		public static const EXTSST			: uint = 0x00FF;
		public static const LABELRANGES		: uint = 0x015F;
		public static const USELFS			: uint = 0x0160;
		public static const DSF				: uint = 0x0161;
		public static const EXTERNALBOOK	: uint = 0x01AE;
		public static const CFHEADER		: uint = 0x01B0;
		public static const CFRULE			: uint = 0x01B1;
		public static const DATAVALIDATIONS	: uint = 0x01B2;
		public static const HYPERLINK		: uint = 0x01B8;
		public static const DATAVALIDATION	: uint = 0x01BE;
		public static const QUICKTIP		: uint = 0x0800;
		public static const SHEETLAYOUT		: uint = 0x0862;
		public static const SHEETPROTECTION	: uint = 0x0867;
		public static const RANGEPROTECTION	: uint = 0x0868;
		

		

		public static function getType(value : uint):String{
			switch(value){
				case DIMENSION : 
					return "DIMENSION";
				case DIMENSION_3458 : 
					return "DIMENSION_3458";
				case BLANK : 
					return "BLANK";
				case BLANK_3458 : 
					return "BLANK_3458";
				case INTEGER : 
					return "INTEGER";
				case NUMBER : 
					return "NUMBER";
				case NUMBER_3458			: 
					return "NUMBER_3458";
				case LABEL : 
					return "LABEL";
				case LABEL_3458 : 
					return "LABEL_3458";
				case BOOLERR : 
					return "BOOLERR";
				case BOOLERR_3458 : 
					return "BOOLERR_3458";
				case FORMULA : 
					return "FORMULA";
				case FORMULA_3 : 
					return "FORMULA_3";
				case FORMULA_4			: 
					return "FORMULA_4";
				case STRING : 
					return "STRING";
				case STRING_3458: 
					return "STRING_3458";
				case ROW : 
					return "ROW";
				case ROW_3458 : 
					return "ROW_3458";
				case BOF : 
					return "BOF";
				case BOF_3 : 
					return "BOF_3";
				case BOF_4 : 
					return "BOF_4";
				case BOF_58 : 
					return "BOF_58";
				case EOF : 
					return "EOF";
				case INDEX : 
					return "INDEX";
				case INDEX_3458 : 
					return "INDEX_3458";
				case CALCCOUNT : 
					return "CALCCOUNT";
				case CALCMODE : 
					return "CALCMODE";
				case PRECISION : 
					return "PRECISION";
				case REFMODE : 
					return "REFMODE";
				case DELTA : 
					return "DELTA";
				case ITERATION		: 
					return "ITERATION";
				case PROTECT			: 
					return "PROTECT";
				case PASSWORD		: 
					return "PASSWORD";
				case HEADER			: 
					return "HEADER";
				case FOOTER			: 
					return "FOOTER";
				case EXTERNCOUNT		: 
					return "EXTERNCOUNT";
				case EXTERNSHEET		: 
					return "EXTERNSHEET";
				case DEFINEDNAME		: 
					return "DEFINEDNAME";
				case DEFINEDNAME_34	: 
					return "DEFINEDNAME_34";
				case WINDOWPROTECT : 
					return "WINDOWPROTECT";
				case VERTICALPAGEBREAKS: 
					return "VERTICALPAGEBREAKS";
				case HORIZONTALPAGEBREAKS: 
					return "HORIZONTALPAGEBREAKS";
				case NOTE : 
					return "NOTE";
				case SELECTION : 
					return "SELECTION";
				case FORMAT : 
					return "FORMAT";
				case FORMAT_458 : 
					return "FORMAT_458";
				case BUILTINFMTCOUNT : 
					return "BUILTINFMTCOUNT";
				case BUILTINFMTCOUNT_34 : 
					return "BUILTINFMTCOUNT_34";
				case COLUMNDEFAULT : 
					return "COLUMNDEFAULT";
				case ARRAY : 
					return "ARRAY";
				case ARRAY_3458 : 
					return "ARRAY_3458";
				case DATEMODE : 
					return "DATEMODE";
				case EXTERNALNAME : 
					return "EXTERNALNAME";
				case EXTERNALNAME_34 : 
					return "EXTERNALNAME_34";
				case COLWIDTH : 
					return "COLWIDTH";
				case DEFAULTROWHEIGHT: 
					return "DEFAULTROWHEIGHT";
				case DEFAULTROWHEIGHT_3458: 
					return "DEFAULTROWHEIGHT_3458";
				case LEFTMARGIN : 
					return "LEFTMARGIN";
				case RIGHTMARGIN : 
					return "RIGHTMARGIN";
				case TOPMARGIN : 
					return "TOPMARGIN";
				case BOTTOMMARGIN : 
					return "BOTTOMMARGIN";
				case PRINTHEADERS	: 
					return "PRINTHEADERS";
				case PRINTGRIDLINES	: 
					return "PRINTGRIDLINES";
				case FILEPASS : 
					return "FILEPASS";
				case FONT : 
					return "FONT";
				case FONT_34 : 
					return "FONT_34";
				case FONT2 : 
					return "FONT2";
				//case TAB;return;return;return;EOP			: 0x36;return;
				//case TABLEOP2		: 0x37;return;
				case DATATABLE : 
					return "DATATABLE";
				case DATATABLE_3458 : 
					return "DATATABLE_3458";
				case DATATABLE2 : 
					return "DATATABLE2";
				case CONTINUE : 
					return "CONTINUE";
				case WINDOW1 : 
					return "WINDOW1";
				case WINDOW2 : 
					return "WINDOW2";
				case WINDOW2_3458 : 
					return "WINDOW2_3458";
				case BACKUP : 
					return "BACKUP";
				case PANE : 
					return "PANE";
				case CODEPAGE : 
					return "CODEPAGE";
				case XF : 
					return "XF";
				case XF_3 : 
					return "XF_3";
				case XF_4 : 
					return "XF_4";
				case XF_58 : 
					return "XF_58";
				case IXFE : 
					return "IXFE";
				case FONTCOLOR : 
					return "FONTCOLOR";
				case PLS :
					return "PLS";
				case DCONREF : 
					return "DCONREF";
				case DEFCOLWIDTH : 
					return "DEFCOLWIDTH";
				
				//New Records in BIFF3
				case XCT : 
					return "XCT";
				case CRN : 
					return "CRN";
				case FILESHARING : 
					return "FILESHARING";
				case WRITEACCESS : 
					return "WRITEACCESS";
				case UNCALCED : 
					return "UNCALCED";
				case SAVERECALC : 
					return "SAVERECALC";
				case OBJECTPROTECT : 
					return "OBJECTPROTECT";
				case COLINFO : 
					return "COLINFO";
				case RK : 
					return "RK";
				case GUTS : 
					return "GUTS";
				case SHEETPR : 
					return "SHEETPR";
				case GRIDSET : 
					return "GRIDSET";
				case HCENTER : 
					return "HCENTER";
				case VCENTER : 
					return "VCENTER";
				case WRITEPROT : 
					return "WRITEPROT";
				case COUNTRY : 
					return "COUNTRY";
				case HIDEOBJ : 
					return "HIDEOBJ";
				case PALETTE : 
					return "PALETTE";
				case STYLE : 
					return "STYLE";
				
				
				//New Records in BIFF4
				case STANDARDWIDTH : 
					return "STANDARDWIDTH";
				case SCL : 
					return "SCL";
				case PAGESETUP : 
					return "PAGESETUP";
				case GCW : 
					return "GCW";
				
				
				//New Records in BIFF5
				case SHEET : 
					return "SHEET";
				case SORT : 
					return "SORT";
				case MULRK : 
					return "MULRK";
				case MULBLANK : 
					return "MULBLANK";
				case RSTRING : 
					return "RSTRING";
				case DBCELL : 
					return "DBCELL";
				case BOOKBOOL : 
					return "BOOKBOOL";
				case SCENPROTECT : 
					return "SCENPROTECT";
				case SHAREDFMLA : 
					return "SHAREDFMLA";
				
				//New Records in BIFF8
				case MERGEDCELLS : 
					return "MERGEDCELLS";
				case BITMAP : 
					return "BITMAP";
				case PHONETICPR : 
					return "PHONETICPR";
				case SST : 
					return "SST";
				case LABELSST : 
					return "LABELSST";
				case EXTSST : 
					return "EXTSST";
				case LABELRANGES : 
					return "LABELRANGES";
				case USELFS : 
					return "USELFS";
				case DSF : 
					return "DSF";
				case EXTERNALBOOK : 
					return "EXTERNALBOOK";
				case CFHEADER : 
					return "CFHEADER";
				case CFRULE : 
					return "CFRULE";
				case DATAVALIDATIONS : 
					return "DATAVALIDATIONS";
				case HYPERLINK : 
					return "HYPERLINK";
				case DATAVALIDATION	: 
					return "DATAVALIDATION";
				case QUICKTIP : 
					return "QUICKTIP";
				case SHEETLAYOUT : 
					return "SHEETLAYOUT";
				case SHEETPROTECTION : 
					return "SHEETPROTECTION";
				case RANGEPROTECTION : 
					return "RANGEPROTECTION";
			}
			return "unknown type: "+ Number(value).toString(16);
		}
		
		
	}
}