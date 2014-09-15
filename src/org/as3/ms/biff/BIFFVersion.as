package org.as3.ms.biff {
	/**
	 * Used to represent the different versions of BIFF files that exist in the wild
	 * 
	 * The Excel file format is named BIFF ( Binary Interchange File Format).
	 * It is used to store all types of documents:
	 * worksheet documents
	 * workbook documents
	 * and workspace documents
	 * There are different versions of this file format, depending on the version of Excel that has written the file and depending on the document type.
	 * 
	 * The following table shows which Excel version writes which file format for worksheet and workbook documents:
	 * Excel version		BIFF version		Document type
	 * Excel 2.x			BIFF2				Worksheet
	 * Excel 3.0			BIFF3				Worksheet
	 * Excel 4.0			BIFF4				Worksheet
	 * Excel 5.0			BIFF5				Workbook
	 * Excel 7.0			BIFF5				Workbook
	 * Excel 8.0			BIFF8				Workbook
	 * Excel 9.0			BIFF8				Workbook
	 * Excel 10.0			BIFF8				Workbook
	 * Excel 11.0			BIFF8				Workbook
	 * 
	 * BIFF8 contains major changes towards older BIFF versions, for instance the handling of Unicode strings.
	 */
	public class BIFFVersion {
		/**
		 * Used by Excel 2.x. It doesn't support multiple sheets, charts, or really anything even remotely fun.
		 */
		public static const BIFF2:uint = 0;
		
		/**
		 *
		 */
		public static const BIFF3:uint = 1;
		
		
		/**
		 * Used by Excel 4.0. Provides support for multiple sheets indirectly via workspaces.
		 */
		public static const BIFF4:uint = 2;
		
		
		/**
		 * Used by Excel 5 and '95. Provides support for multiple sheets natively.
		 */
		public static const BIFF5:uint = 3;
		
		
		
		/**
		 * Used by Excel'97-2003. Generally the stream is wrapped in a CDF file.
		 *
		 *
		 * @see com.as3xls.cdf.CDFReader
		 */
		public static const BIFF8:uint = 4;
	}
}