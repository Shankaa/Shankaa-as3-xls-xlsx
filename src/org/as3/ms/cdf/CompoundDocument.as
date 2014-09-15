package org.as3.ms.cdf
{
	import flash.utils.ByteArray;
	import flash.utils.Endian;
	
	
	/**
	 * <p>
	 * Starting with Excel '97 xls documents are actually CDF (Compound Document Format) files, which can also contain
	 * embedded OLE objects. Their design roughly parallels that of a file system: it starts with a header region,
	 * followed by a sector allocation table and streams (which are like files) spread out over multiple sectors.
	 * Stream names are stored in a directory stream.
	 * </p>
	 *
	 * <p>
	 * The Workbook is stored in a stream called either Workbook or Book, depending (I believe) on the version. Also,
	 * in every Excel file I've seen this stream is conveniently pointed to by the first directory entry. Sometimes I think
	 * Microsoft does love me after all.
	 * </p>
	 * 
	 * 
	 * 
	 * 
	 * <p>
	 * Compound File Binary Format
	 * [source Wikipedia]
	 * Compound File Binary Format (CFBF), also called Compound File or Compound Document, is a file format for storing
	 * numerous files and streams within a single file on a disk. CFBF is developed by Microsoft and is an implementation
	 * of Microsoft COM Structured Storage.
	 * Microsoft has opened the format for use by others and it is now used in a variety of programs from Microsoft Word
	 * and Microsoft Access to Business Objects. It also forms the basis of the Advanced Authoring Format.
	 * 
	 * Overview
	 * At its simplest, the Compound File Binary Format is a container, with little restriction on what can be stored within it.
	 * A CFBF file structure loosely resembles a FAT filesystem. The file is partitioned into Sectors which are chained together
	 * with a File Allocation Table (not to be mistaken with the file system of the same name) which contains chains of sectors
	 * related to each file, a Directory holds information for contained files with a Sector ID (SID) for the starting sector of
	 * a chain and so on.
	 * 
	 * Structure
	 * The CFBF file consists of a 512-Byte header record followed by a number of sectors whose size is defined in the header.
	 * The literature defines Sectors to be either 512 or 4096 bytes in length, although the format is potentially capable of
	 * supporting sectors ranging in size from 128-Bytes upwards in powers of 2 (128, 256, 512, 1024, etc.). The lower limit
	 * of 128 is the minimum required to fit a single directory entry in a Directory Sector.
	 * There are several types of sector that may be present in a CFBF:
	 * - File Allocation Table (FAT) Sector - contains chains of sector indices much as a FAT does in the FAT/FAT32 filesystems
	 * - MiniFAT Sectors - similar to the FAT but storing chains of mini-sectors within the Mini-Stream
	 * - Double-Indirect FAT (DIFAT) Sector - contains chains of FAT sector indices
	 * - Directory Sector - contains directory entries
	 * - Stream Sector - contains arbitrary file data
	 * - Range Lock Sector - contains the byte-range locking area of a large file
	 * 
	 * More detail is given below for the header and each sector type.
	 * 
	 * 
	 * CFBF Header format
	 * 
	 * The CFBF Header occupies the first 512 bytes of the file and information required to interpret the rest of the file.
	 * The C-Style structure declaration below (extracted from the AAFA's Low-Level Container Specification) shows the members
	 * of the CFBF header and their purpose:
	 * 
	 * typedef unsigned long ULONG;    // 4 Bytes
	 * typedef unsigned short USHORT;  // 2 Bytes
	 * typedef short OFFSET;           // 2 Bytes
	 * typedef ULONG SECT;             // 4 Bytes
	 * typedef ULONG FSINDEX;          // 4 Bytes
	 * typedef USHORT FSOFFSET;        // 2 Bytes
	 * typedef USHORT WCHAR;           // 2 Bytes
	 * typedef ULONG DFSIGNATURE;      // 4 Bytes
	 * typedef unsigned char BYTE;     // 1 Byte
	 * typedef unsigned short WORD;    // 2 Bytes
	 * typedef unsigned long DWORD;    // 4 Bytes
	 * typedef ULONG SID;              // 4 Bytes
	 * typedef GUID CLSID;             // 16 Bytes
	 * 
	 * struct StructuredStorageHeader { // [offset from start (bytes), length (bytes)]
	 *     BYTE _abSig[8];             // [00H,08] {0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1,0x1a, 0xe1} for current version
	 *     CLSID _clsid;               // [08H,16] reserved must be zero (WriteClassStg/GetClassFile uses root directory class id)
	 *     USHORT _uMinorVersion;      // [18H,02] minor version of the format: 33 is written by reference implementation
	 *     USHORT _uDllVersion;        // [1AH,02] major version of the dll/format: 3 for 512-byte sectors, 4 for 4 KB sectors
	 *     USHORT _uByteOrder;         // [1CH,02] 0xFFFE: indicates Intel byte-ordering
	 *     USHORT _uSectorShift;       // [1EH,02] size of sectors in power-of-two; typically 9 indicating 512-byte sectors
	 *     USHORT _uMiniSectorShift;   // [20H,02] size of mini-sectors in power-of-two; typically 6 indicating 64-byte mini-sectors
	 *     USHORT _usReserved;         // [22H,02] reserved, must be zero
	 *     ULONG _ulReserved1;         // [24H,04] reserved, must be zero
	 *     FSINDEX _csectDir;          // [28H,04] must be zero for 512-byte sectors, number of SECTs in directory chain for 4 KB sectors
	 *     FSINDEX _csectFat;          // [2CH,04] number of SECTs in the FAT chain
	 *     SECT _sectDirStart;         // [30H,04] first SECT in the directory chain
	 *     DFSIGNATURE _signature;     // [34H,04] signature used for transactions; must be zero. The reference implementation does not support transactions
	 *     ULONG _ulMiniSectorCutoff;  // [38H,04] maximum size for a mini stream; typically 4096 bytes
	 *     SECT _sectMiniFatStart;     // [3CH,04] first SECT in the MiniFAT chain
	 *     FSINDEX _csectMiniFat;      // [40H,04] number of SECTs in the MiniFAT chain
	 *     SECT _sectDifStart;         // [44H,04] first SECT in the DIFAT chain
	 *     FSINDEX _csectDif;          // [48H,04] number of SECTs in the DIFAT chain
	 *     SECT _sectFat[109];         // [4CH,436] the SECTs of first 109 FAT sectors
	 * };
	 * 
	 * 
	 * File Allocation Table (FAT) Sectors
	 * 
	 * When taken together as a single stream the collection of FAT sectors define the status and linkage of every sector in the file.
	 * Each entry in the FAT is 4 bytes in length and contains the sector number of the next sector in a FAT chain or one of the following
	 * special values:
	 *     FREESECT (0xFFFFFFFF) 	- denotes an unused sector
	 *     ENDOFCHAIN (0xFFFFFFFE) 	- marks the last sector in a FAT chain
	 *     FATSECT (0xFFFFFFFD) 	- marks a sector used to store part of the FAT
	 *     DIFSECT (0xFFFFFFFC) 	- marks a sector used to store part of the DIFAT
	 * 
	 * 
	 * http://www.digitalpreservation.gov/formats/digformatspecs/WindowsCompoundBinaryFileFormatSpecification.pdf
	 * The Fat is the main allocator for space within a Compound File. Every sector in the file is represented within the Fat
	 * in some fashion, including those sectors that are unallocated (free). The Fat is a virtual stream made up of one or
	 * more Fat Sectors.
	 * Fat sectors are arrays of SECTs that represent the allocation of space within the file. Each stream is represented in the
	 * Fat by a chain, in much the same fashion as a DOS file-allocation-table (FAT). To elaborate, the set of Fat Sectors can be
	 * considered together to be a single array -- each cell in that array contains the SECT of the next sector in the chain,
	 * and this SECT can be used as an index into the Fat array to continue along the chain. Special values are reserved for
	 * chain terminators (ENDOFCHAIN = 0xFFFFFFFE), free sectors (FREESECT = 0xFFFFFFFF), and sectors that contain storage for
	 * Fat Sectors (FATSECT = 0xFFFFFFFD) or DIF Sectors (DIFSECT = 0xFFFFFFC), which are notchained in the same way as the others.
	 * 
	 * The locations of Fat Sectors are read from the DIF (Double-indirect Fat), which is described below. The Fat is represented in itself,
	 * but not by a chain –a special reserved SECT value (FATSECT = 0xFFFFFFFD) is used to mark sectors allocated to the Fat.
	 * A SECT can be converted into a byte offset into the file by using the following formula: SECT << ssheader._uSectorShift + sizeof(ssheader).
	 * This implies that sector 0 of the file begins at byte offset 512, not at 0.
	 * 
	 * 
	 * MiniFat Sectors
	 * Since space for streams is always allocated in sector - sized blocks, there can be considerable waste when storing objects much smaller than
	 * sectors (typically 512 bytes). As a solution to this problem, we introduced the concept of the MiniFat. The MiniFat is structurally
	 * equivalent to the Fat, but is used in a different way. The virtual sector size for objects represented in the Minifat is 
	 * 1 << ssheader._uMiniSectorShift (typically 64 bytes) instead of 1 << ssheader._uSectorShift (typically 512 bytes).
	 * The storage for these objects comes from a virtual stream within the Multistream (called the Ministream).
	 * The locations for MiniFat sectors are stored in a standard chain in the Fat, with the beginning of the chain stored in the header.
	 * A Minifat sector number can be converted into a byte offset into the ministream by using the following formula: SECT << ssheader._uMiniSectorShift.
	 * (This formula is different from the formula used to convert a SECT into a byte offset in the file, since no header is stored in the Ministream)
	 * The Ministream is chained within the Fat in exactly the same fashion as any normal stream. It is referenced by the first Directory Entry (SID0).
	 * 
	 * DIF Sectors
	 * The Double-Indirect Fat is used to represent storage of the Fat. The DIF is also represented by an array of SECTs,
	 * and is chained by the terminating cell in each sector array (see the diagram above). As an optimization, the first 109 Fat Sectors are represented within the header itself,
	 * so no DIF sectors will be found in a small (< 7 MB) Compound File. The DIF represents the Fat in a different manner than the Fat represents a chain. A given index into the DIF will
	 * contain the SECT of the Fat Sector found at that offset in the Fat virtual stream. For instance, index 3 in the DIF would contain the SECT for Sector #3 of the Fat.
	 * The storage for DIF Sectors is reserved in the Fat, but is not chained there (space for it is reserved by a special SECT value , DIFSECT=0xFFFFFFFC).
	 * The location of the first DIF sector is stored in the header. A value of ENDOFCHAIN=0xFFFFFFFE is stored in the pointer to the next DIF sector of the last DIF sector.
	 * 
	 * http://msdn.microsoft.com/en-us/library/dd941958.aspx
	 * The DIFAT array is used to represent storage of the FAT sectors. The DIFAT is represented by an array of 32-bit sector numbers.
	 * The DIFAT array is stored both in the header and in DIFAT sectors. In the header, the DIFAT array occupies 109 entries,
	 * and in each DIFAT sector, the DIFAT array occupies the entire sector minus 4 bytes (the last field is for chaining the DIFAT sector chain).
	 * 
	 * The DIFAT sectors are linked together by the last field in each DIFAT sector. As an optimization, the first 109 FAT sectors are represented
	 * within the header itself. No DIFAT sectors will be needed in a compound file that is smaller than 6.875 megabyte (MB) for a 512 byte sector
	 * compound file (6.875 MB = (1 header sector + 109 FAT sectors x 128 non-empty entries) × 512 bytes per sector).
	 * The DIFAT represents the FAT sectors in a different manner than the FAT represents a sector chain. A given index, n, into the DIFAT array
	 * will contain the sector number of the (n+1)th FAT sector. For instance, index #3 in the DIFAT contains the sector number for the 4rd FAT sector,
	 * since DIFAT array starts with index #0.
	 * The storage for DIFAT sectors is reserved with the FAT, but it is not chained there. Space for DIFAT sectors is marked by a special sector number, DIFSECT (0xFFFFFFFC).
	 * The location of the first DIFAT sector is stored in the header.
	 * A special value of ENDOFCHAIN (0xFFFFFFFE) is stored in "Next DIFAT Sector Location" field of the last DIFAT sector, or in the header when no DIFAT sectors are needed.
	 * FAT Sector Location (variable): This field specifies the FAT sector number in a DIFAT.
	 * 		If Header Major Version is 3, then there MUST be 127 fields specified to fill a 512-byte sector minus the "Next DIFAT Sector Location" field.
	 * 		If Header Major Version is 4, then there MUST be 1023 fields specified to fill a 4096-byte sector minus the "Next DIFAT Sector Location" field.
	 * 		Next DIFAT Sector Location (4 bytes): This field specifies the next sector number in the DIFAT chain of sectors.
	 * The first DIFAT sector is specified in the Header.
	 * The last DIFAT sector MUST set this field to ENDOFCHAIN (0xFFFFFFFE).
	 *
	 * 
	 * Directory Sectors
	 * The Directory is a structure used to contain per -stream information about the streams in a Compound File, as well as to maintain a tree -
	 * styled containment structure. It is a virtual stream made up of one or more Directory Sectors. The Directory is represented as a standard
	 * chain of sectors within the Fat. The first sector of the Directory chain (the Root Directory Entry) 
	 * 
	 * typedef enum tagSTGTY {
	 * 		STGTY_INVALID	= 0,
	 * 		STGTY_STORAGE	= 1,
	 * 		STGTY_STREAM	= 2,
	 * 		STGTY_LOCKBYTES	= 3,
	 * 		STGTY_PROPERTY	= 4,
	 * 		STGTY_ROOT		= 5,
	 * }STGTY;
	 * typedef enum tagDECOLOR {
	 * 		DE_RED			= 0,
	 * 		DE_BLACK		= 1,
	 * }	DECOLOR;
	 * struct StructuredStorageDirectoryEntry { 	// 	[offset from	start in bytes, length in bytes]
	 * 		BYTE	_ab[32*sizeof(WCHAR)];			//	[000H,64]	64 bytes. The Element name in Unicode, padded with zeros to fill this byte array
	 * 		WORD	_cb;							//	[040H,02]	Length of the Element name in characters, not bytes
	 * 		BYTE	_mse;							//	[042H,01]	Type of object: value taken from the STGTY enumeration
	 * 		BYTE	_bflags;						//	[043H,01]	Value taken from DECOLOR enumeration.
	 * 		SID		_sidLeftSib;					//	[044H,04]	SID of the left	- sibling of this entry in the directory tree
	 * 		SID		_sidRightSib;					//	[048H,04]	SID of the right- sibling of this entry in the directory tree
	 * 		SID		_sidChild;						//	[04CH,04]	SID of the child acting as the root of all the children of this element (if _mse=STGTY_STORAGE)
	 * 		GUID	_clsId;							//	[050H,16]	CLSID of this storage (if _mse=STGTY_STORAGE)
	 * 		DWORD	_dwUserFlags;					//	[060H,04]	User flags of this storage (if _mse=STGTY_STORAGE)
	 * 		TIME_T	_time[2];						//	[064H,16]	Create/Modify time	-stamps (if _mse=STGTY_STORAGE)
	 * 		SECT	_sectStart						//	[074H,04]	starting SECT of the stream (if _mse=STGTY_STREAM)
	 * 		ULONG	_ulSize;						//	[078H,04]	size of stream in bytes (if _mse=STGTY_STREAM)
	 * 		DFPROPTYPE	_dptPropType;				//	[07CH,02]	Reserved for future use. Must be zero.
	 * };
	 * 
	 * Each level of the containment hierarchy (i.e. each set of siblings) is represented as a red-black tree. The parent of this 
	 * set of sibilings will have a pointer to the top of this tree. This red-black tree must maintain the following conditions in order for it to be valid:
	 * 1.The root node must always be black. Since the root directory (see below) does not have siblings, it's color is irrelevant and may therefore be either red or black.
	 * 2.No two consecutive nodes may both be red.
	 * 3.The left child must always be less than the right child. This relationship is defined as:
	 * 		- A node with a shorter name is less than a node with a longer name (i.e. compare the length of the names)
	 * 		- For names with the length names compare names.
	 * The simplest implementation of the above would be to mark every node as black, in which case the tree is simply a binary tree
	 * A Directory Sector is an Array of Directory Entries.
	 * Each user stream within a Compound File is represented by a single Directory Entry. The Directory is considered as
	 * a large Array of Directory Entries. It is usefull to note that the Directory Entry for a stream remains at the same index
	 * in the Directory array for the life of the stream - thus, this index (called an SID) can be used to readily identify a given stream.
	 * The directory entry is then padded out with zeros to make a total size of 128 bytes.Directory entries are grouped
	 * into blocks of four to form Directory Sectors.
	 * 
	 * 
	 * Root Directory Entry
	 * 
	 * The first sector of the Directory chain (also referred to as the first element of the Directory array, or SID0) is known as
	 * the Root Directory Entry and is reserved for two purposes: First, it provides a root parent for all objects stationed at
	 * the root of the multi-stream. Second, its function is overloaded to store the size and starting sector for the Mini-stream.
	 * The Root Directory Entry behaves as both a stream and a storage. All of the fields in the Directory Entry are valid for
	 * the root. The Root Directory Entry’s Name field typically contains the string “RootEntry” in Unicode, although some
	 * versions of structured storage (particularly the preliminary reference implementation and the Macintosh version) store
	 * only the first letter of this string, “R” in the name. This string is always ignored, since the Root Directory Entry is
	 * known by its position at SID0 rather than by its name, and its name is not otherwise used. New implementations should
	 * write “RootEntry” properly in the Root Directory Entry for consistency and support manipulating files created with only the “R” name.
	 * 
	 * 
	 * Other Directory Entries
	 * 
	 * Non-root directory entries are marked as either stream (STGTY_STREAM) or storage (STGTY_STORAGE) elements.
	 * Storage elements have a _clsid, _time[], and _sidChild values; stream elements may not.
	 * Stream elements have valid _sectStart and _ulSize members, whereas these fields are set to zero for storage elements 
	 * (except as noted above for the Root Directory Entry). To determine the physical file location of actual stream data from a stream directory entry,
	 * it is necessary to determine which FAT (normal or mini) the stream exists within.  Streams whose _ulSize member is less
	 * than the _ulMiniSectorCutoff value for the file exist in the ministream, and so the _startSect is used as an index
	 * into the MiniFat (which starts at _sectMiniFatStart ) to track the chain of mini-sectors through the mini-stream
	 * (which is, as noted earlier, the standard (non-mini) stream referred to by the Root Directory Entry’s _sectStart value).
	 * Streams whose _ulSize member is greater than the _ulMiniSectorCutoff value for the file exist as standard streams– their _sectStart value
	 * is used as an index into the standard FAT which describes the chain of full sectors containing their data).
	 * 
	 * 
	 * Range Lock Sector
	 * 
	 * The Range Lock Sector must exist in files greater than 2GB in size, and must not exist in files smaller than 2GB. The Range Lock
	 * Sector must contain the byte range 0x7FFFFF00 to 0x7FFFFFFF in the file. This area is reserved by Microsoft's COM implementation
	 * for storing byte-range locking information for concurrent access.
	 * 
	 * 
	 * Glossary
	 * 
	 *		FAT - File Allocation Table, also known as: SAT - Sector Allocation Table
	 *		DIFAT - Double-Indirect File Allocation Table
	 *		FAT Chain - a group of FAT entries which indicate the sectors allocated to a Stream in the file
	 *		Stream - a virtual file which occupies a number of sectors within the CFBF
	 *		Sector - the unit of allocation within the CFBF, usually 512 or 4096 Bytes in length
	 * 
	 *		globally unique identifier (GUID): A term used interchangeably with universally unique identifier (UUID)
	 * 			in Microsoft protocol technical documents (TDs). Interchanging the usage of these terms does not imply
	 * 			or require a specific algorithm or mechanism to generate the value. Specifically, the use of this term
	 * 			does not imply or require that the algorithms specified in [RFC4122] or [C706] must be used for generating
	 * 			the GUID. See also universally unique identifier (UUID).
	 * 
	 * 		compound file: A structure for storing a file system, similar to a simplified FAT file system inside a single file,
	 * 			 by dividing the single file into sectors.
	 * 
	 * 		creation time: The time, in UTC, when a storage object was created.
	 * 
	 * 		directory entry: A structure that contains a storage object's or stream object'sFileInformation.
	 * 
	 * 		double-indirect file allocation table (DIFAT): A structure used to locate FATsectors in a compound file.
	 * 
	 * 		directory stream: An array of directory entries grouped into sectors.
	 * 
	 * 		header: The structure at the beginning of a compound file.
	 * 
	 * 		mini FAT: A file allocation table (FAT) structure for the mini stream used to allocate space in a small sector size.
	 * 
	 * 		mini stream: A structure that contains all user-defined data for stream objects less than a predefined size limit.
	 * 
	 * 		modification time: The time, in UTC, when a storage object was last modified.
	 * 
	 * 		root storage object: A storage object in a compound file that must be accessed before any other storage objects
	 * 			and stream objects are referenced. It is the uppermost parent object in the storage object and stream object hierarchy.
	 * 		sector chain: A linked list of sectors, where each sector can be located in a different location inside a compound file.
	 * 
	 * 		sector number: A non-negative integer identifying a particular sector located in a compound file.
	 * 
	 * 		sector size: The size in bytes of a sector in a compound file, typically 512 bytes.
	 * 
	 * 		storage object: An object in a compound file analogous to a file systemdirectory. The parent object of a storage object
	 * 			must be another storage object or the root storage object.
	 * 
	 * 		stream object: An object in a compound file analogous to a file systemfile. The parent object of a stream object
	 * 			must be a storage object or the root storage object.
	 * 
	 * 		unallocated free sector: An empty sector that can be allocated to hold data.
	 * 
	 * 		user-defined data: The main stream portion of a stream object.
	 * 
	 * 		CLSID: A GUID representing an object class. In a root storage object or storage object, the object classGUID
	 * 			can be used as a parameter to launch applications.
	 * 
	 * </p>
	 *
	 */
	public class CompoundDocument {
		/**
		 * File Allocation Table (FAT) Sectors
		 * 
		 * When taken together as a single stream the collection of FAT sectors define the status and linkage of every sector in the file.
		 * Each entry in the FAT is 4 bytes in length and contains the sector number of the next sector in a FAT chain or one of the following
		 */
		private static const	FREESECT 		: uint  = 0xFFFFFFFF;	// denotes an unused sector							(int -1) 
		private static const	ENDOFCHAIN		: uint  = 0xFFFFFFFE; 	// marks the last sector in a FAT chain				(int -2)
		private static const	FATSECT			: uint  = 0xFFFFFFFD; 	// marks a sector used to store part of the FAT 	(int -3)  4294967293
		private static const	DIFSECT			: uint  = 0xFFFFFFFC; 	// marks a sector used to store part of the DIFAT	(int -4)
		
		// STGTY
		private static const	STGTY_INVALID	: uint  = 0;
		private static const	STGTY_STORAGE	: uint  = 1;
		private static const	STGTY_STREAM	: uint  = 2;
		private static const	STGTY_LOCKBYTES	: uint  = 3;
		private static const	STGTY_PROPERTY	: uint  = 4;
		private static const	STGTY_ROOT		: uint  = 5;
		
		// Entry color
		private static const	DE_RED			: uint	= 0;
		private static const	DE_BLACK		: uint	= 1;
		
		private static const 	FAT_HEADER_SIZE : uint = 512;
		
		private static const 	DIRECTORY_LENGTH: uint = 16*8;// dir have a 128 byte size
		
		private static const	MAGIC_NUMBER	: Array = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
		
		private var stream 						: ByteArray;
		private var shortStreamContainerStream 	: ByteArray;
		
		private var fatSectorSize 				: uint;
		private var miniFatSectorSize			: uint;
		private var _ulMiniSectorCutoff 		: uint;
		
		private var _sectFat					: Array;
		private var miniFatSectorAllocationTable: Array;
		
		public var dir 							: Array;
		
		
		
		/**
		 * Determines whether a ByteArray contains a CDF file by checking for the presence of the CDF magic value.
		 *
		 * @param stream The ByteArray to check for CDF-ness
		 * @return True if the file is a CDF file; otherwise false
		 *
		 */
		public static function isCompoundDocumentFormatFile(stream : ByteArray):Boolean {
			
			var n 					: uint 	= MAGIC_NUMBER.length;
			if(stream.length < MAGIC_NUMBER.length) {
				return false;
			}
			
			while(n --){
				if(MAGIC_NUMBER[n] != stream[n]) {
					return false;
				}
			}
			
			return true;
		}
		
		
		/**
		 * Wraps the given ByteArray in a CDFReader. It will be rewound and set to LittleEndian.
		 *
		 * @param stream The ByteArray to wrap
		 *
		 */
		public function CompoundDocument(stream : ByteArray) {
			this.stream 					= stream;
			stream.position 				= 0;
			stream.endian 					= Endian.LITTLE_ENDIAN;
			
			loadDocument();
		}
		
		private function getHeaderValueB(index:uint,sizeInBytes:uint):uint{
			var currentStreamPosition 	: uint = stream.position;
			stream.position = index;
			
			var value					: uint = 0;
			switch(sizeInBytes){
				case 1 :
					value = stream.readUnsignedByte();
					break;
				case 2 :
					value = stream.readUnsignedShort();
					break;
				case 4 :
					value = stream.readUnsignedInt();
					break;
				value = 0;
			}
			stream.position = currentStreamPosition;
			return value;
		}
		
		private function getHeaderValue1B(index:uint):uint{
			return getHeaderValueB(index,1);
		}
		
		private function getHeaderValue2B(index:uint):uint{
			return getHeaderValueB(index,2);
		}
		
		private function getHeaderValue4B(index:uint):uint{
			return getHeaderValueB(index,4);
		}
		
		/**
		 * Loads the stream represented by a given directory entry.
		 *
		 * @param dirId The dirId pointing to the stream to load
		 * @return A ByteArray containing the specified directory
		 *
		 */
		public function loadDirectoryEntry(dirId : uint):ByteArray {
			var directory			: Directory = dir[dirId];
			var data 				: ByteArray;
			if (directory.size >= _ulMiniSectorCutoff){
				data 	= loadStream(directory.secId,directory.size);
			} else {
				data 	= loadShortStream(directory.secId,directory.size);
			}
			return data;
		}
		
		
		/**
		 * writes a document into a CompoundDocument
		 *  
		 * @param fileToWrite
		 * @param fileType 0 = xls, 1 ppt
		 * 
		 */
		private function writeDocument(fileToWrite:ByteArray,fileType:uint):ByteArray{
			var fileToWriteLength	: uint		= fileToWrite.length;
			fileToWrite.position				= 0;
			// we take 512B sectores
			var numberOfSectors		: uint		= fileToWriteLength/predefinediniSectorCutoff + ((fileToWriteLength*1.0)%predefinediniSectorCutoff == 0?0:1);
			
			
			
			
			
			
			var stream				: ByteArray = new ByteArray();
			var workFile			: ByteArray = new ByteArray();
			var predefinediniSectorCutoff 	: uint 		= 4096;
			workFile.readBytes(fileToWrite,0,fileToWrite.length);
			
			if(workFile.length > _ulMiniSectorCutoff){
				while(workFile.length%512!=0){
					workFile.writeUnsignedInt(FREESECT);
				}
			}
			
			var _clsid 				: ByteArray = new ByteArray()	// [08H,16] reserved must be zero (WriteClassStg/GetClassFile uses root directory class id)
			_clsid.length = 16;
				
			
			var _uMinorVersion		: uint 		= 62;			// [18H,02] minor version of the format: 33 is written by reference implementation
			var _uDllVersion		: uint 		= 3;			// [1AH,02] major version of the dll/format: 3 for 512-byte sectors, 4 for 4 KB sectors
			var _uByteOrder 		: uint 		= 0xFFFE;		// [1CH,02] 0xFFFE: indicates Intel byte-ordering
			var _uSectorShift		: uint		= 9;			// [1EH,02] size of sectors in power-of-two; typically 9 indicating 512-byte sectors
			
			var _uMiniSectorShift 	: uint		= 6;			// [20H,02] size of mini-sectors in power-of-two; typically 6 indicating 64-byte mini-sectors
			miniFatSectorSize 					= 64;
			
			var _usReserved 		: uint 		= 0; 			// [22H,02] reserved, must be zero
			var _ulReserved1		: uint		= 0;			// [24H,04] reserved, must be zero
			var _csectDir			: uint		= 0;			// [28H,04] must be zero for 512-byte sectors, number of SECTs in directory chain for 4 KB sectors    
			
			var _csectFat			: uint 		= 0;			// [2CH,04] number of SECTs in the FAT chain
			var _sectDirStart 		: uint 		= 0;			// [30H,04] first SECT in the directory chain
			
			var _txSignature		: uint 		= 0;     		// [34H,04] signature used for transactions; must be zero. The reference implementation does not support transactions

			
			var _ulMiniSectorCutoff : uint		= predefinediniSectorCutoff;// [38H,04] maximum size for a mini stream; typically 4096 bytes
			var _sectMiniFatStart 	: uint		= 0;			// [3CH,04] first SECT in the MiniFAT chain
			var _csectMiniFat 		: uint 		= 0;			// [40H,04] number of SECTs in the MiniFAT chain
			
			//double-indirect file allocation table (DIFAT): A structure used to locate FATsectors in a compound file.
			var _sectDifStart 		: uint 		= ENDOFCHAIN;	// [44H,04] first SECT in the DIFAT chain
			var _csectDif 			: uint 		= 0;			// [48H,04] number of SECTs in the DIFAT chain
			
			
			// write header
			for each (var byte		: uint in MAGIC_NUMBER){
				stream.writeByte(byte);
			}
			stream.writeBytes(_clsid,0,16);						// [00H,16] file magic number
			stream.writeShort(_uMinorVersion);					// [08H,16] reserved must be zero (WriteClassStg/GetClassFile uses root directory class id)
			stream.writeShort(_uDllVersion);					// [1AH,02] major version of the dll/format: 3 for 512-byte sectors, 4 for 4 KB sectors
			stream.writeShort(_uByteOrder);						// [1CH,02] 0xFFFE: indicates Intel byte-ordering
			stream.writeShort(_uSectorShift);					// [1EH,02] size of sectors in power-of-two; typically 9 indicating 512-byte sectors
			stream.writeShort(_uMiniSectorShift);				// [20H,02] size of mini-sectors in power-of-two; typically 6 indicating 64-byte mini-sectors
			stream.writeShort(_usReserved);						// [22H,02] reserved, must be zero
			stream.writeUnsignedInt(_ulReserved1);				// [24H,04] reserved, must be zero
			stream.writeUnsignedInt(_csectDir);					// [28H,04] must be zero for 512-byte sectors, number of SECTs in directory chain for 4 KB sectors
			stream.writeUnsignedInt(_csectFat);					// [2CH,04] number of SECTs in the FAT chain
			stream.writeUnsignedInt(_sectDirStart);				// [30H,04] first SECT in the directory chain
			stream.writeUnsignedInt(_txSignature);				// [34H,04] signature used for transactions; must be zero. The reference implementation does not support transactions
			stream.writeUnsignedInt(_ulMiniSectorCutoff);		// [38H,04] maximum size for a mini stream; typically 4096 bytes
			stream.writeUnsignedInt(_sectMiniFatStart);			// [3CH,04] first SECT in the MiniFAT chain
			stream.writeUnsignedInt(_csectMiniFat);				// [40H,04] number of SECTs in the MiniFAT chain
			stream.writeUnsignedInt(_sectDifStart);				// [44H,04] first SECT in the DIFAT chain
			stream.writeUnsignedInt(_csectDif);					// [48H,04] number of SECTs in the DIFAT chain
			
			//write the FAT
			var fatAlloc			: ByteArray = new ByteArray();
			for (var i : int = 0;i<109;i++){
				fatAlloc.writeUnsignedInt(FREESECT);
			}
			
			//write the DIFAT
			
			
			//write the Sectors
			
			
			
			// write directoriy entries
			var elementName	: String			= "RootEntry";
			var cbEleName	: uint				= 22;
			var _dir0		: ByteArray			= new ByteArray();
			_dir0.length						= 128;
			_dir0.position 						= 0;
			_dir0.writeUTFBytes(elementName);
			_dir0.position						= 64;
			
			
			
			
			return stream;
		}
		
		
		
		
		/**
		 * Processes the document and arranges it so that streams can easily be extracted later
		 *
		 */
		private function loadDocument():void {
			// Process the header
			var magic 				: Number 	= stream.readDouble();
			var _clsid 				: ByteArray = new ByteArray();				// [08H,16] reserved must be zero (WriteClassStg/GetClassFile uses root directory class id)
			stream.readBytes(_clsid,stream.position,16);
			//stream.position 					+= 16; 							// Skip past UID
			
			var _uMinorVersion		: uint 		= stream.readUnsignedShort();	// [18H,02] minor version of the format: 33 is written by reference implementation
			var _uDllVersion		: uint 		= stream.readUnsignedShort();	// [1AH,02] major version of the dll/format: 3 for 512-byte sectors, 4 for 4 KB sectors
			var _uByteOrder 		: uint 		= stream.readUnsignedShort();	// [1CH,02] 0xFFFE: indicates Intel byte-ordering
			var _uSectorShift		: uint		= stream.readUnsignedShort();	// [1EH,02] size of sectors in power-of-two; typically 9 indicating 512-byte sectors
			fatSectorSize 						= 1<<_uSectorShift;
			var _uMiniSectorShift 	: uint		= stream.readUnsignedShort();	// [20H,02] size of mini-sectors in power-of-two; typically 6 indicating 64-byte mini-sectors
			miniFatSectorSize 					= 1<<_uMiniSectorShift;
			
			var _usReserved 		: uint 		= stream.readUnsignedShort(); 	// [22H,02] reserved, must be zero
			var _ulReserved1		: uint		= stream.readUnsignedInt();		// [24H,04] reserved, must be zero
			var _csectDir			: uint		= stream.readUnsignedInt();		// [28H,04] must be zero for 512-byte sectors, number of SECTs in directory chain for 4 KB sectors  
			
			var _csectFat			: uint 		= stream.readUnsignedInt();		// [2CH,04] number of SECTs in the FAT chain
			var _sectDirStart 		: uint 		= stream.readUnsignedInt();		// [30H,04] first SECT in the directory chain
			
			var _txSignature 		: uint 		= stream.readUnsignedInt();     // [34H,04] signature used for transactions; must be zero. The reference implementation does not support transactions
			
			_ulMiniSectorCutoff 				= stream.readUnsignedInt();		// [38H,04] maximum size for a mini stream; typically 4096 bytes
			var _sectMiniFatStart 	: uint		= stream.readUnsignedInt();		// [3CH,04] first SECT in the MiniFAT chain
			var _csectMiniFat 		: uint 		= stream.readUnsignedInt();		// [40H,04] number of SECTs in the MiniFAT chain
			
			//double-indirect file allocation table (DIFAT): A structure used to locate FATsectors in a compound file.
			var _sectDifStart 		: uint 		= stream.readUnsignedInt();		// [44H,04] first SECT in the DIFAT chain
			var _csectDif 			: uint 		= stream.readUnsignedInt();		// [48H,04] number of SECTs in the DIFAT chain
			// The DIFAT consists of 32-bit sector numbers that point to the sectors used by the Compound Document File Allocation Table.
			// If there is more than one sector used to contain the contents of the FAT, then those consecutive sectors will be listed
			// sequentially in the DIFAT. DIFAT sectors within the Compound Document file are listed within the FAT as a special value,
			// 0xFFFFFFFC (DIFSECT) and thus are not chained in the FAT. Instead, the last four bytes of a DIFAT sector contain the sector number of the next DIFAT sector. If there are no more DIFAT sectors, then the last four bytes of the DIFAT sector will be the EOC marker (0xFFFFFFFE).
			
			
			var i 					: uint 		= 0;
			
			//_sectFat[109];         											// [4CH,436] the SECTs of first 109 FAT sectors
			// At that point the stream position is 76 (0x4C)
			// Build the sector allocation table on the first 109 sectors
			
			_sectFat = getSectFAT();
			
			/* 
			* Then we add the FATs pointed by the DIFATs
			* The DIFAT array is used to represent storage of the FAT sectors. The DIFAT is represented by an array of 32-bit sector numbers.
			* The DIFAT array is stored both in the header and in DIFAT sectors. In the header, the DIFAT array occupies 109 entries,
			* and in each DIFAT sector, the DIFAT array occupies the entire sector minus 4 bytes (the last field is for chaining the DIFAT sector chain).
			*/
			getSectDIFAT(_sectDifStart,_csectDif,_uDllVersion,_sectFat);
			
			// Now load the directories
			loadDirectory(_sectDirStart);
			
			
			// get the short-stream container stream (sscs)
			var shortStreamContainerSize_dir : Directory = dir[0];
			if (shortStreamContainerSize_dir.type !== STGTY_ROOT) throw new Error("Directory entry type error");
			if (shortStreamContainerSize_dir.secId >= DIFSECT && shortStreamContainerSize_dir.size === 0){
				shortStreamContainerStream = null;
			} else {
				shortStreamContainerStream = loadStream(shortStreamContainerSize_dir.secId,shortStreamContainerSize_dir.size);
			}
			
			// build the short-sector allocation table (ssat)
			miniFatSectorAllocationTable = [];
			if (_csectMiniFat > 0 && shortStreamContainerSize_dir.size === 0) {
				throw new Error("Inconsistency: SSCS size is 0 but SSAT size is non-zero");
			}
			if (shortStreamContainerSize_dir.size > 0) {
				var ba:ByteArray = loadStream(_sectMiniFatStart,shortStreamContainerSize_dir.size);
				for (i = 0; i < ba.length/4; i++){
					miniFatSectorAllocationTable.push(ba.readInt());
				}
			}
		}
		
		
		private function loadDirectory(rootDirSID:uint,relativeIndex:uint = 0):void{
			if(dir == null){
				dir = new Array();
			}
			
			var directory		: Directory = Directory.loadDirectory(stream,relativeIndex,FAT_HEADER_SIZE,fatSectorSize);
			
			// 2 or 5
			if(directory.type == STGTY_STREAM || directory.type == STGTY_ROOT) {
				dir[relativeIndex] = directory;
				
				if(directory._sidLeftSib!= FREESECT){
					loadDirectory(rootDirSID,directory._sidLeftSib);
				}
				
				if(directory._sidRightSib!= FREESECT){
					loadDirectory(rootDirSID,directory._sidRightSib);
				}
				
				if(directory._sidChild!= FREESECT){
					loadDirectory(rootDirSID,directory._sidChild);
				}
			} else {
				
			}
		}
		
		/**
		 * Returns the stream at the given sector id
		 * @param startSecID the first sector id of the stream to extract
		 * 
		 * @return The stream starting at the given sector id as a ByteArray
		 *
		 */
		private function loadStream(startSecID : uint,size : uint, endianness : String = Endian.LITTLE_ENDIAN):ByteArray {
			var defragmentedStream :ByteArray = new ByteArray();
			defragmentedStream.endian = endianness;
			stream.position = sectorOffset(startSecID);
			
			var secId					: uint = startSecID;
			var startOffset				: uint = sectorOffset(secId);
			var endSectOffset			: uint = sectorOffset(secId);
			var offset					: uint = 0;
			var counter 				: uint = 0;
			var amountOfTheSectorToLoad : uint = fatSectorSize;
			
			try{
				while((offset<size && (secId < DIFSECT))) {	//CompoundDocumentFormatReader.FATSECT;
					
					
					stream.position = sectorOffset(secId);
					
					//we load just the amount indicated in the directory when less than fatSectorSize are left
					if(offset + fatSectorSize > size){
						amountOfTheSectorToLoad = size-offset;
					}
					
					stream.readBytes(defragmentedStream, offset, amountOfTheSectorToLoad);
					offset += amountOfTheSectorToLoad;
						
					var nextSecId : uint = _sectFat[secId];
					/*
					if(nextSecId >= DIFSECT){
						while(nextSecId >= DIFSECT){
							nextSecId = _sectFat[++secId];
						}
						nextSecId = secId-1;
					}
					*/
					secId = nextSecId;
					++ counter;
				}
			} catch (error:Error){
				trace(error.getStackTrace());
			}
			
			return defragmentedStream;
		}
		
		/**
		 * Converts a sector id into an absolute offset into the raw ByteArray
		 * File Offset = (Sector Size * Sector Number) + HEADER_SIZE where HEADER_SIZE = 512
		 * 
		 * A SECT can be converted into a byte offset into the file by using the following formula:
		 * SECT << ssheader._uSectorShift + sizeof(ssheader).
		 * This implies that sector 0 of the file begins at byte offset 512, not at 0.
		 * 
		 * 
		 * 
		 * @param secId The sector id to convert
		 * @return The absolute offset of the given sector id
		 *
		 */
		private function sectorOffset(secId:uint):uint {
			//FAT_HEADER_SIZE = 512
			return FAT_HEADER_SIZE + secId * fatSectorSize;
		}
		
		/**
		 * builds the stream from the miniFAT
		 *  
		 * @param startSecID
		 * @return 
		 * 
		 */
		private function loadShortStream(startSecID : uint,size : uint ,endianness:String = Endian.LITTLE_ENDIAN):ByteArray {
			var ret 		: ByteArray 	= new ByteArray();
			ret.endian 						= endianness;
			var secId		: uint 			= startSecID;
			var offset		: uint 			= 0;
			while(offset  < size) {
				shortStreamContainerStream.position 			= secId*miniFatSectorSize;
				shortStreamContainerStream.readBytes(ret, offset, miniFatSectorSize);
				offset 					+= miniFatSectorSize;
				secId 					= miniFatSectorAllocationTable[secId];
			}
			return ret;
		}
		
		
		
		/**
		 * The Fat is the main allocator for space within a Compound File. Every sector in the file is represented within the Fat
		 * in some fashion, including those sectors that are unallocated (free). The Fat is a virtual stream made up of one or
		 * more Fat Sectors.
		 * Fat sectors are arrays of SECTs that represent the allocation of space within the file. Each stream is represented in the
		 * Fat by a chain, in much the same fashion as a DOS file-allocation-table (FAT). To elaborate, the set of Fat Sectors can be
		 * considered together to be a single array -- each cell in that array contains the SECT of the next sector in the chain,
		 * and this SECT can be used as an index into the Fat array to continue along the chain. Special values are reserved for
		 * chain terminators (ENDOFCHAIN = 0xFFFFFFFE), free sectors (FREESECT = 0xFFFFFFFF), and sectors that contain storage for
		 * Fat Sectors (FATSECT = 0xFFFFFFFD) or DIF Sectors (DIFSECT = 0xFFFFFFC), which are notchained in the same way as the others.
		 * 
		 * The locations of Fat Sectors are read from the DIF (Double-indirect Fat), which is described below. The Fat is represented in itself,
		 * but not by a chain –a special reserved SECT value (FATSECT = 0xFFFFFFFD) is used to mark sectors allocated to the Fat.
		 * A SECT can be converted into a byte offset into the file by using the following formula: SECT << ssheader._uSectorShift + sizeof(ssheader).
		 * This implies that sector 0 of the file begins at byte offset 512, not at 0.
		 * 
		 * @param sectFat
		 * @param sectFat_length Compound Document spec size is 109
		 * @return 
		 * 
		 */
		private function getSectFAT(sectFat: Array = null,sectFat_length:uint = 109,majorVersion:uint = 3):Array{
			//_sectFat[109];         											// [4CH,436] the SECTs of first 109 FAT sectors
			// At that point the stream position is 76 (0x4C)
			// Build the sector allocation table
			var initialPosition		: uint 		= 0x4C;
			var difat_109			: Array		= new Array();
			var localSectFat		: Array 	= sectFat;
			var sectorSize 			: uint 		= majorVersion == 3? 512:4096;
			var wordSize			: uint		= 4;
			var fieldsPerSect		: uint 		= sectorSize /wordSize;
			var _secId				: uint;
			var previous 			: uint		= 0;
			var next				: uint		= 0;
			var i 					: uint 		= 0;
			var fat_109_length		: uint 		= 109*128;
			
			if(localSectFat == null){
				localSectFat 					= new Array(109*128);
				while(i<fat_109_length){
					localSectFat[i++] 			= FREESECT;
				}
			}
			
			for(i = 0; i < sectFat_length; i++) {
				stream.position 				= initialPosition + i*wordSize;
				difat_109.push(stream.readUnsignedInt());
			}
			
			var sectAlloc:ByteArray 			= new ByteArray();
			sectAlloc.endian 					= Endian.LITTLE_ENDIAN;
			
			//localSectFat 						= new Array();
			var index 				: uint 		= 0;
			i									= 0;
			for each(_secId in difat_109){
				
				if(_secId >= DIFSECT){
					continue;
				}
				
				stream.position = sectorOffset(_secId);
				index = fieldsPerSect;
				while(index--){
					next = stream.readUnsignedInt();
					localSectFat[i++]	= next;
				}
			}
			return localSectFat;
		}
		
		/**
		 * adds the DIFAT sectors to the 109 FAT
		 * DIFAT sectors exist when more than 109*128 sectors compose the file
		 * 
		 * 
		 * Dif sectors
		 * 			DIF Sector
		 *		_________________________
		 * 		|						|
		 *		|	pointers to FAT		|
		 *		|	sectors				|
		 *		|						|
		 *		|						|
		 *		|						|
		 *		|						|
		 *		|						|
		 *		|				________|
		 *		|_______________|_______| pointer to DIFAT sector
		 * 
		 * 	 
		 * 
		 * The Double-Indirect Fat is used to represent storage of the Fat. The DIF is also represented by an array of SECTs,
		 * and is chained by the terminating cell in each sector array (see the diagram above). As an optimization,
		 * the first 109 Fat Sectors are represented within the header itself, so no DIF sectors will be found in a small (< 7 MB) Compound File.
		 * The DIF represents the Fat in a different manner than the Fat represents a chain. A given index into the DIF will contain the SECT
		 * of the Fat Sector found at that offset in the Fat virtual stream. For instance, index 3 in the DIF would contain the SECT for Sector #3 of the Fat.
		 * The storage for DIF Sectors is reserved in the Fat, but is not chained there (space for it is reserved by a special SECT value , DIFSECT=0xFFFFFFFC).
		 * The location of the first DIF sector is stored in the header.
		 * A value of ENDOFCHAIN=0xFFFFFFFE is stored in the pointer to the next DIF sector of the last DIF sector.
		 * 
		 * 
		 * 
		 * @param firstSECT_inTheDIFAT_chain pointer to a sector that contains a list of FAT ptrs
		 * @param csectDif 	 number of SECTs in the DIFAT chain, these are contiguous sectors
		 * @param majorVersion  Header Major Version version of the dll/format
		 * 				If Header Major Version is 3, then there MUST be 127 fields specified to fill a 512-byte sector minus the "Next DIFAT Sector Location" field.
		 * 				If Header Major Version is 4, then there MUST be 1023 fields specified to fill a 4096-byte sector minus the "Next DIFAT Sector Location" field.
		 * @param sectFat the sectFat table to add sect to
		 * 
		 */		
		private function getSectDIFAT(firstSECT_inTheDIFAT_chain:uint,csectDif:uint,majorVersion:uint,sectFat : Array):void{
			
			if(firstSECT_inTheDIFAT_chain == ENDOFCHAIN || csectDif == 0){
				return;
			}
			
			var _secId			: uint;
			var next			: uint;
			
			//we get the previous from the FAT as if sectors were contiguous
			var previous 		: uint 		= sectFat.length;//supposed to be 109*128 = 13952
			var difat_extension : Array 	= new Array();
			var sectorSize 		: uint 		= majorVersion == 3? 512:4096;
			var fieldsPerSect	: uint 		= sectorSize /4;
			var numberOfDifat	: uint;
			var difIndex 		: uint 		= csectDif;
			var fieldIndex 		: uint 		= fieldsPerSect;
			var nextDifSid		: uint 		= firstSECT_inTheDIFAT_chain;
			var DIFATSectors	: Array		= [];
			var currentDifat	: uint		= 0;
			var starts			: Array		= [];
			
			//get the DIFATs sector Ids
			//-o  stream.position = sectorOffset(nextDifSid);
			fieldIndex = fieldsPerSect;
			while(difIndex--){
				//-n
				stream.position = sectorOffset(nextDifSid);
				//-n
				while(true){
					_secId = stream.readUnsignedInt();
					//if we reach the last field of the a DIFAT we get the next chained difat
					if(--fieldIndex == 0  ){
						//pointer to next DIFAT
						DIFATSectors[currentDifat++] = _secId;
						//-n
						nextDifSid = _secId;
						//++nextDifSid;
						//-n
						
						fieldIndex = fieldsPerSect;
						break;
					}
					difat_extension.push(_secId);
				}
			}
			
			//get the FAT corresponding to the DIFAT
			for each(_secId in difat_extension){
				
				if(_secId >= DIFSECT){
					continue;
				}
				
				stream.position = sectorOffset(_secId);
				fieldIndex = fieldsPerSect;
				while(fieldIndex--){
					next = stream.readUnsignedInt();
					if(next< DIFSECT){
						sectFat[previous] = next;
						starts.push((next+1)*sectorSize);
						previous = next;
					}
				}
			}
		}
	}
}