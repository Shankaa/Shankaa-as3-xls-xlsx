package org.as3.ms.cdf
{
	import flash.utils.ByteArray;

	public class Directory
	{
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
		
		private static const 	DIRECTORY_LENGTH: uint = 16*8;// dir have a 128 byte size
		private static const 	DIRECTORY_SID	: uint = 0x30;
		
		public var name			: String;
		//public var secId		: uint;
		public var size			: uint;				//ULONG		_ulSize;			[078H,04]	size of stream in bytes (if _mse=STGTY_STREAM)
		public var type 		: uint ;			//BYTE		_mse;				[042H,01]	Type of object: value taken from the STGTY enumeration
		
		//protected var _mse 		: uint ;		//BYTE		_mse;				[042H,01]	Type of object: value taken from the STGTY enumeration
		protected var nodeColor : uint;				//BYTE		_bflags;			[043H,01]	Value taken from DECOLOR enumeration.
		public var _sidLeftSib	: uint;				//SID		_sidLeftSib;		[044H,04]	SID of the left	- sibling of this entry in the directory tree
		public var _sidRightSib	: uint;				//SID		_sidRightSib;		[048H,04]	SID of the right- sibling of this entry in the directory tree
		public var _sidChild	: uint;				//SID		_sidChild;			[04CH,04]	SID of the child acting as the root of all the children of this element (if _mse=STGTY_STORAGE)
		
		protected var _clsId		: ByteArray;	//GUID		_clsId;				[050H,16]	CLSID of this storage (if _mse=STGTY_STORAGE)
		protected var _dwUserFlags: uint	;		//DWORD		_dwUserFlags;		[060H,04]	User flags of this storage (if _mse=STGTY_STORAGE)
		protected var _time		: ByteArray;		//TIME_T	_time[2];			[064H,16]	Create/Modify time	-stamps (if _mse=STGTY_STORAGE)
		
		public var secId 		: uint;				//SECT		_sectStart			[074H,04]	starting SECT of the stream (if _mse=STGTY_STREAM)
		//public var _ulSize 		: uint;			
		protected var _dptPropType	: uint;			//DFPROPTYPE _dptPropType;		[07CH,02]	Reserved for future use. Must be zero.
		
		/**
		 * typedef enum tagSTGTY {
	 	 * 		STGTY_INVALID	= 0,
	 	 * 		STGTY_STORAGE	= 1,
	 	 * 		STGTY_STREAM	= 2,
	 	 * 		STGTY_LOCKBYTES	= 3,
	 	 * 		STGTY_PROPERTY	= 4,
		 * 		STGTY_ROOT		= 5,
	 	 * }STGTY; 
		 */
		
		public function Directory(){
			
		}

		
		/**
		 * 
		 * @param dirName
		 * @param type 		Type of object: value taken from the STGTY enumeration
		 * 						STGTY_INVALID	= 0,
		 * 						STGTY_STORAGE	= 1,
		 * 						STGTY_STREAM	= 2,
		 * 						STGTY_LOCKBYTES	= 3,
		 * 						STGTY_PROPERTY	= 4,
		 * 						STGTY_ROOT		= 5,
		 */
		private function writeDirectory(dirName:String,type:uint):void{
			/*
			* 		STGTY_INVALID	= 0,
			* 		STGTY_STORAGE	= 1,
			* 		STGTY_STREAM	= 2,
			* 		STGTY_LOCKBYTES	= 3,
			* 		STGTY_PROPERTY	= 4,
			* 		STGTY_ROOT		= 5,
			*		
			*		DE_RED			= 0,
			* 		DE_BLACK		= 1,
			*/
			var _dir0		: ByteArray			= new ByteArray();
			_dir0.length						= DIRECTORY_LENGTH;
			name= dirName.substr(0,32);
			
		}
		
		/**
		 * 
		 * @param referenceStream
		 * @param rootDirSID
		 * @param relativePosition
		 * @param headerSize
		 * @param fatSectorSize
		 * @return 
		 * 
		 */
		public static function loadDirectory(referenceStream:ByteArray,relativeIndex:uint = 0,headerSize : uint = 512,fatSectorSize:uint = 512):Directory{
			/*
			* 		STGTY_INVALID	= 0,
			* 		STGTY_STORAGE	= 1,
			* 		STGTY_STREAM	= 2,
			* 		STGTY_LOCKBYTES	= 3,
			* 		STGTY_PROPERTY	= 4,
			* 		STGTY_ROOT		= 5,
			*		
			*		DE_RED			= 0,
			* 		DE_BLACK		= 1,
			*/
			
			var directory		: Directory	= new Directory();
			
			var sectorOffset 	: Function 	= function (secId:uint):uint {
				return headerSize + secId * fatSectorSize;
			}
			
			referenceStream.position 		= DIRECTORY_SID;
			var rootDirSID		:uint		= referenceStream.readUnsignedInt();
			
			referenceStream.position		= sectorOffset(rootDirSID);
			referenceStream.position 		+= relativeIndex*DIRECTORY_LENGTH;
			
			
			
			//BYTE	_ab[32*sizeof(WCHAR)];	[000H,64]64 bytes. The Element name in Unicode, padded with zeros to fill this byte array
			directory.name = "";
			for(var i:uint = 0; i < 32; i++) {
				directory.name 				+= String.fromCharCode(referenceStream.readUnsignedByte());	
				referenceStream.position++;
			}
			var nameLen 		: uint		= referenceStream.readUnsignedShort();		//WORD		_cb;				[040H,02] 	Length of the Element name in characters, not bytes
			directory.name 					= directory.name.substr(0, (nameLen-2)/2);
			directory.type		 			= referenceStream.readUnsignedByte();		//BYTE		_mse;				[042H,01]	Type of object: value taken from the STGTY enumeration
			directory.nodeColor 			= referenceStream.readUnsignedByte();		//BYTE		_bflags;			[043H,01]	Value taken from DECOLOR enumeration.
			directory._sidLeftSib		 	= referenceStream.readUnsignedInt();		//SID		_sidLeftSib;		[044H,04]	SID of the left	- sibling of this entry in the directory tree
			directory._sidRightSib		 	= referenceStream.readUnsignedInt();		//SID		_sidRightSib;		[048H,04]	SID of the right- sibling of this entry in the directory tree
			directory._sidChild				= referenceStream.readUnsignedInt();		//SID		_sidChild;			[04CH,04]	SID of the child acting as the root of all the children of this element (if _mse=STGTY_STORAGE)
			
			directory._clsId				= new ByteArray();			
			referenceStream.readBytes(directory._clsId,0,16); 							//GUID		_clsId;				[050H,16]	CLSID of this storage (if _mse=STGTY_STORAGE)
			directory._dwUserFlags			= referenceStream.readUnsignedInt();		//DWORD		_dwUserFlags;		[060H,04]	User flags of this storage (if _mse=STGTY_STORAGE)
			directory._time					= new ByteArray();
			referenceStream.readBytes(directory._clsId,0,16);							//TIME_T	_time[2];			[064H,16]	Create/Modify time	-stamps (if _mse=STGTY_STORAGE)
			
			directory.secId		 			= referenceStream.readUnsignedInt();		//SECT		_sectStart			[074H,04]	starting SECT of the stream (if _mse=STGTY_STREAM)
			directory.size 					= referenceStream.readUnsignedInt();		//ULONG		_ulSize;			[078H,04]	size of stream in bytes (if _mse=STGTY_STREAM)
			directory._dptPropType			= referenceStream.readUnsignedShort();		//DFPROPTYPE _dptPropType;		[07CH,02]	Reserved for future use. Must be zero.
			
			
			return directory;
		}
		
		public function getDirectoryStream(compoundDocumentStream:ByteArray):ByteArray{
			var stream 		: ByteArray = new ByteArray();
			
			
			return stream;
		}
	}
}