package org.as3.ms.ppt
{
	import flash.events.EventDispatcher;
	import flash.utils.ByteArray;
	
	import org.as3.ms.biff.BIFFReader;
	import org.as3.ms.biff.BIFFVersion;
	import org.as3.ms.biff.Record;
	import org.as3.ms.cdf.CompoundDocument;
	
	
	/**
	 * 
	 * @author Hugues Sansen, Shankaa
	 * This PPTFile class mimics the architecture of the ExcelFile for consitency
	 * 
	 * It is based on
	 * Microsoft Office PowerPoint 97-2007 Binary File Format (.ppt) Specification
	 * 
	 * 
	 * MICROSOFT OFFICE POWERPOINT 97-2007 BINARY FILE FORMAT SPECIFICATION
	 * 								[*.ppt]
	 * Includes Binary File Format Documentation Relevant To:
	 * 	Microsoft Office PowerPoint 2007
	 * 	Microsoft Office PowerPoint 2003
	 * 	Microsoft Office PowerPoint 2002
	 * 	Microsoft Office PowerPoint 2000
	 * 	Microsoft Office PowerPoint 1997
	 * 
	 * Current User Stream
	 * The Current User Stream contains a pointer to the latest saved edit in the document stream.
	 * The document stream contains one or more user edit structures.
	 * A graphical representation of this looks like:
	 * 
	 * 									 __________________________________________
	 * 									 |			PowerPoint Document			  |
	 * 									 |________________________________________|
	 *									 __________________________________________
	 *									 | _____________________________________  |	
	 *									 | |									| |
	 *									 | |				UserEdit			|<|-+
	 *									 | |									| |	|
	 *									 | |____________________________________| |	|
	 *									 |										| | |
	 *									 |										| | |
	 *									 | _____________________________________| |	|
	 *									 | |									| |	|
	 *									 | |				UserEdit			| |	|
	 *									 | |									|<|-|---+
	 *									 | |									|_|_|   |
	 *									 | |____________________________________| |		|
	 *		____________________		 | |				lastEdit			| |		|
	 *		|	Current User	|		 | |____________________________________| |		|
	 *		|___________________|		 | _____________________________________  |		|
	 *		| Current User Atom	|___	 | |									| |		|
	 *		|___________________|	|	 | |				UserEdit			| |		|
	 *								+----|>|					 				| |		|
	 *									 | |									|_|_____|
	 *									 | |____________________________________| |
	 *									 | |				lastEdit			| |
	 *									 | |____________________________________| |
	 *									 | |____________________________________| |
	 * 			 						 |										  |
	 *									 |________________________________________|
	 * 
	 * UserEditAtom Structure
	 * The UserEditAtom structure is as follows: 
	 * 		struct PSR_UserEditAtom { 
	 * 				sint4 lastSlideID; // slideID of last viewed slide 
	 * 				uint4 version; // This is major/minor/build which did the edit 
	 * 				uint4 offsetLastEdit; // File offset of last edit 
	 * 				uint4 offsetPersistDirectory; // Offset to PersistPtrs for this edit. 
	 * 				uint4 documentRef; // reference to document atom 
	 * 				uint4 maxPersistWritten; // Addr of last persist ref written to the file (max seen so far). 
	 * 				sint2 lastViewType; // enum view type 
	 * 		};
	 * 
	 * UserEditAtom Element Descriptions
	 * 		* lastSlideID and lastViewType: SlideID of last slide viewed and view type for saved view,
	 * 				respectively. Allow a document window to be opened in its saved configuration.
	 * 		* version: Major/minor/build which did the edit.
	 * 	 	* offsetLastEdit: Pointer to the last user edit. This is a 32 bit fixed offset from the beginning
	 * 				of the file. (This is 0 if no previous edits exist. It is illegal to place a LastEdit structure
	 * 				at offset 0 in the file.)
	 * 		* offsetPersistDirectory: Contains the persistent references (32 bit offset from the beginning
	 * 				of the document stream) in the current user edit. References are number sequentially from 1 
	 * 				(0 is not a valid value) and each user edit will contain a persistent directory.
	 * 				This directory contains only the references made by the current user and the document data
	 * 				included in the edit. To find additional references, PowerPoint begins with the directory of
	 * 				the last edit and then searches recursively through the previous edits until the reference is found.
	 * 
	 * 		The persistent directory is encoded as follows:
	 * 		_____________________________________  __________________________________
	 * 		|	Sequential          Base		|  |	Offset (Sequential times)	|
	 * 		|___________________________________|  |________________________________|
	 * 
	 * 		12 bit value		 20 bit value
	 * 		which is			 indicates current
	 * 		number				 reference number	
	 * 
	 * 		* documentRef: Reverence to the document atom.
	 *		* maxPersistWritten: Address of the last persist ref written to the file. This is the maximum value
	 *			 contained in the file, maintained so that new user edits can be properly numbered.
	 * 
	 * 		Persistent Directory Example
	 * 		Suppose the current save of a PowerPoint document contains the following: Reference File Offset
	 * 				1 	1024
	 * 				2 	2048
	 * 				3 	4096
	 * 				6 	8196
	 * 				8 	10000
	 * 				9 	20000
	 * 
	 * 		The following would be saved to the file: 
	 * 
	 * 		Hex			Decimal			Meaning
	 * 
	 * 		1772		6002 			PST_PersistPtrIncrementalBlock
	 * 		24 			36 				Length of Atom
	 * 		300001 		3145729 		3 consecutive offsets starting at 1 
	 * 		400 		1024 			Offset to ref(1) 
	 * 		800 		2048 			Offset to ref(2) 
	 * 		1000 		4096 			Offset to ref(3)
	 * 		100006 		1048582 		1 consecutive refs starting at 6
	 * 		2000 		8192 			Offset to ref(6)
	 * 		200008 		2087160 		2 consecutive refs starting at 8
	 * 		2710 		10000 			Offset to ref(8) 
	 * 		4E20 		20000 			Offset to ref(9)
	 * 		
	 * 
	 * 		PowerPoint Document Stream
	 * 		The PowerPoint Document Stream keeps all the information about a PowerPoint presentation.
	 * 		A PowerPoint file stores its data in records (see Appendix B). There are two different
	 * 		kinds of records in a file: atoms and containers. We could, as with storages and streams,
	 * 		compare atoms and containers to files and directories, respectively. Atoms, like files,
	 * 		keep the actual information. Containers, just like directories, can contain files and other directories.
	 * 
	 * 		Atoms: 		Records that contain information about a PowerPoint object and are kept inside containers.
	 * 		Containers: Records that keep atoms and other containers in a logical and organized way.

	 * 
	 * 
	 * 
	 * 
	 */
	public class PPTFile extends EventDispatcher
	{
		
		private var biffReader		: BIFFReader;
		private var version			: uint;
		private var dateMode		: uint;
		
		private const readHandlers		: Array = initReadHandlers();
		
		private function initReadHandlers():Array {
			var handlers:Array = new Array();
			handlers[Type.Unknown]									= read_Unknown;
			handlers[Type.SubContainerCompleted]					= read_SubContainerCompleted;
			handlers[Type.IRRAtom]									= read_IRRAtom;
			handlers[Type.PSS]										= read_PSS;
			handlers[Type.SubContainerException]					= read_SubContainerException;
			handlers[Type.ClientSignal1]							= read_ClientSignal1;
			handlers[Type.ClientSignal2]							= read_ClientSignal2;
			handlers[Type.PowerPointStateInfoAtom]					= read_PowerPointStateInfoAtom;
			handlers[Type.Document]									= read_Document;
			handlers[Type.DocumentAtom]								= read_DocumentAtom;
			handlers[Type.EndDocument]								= read_EndDocument;
			handlers[Type.SlidePersist]								= read_SlidePersist;
			handlers[Type.SlideBase]								= read_SlideBase;
			handlers[Type.SlideBaseAtom]							= read_SlideBaseAtom;
			handlers[Type.Slide]									= read_Slide;
			handlers[Type.SlideAtom]								= read_SlideAtom;
			handlers[Type.Notes]									= read_Notes;
			handlers[Type.NotesAtom]								= read_NotesAtom;
			handlers[Type.Environment]								= read_Environment;
			handlers[Type.SlidePersistAtom]							= read_SlidePersistAtom;
			handlers[Type.Scheme]									= read_Scheme;
			handlers[Type.SchemeAtom]								= read_SchemeAtom;
			handlers[Type.DocViewInfo]								= read_DocViewInfo;
			handlers[Type.SslideLayoutAtom]							= read_SslideLayoutAtom;
			handlers[Type.MainMaster]								= read_MainMaster;
			handlers[Type.SSSlideInfoAtom]							= read_SSSlideInfoAtom;
			handlers[Type.SlideViewInfo]							= read_SlideViewInfo;
			handlers[Type.GuideAtom]								= read_GuideAtom;
			handlers[Type.ViewInfo]									= read_ViewInfo;
			handlers[Type.ViewInfoAtom]								= read_ViewInfoAtom;
			handlers[Type.SlideViewInfoAtom]						= read_SlideViewInfoAtom;
			handlers[Type.VBAInfo]									= read_VBAInfo;
			handlers[Type.VBAInfoAtom]								= read_VBAInfoAtom;
			handlers[Type.SSDocInfoAtom]							= read_SSDocInfoAtom;
			handlers[Type.Summary]									= read_Summary;
			handlers[Type.Texture]									= read_Texture;
			handlers[Type.VBASlideInfo]								= read_VBASlideInfo;
			handlers[Type.VBASlideInfoAtom]							= read_VBASlideInfoAtom;
			handlers[Type.DocRoutingSlip]							= read_DocRoutingSlip;
			handlers[Type.OutlineViewInfo]							= read_OutlineViewInfo;
			handlers[Type.SorterViewInfo]							= read_SorterViewInfo;
			handlers[Type.ExObjList]								= read_ExObjList;
			handlers[Type.ExObjListAtom]							= read_ExObjListAtom;
			handlers[Type.PPDrawingGroup]							= read_PPDrawingGroup;
			handlers[Type.PPDrawing]								= read_PPDrawing;
			handlers[Type.Theme]									= read_Theme;
			handlers[Type.ColorMapping]								= read_ColorMapping;
			handlers[Type.NamedShows]								= read_NamedShows;
			handlers[Type.NamedShow]								= read_NamedShow;
			handlers[Type.NamedShowSlides]							= read_NamedShowSlides;
			handlers[Type.OriginalMainMasterId]						= read_OriginalMainMasterId;
			handlers[Type.CompositeMasterId]						= read_CompositeMasterId;
			handlers[Type.RoundTripContentMasterInfo12]				= read_RoundTripContentMasterInfo12;
			handlers[Type.RoundTripShapeId12]						= read_RoundTripShapeId12;
			handlers[Type.RoundTripHFPlaceholder12]					= read_RoundTripHFPlaceholder12;
			handlers[Type.RoundTripContentMasterId12]				= read_RoundTripContentMasterId12;
			handlers[Type.RoundTripOArtTextStyles12]				= read_RoundTripOArtTextStyles12;
			handlers[Type.HeaderFooterDefaults12]					= read_HeaderFooterDefaults12;
			handlers[Type.DocFlags12]								= read_DocFlags12;
			handlers[Type.RoundTripShapeCheckSumForCustomLayouts12]	= read_RoundTripShapeCheckSumForCustomLayouts12;
			handlers[Type.RoundTripNotesMasterTextStyles12]			= read_RoundTripNotesMasterTextStyles12;
			handlers[Type.RoundTripCustomTableStyles12]				= read_RoundTripCustomTableStyles12;
			handlers[Type.List]										= read_List;
			handlers[Type.FontCollection]							= read_FontCollection;
			handlers[Type.ListPlaceholder]							= read_ListPlaceholder;
			handlers[Type.BookmarkCollection]						= read_BookmarkCollection;
			handlers[Type.SoundCollection]							= read_SoundCollection;
			handlers[Type.SoundCollAtom]							= read_SoundCollAtom;
			handlers[Type.Sound]									= read_Sound;
			handlers[Type.SoundData]								= read_SoundData;
			handlers[Type.BookmarkSeedAtom]							= read_BookmarkSeedAtom;
			handlers[Type.GuideList]								= read_GuideList;
			handlers[Type.RunArray]									= read_RunArray;
			handlers[Type.RunArrayAtom]								= read_RunArrayAtom;
			handlers[Type.ArrayElementAtom]							= read_ArrayElementAtom;
			handlers[Type.Int4ArrayAtom]							= read_Int4ArrayAtom;
			handlers[Type.ColorSchemeAtom]							= read_ColorSchemeAtom;
			handlers[Type.OEShape]									= read_OEShape;
			handlers[Type.ExObjRefAtom]								= read_ExObjRefAtom;
			handlers[Type.OEPlaceholderAtom]						= read_OEPlaceholderAtom;
			handlers[Type.GrColor]									= read_GrColor;
			handlers[Type.GrectAtom]								= read_GrectAtom;
			handlers[Type.GratioAtom]								= read_GratioAtom;
			handlers[Type.Gscaling]									= read_Gscaling;
			handlers[Type.GpointAtom]								= read_GpointAtom;
			handlers[Type.OEShapeAtom]								= read_OEShapeAtom;
			handlers[Type.OEPlaceholderNewPlaceholderId12]			= read_OEPlaceholderNewPlaceholderId12;
			handlers[Type.OutlineTextRefAtom]						= read_OutlineTextRefAtom;
			handlers[Type.TextHeaderAtom]							= read_TextHeaderAtom;
			handlers[Type.TextCharsAtom]							= read_TextCharsAtom;
			handlers[Type.StyleTextPropAtom]						= read_StyleTextPropAtom;
			handlers[Type.BaseTextPropAtom]							= read_BaseTextPropAtom;
			handlers[Type.TxMasterStyleAtom]						= read_TxMasterStyleAtom;
			handlers[Type.TxCFStyleAtom]							= read_TxCFStyleAtom;
			handlers[Type.TxPFStyleAtom]							= read_TxPFStyleAtom;
			handlers[Type.TextRulerAtom]							= read_TextRulerAtom;
			handlers[Type.TextBookmarkAtom]							= read_TextBookmarkAtom;
			handlers[Type.TextBytesAtom]							= read_TextBytesAtom;
			handlers[Type.TxSIStyleAtom]							= read_TxSIStyleAtom;
			handlers[Type.TextSpecInfoAtom]							= read_TextSpecInfoAtom;
			handlers[Type.DefaultRulerAtom]							= read_DefaultRulerAtom;
			handlers[Type.FontEntityAtom]							= read_FontEntityAtom;
			handlers[Type.FontEmbedData]							= read_FontEmbedData;
			handlers[Type.TypeFace]									= read_TypeFace;
			handlers[Type.CString]									= read_CString;
			handlers[Type.ExternalObject]							= read_ExternalObject;
			handlers[Type.MetaFile]									= read_MetaFile;
			handlers[Type.ExOleObj]									= read_ExOleObj;
			handlers[Type.ExOleObjAtom]								= read_ExOleObjAtom;
			handlers[Type.ExPlainLinkAtom]							= read_ExPlainLinkAtom;
			handlers[Type.CorePict]									= read_CorePict;
			handlers[Type.CorePictAtom]								= read_CorePictAtom;
			handlers[Type.ExPlainAtom]								= read_ExPlainAtom;
			handlers[Type.SrKinsoku]								= read_SrKinsoku;
			handlers[Type.Handout]									= read_Handout;
			handlers[Type.ExEmbed]									= read_ExEmbed;
			handlers[Type.ExEmbedAtom]								= read_ExEmbedAtom;
			handlers[Type.ExLink]									= read_ExLink;
			handlers[Type.ExLinkAtom_old]							= read_ExLinkAtom_old;
			handlers[Type.BookmarkEntityAtom]						= read_BookmarkEntityAtom;
			handlers[Type.ExLinkAtom]								= read_ExLinkAtom;
			handlers[Type.SrKinsokuAtom]							= read_SrKinsokuAtom;
			handlers[Type.ExHyperlinkAtom]							= read_ExHyperlinkAtom;
			handlers[Type.ExPlain]									= read_ExPlain;
			handlers[Type.ExPlainLink]								= read_ExPlainLink;
			handlers[Type.ExHyperlink]								= read_ExHyperlink;
			handlers[Type.SlideNumberMCAtom]						= read_SlideNumberMCAtom;
			handlers[Type.HeadersFooters]							= read_HeadersFooters;
			handlers[Type.HeadersFootersAtom]						= read_HeadersFootersAtom;
			handlers[Type.RecolorEntryAtom]							= read_RecolorEntryAtom;
			handlers[Type.TxInteractiveInfoAtom]					= read_TxInteractiveInfoAtom;
			handlers[Type.EmFormatAtom]								= read_EmFormatAtom;
			handlers[Type.CharFormatAtom]							= read_CharFormatAtom;
			handlers[Type.ParaFormatAtom]							= read_ParaFormatAtom;
			handlers[Type.MasterText]								= read_MasterText;
			handlers[Type.RecolorInfoAtom]							= read_RecolorInfoAtom;
			handlers[Type.ExQuickTime]								= read_ExQuickTime;
			handlers[Type.ExQuickTimeMovie]							= read_ExQuickTimeMovie;
			handlers[Type.ExQuickTimeMovieData]						= read_ExQuickTimeMovieData;
			handlers[Type.ExSubscription]							= read_ExSubscription;
			handlers[Type.ExSubscriptionSection]					= read_ExSubscriptionSection;
			handlers[Type.ExControl]								= read_ExControl;
			handlers[Type.ExControlAtom]							= read_ExControlAtom;
			handlers[Type.SlideListWithText]						= read_SlideListWithText;
			handlers[Type.AnimationInfoAtom]						= read_AnimationInfoAtom;
			handlers[Type.InteractiveInfo]							= read_InteractiveInfo;
			handlers[Type.InteractiveInfoAtom]						= read_InteractiveInfoAtom;
			handlers[Type.SlideList]								= read_SlideList;
			handlers[Type.UserEditAtom]								= read_UserEditAtom;
			handlers[Type.CurrentUserAtom]							= read_CurrentUserAtom;
			handlers[Type.DateTimeMCAtom]							= read_DateTimeMCAtom;
			handlers[Type.GenericDateMCAtom]						= read_GenericDateMCAtom;
			handlers[Type.HeaderMCAtom]								= read_HeaderMCAtom;
			handlers[Type.FooterMCAtom]								= read_FooterMCAtom;
			handlers[Type.ExMediaAtom]								= read_ExMediaAtom;
			handlers[Type.ExVideo]									= read_ExVideo;
			handlers[Type.ExAviMovie]								= read_ExAviMovie;
			handlers[Type.ExMCIMovie]								= read_ExMCIMovie;
			handlers[Type.ExMIDIAudio]								= read_ExMIDIAudio;
			handlers[Type.ExCDAudio]								= read_ExCDAudio;
			handlers[Type.ExWAVAudioEmbedded]						= read_ExWAVAudioEmbedded;
			handlers[Type.ExWAVAudioLink]							= read_ExWAVAudioLink;
			handlers[Type.ExOleObjStg]								= read_ExOleObjStg;
			handlers[Type.ExCDAudioAtom]							= read_ExCDAudioAtom;
			handlers[Type.ExWAVAudioEmbeddedAtom]					= read_ExWAVAudioEmbeddedAtom;
			handlers[Type.AnimationInfo]							= read_AnimationInfo;
			handlers[Type.RTFDateTimeMCAtom]						= read_RTFDateTimeMCAtom;
			handlers[Type.ProgTags]									= read_ProgTags;
			handlers[Type.ProgStringTag]							= read_ProgStringTag;
			handlers[Type.ProgBinaryTag]							= read_ProgBinaryTag;
			handlers[Type.BinaryTagData]							= read_BinaryTagData;
			handlers[Type.PrintOptions]								= read_PrintOptions;
			handlers[Type.PersistPtrFullBlock]						= read_PersistPtrFullBlock;
			handlers[Type.PersistPtrIncrementalBlock]				= read_PersistPtrIncrementalBlock;
			handlers[Type.RulerIndentAtom]							= read_RulerIndentAtom;
			handlers[Type.GscalingAtom]								= read_GscalingAtom;
			handlers[Type.GrColorAtom]								= read_GrColorAtom;
			handlers[Type.GLPointAtom]								= read_GLPointAtom;
			handlers[Type.GlineAtom]								= read_GlineAtom;
			handlers[Type.AnimationAtom12]							= read_AnimationAtom12;
			handlers[Type.AnimationHashAtom12]						= read_AnimationHashAtom12;
			handlers[Type.SlideSyncInfo12]							= read_SlideSyncInfo12;
			handlers[Type.SlideSyncInfoAtom12]						= read_SlideSyncInfoAtom12;
			
			return handlers;
		}
		
		
		
		public function PPTFile()
		{
			super(null);
		}
		
		
		public function loadFromByteArray(ppt:ByteArray):void{
			
			if(CompoundDocument.isCompoundDocumentFormatFile(ppt)) {
				var cdf:CompoundDocument = new CompoundDocument(ppt);
				ppt = cdf.loadDirectoryEntry(1);
			}
			
			biffReader = new BIFFReader(ppt);
			
			var unknown:Array = [];
			var r:Record;
			
			
			while((r = biffReader.readTag()) != null) {
				
				/*
				if(r.type != Type.CONTINUE) {
					lastRecordType = r.type;
				}
				*/
				if(readHandlers[r.type] is Function) {
					
					
					(readHandlers[r.type] as Function).call(this, r);
				} else {
					unknown.push(r.type);
				}
				
				
			}
		}
		
	//************************************************************************************************************/		
	//	Reader
 	//	 
 	//************************************************************************************************************/
		
		private function read_Unknown(pptRecord : Record):void{
		} 
		private function read_SubContainerCompleted(pptRecord : Record):void{
		}
		private function read_IRRAtom(pptRecord : Record):void{
		}
		private function read_PSS(pptRecord : Record):void{
		}
		private function read_SubContainerException(pptRecord : Record):void{
		}
		private function read_ClientSignal1(pptRecord : Record):void{
		}
		private function read_ClientSignal2(pptRecord : Record):void{
		}
		private function read_PowerPointStateInfoAtom(pptRecord : Record):void{
		}
		private function read_Document(pptRecord : Record):void{
		}
		private function read_DocumentAtom(pptRecord : Record):void{
		}
		private function read_EndDocument(pptRecord : Record):void{
		}
		private function read_SlidePersist(pptRecord : Record):void{
		}
		private function read_SlideBase(pptRecord : Record):void{
		}
		private function read_SlideBaseAtom(pptRecord : Record):void{
		}
		private function read_Slide(pptRecord : Record):void{
		}
		private function read_SlideAtom(pptRecord : Record):void{
		}
		private function read_Notes(pptRecord : Record):void{
		}
		private function read_NotesAtom(pptRecord : Record):void{
		}
		private function read_Environment(pptRecord : Record):void{
		}
		private function read_SlidePersistAtom(pptRecord : Record):void{
		}
		private function read_Scheme(pptRecord : Record):void{
		}
		private function read_SchemeAtom(pptRecord : Record):void{
		}
		private function read_DocViewInfo(pptRecord : Record):void{
		}
		private function read_SslideLayoutAtom(pptRecord : Record):void{
		}
		private function read_MainMaster(pptRecord : Record):void{
		}
		private function read_SSSlideInfoAtom(pptRecord : Record):void{
		}
		private function read_SlideViewInfo(pptRecord : Record):void{
		}
		private function read_GuideAtom(pptRecord : Record):void{
		}
		private function read_ViewInfo(pptRecord : Record):void{
		}
		private function read_ViewInfoAtom(pptRecord : Record):void{
		}
		private function read_SlideViewInfoAtom(pptRecord : Record):void{
		}
		private function read_VBAInfo(pptRecord : Record):void{
		}
		private function read_VBAInfoAtom(pptRecord : Record):void{
		}
		private function read_SSDocInfoAtom(pptRecord : Record):void{
		}
		private function read_Summary(pptRecord : Record):void{
		}
		private function read_Texture(pptRecord : Record):void{
		}
		private function read_VBASlideInfo(pptRecord : Record):void{
		}
		private function read_VBASlideInfoAtom(pptRecord : Record):void{
		}
		private function read_DocRoutingSlip(pptRecord : Record):void{
		}
		private function read_OutlineViewInfo(pptRecord : Record):void{
		}
		private function read_SorterViewInfo(pptRecord : Record):void{
		}
		private function read_ExObjList(pptRecord : Record):void{
		}
		private function read_ExObjListAtom(pptRecord : Record):void{
		}
		private function read_PPDrawingGroup(pptRecord : Record):void{
		}
		private function read_PPDrawing(pptRecord : Record):void{
		}
		private function read_Theme(pptRecord : Record):void{
		}
		private function read_ColorMapping(pptRecord : Record):void{
		}
		private function read_NamedShows(pptRecord : Record):void{
		}
		private function read_NamedShow(pptRecord : Record):void{
		}
		private function read_NamedShowSlides(pptRecord : Record):void{
		}
		private function read_OriginalMainMasterId(pptRecord : Record):void{
		}
		private function read_CompositeMasterId(pptRecord : Record):void{
		}
		private function read_RoundTripContentMasterInfo12(pptRecord : Record):void{
		}
		private function read_RoundTripShapeId12(pptRecord : Record):void{
		}
		private function read_RoundTripHFPlaceholder12(pptRecord : Record):void{
		}
		private function read_RoundTripContentMasterId12(pptRecord : Record):void{
		}
		private function read_RoundTripOArtTextStyles12(pptRecord : Record):void{
		}
		private function read_HeaderFooterDefaults12(pptRecord : Record):void{
		}
		private function read_DocFlags12(pptRecord : Record):void{
		}
		private function read_RoundTripShapeCheckSumForCustomLayouts12(pptRecord : Record):void{
		}
		private function read_RoundTripNotesMasterTextStyles12(pptRecord : Record):void{
		}
		private function read_RoundTripCustomTableStyles12(pptRecord : Record):void{
		}
		private function read_List(pptRecord : Record):void{
		}
		private function read_FontCollection(pptRecord : Record):void{
		}
		private function read_ListPlaceholder(pptRecord : Record):void{
		}
		private function read_BookmarkCollection(pptRecord : Record):void{
		}
		private function read_SoundCollection(pptRecord : Record):void{
		}
		private function read_SoundCollAtom(pptRecord : Record):void{
		}
		private function read_Sound(pptRecord : Record):void{
		}
		private function read_SoundData(pptRecord : Record):void{
		}
		private function read_BookmarkSeedAtom(pptRecord : Record):void{
		}
		private function read_GuideList(pptRecord : Record):void{
		}
		private function read_RunArray(pptRecord : Record):void{
		}
		private function read_RunArrayAtom(pptRecord : Record):void{
		}
		private function read_ArrayElementAtom(pptRecord : Record):void{
		}
		private function read_Int4ArrayAtom(pptRecord : Record):void{
		}
		private function read_ColorSchemeAtom(pptRecord : Record):void{
		}
		private function read_OEShape(pptRecord : Record):void{
		}
		private function read_ExObjRefAtom(pptRecord : Record):void{
		}
		private function read_OEPlaceholderAtom(pptRecord : Record):void{
		}
		private function read_GrColor(pptRecord : Record):void{
		}
		private function read_GrectAtom(pptRecord : Record):void{
		}
		private function read_GratioAtom(pptRecord : Record):void{
		}
		private function read_Gscaling(pptRecord : Record):void{
		}
		private function read_GpointAtom(pptRecord : Record):void{
		}
		private function read_OEShapeAtom(pptRecord : Record):void{
		}
		private function read_OEPlaceholderNewPlaceholderId12(pptRecord : Record):void{
		}
		private function read_OutlineTextRefAtom(pptRecord : Record):void{
		}
		private function read_TextHeaderAtom(pptRecord : Record):void{
		}
		private function read_TextCharsAtom(pptRecord : Record):void{
		}
		private function read_StyleTextPropAtom(pptRecord : Record):void{
		}
		private function read_BaseTextPropAtom(pptRecord : Record):void{
		}
		private function read_TxMasterStyleAtom(pptRecord : Record):void{
		}
		private function read_TxCFStyleAtom(pptRecord : Record):void{
		}
		private function read_TxPFStyleAtom(pptRecord : Record):void{
		}
		private function read_TextRulerAtom(pptRecord : Record):void{
		}
		private function read_TextBookmarkAtom(pptRecord : Record):void{
		}
		private function read_TextBytesAtom(pptRecord : Record):void{
		}
		private function read_TxSIStyleAtom(pptRecord : Record):void{
		}
		private function read_TextSpecInfoAtom(pptRecord : Record):void{
		}
		private function read_DefaultRulerAtom(pptRecord : Record):void{
		}
		private function read_FontEntityAtom(pptRecord : Record):void{
		}
		private function read_FontEmbedData(pptRecord : Record):void{
		}
		private function read_TypeFace(pptRecord : Record):void{
		}
		private function read_CString(pptRecord : Record):void{
		}
		private function read_ExternalObject(pptRecord : Record):void{
		}
		private function read_MetaFile(pptRecord : Record):void{
		}
		private function read_ExOleObj(pptRecord : Record):void{
		}
		private function read_ExOleObjAtom(pptRecord : Record):void{
		}
		private function read_ExPlainLinkAtom(pptRecord : Record):void{
		}
		private function read_CorePict(pptRecord : Record):void{
		}
		private function read_CorePictAtom(pptRecord : Record):void{
		}
		private function read_ExPlainAtom(pptRecord : Record):void{
		}
		private function read_SrKinsoku(pptRecord : Record):void{
		}
		private function read_Handout(pptRecord : Record):void{
		}
		private function read_ExEmbed(pptRecord : Record):void{
		}
		private function read_ExEmbedAtom(pptRecord : Record):void{
		}
		private function read_ExLink(pptRecord : Record):void{
		}
		private function read_ExLinkAtom_old(pptRecord : Record):void{
		}
		private function read_BookmarkEntityAtom(pptRecord : Record):void{
		}
		private function read_ExLinkAtom(pptRecord : Record):void{
		}
		private function read_SrKinsokuAtom(pptRecord : Record):void{
		}
		private function read_ExHyperlinkAtom(pptRecord : Record):void{
		}
		private function read_ExPlain(pptRecord : Record):void{
		}
		private function read_ExPlainLink(pptRecord : Record):void{
		}
		private function read_ExHyperlink(pptRecord : Record):void{
		}
		private function read_SlideNumberMCAtom(pptRecord : Record):void{
		}
		private function read_HeadersFooters(pptRecord : Record):void{
		}
		private function read_HeadersFootersAtom(pptRecord : Record):void{
		}
		private function read_RecolorEntryAtom(pptRecord : Record):void{
		}
		private function read_TxInteractiveInfoAtom(pptRecord : Record):void{
		}
		private function read_EmFormatAtom(pptRecord : Record):void{
		}
		private function read_CharFormatAtom(pptRecord : Record):void{
		}
		private function read_ParaFormatAtom(pptRecord : Record):void{
		}
		private function read_MasterText(pptRecord : Record):void{
		}
		private function read_RecolorInfoAtom(pptRecord : Record):void{
		}
		private function read_ExQuickTime(pptRecord : Record):void{
		}
		private function read_ExQuickTimeMovie(pptRecord : Record):void{
		}
		private function read_ExQuickTimeMovieData(pptRecord : Record):void{
		}
		private function read_ExSubscription(pptRecord : Record):void{
		}
		private function read_ExSubscriptionSection(pptRecord : Record):void{
		}
		private function read_ExControl(pptRecord : Record):void{
		}
		private function read_ExControlAtom(pptRecord : Record):void{
		}
		private function read_SlideListWithText(pptRecord : Record):void{
			
		}
		private function read_AnimationInfoAtom(pptRecord : Record):void{
		}
		private function read_InteractiveInfo(pptRecord : Record):void{
		}
		private function read_InteractiveInfoAtom(pptRecord : Record):void{
		}
		private function read_SlideList(pptRecord : Record):void{
		}
		private function read_UserEditAtom(pptRecord : Record):void{
		}
		private function read_CurrentUserAtom(pptRecord : Record):void{
		}
		private function read_DateTimeMCAtom(pptRecord : Record):void{
		}
		private function read_GenericDateMCAtom(pptRecord : Record):void{
		}
		private function read_HeaderMCAtom(pptRecord : Record):void{
		}
		private function read_FooterMCAtom(pptRecord : Record):void{
		}
		private function read_ExMediaAtom(pptRecord : Record):void{
		}
		private function read_ExVideo(pptRecord : Record):void{
		}
		private function read_ExAviMovie(pptRecord : Record):void{
		}
		private function read_ExMCIMovie(pptRecord : Record):void{
		}
		private function read_ExMIDIAudio(pptRecord : Record):void{
		}
		private function read_ExCDAudio(pptRecord : Record):void{
		}
		private function read_ExWAVAudioEmbedded(pptRecord : Record):void{
		}
		private function read_ExWAVAudioLink(pptRecord : Record):void{
		}
		private function read_ExOleObjStg(pptRecord : Record):void{
		}
		private function read_ExCDAudioAtom(pptRecord : Record):void{
		}
		private function read_ExWAVAudioEmbeddedAtom(pptRecord : Record):void{
		}
		private function read_AnimationInfo(pptRecord : Record):void{
		}
		private function read_RTFDateTimeMCAtom(pptRecord : Record):void{
		}
		private function read_ProgTags(pptRecord : Record):void{
		}
		private function read_ProgStringTag(pptRecord : Record):void{
		}
		private function read_ProgBinaryTag(pptRecord : Record):void{
		}
		private function read_BinaryTagData(pptRecord : Record):void{
		}
		private function read_PrintOptions(pptRecord : Record):void{
		}
		private function read_PersistPtrFullBlock(pptRecord : Record):void{
		}
		private function read_PersistPtrIncrementalBlock(pptRecord : Record):void{
		}
		private function read_RulerIndentAtom(pptRecord : Record):void{
		}
		private function read_GscalingAtom(pptRecord : Record):void{
		}
		private function read_GrColorAtom(pptRecord : Record):void{
		}
		private function read_GLPointAtom(pptRecord : Record):void{
		}
		private function read_GlineAtom(pptRecord : Record):void{
		}
		private function read_AnimationAtom12(pptRecord : Record):void{
		}
		private function read_AnimationHashAtom12(pptRecord : Record):void{
		}
		private function read_SlideSyncInfo12(pptRecord : Record):void{
		}
		private function read_SlideSyncInfoAtom12(pptRecord : Record):void{
		}
		
		
		
		//************************************************************************************************************/		
		//	Write
		//	 
		//************************************************************************************************************/
		
		private function write_Unknown(pptRecord : Record):void{
		} 
		private function write_SubContainerCompleted(pptRecord : Record):void{
		}
		private function write_IRRAtom(pptRecord : Record):void{
		}
		private function write_PSS(pptRecord : Record):void{
		}
		private function write_SubContainerException(pptRecord : Record):void{
		}
		private function write_ClientSignal1(pptRecord : Record):void{
		}
		private function write_ClientSignal2(pptRecord : Record):void{
		}
		private function write_PowerPointStateInfoAtom(pptRecord : Record):void{
		}
		private function write_Document(pptRecord : Record):void{
		}
		private function write_DocumentAtom(pptRecord : Record):void{
		}
		private function write_EndDocument(pptRecord : Record):void{
		}
		private function write_SlidePersist(pptRecord : Record):void{
		}
		private function write_SlideBase(pptRecord : Record):void{
		}
		private function write_SlideBaseAtom(pptRecord : Record):void{
		}
		private function write_Slide(pptRecord : Record):void{
		}
		private function write_SlideAtom(pptRecord : Record):void{
		}
		private function write_Notes(pptRecord : Record):void{
		}
		private function write_NotesAtom(pptRecord : Record):void{
		}
		private function write_Environment(pptRecord : Record):void{
		}
		private function write_SlidePersistAtom(pptRecord : Record):void{
		}
		private function write_Scheme(pptRecord : Record):void{
		}
		private function write_SchemeAtom(pptRecord : Record):void{
		}
		private function write_DocViewInfo(pptRecord : Record):void{
		}
		private function write_SslideLayoutAtom(pptRecord : Record):void{
		}
		private function write_MainMaster(pptRecord : Record):void{
		}
		private function write_SSSlideInfoAtom(pptRecord : Record):void{
		}
		private function write_SlideViewInfo(pptRecord : Record):void{
		}
		private function write_GuideAtom(pptRecord : Record):void{
		}
		private function write_ViewInfo(pptRecord : Record):void{
		}
		private function write_ViewInfoAtom(pptRecord : Record):void{
		}
		private function write_SlideViewInfoAtom(pptRecord : Record):void{
		}
		private function write_VBAInfo(pptRecord : Record):void{
		}
		private function write_VBAInfoAtom(pptRecord : Record):void{
		}
		private function write_SSDocInfoAtom(pptRecord : Record):void{
		}
		private function write_Summary(pptRecord : Record):void{
		}
		private function write_Texture(pptRecord : Record):void{
		}
		private function write_VBASlideInfo(pptRecord : Record):void{
		}
		private function write_VBASlideInfoAtom(pptRecord : Record):void{
		}
		private function write_DocRoutingSlip(pptRecord : Record):void{
		}
		private function write_OutlineViewInfo(pptRecord : Record):void{
		}
		private function write_SorterViewInfo(pptRecord : Record):void{
		}
		private function write_ExObjList(pptRecord : Record):void{
		}
		private function write_ExObjListAtom(pptRecord : Record):void{
		}
		private function write_PPDrawingGroup(pptRecord : Record):void{
		}
		private function write_PPDrawing(pptRecord : Record):void{
		}
		private function write_Theme(pptRecord : Record):void{
		}
		private function write_ColorMapping(pptRecord : Record):void{
		}
		private function write_NamedShows(pptRecord : Record):void{
		}
		private function write_NamedShow(pptRecord : Record):void{
		}
		private function write_NamedShowSlides(pptRecord : Record):void{
		}
		private function write_OriginalMainMasterId(pptRecord : Record):void{
		}
		private function write_CompositeMasterId(pptRecord : Record):void{
		}
		private function write_RoundTripContentMasterInfo12(pptRecord : Record):void{
		}
		private function write_RoundTripShapeId12(pptRecord : Record):void{
		}
		private function write_RoundTripHFPlaceholder12(pptRecord : Record):void{
		}
		private function write_RoundTripContentMasterId12(pptRecord : Record):void{
		}
		private function write_RoundTripOArtTextStyles12(pptRecord : Record):void{
		}
		private function write_HeaderFooterDefaults12(pptRecord : Record):void{
		}
		private function write_DocFlags12(pptRecord : Record):void{
		}
		private function write_RoundTripShapeCheckSumForCustomLayouts12(pptRecord : Record):void{
		}
		private function write_RoundTripNotesMasterTextStyles12(pptRecord : Record):void{
		}
		private function write_RoundTripCustomTableStyles12(pptRecord : Record):void{
		}
		private function write_List(pptRecord : Record):void{
		}
		private function write_FontCollection(pptRecord : Record):void{
		}
		private function write_ListPlaceholder(pptRecord : Record):void{
		}
		private function write_BookmarkCollection(pptRecord : Record):void{
		}
		private function write_SoundCollection(pptRecord : Record):void{
		}
		private function write_SoundCollAtom(pptRecord : Record):void{
		}
		private function write_Sound(pptRecord : Record):void{
		}
		private function write_SoundData(pptRecord : Record):void{
		}
		private function write_BookmarkSeedAtom(pptRecord : Record):void{
		}
		private function write_GuideList(pptRecord : Record):void{
		}
		private function write_RunArray(pptRecord : Record):void{
		}
		private function write_RunArrayAtom(pptRecord : Record):void{
		}
		private function write_ArrayElementAtom(pptRecord : Record):void{
		}
		private function write_Int4ArrayAtom(pptRecord : Record):void{
		}
		private function write_ColorSchemeAtom(pptRecord : Record):void{
		}
		private function write_OEShape(pptRecord : Record):void{
		}
		private function write_ExObjRefAtom(pptRecord : Record):void{
		}
		private function write_OEPlaceholderAtom(pptRecord : Record):void{
		}
		private function write_GrColor(pptRecord : Record):void{
		}
		private function write_GrectAtom(pptRecord : Record):void{
		}
		private function write_GratioAtom(pptRecord : Record):void{
		}
		private function write_Gscaling(pptRecord : Record):void{
		}
		private function write_GpointAtom(pptRecord : Record):void{
		}
		private function write_OEShapeAtom(pptRecord : Record):void{
		}
		private function write_OEPlaceholderNewPlaceholderId12(pptRecord : Record):void{
		}
		private function write_OutlineTextRefAtom(pptRecord : Record):void{
		}
		private function write_TextHeaderAtom(pptRecord : Record):void{
		}
		private function write_TextCharsAtom(pptRecord : Record):void{
		}
		private function write_StyleTextPropAtom(pptRecord : Record):void{
		}
		private function write_BaseTextPropAtom(pptRecord : Record):void{
		}
		private function write_TxMasterStyleAtom(pptRecord : Record):void{
		}
		private function write_TxCFStyleAtom(pptRecord : Record):void{
		}
		private function write_TxPFStyleAtom(pptRecord : Record):void{
		}
		private function write_TextRulerAtom(pptRecord : Record):void{
		}
		private function write_TextBookmarkAtom(pptRecord : Record):void{
		}
		private function write_TextBytesAtom(pptRecord : Record):void{
		}
		private function write_TxSIStyleAtom(pptRecord : Record):void{
		}
		private function write_TextSpecInfoAtom(pptRecord : Record):void{
		}
		private function write_DefaultRulerAtom(pptRecord : Record):void{
		}
		private function write_FontEntityAtom(pptRecord : Record):void{
		}
		private function write_FontEmbedData(pptRecord : Record):void{
		}
		private function write_TypeFace(pptRecord : Record):void{
		}
		private function write_CString(pptRecord : Record):void{
		}
		private function write_ExternalObject(pptRecord : Record):void{
		}
		private function write_MetaFile(pptRecord : Record):void{
		}
		private function write_ExOleObj(pptRecord : Record):void{
		}
		private function write_ExOleObjAtom(pptRecord : Record):void{
		}
		private function write_ExPlainLinkAtom(pptRecord : Record):void{
		}
		private function write_CorePict(pptRecord : Record):void{
		}
		private function write_CorePictAtom(pptRecord : Record):void{
		}
		private function write_ExPlainAtom(pptRecord : Record):void{
		}
		private function write_SrKinsoku(pptRecord : Record):void{
		}
		private function write_Handout(pptRecord : Record):void{
		}
		private function write_ExEmbed(pptRecord : Record):void{
		}
		private function write_ExEmbedAtom(pptRecord : Record):void{
		}
		private function write_ExLink(pptRecord : Record):void{
		}
		private function write_ExLinkAtom_old(pptRecord : Record):void{
		}
		private function write_BookmarkEntityAtom(pptRecord : Record):void{
		}
		private function write_ExLinkAtom(pptRecord : Record):void{
		}
		private function write_SrKinsokuAtom(pptRecord : Record):void{
		}
		private function write_ExHyperlinkAtom(pptRecord : Record):void{
		}
		private function write_ExPlain(pptRecord : Record):void{
		}
		private function write_ExPlainLink(pptRecord : Record):void{
		}
		private function write_ExHyperlink(pptRecord : Record):void{
		}
		private function write_SlideNumberMCAtom(pptRecord : Record):void{
		}
		private function write_HeadersFooters(pptRecord : Record):void{
		}
		private function write_HeadersFootersAtom(pptRecord : Record):void{
		}
		private function write_RecolorEntryAtom(pptRecord : Record):void{
		}
		private function write_TxInteractiveInfoAtom(pptRecord : Record):void{
		}
		private function write_EmFormatAtom(pptRecord : Record):void{
		}
		private function write_CharFormatAtom(pptRecord : Record):void{
		}
		private function write_ParaFormatAtom(pptRecord : Record):void{
		}
		private function write_MasterText(pptRecord : Record):void{
		}
		private function write_RecolorInfoAtom(pptRecord : Record):void{
		}
		private function write_ExQuickTime(pptRecord : Record):void{
		}
		private function write_ExQuickTimeMovie(pptRecord : Record):void{
		}
		private function write_ExQuickTimeMovieData(pptRecord : Record):void{
		}
		private function write_ExSubscription(pptRecord : Record):void{
		}
		private function write_ExSubscriptionSection(pptRecord : Record):void{
		}
		private function write_ExControl(pptRecord : Record):void{
		}
		private function write_ExControlAtom(pptRecord : Record):void{
		}
		private function write_SlideListWithText(pptRecord : Record):void{
			
		}
		private function write_AnimationInfoAtom(pptRecord : Record):void{
		}
		private function write_InteractiveInfo(pptRecord : Record):void{
		}
		private function write_InteractiveInfoAtom(pptRecord : Record):void{
		}
		private function write_SlideList(pptRecord : Record):void{
		}
		private function write_UserEditAtom(pptRecord : Record):void{
		}
		private function write_CurrentUserAtom(pptRecord : Record):void{
		}
		private function write_DateTimeMCAtom(pptRecord : Record):void{
		}
		private function write_GenericDateMCAtom(pptRecord : Record):void{
		}
		private function write_HeaderMCAtom(pptRecord : Record):void{
		}
		private function write_FooterMCAtom(pptRecord : Record):void{
		}
		private function write_ExMediaAtom(pptRecord : Record):void{
		}
		private function write_ExVideo(pptRecord : Record):void{
		}
		private function write_ExAviMovie(pptRecord : Record):void{
		}
		private function write_ExMCIMovie(pptRecord : Record):void{
		}
		private function write_ExMIDIAudio(pptRecord : Record):void{
		}
		private function write_ExCDAudio(pptRecord : Record):void{
		}
		private function write_ExWAVAudioEmbedded(pptRecord : Record):void{
		}
		private function write_ExWAVAudioLink(pptRecord : Record):void{
		}
		private function write_ExOleObjStg(pptRecord : Record):void{
		}
		private function write_ExCDAudioAtom(pptRecord : Record):void{
		}
		private function write_ExWAVAudioEmbeddedAtom(pptRecord : Record):void{
		}
		private function write_AnimationInfo(pptRecord : Record):void{
		}
		private function write_RTFDateTimeMCAtom(pptRecord : Record):void{
		}
		private function write_ProgTags(pptRecord : Record):void{
		}
		private function write_ProgStringTag(pptRecord : Record):void{
		}
		private function write_ProgBinaryTag(pptRecord : Record):void{
		}
		private function write_BinaryTagData(pptRecord : Record):void{
		}
		private function write_PrintOptions(pptRecord : Record):void{
		}
		private function write_PersistPtrFullBlock(pptRecord : Record):void{
		}
		private function write_PersistPtrIncrementalBlock(pptRecord : Record):void{
		}
		private function write_RulerIndentAtom(pptRecord : Record):void{
		}
		private function write_GscalingAtom(pptRecord : Record):void{
		}
		private function write_GrColorAtom(pptRecord : Record):void{
		}
		private function write_GLPointAtom(pptRecord : Record):void{
		}
		private function write_GlineAtom(pptRecord : Record):void{
		}
		private function write_AnimationAtom12(pptRecord : Record):void{
		}
		private function write_AnimationHashAtom12(pptRecord : Record):void{
		}
		private function write_SlideSyncInfo12(pptRecord : Record):void{
		}
		private function write_SlideSyncInfoAtom12(pptRecord : Record):void{
		}
		
	}
}