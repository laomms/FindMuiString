Imports System.Runtime.InteropServices

Module Module1

	Public Enum LpType
		RT_CURSOR = 1 '由硬件支持的光标资源
		RT_BITMAP = 2 '位图资源
		RT_ICON = 3 '由硬件支持的图标资源
		RT_MENU = 4 '菜单资源
		RT_DIALOG = 5 '对话框
		RT_STRING = 6 '字符表入口
		RT_FONTDIR = 7 '字体目录资源
		RT_FONT = 8 '字体资源
		RT_ACCELERATOR = 9 '加速器表
		RT_MESSAGETABLE = 11 '消息表的入口
		RT_GROUP_CURSOR = 12 '与硬件无关的光标资源
		RT_GROUP_ICON = 14 ' 与硬件无关的目标资源
		RT_VERSION = 16 '版本资源
		RT_PLUGPLAY = 19 '即插即用资源
		RT_VXD = 20 'VXD
		RT_ANICURSOR = 21 '动态光标
		RT_ANIICON = 22 '动态图标
		RT_HTML = 23 'HTML文档
		RT_RCDATA = 10 '原始数据或自定义资源
	End Enum
	'#Disable Warning BC0649
	<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto, Pack:=2)>
	Public Structure ControlData
		Public Style As Integer
		Public ExStyle As Integer
		Public x As Short
		Public y As Short
		Public cx As Short
		Public cy As Short
		Public id As Short
	End Structure
	<StructLayout(LayoutKind.Sequential, Pack:=2)>
	Friend Structure ControlDataEx
		Public helpId As Integer
		Public exStyle As Integer
		Public style As Integer
		Public x As Short
		Public y As Short
		Public cx As Short
		Public cy As Short
		Public id As Short
	End Structure
	<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto, Pack:=2)>
	Public Structure DialogBoxHeader
		Public Style As Integer
		Public ExStyle As Integer
		Public cdit As UShort 'control amount
		Public x As Short
		Public y As Short
		Public cx As Short
		Public cy As Short
		'<MarshalAs(UnmanagedType.ByValArray, SizeConst:=2)>
		'Public menuName() As Byte
	End Structure
	<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto, Pack:=2)>
	Public Structure DialogFont
		Public wPointSize As Short
		'<MarshalAs(UnmanagedType.ByValArray, SizeConst:=2)>
		'Public FontName() As Byte
	End Structure
	<StructLayout(LayoutKind.Sequential, Pack:=2)>
	Public Structure DialogFontEx
		Public wPointSize As Short
		Public Weight As Short
		Public Italic As Byte
		Public CharSet As Byte
		'<MarshalAs(UnmanagedType.ByValArray, SizeConst:=2)>
		'Public FontName() As Byte
	End Structure
	<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto, Pack:=2)>
	Public Structure DialogBoxHeaderEx
		Public SignEx As Integer '0xFFFF0001
		Public Version As Integer
		Public Style As Integer
		Public ExStyle As Integer
		Public DlgItems As Short 'control amount
		Public x As Short
		Public y As Short
		Public cx As Short
		Public cy As Short
		'<MarshalAs(UnmanagedType.ByValArray, SizeConst:=2)>
		'Public menuName() As Byte
	End Structure
	'<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
	'Public Structure DialogBoxHeader
	'	Public style As UInteger
	'	Public dwExtendedStyle As UInteger
	'	Public cdit As UShort
	'	Public x As Short
	'	Public y As Short
	'	Public cx As Short
	'	Public cy As Short
	'End Structure
	'<StructLayout(LayoutKind.Sequential, Pack:=2)>
	'Friend Structure DlgItemTemplate
	'	Public style As UInteger
	'	Public dwExtendedStyle As UInteger
	'	Public x As Short
	'	Public y As Short
	'	Public cx As Short
	'	Public cy As Short
	'	Public id As UShort
	'	Public exWindowClass As Object
	'	Public exTitle As Object
	'	Public exCreationData() As Byte
	'End Structure

	'<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
	'Public Structure DialogBoxHeaderEx
	'	Public dlgVer As UShort
	'	Public signature As UShort
	'	Public helpID As UInteger
	'	Public exStyle As UInteger
	'	Public style As UInteger
	'	Public cDlgItems As UShort
	'	Public x As Short
	'	Public y As Short
	'	Public cx As Short
	'	Public cy As Short
	'	Public menu As Object
	'	Public windowClass As Object
	'	Public title As String
	'	Public pointsize As UShort
	'	Public weight As UShort
	'	Public italic As Byte
	'	Public charset As Byte
	'	<MarshalAs(UnmanagedType.ByValArray, SizeConst:=2)>
	'	Public typeface() As Byte
	'End Structure
	'<StructLayout(LayoutKind.Sequential, Pack:=2)>
	'Friend Structure DlgItemTemplateEx
	'	Public helpId As UInteger
	'	Public exStyle As UInteger
	'	Public style As UInteger
	'	Public x As Short
	'	Public y As Short
	'	Public cx As Short
	'	Public cy As Short
	'	Public id As UInteger
	'	Public windowClass As Object
	'	Public title As Object
	'	Public extraCount As UShort
	'	Public creationData() As Byte
	'End Structure


	'#Enable Warning BC0649
	Public Enum DialogBoxStyles
		DS_SETFOREGROUND = &H200
		DS_NOFAILCREATE = &H10
		DS_CONTEXTHELP = &H2000
		DS_CENTERMOUSE = &H1000
		DS_MODALFRAME = &H80
		DS_S_SUCCESS = 0
		DS_SHELLFONT = DS_SETFONT Or DS_FIXEDSYS
		DS_NOIDLEMSG = &H100
		DS_LOCALEDIT = &H20
		DS_SYSMODAL = &H2
		DS_FIXEDSYS = &H8
		DS_ABSALIGN = &H1
		DS_SETFONT = &H40
		DS_CONTROL = &H400
		DS_CENTER = &H800
		DS_3DLOOK = &H4
	End Enum
	Public Enum WindowExStyles
		WS_EX_DLGMODALFRAME = &H1
		WS_EX_NOPARENTNOTIFY = &H4
		WS_EX_TOPMOST = &H8
		WS_EX_ACCEPTFILES = &H10
		WS_EX_TRANSPARENT = &H20
		WS_EX_MDICHILD = &H40
		WS_EX_TOOLWINDOW = &H80
		WS_EX_WINDOWEDGE = &H100
		WS_EX_CLIENTEDGE = &H200
		WS_EX_CONTEXTHELP = &H400
		WS_EX_RIGHT = &H1000
		WS_EX_LEFT = &H0
		WS_EX_RTLREADING = &H2000
		WS_EX_LTRREADING = &H0
		WS_EX_LEFTSCROLLBAR = &H4000
		WS_EX_RIGHTSCROLLBAR = &H0
		WS_EX_CONTROLPARENT = &H10000
		WS_EX_STATICEDGE = &H20000
		WS_EX_APPWINDOW = &H40000
		WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
		WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
		WS_EX_LAYERED = &H80000
		WS_EX_NOREDIRECTIONBITMAP = &H200000
		WS_EX_NOINHERITLAYOUT = &H100000
		WS_EX_NOACTIVATE = &H8000000
		WS_EX_LAYOUTRTL = &H400000
		WS_EX_COMPOSITED = &H2000000

	End Enum
	Public Enum WindowStyles As UInteger
		WS_OVERLAPPED = &H0
		WS_POPUP = &H80000000UI
		WS_CHILD = &H40000000
		WS_MINIMIZE = &H20000000
		WS_VISIBLE = &H10000000
		WS_DISABLED = &H8000000
		WS_CLIPSIBLINGS = &H4000000
		WS_CLIPCHILDREN = &H2000000
		WS_MAXIMIZE = &H1000000
		WS_CAPTION = &HC00000
		WS_BORDER = &H800000
		WS_DLGFRAME = &H400000
		WS_VSCROLL = &H200000
		WS_HSCROLL = &H100000
		WS_SYSMENU = &H80000
		WS_THICKFRAME = &H40000
		WS_GROUP = &H20000
		WS_TABSTOP = &H10000
		WS_MINIMIZEBOX = &H20000
		WS_MAXIMIZEBOX = &H10000
		WS_TILED = &H0
		WS_ICONIC = &H20000000
		WS_SIZEBOX = &H40000
		WS_TILEDWINDOW = &HCF0000
		WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
		WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
		WS_CHILDWINDOW = (WS_CHILD)

		'		
		'		 * Edit Control Styles
		'		 
		ES_LEFT = &H0
		ES_CENTER = &H1
		ES_RIGHT = &H2
		ES_MULTILINE = &H4
		ES_UPPERCASE = &H8
		ES_LOWERCASE = &H10
		ES_PASSWORD = &H20
		ES_AUTOVSCROLL = &H40
		ES_AUTOHSCROLL = &H80
		ES_NOHIDESEL = &H100
		ES_OEMCONVERT = &H400
		ES_READONLY = &H800
		ES_WANTRETURN = &H1000
		ES_NUMBER = &H2000


		'		
		'		 * Button Control Styles
		'		 
		BS_PUSHBUTTON = &H0
		BS_DEFPUSHBUTTON = &H1
		BS_CHECKBOX = &H2
		BS_AUTOCHECKBOX = &H3
		BS_RADIOBUTTON = &H4
		BS_3STATE = &H5
		BS_AUTO3STATE = &H6
		BS_GROUPBOX = &H7
		BS_USERBUTTON = &H8
		BS_AUTORADIOBUTTON = &H9
		BS_PUSHBOX = &HA
		BS_OWNERDRAW = &HB
		BS_TYPEMASK = &HF
		BS_LEFTTEXT = &H20
		BS_TEXT = &H0
		BS_ICON = &H40
		BS_BITMAP = &H80
		BS_LEFT = &H100
		BS_RIGHT = &H200
		BS_CENTER = &H300
		BS_TOP = &H400
		BS_BOTTOM = &H800
		BS_VCENTER = &HC00
		BS_PUSHLIKE = &H1000
		BS_MULTILINE = &H2000
		BS_NOTIFY = &H4000
		BS_FLAT = &H8000
		BS_RIGHTBUTTON = BS_LEFTTEXT


		'		
		'		 * Static Control Constants
		'		 
		SS_LEFT = &H0
		SS_CENTER = &H1
		SS_RIGHT = &H2
		SS_ICON = &H3
		SS_BLACKRECT = &H4
		SS_GRAYRECT = &H5
		SS_WHITERECT = &H6
		SS_BLACKFRAME = &H7
		SS_GRAYFRAME = &H8
		SS_WHITEFRAME = &H9
		SS_USERITEM = &HA
		SS_SIMPLE = &HB
		SS_LEFTNOWORDWRAP = &HC
		SS_OWNERDRAW = &HD
		SS_BITMAP = &HE
		SS_ENHMETAFILE = &HF
		SS_ETCHEDHORZ = &H10
		SS_ETCHEDVERT = &H11
		SS_ETCHEDFRAME = &H12
		SS_TYPEMASK = &H1F
		SS_REALSIZECONTROL = &H40
		SS_NOPREFIX = &H80
		SS_NOTIFY = &H100
		SS_CENTERIMAGE = &H200
		SS_RIGHTJUST = &H400
		SS_REALSIZEIMAGE = &H800
		SS_SUNKEN = &H1000
		SS_EDITCONTROL = &H2000
		SS_ENDELLIPSIS = &H4000
		SS_PATHELLIPSIS = &H8000
		SS_WORDELLIPSIS = &HC000
		SS_ELLIPSISMASK = &HC000


		' Dialog Styles
		DS_ABSALIGN = &H1
		DS_SYSMODAL = &H2
		DS_LOCALEDIT = &H20
		DS_SETFONT = &H40
		DS_MODALFRAME = &H80
		DS_NOIDLEMSG = &H100
		DS_SETFOREGROUND = &H200
		DS_3DLOOK = &H4
		DS_FIXEDSYS = &H8
		DS_NOFAILCREATE = &H10
		DS_CONTROL = &H400
		DS_CENTER = &H800
		DS_CENTERMOUSE = &H1000
		DS_CONTEXTHELP = &H2000
		DS_SHELLFONT = (DS_SETFONT Or DS_FIXEDSYS)

		' Listbox Styles
		LBS_NOTIFY = &H1
		LBS_SORT = &H2
		LBS_NOREDRAW = &H4
		LBS_MULTIPLESEL = &H8
		LBS_OWNERDRAWFIXED = &H10
		LBS_OWNERDRAWVARIABLE = &H20
		LBS_HASSTRINGS = &H40
		LBS_USETABSTOPS = &H80
		LBS_NOINTEGRALHEIGHT = &H100
		LBS_MULTICOLUMN = &H200
		LBS_WANTKEYBOARDINPUT = &H400
		LBS_EXTENDEDSEL = &H800
		LBS_DISABLENOSCROLL = &H1000
		LBS_NODATA = &H2000
		LBS_NOSEL = &H4000
		LBS_COMBOBOX = &H8000
		LBS_STANDARD = (LBS_NOTIFY Or LBS_SORT Or WS_VSCROLL Or WS_BORDER)

		' combobox styles
		CBS_SIMPLE = &H1
		CBS_DROPDOWN = &H2
		CBS_DROPDOWNLIST = &H3
		CBS_OWNERDRAWFIXED = &H10
		CBS_OWNERDRAWVARIABLE = &H20
		CBS_AUTOHSCROLL = &H40
		CBS_OEMCONVERT = &H80
		CBS_SORT = &H100
		CBS_HASSTRINGS = &H200
		CBS_NOINTEGRALHEIGHT = &H400
		CBS_DISABLENOSCROLL = &H800
		CBS_UPPERCASE = &H2000
		CBS_LOWERCASE = &H4000


		' Scroll Bar Styles
		SBS_HORZ = &H0
		SBS_VERT = &H1
		SBS_TOPALIGN = &H2
		SBS_LEFTALIGN = &H2
		SBS_BOTTOMALIGN = &H4
		SBS_RIGHTALIGN = &H4
		SBS_SIZEBOXTOPLEFTALIGN = &H2
		SBS_SIZEBOXBOTTOMRIGHTALIGN = &H4
		SBS_SIZEBOX = &H8
		SBS_SIZEGRIP = &H10

		'treeview styles

		TVS_HASBUTTONS = &H1
		TVS_HASLINES = &H2
		TVS_LINESATROOT = &H4
		TVS_EDITLABELS = &H8
		TVS_DISABLEDRAGDROP = &H10
		TVS_SHOWSELALWAYS = &H20
		TVS_RTLREADING = &H40
		TVS_NOTOOLTIPS = &H80
		TVS_CHECKBOXES = &H100
		TVS_TRACKSELECT = &H200
		TVS_SINGLEEXPAND = &H400
		TVS_INFOTIP = &H800
		TVS_FULLROWSELECT = &H1000
		TVS_NOSCROLL = &H2000
		TVS_NONEVENHEIGHT = &H4000
		TVS_NOHSCROLL = &H8000


		'tab control styles
		TCS_SCROLLOPPOSITE = &H1
		TCS_BOTTOM = &H2
		TCS_RIGHT = &H2
		TCS_MULTISELECT = &H4
		TCS_FLATBUTTONS = &H8
		TCS_FORCEICONLEFT = &H10
		TCS_FORCELABELLEFT = &H20
		TCS_HOTTRACK = &H40
		TCS_VERTICAL = &H80
		TCS_TABS = &H0
		TCS_BUTTONS = &H100
		TCS_SINGLELINE = &H0
		TCS_MULTILINE = &H200
		TCS_RIGHTJUSTIFY = &H0
		TCS_FIXEDWIDTH = &H400
		TCS_RAGGEDRIGHT = &H800
		TCS_FOCUSONBUTTONDOWN = &H1000
		TCS_OWNERDRAWFIXED = &H2000
		TCS_TOOLTIPS = &H4000
		TCS_FOCUSNEVER = &H8000

		'spin styles
		UDS_WRAP = &H1
		UDS_SETBUDDYINT = &H2
		UDS_ALIGNRIGHT = &H4
		UDS_ALIGNLEFT = &H8
		UDS_AUTOBUDDY = &H10
		UDS_ARROWKEYS = &H20
		UDS_HORZ = &H40
		UDS_NOTHOUSANDS = &H80
		UDS_HOTTRACK = &H100

		'progress styles
		PBS_SMOOTH = &H1
		PBS_VERTICAL = &H4


		'list view styles
		LVS_ICON = &H0
		LVS_REPORT = &H1
		LVS_SMALLICON = &H2
		LVS_LIST = &H3
		LVS_TYPEMASK = &H3
		LVS_SINGLESEL = &H4
		LVS_SHOWSELALWAYS = &H8
		LVS_SORTASCENDING = &H10
		LVS_SORTDESCENDING = &H20
		LVS_SHAREIMAGELISTS = &H40
		LVS_NOLABELWRAP = &H80
		LVS_AUTOARRANGE = &H100
		LVS_EDITLABELS = &H200
		LVS_OWNERDATA = &H1000
		LVS_NOSCROLL = &H2000
		LVS_TYPESTYLEMASK = &HFC00
		LVS_ALIGNTOP = &H0
		LVS_ALIGNLEFT = &H800
		LVS_ALIGNMASK = &HC00
		LVS_OWNERDRAWFIXED = &H400
		LVS_NOCOLUMNHEADER = &H4000
		LVS_NOSORTHEADER = &H8000

		ACS_CENTER = &H1
		ACS_TRANSPARENT = &H2
		ACS_AUTOPLAY = &H4
		ACS_TIMER = &H8
	End Enum

	<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto, Pack:=2)>
	Public Class MenuHeader
		Public wVersion As Short
		Public cbHeaderSize As Short
	End Class
	<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto, Pack:=2)>
	Public Class MenuExHeader
		Public wVersion As Short
		Public cbHeaderSize As Short
		Public dwHelpId As Integer
	End Class
	<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto, Pack:=2)>
	Structure PopupMenuItem
		Public resInfo As UShort
		<MarshalAs(UnmanagedType.ByValArray, SizeConst:=2)>
		Public szItemText() As Byte  '菜单名
	End Structure
	<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto, Pack:=2)>
	Structure MenuItemHeader
		Public resInfo As UShort
		Public mtId As UShort
		'<MarshalAs(UnmanagedType.ByValArray, SizeConst:=2)>
		'Public szItemText() As Byte  '菜单名
	End Structure


	<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto, Pack:=2)>
	Friend Structure MenuExItem
		Public dwType As UInteger
		Public dwState As UInteger
		Public mtId As Integer
		Public resInfo As UShort
		'<MarshalAs(UnmanagedType.ByValArray, SizeConst:=2)>
		'Public szItemText() As Byte  '菜单名

	End Structure

	Public Enum MenuItemOptions As UShort
		MF_GRAYED = &H1
		MF_DISABLED = &H2
		MF_BITMAP = &H4
		MF_CHECKED = &H8
		MF_MENUBARBREAK = &H20
		MF_MENUBREAK = &H40
		MF_OWNERDRAW = &H100
		MF_HELP = &H4000
		MF_POPUP = &H10
		MF_END = &H80
	End Enum
	Enum MenuExItemInfo As UShort
		LastItem = &H80
		HasChildren = &H1
	End Enum
	Enum MenuExItemType As UInteger
		MF_BITMAP = &H4
		MF_MENUBARBREAK = &H20
		MF_MENUBREAK = &H40
		MF_OWNERDRAW = &H100
		MF_USECHECKBITMAPS = &H200
		MF_RIGHTJUSTIFY = &H4000
		MF_RIGHTORDER = &H2000
		MF_SEPARATOR = &H800
		MF_STRING = &H0
	End Enum
	Enum MenuExItemState As Integer
		MF_CHECKED = &H8
		MF_DEFAULT = &H1000
		MF_DISABLED = &H1
		MF_ENABLED = &H0
		MF_GRAYED = &H3
		MF_HILITE = &H80
		MF_UNCHECKED = &H0
		MF_UNHILITE = &H0
	End Enum

End Module
