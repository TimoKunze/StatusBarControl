// DispIDs.h: Defines a DISPID for each COM property and method we're providing.

// IStatusBar
// properties
#define DISPID_STATBAR_APPEARANCE													1
#define DISPID_STATBAR_BACKCOLOR													2
#define DISPID_STATBAR_BORDERSTYLE												3
#define DISPID_STATBAR_CUSTOMCAPSLOCKTEXT									4
#define DISPID_STATBAR_CUSTOMINSERTKEYTEXT								5
#define DISPID_STATBAR_CUSTOMKANALOCKTEXT									6
#define DISPID_STATBAR_CUSTOMNUMLOCKTEXT									7
#define DISPID_STATBAR_CUSTOMSCROLLLOCKTEXT								8
#define DISPID_STATBAR_DISABLEDEVENTS											9
#define DISPID_STATBAR_DONTREDRAW													10
#define DISPID_STATBAR_ENABLED														DISPID_ENABLED
#define DISPID_STATBAR_FONT																11
#define DISPID_STATBAR_FORCESIZEGRIPPERDISPLAY						12
#define DISPID_STATBAR_HOVERTIME													13
#define DISPID_STATBAR_HWND																14
#define DISPID_STATBAR_MINIMUMHEIGHT											15
#define DISPID_STATBAR_MOUSEICON													16
#define DISPID_STATBAR_MOUSEPOINTER												17
#define DISPID_STATBAR_PANELS															18
#define DISPID_STATBAR_PANELTOAUTOSIZE										19
#define DISPID_STATBAR_REGISTERFOROLEDRAGDROP							20
#define DISPID_STATBAR_RIGHTTOLEFTLAYOUT									21
#define DISPID_STATBAR_SHOWTOOLTIPS												22
#define DISPID_STATBAR_SIMPLEMODE													23
#define DISPID_STATBAR_SIMPLEPANEL												DISPID_VALUE
#define DISPID_STATBAR_SUPPORTOLEDRAGIMAGES								24
#define DISPID_STATBAR_USESYSTEMFONT											25
#define DISPID_STATBAR_VERSION														26
// fingerprint
#define DISPID_STATBAR_APPID															500
#define DISPID_STATBAR_APPNAME														501
#define DISPID_STATBAR_APPSHORTNAME												502
#define DISPID_STATBAR_BUILD															503
#define DISPID_STATBAR_CHARSET														504
#define DISPID_STATBAR_ISRELEASE													505
#define DISPID_STATBAR_PROGRAMMER													506
#define DISPID_STATBAR_TESTER															507
// methods
#define DISPID_STATBAR_ABOUT															DISPID_ABOUTBOX
#define DISPID_STATBAR_GETBORDERS													27
#define DISPID_STATBAR_HITTEST														28
#define DISPID_STATBAR_LOADSETTINGSFROMFILE								29
#define DISPID_STATBAR_REFRESH														DISPID_REFRESH
#define DISPID_STATBAR_SAVESETTINGSTOFILE									30
#define DISPID_STATBAR_FINISHOLEDRAGDROP									31


// _IStatusBarEvents
// methods
#define DISPID_STATBARE_CLICK															DISPID_VALUE
#define DISPID_STATBARE_CONTEXTMENU												1
#define DISPID_STATBARE_DBLCLICK													2
#define DISPID_STATBARE_DESTROYEDCONTROLWINDOW						3
#define DISPID_STATBARE_MCLICK														4
#define DISPID_STATBARE_MDBLCLICK													5
#define DISPID_STATBARE_MOUSEDOWN													6
#define DISPID_STATBARE_MOUSEENTER												7
#define DISPID_STATBARE_MOUSEHOVER												8
#define DISPID_STATBARE_MOUSELEAVE												9
#define DISPID_STATBARE_MOUSEMOVE													10
#define DISPID_STATBARE_MOUSEUP														11
#define DISPID_STATBARE_OLEDRAGDROP												12
#define DISPID_STATBARE_OLEDRAGENTER											13
#define DISPID_STATBARE_OLEDRAGLEAVE											14
#define DISPID_STATBARE_OLEDRAGMOUSEMOVE									15
#define DISPID_STATBARE_OWNERDRAWPANEL										16
#define DISPID_STATBARE_PANELMOUSEENTER										17
#define DISPID_STATBARE_PANELMOUSELEAVE										18
#define DISPID_STATBARE_RCLICK														19
#define DISPID_STATBARE_RDBLCLICK													20
#define DISPID_STATBARE_RECREATEDCONTROLWINDOW						21
#define DISPID_STATBARE_RESIZEDCONTROLWINDOW							22
#define DISPID_STATBARE_SETUPTOOLTIPWINDOW								23
#define DISPID_STATBARE_TOGGLEDSIMPLEMODE									24
#define DISPID_STATBARE_XCLICK														25
#define DISPID_STATBARE_XDBLCLICK													26


// IStatusBarPanel
// properties
#define DISPID_SBP_ALIGNMENT															1
#define DISPID_SBP_BORDERSTYLE														2
#define DISPID_SBP_CONTENT																3
#define DISPID_SBP_CURRENTWIDTH														4
#define DISPID_SBP_ENABLED																5
#define DISPID_SBP_FORECOLOR															6
#define DISPID_SBP_HICON																	7
#define DISPID_SBP_INDEX																	8
#define DISPID_SBP_MINIMUMWIDTH														9
#define DISPID_SBP_PANELDATA															10
#define DISPID_SBP_PARSETABS															11
#define DISPID_SBP_PREFERREDWIDTH													12
#define DISPID_SBP_RIGHTTOLEFTTEXT												13
#define DISPID_SBP_TEXT																		DISPID_VALUE
#define DISPID_SBP_TOOLTIPTEXT														14
// methods
#define DISPID_SBP_GETRECTANGLE														15


// IStatusBarPanels
// properties
#define DISPID_SBPS_ITEM																	DISPID_VALUE
#define DISPID_SBPS__NEWENUM															DISPID_NEWENUM
// methods
#define DISPID_SBPS_ADD																		1
#define DISPID_SBPS_COUNT																	2
#define DISPID_SBPS_REMOVE																3
#define DISPID_SBPS_REMOVEALL															4


// IOLEDataObject
// properties
// methods
#define DISPID_ODO_CLEAR																	1
#define DISPID_ODO_GETCANONICALFORMAT											2
#define DISPID_ODO_GETDATA																3
#define DISPID_ODO_GETFORMAT															4
#define DISPID_ODO_SETDATA																5
#define DISPID_ODO_GETDROPDESCRIPTION											6
#define DISPID_ODO_SETDROPDESCRIPTION											7
