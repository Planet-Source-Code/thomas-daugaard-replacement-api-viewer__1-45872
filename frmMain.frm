VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "APIViewer"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12555
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   12555
   StartUpPosition =   2  'CenterScreen
   Begin APIViewer.XPCheckBox chkNoComments 
      Height          =   270
      Left            =   8055
      TabIndex        =   23
      Top             =   5430
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   476
      Caption         =   "Do not show comments"
      Checked         =   -1  'True
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Direction       =   2
   End
   Begin APIViewer.XPButton cmdRemove 
      Height          =   390
      Left            =   6075
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   688
      Caption         =   "Remove"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin APIViewer.XPButton cmdShow 
      Height          =   390
      Left            =   6075
      TabIndex        =   20
      Top             =   6630
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   688
      Caption         =   "Show"
      ButtonType      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cmbCategories 
      Height          =   330
      Left            =   5370
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1185
      Width           =   2820
   End
   Begin VB.ListBox lstSelectedItems 
      Height          =   1260
      IntegralHeight  =   0   'False
      Left            =   8520
      TabIndex        =   16
      Top             =   3960
      Width           =   2580
   End
   Begin APIViewer.XPButton cmdClear 
      Height          =   390
      Left            =   6075
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5295
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   688
      Caption         =   "Clear"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtItemText 
      Height          =   1455
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   5175
      Width           =   3705
   End
   Begin VB.Frame freDeclareScope 
      Caption         =   "Declare Scope"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   7290
      TabIndex        =   9
      Top             =   2415
      Width           =   1680
      Begin APIViewer.XPCheckBox chkScope 
         Height          =   270
         Index           =   0
         Left            =   255
         TabIndex        =   10
         Top             =   345
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   476
         Caption         =   "Public"
         Checked         =   -1  'True
         AutoSize        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Type            =   2
      End
      Begin APIViewer.XPCheckBox chkScope 
         Height          =   270
         Index           =   1
         Left            =   255
         TabIndex        =   11
         Top             =   675
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   476
         Caption         =   "Private"
         AutoSize        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Type            =   2
      End
   End
   Begin VB.ListBox lstItems 
      Height          =   2160
      Left            =   90
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   2670
      Width           =   5595
   End
   Begin VB.TextBox txtSearch 
      Height          =   315
      Left            =   90
      TabIndex        =   7
      Top             =   1980
      Width           =   5205
   End
   Begin VB.ComboBox cmbLibraryList 
      Height          =   330
      Left            =   2295
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1185
      Width           =   2820
   End
   Begin VB.ComboBox cmbEntryType 
      Height          =   330
      ItemData        =   "frmMain.frx":0000
      Left            =   90
      List            =   "frmMain.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1185
      Width           =   1665
   End
   Begin APIViewer.XPButton cmdCopy 
      Height          =   390
      Left            =   6075
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6045
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   688
      Caption         =   "Copy"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin APIViewer.XPButton cmdLoadFile 
      Height          =   390
      Left            =   6075
      TabIndex        =   15
      Top             =   4275
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   688
      Caption         =   "Reload API file"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin APIViewer.XPButton cmdInsert 
      Height          =   390
      Left            =   6075
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5670
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   688
      Caption         =   "Insert into VB"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin APIViewer.XPHeader xphdrMain 
      Align           =   1  'Align Top
      Height          =   795
      Left            =   0
      Top             =   0
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   1402
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Windows API Viewer"
      Description     =   "View Win32 API Constants, Declares and Types"
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Written by and Copyright (C) Thomas Daugaard, 2003."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   180
      Left            =   8130
      TabIndex        =   24
      Top             =   855
      Width           =   3450
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Double-click an item to add it to ""Selected Items"""
      Height          =   210
      Index           =   6
      Left            =   6660
      TabIndex        =   19
      Top             =   1995
      Width           =   3435
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Fliter by category:"
      Height          =   210
      Index           =   5
      Left            =   5370
      TabIndex        =   18
      Top             =   870
      Width           =   1305
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Filter by library:"
      Height          =   210
      Index           =   4
      Left            =   2295
      TabIndex        =   5
      Top             =   870
      Width           =   1110
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Selected items:"
      Height          =   210
      Index           =   3
      Left            =   90
      TabIndex        =   4
      Top             =   4890
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Available items:"
      Height          =   210
      Index           =   2
      Left            =   90
      TabIndex        =   3
      Top             =   2370
      UseMnemonic     =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Type the first few letters of the word you are looking for:"
      Height          =   210
      Index           =   1
      Left            =   90
      TabIndex        =   2
      Top             =   1665
      Width           =   4140
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "View API entries of type:"
      Height          =   210
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   870
      Width           =   1815
   End
   Begin VB.Menu menu_showpopup 
      Caption         =   "show_popup"
      Begin VB.Menu menu_showpopup_show 
         Caption         =   "Line Item"
         Index           =   0
      End
      Begin VB.Menu menu_showpopup_show 
         Caption         =   "Full Item"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type apiEntry
    DisplayName As String   ' Displayed name (Searchable)
    Library As String       ' Library in which a Sub/Function is located
    Code As String          ' Complete entry code
    EntryType As Integer    ' 1 = Constant, 2 = Declare, 3 = Type
End Type

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long

Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNFACE = 15
Private Const COLOR_WINDOW = 5

Private Const WM_KEYDOWN = &H100

Private vCategories As Variant
Private vCategoryDeclares(61) As Variant

Private APIEntries() As apiEntry
Private DLLEntries() As String

Private APIFileLoaded As Boolean
Private ListSearchIndex As Long

Private Sub chkNoComments_Click()
    UpdateFullTextView
End Sub

Private Sub chkScope_Click(Index As Integer)
    UpdateFullTextView
End Sub

Private Sub cmbCategories_Click()
    ChangeListingType cmbEntryType.ListIndex
End Sub

Private Sub cmbEntryType_Click()
    SetCategoriesEnabled cmbEntryType.ListIndex = 1
    ChangeListingType cmbEntryType.ListIndex
End Sub

Private Sub cmbLibraryList_Click()
    ChangeListingType cmbEntryType.ListIndex
End Sub

Private Sub cmdClear_Click()
    lstSelectedItems.Clear

    cmdClear.Enabled = False
    cmdCopy.Enabled = False
    cmdInsert.Enabled = False
    cmdRemove.Enabled = False

    UpdateFullTextView
End Sub

Private Sub cmdCopy_Click()
    Clipboard.SetText CStr(txtItemText.Text), vbCFText
End Sub

Private Sub cmdLoadFile_Click()
    LoadAPIFile App.Path & "\win32api.txt"
End Sub

Private Sub cmdRemove_Click()
    Dim dwSelIndex As Long

    dwSelIndex = lstSelectedItems.ListIndex
    If dwSelIndex > -1 Then
        lstSelectedItems.RemoveItem dwSelIndex
        If dwSelIndex > 0 And lstSelectedItems.ListCount > 0 Then
            dwSelIndex = dwSelIndex - 1
        ElseIf lstSelectedItems.ListCount > 0 Then
            dwSelIndex = 0
        Else
            dwSelIndex = -1
        End If

        If dwSelIndex > -1 Then
            lstSelectedItems.ListIndex = dwSelIndex
        Else
            cmdClear.Enabled = False
            cmdCopy.Enabled = False
            cmdInsert.Enabled = False
            cmdRemove.Enabled = False
        End If

        UpdateFullTextView
    End If
End Sub

Private Sub cmdShow_Click()
    PopupMenu menu_showpopup, , cmdShow.Left, cmdShow.Top + cmdShow.Height
End Sub

Private Sub Form_Load()
    On Local Error Resume Next

    Dim intIndex As Integer

    vCategories = Array("(None)", "Accessibility", "Audio", "Bitmaps", "Brushes", "Console", "Common Controls", "Common Dialog", "Communication", "Cursor", "Devices", "Dialog Boxes", "Errors", "Events", "Event Logs", "Files & Directories", "File System", "Filled Shapes", "Fonts & Text", "Handles", "Help", "Icons", "INI Files", "Input (General)", "Joysticks", "Keyboard", "Lines & Curves", "Libraries", "Math", "Media Control Interface (MCI)", "Memory", "Menus", "Messages", "Mouse", "National Language Support", "OLE", "Painting & Drawing", "Pens", "Pointers", "Printers", "Pipes", "Processes & Threads", "Rectangles", "Regions", "Registry", "Resources", "Shell", "Shutdown", "Strings", "SIDs", "Synchronization", "System Information", "Tape Devices", "Time", "Timers", "Tool Help", "Tokens", "Window Classes", "Window Procedures", "Window Properties", "Windows", "Winsock", "Other")

    vCategoryDeclares(0) = " getsystemmetrics systemparametersinfo "
    vCategoryDeclares(1) = " auxgetdevcaps auxgetnumdevs auxgetvolume auxsetvolume playsound sndplaysound waveoutgetdevcaps waveoutgetnumdevs waveoutgetvolume waveoutsetvolume "
    vCategoryDeclares(2) = " bitblt extfloodfill getpixel setpixel setpixelv stretchblt "
    vCategoryDeclares(3) = " createhatchbrush createsolidbrush getbrushorgex setbrushorgex "
    vCategoryDeclares(4) = " readconsoleoutput writeconsoleoutput readconsoleoutputcharacter readconsoleoutputattribute writeconsoleoutputcharacter writeconsoleoutputattribute fillconsoleoutputcharacter fillconsoleoutputattribute getconsolemode getnumberofconsoleinputevents getconsolescreenbufferinfo getlargestconsolewindowsize getconsolecursorinfo getnumberofconsolemousebuttons setconsolemode setconsoleactivescreenbuffer flushconsoleinputbuffer setconsolescreenbuffersize setconsolecursorposition setconsolecursorinfo scrollconsolescreenbuffer setconsolewindowinfo setconsoletextattribute setconsolectrlhandler generateconsolectrlevent allocconsole freeconsole getconsoletitle setconsoletitle readconsole writeconsole createconsolescreenbuffer getconsolecp setconsolecp getconsoleoutputcp setconsoleoutputcp "
    vCategoryDeclares(5) = " initcommoncontrolsex "
    vCategoryDeclares(6) = " choosecolor choosefont commdlgextendederror getopenfilename getsavefilename printdlg "
    vCategoryDeclares(7) = " setcommstate setcommtimeouts getcommstate getcommtimeouts purgecomm buildcommdcb buildcommdcbandtimeouts transmitcommchar setcommbreak setcommmask clearcommbreak clearcommerror setupcomm escapecommfunction getcommmask getcommproperties getcommmodemstatus waitcommevent "
    vCategoryDeclares(8) = " clipcursor createcursor destroycursor getclipcursor getcursor getcursorpos loadcursor loadcursorfromfile setcursor setcursorpos setsystemcursor showcursor "
    vCategoryDeclares(9) = " createdc deletedc deleteobject getdc getstockobject releasedc selectobject "
    vCategoryDeclares(10) = " messagebox messageboxex messageboxindirect "
    vCategoryDeclares(11) = " beep getlasterror messagebeep setlasterror setlasterrorex "
    vCategoryDeclares(12) = " setevent resetevent pulseevent deregistereventsource registereventsource reportevent "
    vCategoryDeclares(13) = " cleareventlog backupeventlog closeeventlog openeventlog openbackupeventlog readeventlog getnumberofeventlogrecords getoldesteventlogrecord "
    vCategoryDeclares(14) = " copyfile createdirectory createdirectoryex createfile deletefile findclose findfirstfile findnextfile getdiskfreespace getdiskfreespaceex getdrivetype getfileattributes getfileinformationbyhandle getfilesize getfiletime getfileversioninfo getfileversioninfosize getfullpathname getlogicaldrives getlogicaldrivestrings getshortpathname gettempfilename movefile readfile removedirectory setfileattributes setfilepointer setfiletime verqueryvalue writefile openfile sethandlecount lockfile unlockfile lockfileex unlockfileex getfiletype getstdhandle setstdhandle flushfilebuffers deviceiocontrol setendoffile duplicatehandle setcurrentdirectory getcurrentdirectory movefileex findfirstchangenotification findnextchangenotification findclosechangenotification lopen lclose lcreat llseek lread lwrite hread hwrite "
    vCategoryDeclares(15) = " getvolumeinformation setvolumelabel "
    vCategoryDeclares(16) = " chord ellipse fillrect framerect invertrect pie polygon polypolygon rectangle roundrect "
    vCategoryDeclares(17) = " createfont createfontindirect enumfontfamilies enumfontfamiliesex gettextalign settextalign textout "
    vCategoryDeclares(18) = " closehandle "
    vCategoryDeclares(19) = " winhelp "
    vCategoryDeclares(20) = " destroyicon drawicon drawiconex extracticon extracticonex "
    vCategoryDeclares(21) = " getprivateprofileint getprivateprofilestring getprofileint getprofilestring writeprivateprofilestring writeprofilestring getprofilesection writeprofilesection getprivateprofilesection writeprivateprofilesection "
    vCategoryDeclares(22) = " sendinput "
    vCategoryDeclares(23) = " joygetdevcaps joygetnumdevs joygetpos "
    vCategoryDeclares(24) = " getasynckeystate getkeyboardstate getkeystate keybd_event setkeyboardstate "
    vCategoryDeclares(25) = " anglearc arc arcto getarcdirection lineto movetoex polybezier polybezierto polyline polylineto polypolyline setarcdirection "
    vCategoryDeclares(26) = " loadlibrary loadlibraryex loadmodule freelibrary "
    vCategoryDeclares(27) = " muldiv "
    vCategoryDeclares(28) = " mcigeterrorstring mcisendstring "
    vCategoryDeclares(29) = " copymemory fillmemory globalalloc globalfree globallock globalmemorystatus globalmemorystatusex globalunlock movememory zeromemory globalhandle globalrealloc globalsize globalflags localalloc localfree localhandle locallock localrealloc localsize localunlock localflags virtualalloc virtualfree virtualprotect virtualquery virtualprotectex virtualqueryex heapcreate heapdestroy heapalloc heaprealloc heapfree heapsize "
    vCategoryDeclares(30) = " createpopupmenu destroymenu getmenu getmenuitemcount getmenuiteminfo getsystemmenu insertmenuitem removemenu setmenuiteminfo trackpopupmenu trackpopupmenuex "
    vCategoryDeclares(31) = " sendmessage "
    vCategoryDeclares(32) = " getcapture getdoubleclicktime mouse_event releasecapture setcapture setdoubleclicktime swapmousebutton "
    vCategoryDeclares(33) = " getcurrencyformat getdateformat getnumberformat getthreadlocale gettimeformat setthreadlocale "
    vCategoryDeclares(34) = " cotaskmemfree "
    vCategoryDeclares(35) = " getwindowrgn setwindowrgn "
    vCategoryDeclares(36) = " createpen createpenindirect "
    vCategoryDeclares(37) = " isbadreadptr isbadwriteptr isbadstringptr isbadhugereadptr isbadhugewriteptr "
    vCategoryDeclares(38) = " closeprinter enddoc endpage enumjobs enumprinters openprinter startdoc startpage "
    vCategoryDeclares(39) = " createnamedpipe getnamedpipehandlestate callnamedpipe waitnamedpipe createpipe connectnamedpipe disconnectnamedpipe setnamedpipehandlestate getnamedpipeinfo peeknamedpipe transactnamedpipe "
    vCategoryDeclares(40) = " getenvironmentvariable setenvironmentvariable createprocess setprocessshutdownparameters getprocessshutdownparameters getmodulefilename getmodulehandle getprocessheap getprocesstimes openprocess getcurrentprocess getcurrentprocessid exitprocess terminateprocess getexitcodeprocess readprocessmemory writeprocessmemory getthreadcontext setthreadcontext suspendthread resumethread openprocesstoken openthreadtoken createthread createremotethread getcurrentthread getcurrentthreadid setthreadpriority getthreadpriority getthreadtimes exitthread terminatethread getexitcodethread getthreadselectorentry "
    vCategoryDeclares(41) = " copyrect equalrect inflaterect intersectrect isrectempty offsetrect ptinrect setrect setrectempty subtractrect unionrect "
    vCategoryDeclares(42) = " combinergn createellipticrgn createellipticrgnindirect createpolygonrgn createpolypolygonrgn createrectrgn createrectrgnindirect createroundrectrgn equalrgn fillrgn framergn getpolyfillmode getrgnbox invertrgn offsetrgn ptinregion rectinregion setpolyfillmode "
    vCategoryDeclares(43) = " regclosekey regcreatekeyex regdeletekey regdeletevalue regenumkeyex regenumvalue regopenkeyex regqueryvalueex regsetvalueex "
    vCategoryDeclares(44) = " findresource findresourceex beginupdateresource updateresource endupdateresource loadresource lockresource sizeofresource "
    vCategoryDeclares(45) = " exitwindowsdialog pickicondlg restartdialog shaddtorecentdocs shbrowseforfolder shell_notifyicon shellexecute shellexecuteex shemptyrecyclebin shfileoperation shfreenamemappings shgetfileinfo shgetfolderlocation shgetfolderpath shgetpathfromidlist shgetspecialfolderlocation shgetspecialfolderpath shqueryrecyclebin shupdaterecyclebinicon "
    vCategoryDeclares(46) = " lockworkstation "
    vCategoryDeclares(47) = " charlower charupper comparestring lstrcmp lstrcmpi lstrcpy lstrcpyn lstrlen "
    vCategoryDeclares(48) = " isvalidsid equalsid equalprefixsid getsidlengthrequired allocateandinitializesid freesid initializesid getsididentifierauthority getsidsubauthority getsidsubauthoritycount getlengthsid copysid "
    vCategoryDeclares(49) = " waitforsingleobject "
    vCategoryDeclares(50) = " getcomputername getsyscolor getsystemdirectory gettemppath getusername getversionex getwindowsdirectory setsyscolors "
    vCategoryDeclares(51) = " settapeposition gettapeposition preparetape erasetape createtapepartition writetapemark gettapestatus gettapeparameters settapeparameters "
    vCategoryDeclares(52) = " comparefiletime filetimetolocalfiletime filetimetosystemtime getlocaltime getsystemtime getsystemtimeasfiletime gettickcount gettimezoneinformation localfiletimetofiletime setsystemtime systemtimetofiletime getcurrenttime setlocaltime getsysteminfo settimezoneinformation filetimetodosdatetime dosdatetimetofiletime enumtimeformats enumdateformats "
    vCategoryDeclares(53) = " killtimer queryperformancecounter queryperformancefrequency settimer "
    vCategoryDeclares(54) = " createtoolhelp32snapshot process32first process32next "
    vCategoryDeclares(55) = " duplicatetoken gettokeninformation settokeninformation adjusttokenprivileges adjusttokengroups "
    vCategoryDeclares(56) = " getclassinfo getclassinfoex getclasslong getclassname getwindowlong registerclass registerclassex setclasslong setwindowlong unregisterclass "
    vCategoryDeclares(57) = " callwindowproc defwindowproc "
    vCategoryDeclares(58) = " enumpropsex getprop removeprop setprop "
    vCategoryDeclares(59) = " bringwindowtotop createwindowex destroywindow enablewindow enumchildwindows enumthreadwindows enumwindows findwindow findwindowex flashwindow getactivewindow getdesktopwindow getfocus getforegroundwindow getparent gettopwindow getwindow getwindowrect getwindowtext getwindowtextlength getwindowthreadprocessid ischild isiconic iswindow iswindowenabled iszoomed movewindow setactivewindow setfocus setforegroundwindow setparent setwindowpos setwindowtext showwindow windowfrompoint "
    vCategoryDeclares(60) = " closesocket connect gethostbyaddr gethostbyname gethostname htonl htons inet_addr inet_nota recv send socket wsacleanup wsagetlasterror wsastartup "
    vCategoryDeclares(61) = " exitwindowsex sleep sleepex "

    For intIndex = 0 To UBound(vCategories)
        cmbCategories.AddItem vCategories(intIndex)
    Next

    cmbLibraryList.AddItem "(None)"
    cmbLibraryList.ListIndex = 0
    cmbCategories.ListIndex = 0
    cmbEntryType.ListIndex = 1

    menu_showpopup.Visible = False
    LoadAPIFile App.Path & "\win32api.txt"

    Call menu_showpopup_show_Click(1)
End Sub

Private Sub SetCategoriesEnabled(State As Boolean)
    lblInfo(4).ForeColor = IIf(State, 0, GetSysColor(COLOR_BTNSHADOW))
    cmbCategories.BackColor = IIf(State, GetSysColor(COLOR_WINDOW), GetSysColor(COLOR_BTNFACE))
    cmbCategories.Enabled = State

    lblInfo(5).ForeColor = lblInfo(4).ForeColor
    cmbLibraryList.BackColor = cmbCategories.BackColor
    cmbLibraryList.Enabled = State
End Sub

Private Sub SetCategoriesVisible(State As Boolean)
    cmbCategories.Visible = State
    cmbLibraryList.Visible = State

    lblInfo(4).Visible = State
    lblInfo(5).Visible = State
End Sub

Private Sub Form_Resize()
    On Local Error Resume Next

    freDeclareScope.Move Me.ScaleWidth - freDeclareScope.Width - 90, lstItems.Top
    
    lstItems.Width = freDeclareScope.Left - (lstItems.Left * 2) - 60
    txtItemText.Width = lstItems.Width
    txtItemText.Height = Me.ScaleHeight - txtItemText.Top - 90

    txtSearch.Width = lstItems.Width

    cmdLoadFile.Move freDeclareScope.Left, (lstItems.Top + lstItems.Height) - cmdLoadFile.Height

    cmdRemove.Move freDeclareScope.Left, txtItemText.Top
    cmdClear.Move freDeclareScope.Left, cmdRemove.Top + cmdRemove.Height + 30
    cmdInsert.Move freDeclareScope.Left, cmdClear.Top + cmdClear.Height + 30
    cmdCopy.Move freDeclareScope.Left, cmdInsert.Top + cmdInsert.Height + 30

    cmdShow.Move freDeclareScope.Left, (txtItemText.Top + txtItemText.Height) - cmdShow.Height

    lblInfo(6).Move (lstItems.Left + lstItems.Width) - lblInfo(6).Width, lblInfo(2).Top
    lstSelectedItems.Move txtItemText.Left, txtItemText.Top, txtItemText.Width, txtItemText.Height
    
    chkNoComments.Move (lstSelectedItems.Left + lstSelectedItems.Width) - chkNoComments.Width, lblInfo(3).Top
    lblCopyright.Move frmMain.ScaleWidth - lblCopyright.Width - 60, lblInfo(0).Top
End Sub

Private Sub ChangeListingType(intListingType As Integer)
    Dim intIndex As Integer
    Dim dwViewCount As Long, dwTotalCount As Long
    Dim strLibraryFilter As String
    Dim intCatID As Integer
    Dim blnAddItem As Boolean

    If APIFileLoaded Then
        lstItems.Clear
        Call LockWindowUpdate(lstItems.hWnd)

        strLibraryFilter = DLLEntries(cmbLibraryList.ItemData(cmbLibraryList.ListIndex))
        intCatID = cmbCategories.ListIndex

        For intIndex = 0 To UBound(APIEntries)
            If APIEntries(intIndex).EntryType = intListingType + 1 Then
                blnAddItem = True

                If strLibraryFilter <> "(None)" Then blnAddItem = (APIEntries(intIndex).Library = strLibraryFilter)
                If intCatID > 0 Then blnAddItem = blnAddItem And (InStr(vCategoryDeclares(intCatID - 1), " " & LCase(APIEntries(intIndex).DisplayName) & " ") > 0)

                If blnAddItem Then
                    lstItems.AddItem APIEntries(intIndex).DisplayName
                    lstItems.ItemData(lstItems.NewIndex) = intIndex
                    dwViewCount = dwViewCount + 1
                End If

                dwTotalCount = dwTotalCount + 1
            End If
        Next

        Call LockWindowUpdate(0)
        
        Dim strCaption  As String
        strCaption = "Displaying " & dwViewCount & " items"

        If strLibraryFilter <> "(None)" Then
            strCaption = strCaption & " in """ & cmbLibraryList.Text & """"
            If intCatID > 0 Then strCaption = strCaption & ", """ & cmbCategories.Text & """"
        Else
            If intCatID > 0 Then strCaption = strCaption & " in """ & cmbCategories.Text & """"
        End If

        lblInfo(2).Caption = strCaption & ":"
    End If
End Sub

Private Sub LoadAPIFile(strFile As String)
    If Dir$(strFile) = "" Then
        MsgBox "Selected file does not exist!"
        Exit Sub
    End If

    Dim intFH As Integer, intPos As Integer, intCommentPos As Integer, intQuotePos As Integer
    Dim intDLLIndex As Integer
    Dim strLine As String, strToken As String, strRest As String, strLibraries As String
    Dim blnAdvanceNext As Boolean
    Dim dwEntryIndex As Long
    Dim vDeclare

    dwEntryIndex = 0
    intDLLIndex = 1
    strLibraries = "|"
    intFH = FreeFile

    Open strFile For Input As #intFH
    ReDim APIEntries(100) As apiEntry
    ReDim DLLEntries(50) As String

    DLLEntries(0) = "(None)"

    frmMain.MousePointer = vbHourglass
    Do Until EOF(intFH)
        Line Input #1, strLine
        strLine = Trim(strLine)

        If Not (Left(strLine, 1) = "'" Or Left(strLine, 1) = "") Then
            intPos = InStr(strLine, " ")
            If intPos > 0 Then
                strToken = Left(strLine, intPos - 1)
                strRest = Right(strLine, Len(strLine) - intPos)
                intCommentPos = InStrRev(strRest, "'")

                If intCommentPos > 0 Then
                    intQuotePos = InStrRev(strRest, Chr(34))
                    If intCommentPos > intQuotePos Or intQuotePos = 0 Then
                        strRest = Left(strRest, intCommentPos - 1)
                    End If
                End If

                blnAdvanceNext = False
                Select Case LCase(strToken)
                    Case "const":
                        ' Const SHUTDOWN_NORETRY = &H1

                        With APIEntries(dwEntryIndex)
                            .DisplayName = strRest
                            .Code = strLine
                            .EntryType = 1
                        End With

                    Case "declare":
                        ' Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Long, ByVal lpName As String) As Long

                        vDeclare = Split(strLine, " ", 6)
                        vDeclare(4) = Replace(LCase(UnQuote(vDeclare(4))), ".dll", "")
                        If InStr(strLibraries, "|" & vDeclare(4) & "|") = 0 Then
                            strLibraries = strLibraries & vDeclare(4) & "|"
                            DLLEntries(intDLLIndex) = vDeclare(4)

                            intDLLIndex = intDLLIndex + 1
                            If intDLLIndex > UBound(DLLEntries) Then
                                ReDim Preserve DLLEntries(UBound(DLLEntries) + 50) As String
                            End If
                        End If
                        
                        With APIEntries(dwEntryIndex)
                            .DisplayName = vDeclare(2)
                            .Library = vDeclare(4)
                            .Code = strLine
                            .EntryType = 2
                        End With

                    Case "type":
                        ' Type WIN32_FIND_DATA

                        With APIEntries(dwEntryIndex)
                            .DisplayName = strRest
                            .Code = strLine & vbCrLf
                            .EntryType = 3
                        End With

                    Case Else:
                        ' Continued code

                        If LCase(strLine) = "end type" Then blnAdvanceNext = True
                        APIEntries(dwEntryIndex).Code = APIEntries(dwEntryIndex).Code & IIf(blnAdvanceNext, "", Space$(4)) & Trim(strLine) & vbCrLf
                End Select
                
                If APIEntries(dwEntryIndex).EntryType <> 3 Or blnAdvanceNext Then
                    dwEntryIndex = dwEntryIndex + 1
                    If dwEntryIndex > UBound(APIEntries) Then ReDim Preserve APIEntries(UBound(APIEntries) + 100) As apiEntry
                End If
            End If
        End If
        
        DoEvents
    Loop

    ReDim Preserve APIEntries(dwEntryIndex - 1) As apiEntry
    Dim intIndex As Integer

    cmbLibraryList.Clear

    For intIndex = 0 To intDLLIndex - 1
        cmbLibraryList.AddItem DLLEntries(intIndex) & IIf(InStr(DLLEntries(intIndex), ".") = 0 And Left(DLLEntries(intIndex), 1) <> "(", ".dll", "")
        cmbLibraryList.ItemData(cmbLibraryList.NewIndex) = intIndex
    Next

    ReDim Preserve DLLEntries(intDLLIndex - 1) As String

    strLibraries = ""
    cmbLibraryList.ListIndex = 0
    frmMain.MousePointer = vbDefault
    APIFileLoaded = True
    
    Call cmbEntryType_Click

    Close #intFH
End Sub

Private Function UnQuote(ByVal strStr As String) As String
    If Left(strStr, 1) = Chr(34) And Right(strStr, 1) = Chr(34) Then
        strStr = Mid(strStr, 2, Len(strStr) - 2)
    End If

    UnQuote = strStr
End Function

Private Sub lstItems_DblClick()
    Dim dwIndex As Long
    Dim blnAddItem As Boolean
    Dim szPrefix As String, szText As String
    Dim vSplit As Variant
    
    Select Case cmbEntryType.ListIndex
        Case 0: szPrefix = "Const: "
        Case 1: szPrefix = "Declare: "
        Case 2: szPrefix = "Type: "
    End Select

    szText = lstItems.List(lstItems.ListIndex)
    If InStr(szText, " ") > 0 Then
        vSplit = Split(szText, " ", 5)
        szText = vSplit(IIf(LCase(vSplit(1)) = "declare", 2, 0))
    End If
    
    blnAddItem = True
    For dwIndex = 0 To lstSelectedItems.ListCount - 1
        If lstSelectedItems.List(dwIndex) = szPrefix & szText Then
            blnAddItem = False
            Exit For
        End If
    Next
    
    If blnAddItem Then
        lstSelectedItems.AddItem szPrefix & szText
        lstSelectedItems.ItemData(lstSelectedItems.NewIndex) = lstItems.ItemData(lstItems.ListIndex)
    End If

    cmdClear.Enabled = True
    cmdCopy.Enabled = True
    cmdInsert.Enabled = True
    cmdRemove.Enabled = True

    UpdateFullTextView
End Sub

Private Sub lstItems_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call lstItems_DblClick
End Sub

Private Sub menu_showpopup_show_Click(Index As Integer)
    menu_showpopup_show(0).Checked = (Index = 0)
    menu_showpopup_show(1).Checked = (Index = 1)

    cmdShow.Caption = "Show [" & IIf(Index = 0, "Line", "Full") & "]"

    Select Case Index
        Case 0: lstSelectedItems.ZOrder
        Case 1: txtItemText.ZOrder
    End Select
End Sub

Private Sub txtSearch_Change()
    Dim dwIndex As Long, dwIdx As Long

    For dwIndex = ListSearchIndex To lstItems.ListCount - 1
        If LCase(Left(lstItems.List(dwIndex), Len(txtSearch.Text))) = LCase(txtSearch.Text) Then
            If dwIndex + 9 < lstItems.ListCount - 1 Then dwIdx = dwIndex + 9 Else dwIdx = dwIndex

            lstItems.ListIndex = dwIdx
            Exit For
        End If
    Next

    If dwIndex >= lstItems.ListCount Then Exit Sub

    ' Set search index (optimized searching)
    ListSearchIndex = dwIndex - 1
    lstItems.ListIndex = dwIndex
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageDown Or KeyCode = vbKeyPageUp Then
        Call SendMessage(lstItems.hWnd, WM_KEYDOWN, KeyCode, &O0)
        KeyCode = 0
    ElseIf KeyCode = vbKeyReturn Then
        Call lstItems_DblClick
    ElseIf KeyCode = vbKeyBack Then
        ' Reset search index
        ListSearchIndex = 0
    End If
End Sub

Private Sub UpdateFullTextView()
    Dim dwIndex As Long, dwCodeID As Long, dwCodeIdx As Long
    Dim szScope As String, szCodeView As String
    Dim intCommentPos As Integer, intQuotePos As Integer
    Dim vCode As Variant

    szScope = IIf(chkScope(0).Checked, "Public ", "Private ")
    txtItemText.Text = ""

    If lstSelectedItems.ListCount > -1 Then
        For dwIndex = 0 To lstSelectedItems.ListCount - 1
            dwCodeID = lstSelectedItems.ItemData(dwIndex)
            vCode = Split(APIEntries(dwCodeID).Code, vbCrLf)

            For dwCodeIdx = 0 To UBound(vCode)
                vCode(dwCodeIdx) = LTrimNL(RTrimNL(CStr(vCode(dwCodeIdx))))
                
                If chkNoComments.Checked Then
                    intCommentPos = InStrRev(vCode(dwCodeIdx), "'")

                    If intCommentPos > 0 Then
                        intQuotePos = InStrRev(vCode(dwCodeIdx), Chr(34))
                        If intCommentPos > intQuotePos Or intQuotePos = 0 Then
                            vCode(dwCodeIdx) = Left(vCode(dwCodeIdx), intCommentPos - 1)
                        End If
                    End If
                End If

                vCode(dwCodeIdx) = RTrim(vCode(dwCodeIdx))
            Next
            
            szCodeView = szCodeView & szScope & RTrimNL(Join(vCode, vbCrLf)) & vbCrLf & vbCrLf
        Next
    End If

    txtItemText.Text = szCodeView
    txtItemText.SelStart = Len(txtItemText.Text)
End Sub

Private Sub txtItemText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Local Error Resume Next

    Dim szText As String, szPrefix As String, szScope As String
    Dim intPos As Integer
    Dim vSplit As Variant

    szScope = IIf(chkScope(0).Checked, "Public ", "Private ")
    szText = txtItemText.Text

    If szText > "" Then
        intPos = InStr(txtItemText.SelStart, szText, vbCrLf)
        intPos = IIf(intPos > 0, intPos, Len(txtItemText.Text))
        szText = RTrimNL(Left(szText, intPos + 1))
        
        intPos = InStrRev(szText, szScope)
        szText = RTrimNL(Right(szText, Len(szText) - intPos + 1))
        
        intPos = InStr(szText, vbCrLf)
        intPos = IIf(intPos > 0, intPos - 1, Len(szText))
        szText = Left(szText, intPos)

        vSplit = Split(szText, " ", 6)

        szText = vSplit(IIf(LCase(vSplit(1)) = "declare", 3, 2))
        szText = StrConv(vSplit(1), vbProperCase) & ": " & RTrimNL(szText)

        For intPos = 0 To lstSelectedItems.ListCount - 1
            If lstSelectedItems.List(intPos) = szText Then Exit For
        Next

        lstSelectedItems.ListIndex = intPos
    End If
End Sub

Private Function RTrimNL(szStr As String) As String
    Dim intPos As Integer

    If Len(szStr) > 0 Then
        For intPos = Len(szStr) To 1 Step -1
            If Asc(Mid(szStr, intPos, 1)) > 31 Then Exit For
        Next
    
        RTrimNL = IIf(intPos > 0, Left(szStr, intPos), szStr)
    Else
        RTrimNL = ""
    End If
End Function

Private Function LTrimNL(szStr As String) As String
    Dim intPos As Integer

    For intPos = 1 To Len(szStr)
        If Asc(Mid(szStr, intPos, 1)) > 31 Then
            Exit For
        End If
    Next

    
    LTrimNL = IIf(intPos > 1, Right(szStr, Len(szStr) - intPos + 1), szStr)
End Function

