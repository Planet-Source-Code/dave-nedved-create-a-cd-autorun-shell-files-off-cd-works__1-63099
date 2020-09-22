VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "David Nedveds MIDI Assesment"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3600
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin DavidsMIDI.chameleonButton chameleonButton1 
      Height          =   495
      Left            =   180
      TabIndex        =   0
      Top             =   195
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   873
      BTYPE           =   9
      TX              =   "Reason Project"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":0ECA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin DavidsMIDI.chameleonButton chameleonButton2 
      Height          =   495
      Left            =   180
      TabIndex        =   1
      Top             =   915
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   873
      BTYPE           =   9
      TX              =   "Audio Mixdown"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":0EE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin DavidsMIDI.chameleonButton chameleonButton3 
      Height          =   495
      Left            =   180
      TabIndex        =   2
      Top             =   1635
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   873
      BTYPE           =   9
      TX              =   "Reason Document"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":0F02
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin DavidsMIDI.chameleonButton chameleonButton4 
      Height          =   495
      Left            =   180
      TabIndex        =   3
      Top             =   2355
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   873
      BTYPE           =   9
      TX              =   "Explore Files"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":0F1E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin DavidsMIDI.chameleonButton chameleonButton5 
      Height          =   495
      Left            =   180
      TabIndex        =   4
      Top             =   3075
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   873
      BTYPE           =   9
      TX              =   "Quit NOW!"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":0F3A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Shape shaBDR 
      BorderColor     =   &H00E0E0E0&
      Height          =   645
      Index           =   4
      Left            =   120
      Top             =   3000
      Width           =   3375
   End
   Begin VB.Shape shaBDR 
      BorderColor     =   &H00E0E0E0&
      Height          =   645
      Index           =   3
      Left            =   120
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Shape shaBDR 
      BorderColor     =   &H00E0E0E0&
      Height          =   645
      Index           =   2
      Left            =   120
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Shape shaBDR 
      BorderColor     =   &H00E0E0E0&
      Height          =   645
      Index           =   1
      Left            =   120
      Top             =   840
      Width           =   3375
   End
   Begin VB.Shape shaBDR 
      BorderColor     =   &H00E0E0E0&
      Height          =   645
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem // ----------------------------------------------------------------------
Rem // | This project is a Simple Autorun by David Nedved                   |
Rem // |                                                                    |
Rem // | EM. dnedved@datosoftware.com                                       |
Rem // | WS. www.datosoftware.com                                           |
Rem // |                                                                    |
Rem // | The Button Control Code was not by Me, but i used it as            |
Rem // | This project was for a non Profit, Assignment, that i used         |
Rem // | on a Mixed CD (Data CD & Audio CD)                                 |
Rem // |                                                                    |
Rem // | Button Control Made by gonchuki                                    |
Rem // | I Decided to Upload this on to the Web as There aren't Many        |
Rem // | Examples on how to make a Basic Autorun.                           |
Rem // |                                                                    |
Rem // | One of the things that i first struggled with was to               |
Rem // | Open A file, when it is on a CD                                    |
Rem // | e.g. 'Shell App.Path & "\Data\Setup.exe' would not work when       |
Rem // | used as an Autorun.                                                |
Rem // |                                                                    |
Rem // | I Stuffed around for a While with the Windows Shell Execute        |
Rem // | And this is what I came up with.                                   |
Rem // | You can use this code to Shell to a website, File, Folder, etc...  |
Rem // |                                                                    |
Rem // | Have fun, and please Vote if you learnt anything                   |
Rem // | about making a Basic Autorun.                                      |
Rem // ----------------------------------------------------------------------


Option Explicit
 Rem // Declare the Functions that are 'Outside' Of Visual Basic.
 Rem // Declare the Function to SetWindowPos, that will be used to create the Autorun 'Always On Top'
 Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
 Rem // Declare the Function to ShellExecute to a Location (This function is inside the "shell32.dll" in the Windows System DIR
 Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub chameleonButton1_Click()
Rem // Shell a file on the Local System, or CD Rom
ShellExecute Me.hwnd, "open", App.Path & "\Use MIDI task (2)\EXPORTED\David Nedveds Use MIDI Assignment.rps", 0&, "", vbNormalFocus
End Sub

Private Sub chameleonButton2_Click()
Rem // Shell a file on the Local System, or CD Rom
ShellExecute Me.hwnd, "open", App.Path & "\Use MIDI task (2)\EXPORTED\David Nedveds Use MIDI Assignment.mp3", 0&, "", vbNormalFocus
End Sub

Private Sub chameleonButton3_Click()
Rem // Shell a file on the Local System, or CD Rom
ShellExecute Me.hwnd, "open", App.Path & "\Use MIDI task (2)\DOC\Reason Screenshots.doc", 0&, "", vbNormalFocus
End Sub

Private Sub chameleonButton4_Click()
Rem // Shell a folder on the Local System, or CD Rom
Rem // If you wanted to shell to a Website you could just type in the Website address e.g.
Rem // ShellExecute Me.hwnd, "open", "http://www.datosoftware.com", 0&, "", vbNormalFocus
ShellExecute Me.hwnd, "open", App.Path & "\Use MIDI task (2)", 0&, "", vbNormalFocus
End Sub

Private Sub chameleonButton5_Click()
Rem // Exits the Application
End
End Sub

Private Sub Form_Load()
Rem // This Code can be used to Create the Form 'Always On Top' of others, Like wise you can Reverse the code and the Form will be 'Always on Bottom'
SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
End Sub
