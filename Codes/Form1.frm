VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Naveed ASCII City"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   404
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CDLSave 
      Left            =   3135
      Top             =   5460
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Save ASCII Picture to"
      Filter          =   "HTML File|*.htm;*.html|All Files|*.*"
   End
   Begin VB.CommandButton cmdWriteTo 
      Caption         =   "..."
      Height          =   375
      Left            =   3750
      TabIndex        =   29
      Top             =   5550
      Width           =   450
   End
   Begin VB.TextBox txtWriteTo 
      Height          =   285
      Left            =   795
      TabIndex        =   28
      Top             =   5595
      Width           =   2865
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   5760
      TabIndex        =   31
      Top             =   5550
      Width           =   1185
   End
   Begin VB.CommandButton cmdMake 
      Caption         =   "&Make"
      Default         =   -1  'True
      Height          =   375
      Left            =   4380
      TabIndex        =   30
      Top             =   5550
      Width           =   1185
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5205
      Left            =   135
      TabIndex        =   32
      Top             =   195
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   9181
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Picture"
      TabPicture(0)   =   "Form1.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "HTML Settings"
      TabPicture(1)   =   "Form1.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Preview"
      TabPicture(2)   =   "Form1.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame4 
         Height          =   4575
         Left            =   -74790
         TabIndex        =   37
         Top             =   405
         Width           =   6480
         Begin SHDocVwCtl.WebBrowser WWW 
            Height          =   4245
            Left            =   120
            TabIndex        =   26
            Top             =   210
            Width           =   6225
            ExtentX         =   10980
            ExtentY         =   7488
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4575
         Left            =   -74790
         TabIndex        =   36
         Top             =   405
         Width           =   6480
         Begin MSComDlg.CommonDialog CDLBrowseB 
            Left            =   3615
            Top             =   3735
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            CancelError     =   -1  'True
            DialogTitle     =   "Codes after ASCII Picture"
            Filter          =   "HTML File|*.htm;*.html|All Files|*.*"
         End
         Begin MSComDlg.CommonDialog CDLBrowseA 
            Left            =   3660
            Top             =   2895
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            CancelError     =   -1  'True
            DialogTitle     =   "Codes before ASCII Picture"
            Filter          =   "HTML File|*.htm;*.html|All Files|*.*"
         End
         Begin MSComDlg.CommonDialog CDLColor 
            Left            =   1890
            Top             =   1380
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            CancelError     =   -1  'True
         End
         Begin VB.TextBox txtSize 
            Height          =   285
            Left            =   5850
            MaxLength       =   50
            TabIndex        =   16
            Text            =   "1"
            Top             =   1365
            Width           =   420
         End
         Begin VB.TextBox txtTitle 
            Height          =   285
            Left            =   2445
            TabIndex        =   10
            Text            =   "ASCII Picture Created By Naveed ASCII City (Feedback: neenojee@hotmail.com)"
            Top             =   945
            Width           =   3825
         End
         Begin VB.CheckBox chkCenter 
            Caption         =   "&Center."
            Height          =   195
            Left            =   2445
            TabIndex        =   17
            Top             =   1830
            Value           =   1  'Checked
            Width           =   1260
         End
         Begin VB.OptionButton optUseQucik 
            Caption         =   "Use q&uick options:"
            Height          =   195
            Left            =   285
            TabIndex        =   8
            Top             =   450
            Value           =   -1  'True
            Width           =   4320
         End
         Begin VB.CommandButton cmdBackColor 
            BackColor       =   &H00000000&
            Height          =   285
            Left            =   2445
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   1365
            Width           =   270
         End
         Begin VB.CheckBox chkBold 
            Caption         =   "B&old."
            Height          =   195
            Left            =   4005
            TabIndex        =   18
            Top             =   1845
            Value           =   1  'Checked
            Width           =   1260
         End
         Begin VB.ComboBox cmbFonts 
            Height          =   315
            Left            =   3825
            TabIndex        =   14
            Text            =   "Terminal"
            Top             =   1350
            Width           =   1440
         End
         Begin VB.OptionButton optUseAdv 
            Caption         =   "Use ad&vance method:"
            Height          =   195
            Left            =   285
            TabIndex        =   19
            Top             =   2355
            Width           =   4320
         End
         Begin VB.TextBox txtFileA 
            Height          =   300
            Left            =   1680
            OLEDropMode     =   1  'Manual
            TabIndex        =   21
            Text            =   "<Browse File>"
            Top             =   3030
            Width           =   3240
         End
         Begin VB.CommandButton cmdBrowseA 
            Caption         =   "Browse"
            Height          =   375
            Left            =   5085
            TabIndex        =   22
            Top             =   3000
            Width           =   1185
         End
         Begin VB.TextBox txtFileB 
            Height          =   300
            Left            =   1680
            OLEDropMode     =   1  'Manual
            TabIndex        =   24
            Text            =   "<Browse File>"
            Top             =   3870
            Width           =   3240
         End
         Begin VB.CommandButton cmdBrowseB 
            Caption         =   "Browse"
            Height          =   375
            Left            =   5085
            TabIndex        =   25
            Top             =   3840
            Width           =   1185
         End
         Begin VB.Label lblOPTUseQuick 
            AutoSize        =   -1  'True
            Caption         =   "Web page &title:"
            Height          =   195
            Index           =   0
            Left            =   1005
            TabIndex        =   9
            Top             =   975
            Width           =   1080
         End
         Begin VB.Label lblOPTUseQuick 
            AutoSize        =   -1  'True
            Caption         =   "Backgroun&d color:"
            Height          =   195
            Index           =   1
            Left            =   1005
            TabIndex        =   11
            Top             =   1410
            Width           =   1305
         End
         Begin VB.Label lblOPTUseQuick 
            AutoSize        =   -1  'True
            Caption         =   "Fo&nt face:"
            Height          =   195
            Index           =   2
            Left            =   2985
            TabIndex        =   13
            Top             =   1410
            Width           =   720
         End
         Begin VB.Label lblOPTUseQuick 
            AutoSize        =   -1  'True
            Caption         =   "Si&ze:"
            Height          =   195
            Index           =   3
            Left            =   5415
            TabIndex        =   15
            Top             =   1410
            Width           =   345
         End
         Begin VB.Label lblOPTUseAdv 
            AutoSize        =   -1  'True
            Caption         =   "HTML c&odes after <HTML> tag and before ASCII picture codes:"
            Height          =   195
            Index           =   0
            Left            =   1005
            TabIndex        =   20
            Top             =   2700
            Width           =   4545
         End
         Begin VB.Label lblOPTUseAdv 
            AutoSize        =   -1  'True
            Caption         =   "File:"
            Height          =   195
            Index           =   1
            Left            =   1005
            TabIndex        =   39
            Top             =   3090
            Width           =   285
         End
         Begin VB.Label lblOPTUseAdv 
            AutoSize        =   -1  'True
            Caption         =   "HTML codes after ASCII p&icture codes and before </HTML> tag:"
            Height          =   195
            Index           =   2
            Left            =   1005
            TabIndex        =   23
            Top             =   3540
            Width           =   4620
         End
         Begin VB.Label lblOPTUseAdv 
            AutoSize        =   -1  'True
            Caption         =   "File:"
            Height          =   195
            Index           =   3
            Left            =   1005
            TabIndex        =   38
            Top             =   3930
            Width           =   285
         End
         Begin VB.Line Line1 
            X1              =   375
            X2              =   375
            Y1              =   2190
            Y2              =   750
         End
         Begin VB.Line Line2 
            X1              =   390
            X2              =   390
            Y1              =   2655
            Y2              =   4260
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4575
         Left            =   210
         TabIndex        =   33
         Top             =   405
         Width           =   6480
         Begin VB.ComboBox cmbQuality 
            Height          =   315
            ItemData        =   "Form1.frx":0496
            Left            =   3000
            List            =   "Form1.frx":0498
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1635
            Width           =   1680
         End
         Begin VB.TextBox txtASCIICHAR 
            Height          =   285
            Left            =   1155
            TabIndex        =   7
            Text            =   "o"
            Top             =   3030
            Width           =   5100
         End
         Begin VB.PictureBox picProperties 
            BorderStyle     =   0  'None
            Height          =   645
            Left            =   180
            ScaleHeight     =   645
            ScaleWidth      =   6135
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   3810
            Width           =   6135
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Caption         =   "FlyFusion Silicon Release"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000985F4&
               Height          =   195
               Left            =   2595
               TabIndex        =   46
               Top             =   330
               Width           =   3345
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "www.geocities.com/kyriakosnicola/ff_setup.zip"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00CF4212&
               Height          =   195
               Left            =   2595
               MousePointer    =   3  'I-Beam
               TabIndex        =   45
               Top             =   30
               Width           =   3345
            End
            Begin VB.Label lblDimensions 
               AutoSize        =   -1  'True
               Caption         =   "Dimensions:"
               Height          =   195
               Left            =   195
               TabIndex        =   41
               Top             =   105
               Width           =   855
            End
         End
         Begin VB.PictureBox picPic 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   705
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   35
            Top             =   1665
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.CommandButton cmdPaste 
            Caption         =   "&Paste"
            Height          =   375
            Left            =   5070
            TabIndex        =   5
            Top             =   1605
            Width           =   1185
         End
         Begin VB.TextBox txtFile 
            Height          =   300
            Left            =   210
            OLEDropMode     =   1  'Manual
            TabIndex        =   1
            Text            =   "Created by Naveed. neenojee@hotmail.com"
            Top             =   735
            Width           =   6060
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "&Browse"
            Height          =   375
            Left            =   5070
            TabIndex        =   2
            Top             =   1155
            Width           =   1185
         End
         Begin MSComDlg.CommonDialog CDLOpen 
            Left            =   5010
            Top             =   1125
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            CancelError     =   -1  'True
            DialogTitle     =   "Open Picture"
            Filter          =   "All Picture Files|*.bmp;*.dib;*.jpg;*.jpeg;*.gif;*.ico;*.cur|All Files|*.*"
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Download the fastest downloading accelerator:"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   2760
            TabIndex        =   44
            Top             =   3540
            Width           =   3330
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "&Quality:"
            Height          =   195
            Left            =   2340
            TabIndex        =   3
            Top             =   1695
            Width           =   525
         End
         Begin VB.Image imgPIC 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   1365
            Left            =   225
            Top             =   1440
            Width           =   1770
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "&ASCII text:"
            Height          =   195
            Left            =   240
            TabIndex        =   6
            Top             =   3075
            Width           =   750
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "(File is not a picture file. It is of invalid format so please select some other)"
            Height          =   195
            Left            =   630
            TabIndex        =   43
            Top             =   4080
            Width           =   5145
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Picture Properties:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   135
            TabIndex        =   42
            Top             =   3465
            Width           =   2235
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Picture &file name:"
            Height          =   195
            Left            =   240
            TabIndex        =   0
            Top             =   450
            Width           =   1215
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "(You can drag and drop file from explorer)"
            Height          =   195
            Left            =   210
            TabIndex        =   34
            Top             =   1140
            Width           =   2910
         End
      End
   End
   Begin VB.Label lblWriteTo 
      AutoSize        =   -1  'True
      Caption         =   "&Write to:"
      Height          =   195
      Left            =   135
      TabIndex        =   27
      Top             =   5640
      Width           =   600
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Naveed ur Rehman
'Email: neenojee@hotmail.com
'Thank you very much for all those who appritiate my
'previous release of the same codes which contain a
'few bugs. But all of them are perfectly removed.
'If you still find any error then report it to me.
'Feel free to email me :D I love your emails.
'All codes are now yours, change them if you need
'but please try informing me the good and successfull
'changes you make (or tried to make).
'
'Dont forget to vote me on www.planet-source-code.com
'And also see my other subbmissions there too !
'
'Thank you

Option Explicit

Dim ASCIICHAR As String, _
    POINTERCHAR As Integer

Private Sub cmbFonts_Click()
    
    On Error GoTo HELLERROR
    
    txtASCIICHAR.Font = cmbFonts.Text
    
    Exit Sub

HELLERROR:

    txtASCIICHAR.Font = "Terminal"
    
End Sub

Private Sub cmdBackColor_Click()
    
    On Error GoTo HELLERROR
        
    CDLColor.Color = cmdBackColor.BackColor
    CDLColor.ShowColor
    
    cmdBackColor.BackColor = CDLColor.Color
    
HELLERROR:

End Sub

Private Sub cmdBrowse_Click()

    On Error GoTo HELLERROR
    
    CDLOpen.Filename = txtFile.Text
    CDLOpen.ShowOpen
    
    txtFile.Text = CDLOpen.Filename
    
    Call SetWriteToIntelligence 'Fillup txtWriteTo textbox
    
    'Now see txtFile_change()
   
HELLERROR:
    
End Sub

Private Sub cmdBrowseA_Click()
    
    On Error GoTo HELLERROR
    
    CDLBrowseA.ShowSave
    
    txtFileA.Text = CDLBrowseA.Filename
    
HELLERROR:

End Sub

Private Sub cmdBrowseB_Click()
    
    On Error GoTo HELLERROR
    
    CDLBrowseB.ShowSave
    
    txtFileB.Text = CDLBrowseB.Filename
    
HELLERROR:

End Sub

Private Sub cmdClose_Click()

    End
    
End Sub

Private Sub cmdPaste_Click()

On Error GoTo HELLERROR

    If Clipboard.GetFormat(vbCFBitmap) = True Then
    
        Dim TempFileName
        
        TempFileName = XBuildPath(TempDir, GetRandom & ".bmp")
        SavePicture Clipboard.GetData(vbCFBitmap), TempFileName
        
        txtFile.Text = TempFileName
    
        Call SetWriteToIntelligence 'Fillup txtWriteTo textbox
        
    End If
    
HELLERROR:

End Sub

Private Sub cmdWriteTo_Click()

    On Error GoTo HELLERROR
    
    CDLSave.ShowSave
    
    txtWriteTo.Text = CDLSave.Filename
    
HELLERROR:

End Sub

Private Sub Form_Load()
    
    Dim i
    
    Call UpdateFormSettings  'Updating form settings (due to option buttons)
    
    'Setting Flags of CDL boxes:
    CDLOpen.Flags = CDLOpen.Flags Or cdlOFNExtensionDifferent _
                    Or cdlOFNFileMustExist Or cdlOFNHideReadOnly _
                    Or cdlOFNLongNames Or cdlOFNShareAware Or _
                    cdlOFNPathMustExist
    CDLBrowseA.Flags = CDLBrowseA.Flags Or cdlOFNExtensionDifferent _
                    Or cdlOFNFileMustExist Or cdlOFNHideReadOnly _
                    Or cdlOFNLongNames Or cdlOFNShareAware Or _
                    cdlOFNPathMustExist
    CDLBrowseB.Flags = CDLBrowseB.Flags Or cdlOFNExtensionDifferent _
                    Or cdlOFNFileMustExist Or cdlOFNHideReadOnly _
                    Or cdlOFNLongNames Or cdlOFNShareAware Or _
                    cdlOFNPathMustExist
    CDLSave.Flags = CDLSave.Flags Or cdlOFNExplorer Or _
                    cdlOFNHideReadOnly Or cdlOFNLongNames _
                    Or cdlOFNOverwritePrompt
    
   
    'Filling fonts list in cmbFonts
    For i = 0 To Screen.FontCount - 1
        cmbFonts.AddItem Screen.Fonts(i)
    Next i
    
    'Qualities:
    cmbQuality.AddItem "Excellent"
    cmbQuality.AddItem "Good"
    cmbQuality.AddItem "Bad"
    cmbQuality.AddItem "Worst"
    
    cmbQuality.ListIndex = 0
    
    Call LoadBlankPreview 'Loading my info page
    txtFile.Text = ""
    
    'Reading Command$ if any
    If Command$ <> "" Then
        
        If Left(Command$, 1) = Chr(34) Then
            txtFile.Text = Mid(Command$, 2, Len(Command$) - 2)
        Else
            txtFile.Text = Command$
        End If
        
        Call SetWriteToIntelligence 'Fillup txtWriteTo textbox
        
    End If
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call cmdClose_Click
    
End Sub

Private Sub Label8_Click()
StartDoc "http://www.geocities.com/kyriakosnicola/ff_setup.zip"
End Sub

Private Sub optUseAdv_Click()

    Call UpdateFormSettings  'Updating form settings
    
End Sub

Private Sub optUseQucik_Click()
    
    Call UpdateFormSettings  'Updating form settings
    
End Sub


Sub UpdateFormSettings()
    
    Dim i
    
    For i = 0 To lblOPTUseQuick.Count - 1
        lblOPTUseQuick(i).Enabled = optUseQucik.Value
    Next i
    txtTitle.Enabled = optUseQucik.Value
    cmdBackColor.Enabled = optUseQucik.Value
    cmbFonts.Enabled = optUseQucik.Value
    chkCenter.Enabled = optUseQucik.Value
    chkBold.Enabled = optUseQucik.Value
    
    For i = 0 To lblOPTUseQuick.Count - 1
        lblOPTUseAdv(i).Enabled = optUseAdv.Value
    Next i
    txtFileA.Enabled = optUseAdv.Value
    txtFileB.Enabled = optUseAdv.Value
    cmdBrowseA.Enabled = optUseAdv.Value
    cmdBrowseB.Enabled = optUseAdv.Value
    
    
End Sub

Sub LoadProperties()
    
    On Error GoTo HELLERROR
    
    DoEvents
    
    picPic.Picture = LoadPicture()
    picPic.BackColor = cmdBackColor.BackColor
    
    picPic.Picture = LoadPicture(txtFile.Text)
    
    imgPIC.Stretch = True
    imgPIC.Height = 1365
    imgPIC.Width = 1770
    imgPIC.Picture = picPic.Picture
    If imgPIC.Picture.Height < 1365 Then imgPIC.Height = imgPIC.Picture.Height
    If imgPIC.Picture.Width < 1770 Then imgPIC.Width = imgPIC.Picture.Width
    
    lblDimensions.Caption = "Dimensions: " & picPic.ScaleHeight & "x" & picPic.ScaleWidth & " (pix)"
    
    picProperties.Visible = True

    Exit Sub

HELLERROR:  'Error occured means not a valid picture file
    
    picProperties.Visible = False
    
End Sub

Sub SetWriteToIntelligence()

    If IsFileExist(txtFile) = True Then
    
        txtWriteTo.Text = txtFile & "_asciipic.htm"
        
    End If
    
End Sub

Private Sub txtFile_Change()

    cmdMake.Enabled = txtFile.Text <> ""
    
    Call LoadProperties 'This will load picture properties
    'Don't worry, it will also handle if the file is not a picture

End Sub

Private Sub cmdMake_Click()

On Error GoTo HELLERROR

If cmdMake.Caption = "&Make" Then
    SSTab1.Enabled = False
    lblWriteTo.Enabled = False
    txtWriteTo.Enabled = False
    cmdWriteTo.Enabled = False
    cmdClose.Enabled = False
    cmdMake.Caption = "Cancel"
    DoEvents
Else
    SSTab1.Enabled = True
    lblWriteTo.Enabled = True
    txtWriteTo.Enabled = True
    cmdWriteTo.Enabled = True
    cmdClose.Enabled = True
    cmdMake.Caption = "&Make"
    DoEvents
    Exit Sub
End If

    LoadProperties
    
'First writing a temporary file then copy it to txtSaveTo specified path

    Dim F, F1, F2
    Dim SomeString As String
    
    F = FreeFile

    Me.Caption = "Naveed ASCII City - Writing file..."
    
 Open txtWriteTo For Output As #F
    
    Print #F, "<!--- ASCII Picture created by Naveed ASCII City on " & Now & " ---!>"
    Print #F, "<HTML>"
    
    If optUseQucik.Value = True Then
        
        Print #F, "<HEAD><TITLE>" & txtTitle.Text & "</TITLE></HEAD>"
        Print #F, "<BODY BGCOLOR=" & GetHexColor(cmdBackColor.BackColor) & ">"
            
            If chkCenter.Value = 1 Then _
        Print #F, "<CENTER>"
            If chkBold.Value = 1 Then _
        Print #F, "<B>"
        
        If cmbFonts.Text = "" Then cmbFonts.Text = "FIXDYS"
        If txtSize.Text = "" Then txtSize.Text = "1"
        
        Print #F, "<FONT FACE=" & Chr(34) & cmbFonts.Text & Chr(34) & " SIZE=" & Chr(34) & txtSize.Text & Chr(34) & ">"
        
        Print #F, "<BR>"
        
        Else
        
        F2 = FreeFile
        
        Open txtFileA.Text For Input As #F2
            While EOF(F2) = False
             Line Input #F2, SomeString
             Print #F, SomeString
            Wend
        Close #F2
        
    End If

    '----------------------------------------------+
    Call WriteASCIIPicture(F)  'wowowowowow  !!!   |
    '----------------------------------------------+

    If optUseQucik.Value = True Then
        
        Print #F, "<BR>"
            If chkCenter.Value = 1 Then _
        Print #F, "</CENTER>"
            If chkBold.Value = 1 Then _
        Print #F, "</B>"
        Print #F, "</BODY>"
            If chkCenter.Value = 1 Then _
        Print #F, "<CENTER>"
            If chkBold.Value = 1 Then _
        Print #F, "<B>"
                
        If cmbFonts.Text = "" Then cmbFonts.Text = "FIXDYS"
        If txtSize.Text = "" Then txtSize.Text = "1"
        
        Print #F, "<FONT FACE=" & Chr(34) & cmbFonts.Text & Chr(34) & " SIZE=" & Chr(34) & txtSize.Text & Chr(34) & ">"
        
        Print #F, "<BR>"
        
        Else
        
        F2 = FreeFile
        
        Open txtFileB.Text For Input As #F2
            While EOF(F2) = False
             Line Input #F2, SomeString
             Print #F, SomeString
            Wend
        Close #F2
        
    End If
    
    Print #F, "</HTML>"
    Print #F, "<!--- ASCII Picture created by Naveed ASCII City on " & Now & " ---!>"
    Print #F, "<!--- Feedback: neenojee@hotmail.com ---!>"
    
Close #F

    Me.Caption = "Naveed ASCII City"
    
    SSTab1.Enabled = True
    lblWriteTo.Enabled = True
    txtWriteTo.Enabled = True
    cmdWriteTo.Enabled = True
    cmdClose.Enabled = True
    cmdMake.Caption = "&Make"
    DoEvents
    
    WWW.Navigate txtWriteTo
    
    Exit Sub

HELLERROR:

    Me.Caption = "Naveed ASCII City"
    
    SSTab1.Enabled = True
    lblWriteTo.Enabled = True
    txtWriteTo.Enabled = True
    cmdWriteTo.Enabled = True
    cmdClose.Enabled = True
    cmdMake.Caption = "&Make"
    DoEvents
    
    MsgBox "An error is occured while performing the task." & vbCrLf & Err.Description & vbCrLf & vbCrLf, vbExclamation, Me.Caption
    
    Close
    
End Sub

Sub WriteASCIIPicture(ByVal F)
    
    On Error GoTo HELLERROR
    
    Dim P As Long, PC As Long
        'P=Old Color, PC = New Color
        'Both above variables use to compress HTML file size
        'by not writing <Font> tag again and again.
        
    Dim X, Y
    Dim BUF As String   'BUFFER to hold data and writing faster to disk.
    Dim hxclr As String 'hex of new color
    Dim ppt             'percentage
    Dim QualityNumber, XQuality, YQuality 'Handle qualities
    
    P = cmdBackColor.BackColor
    ASCIICHAR = txtASCIICHAR.Text
    POINTERCHAR = 1

    BUF = "<FONT COLOR=" & GetHexColor(P) & ">"
 
    QualityNumber = 100 - (cmbQuality.ListIndex)
    YQuality = picPic.ScaleHeight - Round(picPic.ScaleHeight / 100 * QualityNumber, 0) + 1
    XQuality = picPic.ScaleWidth - Round(picPic.ScaleWidth / 100 * QualityNumber, 0) + 1

    For Y = 0 To picPic.ScaleHeight - 1 Step YQuality
    For X = 0 To picPic.ScaleWidth - 1 Step XQuality

        PC = GetPixel(picPic.hdc, X, Y)
        
        If P = PC Then  'Old PC
            BUF = BUF & GetASCIICHAR
        Else
            hxclr = GetHexColor(PC)
            BUF = BUF & "</FONT>" & "<FONT COLOR=" & hxclr & ">" & GetASCIICHAR
    End If
        
    P = PC  'Necessary Settings
    
    If cmdMake.Caption <> "Cancel" Then Exit Sub
    DoEvents
    
    Next X
    
    BUF = BUF & "<BR>"  'Break line after completing one horizontal line

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' The folowing was the Big Bug       |
'But it is perfectly removed now !!! |
'===================================<
    If Len(BUF) > 5000 Then         '|
        Print #F, BUF               '|
        BUF = ""                    '|
    End If                          '|
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

    'Sending progress:
    ppt = Round(Y / picPic.ScaleHeight * 100, 2) & "%"
    Me.Caption = ppt & Space(8 - Len(ppt)) & " - Naveed ASCII City - Writing file..."
    
    
    Next Y
    
    BUF = BUF & "</FONT>"
    Print #F, BUF

Exit Sub

HELLERROR:

End Sub

Function GetASCIICHAR() As String
If Len(ASCIICHAR) = 1 Then
GetASCIICHAR = ASCIICHAR
Else
GetASCIICHAR = Mid(ASCIICHAR, POINTERCHAR, 1)

POINTERCHAR = POINTERCHAR + 1
If POINTERCHAR > Len(ASCIICHAR) Then POINTERCHAR = 1
End If
End Function

Private Sub txtFile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo HELLERROR
    
    txtFile.Text = Data.Files(1)
    
    Call SetWriteToIntelligence 'Fillup txtWriteTo textbox
    
    'Now see txtFile_change()
   
HELLERROR:

End Sub

Private Sub txtFileA_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo HELLERROR
    
    txtFileA.Text = Data.Files(1)

HELLERROR:

End Sub

Private Sub txtFileB_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo HELLERROR
    
    txtFileB.Text = Data.Files(1)

HELLERROR:

End Sub

Sub LoadBlankPreview()

    On Error Resume Next
    
    If IsFileExist(XBuildPath(App.Path, "PreviewInfo_.html")) = True Then WWW.Navigate XBuildPath(App.Path, "PreviewInfo_.html")

End Sub
