VERSION 5.00
Object = "*\AProjectStdSS.vbp"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Previev of 3nity SS ocx control...."
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin ProjectStdSS.stdSS stdSS1 
      Height          =   5340
      Left            =   75
      TabIndex        =   43
      Top             =   75
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   9419
      PlayedBorderColor=   33023
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   6675
      Top             =   2100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Spremeni Barvo"
   End
   Begin VB.PictureBox Picture13 
      Height          =   315
      Left            =   10200
      ScaleHeight     =   255
      ScaleWidth      =   480
      TabIndex        =   40
      Top             =   4950
      Width           =   540
   End
   Begin VB.PictureBox Picture12 
      Height          =   315
      Left            =   10200
      ScaleHeight     =   255
      ScaleWidth      =   480
      TabIndex        =   38
      Top             =   4575
      Width           =   540
   End
   Begin VB.PictureBox Picture11 
      Height          =   315
      Left            =   10200
      ScaleHeight     =   255
      ScaleWidth      =   480
      TabIndex        =   36
      Top             =   4200
      Width           =   540
   End
   Begin VB.PictureBox Picture10 
      Height          =   315
      Left            =   10200
      ScaleHeight     =   255
      ScaleWidth      =   480
      TabIndex        =   34
      Top             =   3825
      Width           =   540
   End
   Begin VB.PictureBox Picture9 
      Height          =   315
      Left            =   10200
      ScaleHeight     =   255
      ScaleWidth      =   480
      TabIndex        =   32
      Top             =   3450
      Width           =   540
   End
   Begin VB.PictureBox Picture8 
      Height          =   315
      Left            =   10200
      ScaleHeight     =   255
      ScaleWidth      =   480
      TabIndex        =   30
      Top             =   3075
      Width           =   540
   End
   Begin VB.PictureBox Picture7 
      Height          =   315
      Left            =   10200
      ScaleHeight     =   255
      ScaleWidth      =   480
      TabIndex        =   28
      Top             =   2700
      Width           =   540
   End
   Begin VB.PictureBox Picture6 
      Height          =   315
      Left            =   10200
      ScaleHeight     =   255
      ScaleWidth      =   480
      TabIndex        =   26
      Top             =   2325
      Width           =   540
   End
   Begin VB.PictureBox Picture5 
      Height          =   315
      Left            =   10200
      ScaleHeight     =   255
      ScaleWidth      =   480
      TabIndex        =   24
      Top             =   1950
      Width           =   540
   End
   Begin VB.PictureBox Picture4 
      Height          =   315
      Left            =   10200
      ScaleHeight     =   255
      ScaleWidth      =   480
      TabIndex        =   22
      Top             =   1575
      Width           =   540
   End
   Begin VB.PictureBox Picture3 
      Height          =   315
      Left            =   10200
      ScaleHeight     =   255
      ScaleWidth      =   480
      TabIndex        =   20
      Top             =   1200
      Width           =   540
   End
   Begin VB.PictureBox Picture2 
      Height          =   315
      Left            =   10200
      ScaleHeight     =   255
      ScaleWidth      =   480
      TabIndex        =   18
      Top             =   825
      Width           =   540
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   5925
      TabIndex        =   17
      Text            =   "New Title"
      Top             =   3525
      Width           =   2265
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Refresh Title"
      Height          =   315
      Left            =   4350
      TabIndex        =   16
      Top             =   3525
      Width           =   1440
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   6750
      TabIndex        =   15
      Text            =   "147"
      Top             =   3150
      Width           =   1440
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Refresh Time (only numbers!!!)"
      Height          =   315
      Left            =   4350
      TabIndex        =   14
      Top             =   3150
      Width           =   2265
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   6300
      TabIndex        =   13
      Text            =   "File2"
      Top             =   2700
      Width           =   1890
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Add File2 (to selected)"
      Height          =   315
      Left            =   4350
      TabIndex        =   12
      Top             =   2700
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   1035
      ItemData        =   "frmMain.frx":0000
      Left            =   4275
      List            =   "frmMain.frx":0002
      TabIndex        =   11
      Top             =   4350
      Width           =   3990
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Get Selected Data"
      Height          =   315
      Left            =   4275
      TabIndex        =   10
      Top             =   3975
      Width           =   3990
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   6750
      Picture         =   "frmMain.frx":0004
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   91
      TabIndex        =   0
      Top             =   1425
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Always Show Scroll Bar"
      Height          =   315
      Left            =   4350
      TabIndex        =   9
      Top             =   1425
      Value           =   1  'Checked
      Width           =   2640
   End
   Begin VB.CheckBox Check2 
      Caption         =   "AutoResize Scroller"
      Height          =   315
      Left            =   4350
      TabIndex        =   8
      Top             =   1725
      Value           =   1  'Checked
      Width           =   2640
   End
   Begin VB.CheckBox Check3 
      Caption         =   "ShowUP"
      Height          =   315
      Left            =   4350
      TabIndex        =   7
      Top             =   2025
      Value           =   1  'Checked
      Width           =   2640
   End
   Begin VB.CheckBox Check4 
      Caption         =   "ShowDown"
      Height          =   315
      Left            =   4350
      TabIndex        =   6
      Top             =   2325
      Value           =   1  'Checked
      Width           =   2640
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Remove All"
      Height          =   315
      Left            =   4275
      TabIndex        =   5
      Top             =   975
      Width           =   3990
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove Selected (Del)"
      Height          =   315
      Left            =   4275
      TabIndex        =   4
      Top             =   600
      Width           =   3990
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   7725
      TabIndex        =   3
      Text            =   "1:23"
      Top             =   75
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   5625
      TabIndex        =   2
      Text            =   "Item - Title"
      Top             =   75
      Width           =   1965
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Item"
      Height          =   315
      Left            =   4275
      TabIndex        =   1
      Top             =   75
      Width           =   1290
   End
   Begin VB.Label Label13 
      Caption         =   "Click on the PictureBox to change the spcified color:"
      Height          =   465
      Left            =   8775
      TabIndex        =   42
      Top             =   300
      Width           =   2115
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Played Time Text"
      Height          =   240
      Left            =   8625
      TabIndex        =   41
      Top             =   5025
      Width           =   1440
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Played Title Text:"
      Height          =   240
      Left            =   8625
      TabIndex        =   39
      Top             =   4650
      Width           =   1440
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Played Border:"
      Height          =   240
      Left            =   8625
      TabIndex        =   37
      Top             =   4275
      Width           =   1440
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Selected Time Text:"
      Height          =   240
      Left            =   8625
      TabIndex        =   35
      Top             =   3900
      Width           =   1440
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Selected Title Text:"
      Height          =   240
      Left            =   8550
      TabIndex        =   33
      Top             =   3525
      Width           =   1515
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Selected Time Back:"
      Height          =   240
      Left            =   8550
      TabIndex        =   31
      Top             =   3150
      Width           =   1515
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Selected Title Back:"
      Height          =   240
      Left            =   8625
      TabIndex        =   29
      Top             =   2775
      Width           =   1440
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Selected Border:"
      Height          =   240
      Left            =   8625
      TabIndex        =   27
      Top             =   2400
      Width           =   1440
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Time Text:"
      Height          =   240
      Left            =   8625
      TabIndex        =   25
      Top             =   2025
      Width           =   1440
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Title Text:"
      Height          =   240
      Left            =   8625
      TabIndex        =   23
      Top             =   1650
      Width           =   1440
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Time Background:"
      Height          =   240
      Left            =   8625
      TabIndex        =   21
      Top             =   1275
      Width           =   1440
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Background:"
      Height          =   240
      Left            =   8625
      TabIndex        =   19
      Top             =   900
      Width           =   1440
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is just a simple example... You can use it in
'your medaiplayers and so on...

Option Explicit

Private Sub Check1_Click()
'Shows the right scroller if true for all the time else just when needed

If Check1.Value = 1 Then
    Me.stdSS1.AlwaysShowScroller True
Else
    Me.stdSS1.AlwaysShowScroller False
End If


End Sub

Private Sub Check2_Click()
'autosizes the scroller depending on the lenght of the list if true

Dim CHK2 As Boolean
Dim CHK3 As Boolean
Dim CHK4 As Boolean

If Check2.Value = 1 Then
    CHK2 = True
Else
    CHK2 = False
End If

'shows btnUp in scroll bar if true

If Check3.Value = 1 Then
    CHK3 = True
Else
    CHK3 = False
End If

'shows btnDown in scroll bar if true

If Check4.Value = 1 Then
    CHK4 = True
Else
    CHK4 = False
End If

Me.stdSS1.SetScroller stdSS1.GUI_SSOzadjeŠirina, CHK3, CHK4, CHK2, stdSS1.GUI_SSGorŠirina, stdSS1.GUI_SSGorVišina, stdSS1.GUI_SSDolŠirina, stdSS1.GUI_SSDolVišina, stdSS1.GUI_SSDrsnikMiniVišina


End Sub

Private Sub Check3_Click()
Check2_Click

End Sub

Private Sub Check4_Click()
Check2_Click

End Sub

Private Sub Command1_Click()
'add entry (filename, filenam2, title,time,timeinseconds)
Me.stdSS1.AddItem "", "", Text1, Text2, 0

End Sub

Private Sub Command2_Click()
'removes selected entry
Me.stdSS1.Remove (Me.stdSS1.Selected)

End Sub

Private Sub Command3_Click()
'clears the list
Me.stdSS1.Clear

End Sub

Private Sub Command4_Click()
'shows data from selected string u can set any index as well
List1.Clear
If Me.stdSS1.Selected > 0 Then
    Me.stdSS1.GetData Me.stdSS1.Selected
    
    List1.AddItem "Title: " & Me.stdSS1.gTitle
    List1.AddItem "File1: " & Me.stdSS1.gFileName
    List1.AddItem "File2: " & Me.stdSS1.gFileName2
    List1.AddItem "Time (seconds): " & Me.stdSS1.gTimeInSeconds
    List1.AddItem "Time: " & Me.stdSS1.gTime
    
End If

End Sub

Private Sub Command5_Click()
'adds a second filename - i use it for subtitles in my player
If Me.stdSS1.Selected > 0 Then
    Me.stdSS1.AddFileName2 Text3, Me.stdSS1.Selected
    
End If

End Sub


Private Sub Command6_Click()
'updates time in selected entry
If Me.stdSS1.Selected > 0 Then
Dim cc As String
Dim mm As Long
Dim ss As Integer
Dim seconds As Long
On Error GoTo err
seconds = Text4

mm = seconds / 60

If seconds - mm + 60 < 0 Then mm = mm - 1

ss = seconds - mm * 60

If ss < 10 Then
    cc = mm & ":0" & ss
Else
    cc = mm & ":" & ss
End If

    Me.stdSS1.RefreshTime Text4, cc, Me.stdSS1.Selected
    
End If
Exit Sub
err:
Text4 = 0

End Sub


Private Sub Command7_Click()
'updates title in selcted entry
If Me.stdSS1.Selected > 0 Then
    Me.stdSS1.RefreshTitle Text5, Me.stdSS1.Selected
    
End If

End Sub


Private Sub Form_Load()
'picture property must be set befor calling gui
Set Me.stdSS1.PictureData = Picture1

'constants must be set before calling the gui

LoadGUIConstants
Me.stdSS1.AlwaysShowScroller True

Me.stdSS1.GUI
'this must be call each time the list resizes!!!
Me.stdSS1.SetScroller stdSS1.GUI_SSOzadjeŠirina, True, True, True, stdSS1.GUI_SSGorŠirina, stdSS1.GUI_SSGorVišina, stdSS1.GUI_SSDolŠirina, stdSS1.GUI_SSDolVišina, stdSS1.GUI_SSDrsnikMiniVišina


'This is not important!!!
'sets the colors on  the right of the form...
Picture2.BackColor = Me.stdSS1.BackgroundColor
Picture3.BackColor = Me.stdSS1.BackgroundTimeColor
Picture4.BackColor = Me.stdSS1.TitleTextColor
Picture5.BackColor = Me.stdSS1.TimeTextColor
Picture6.BackColor = Me.stdSS1.SelectedBorderColor
Picture7.BackColor = Me.stdSS1.SelectedTitleBackColor
Picture8.BackColor = Me.stdSS1.SelectedTimeBackColor
Picture9.BackColor = Me.stdSS1.SelectedTitleTextColor
Picture10.BackColor = Me.stdSS1.SelectedTimeTextColor
Picture11.BackColor = Me.stdSS1.PlayedBorderColor
Picture12.BackColor = Me.stdSS1.PlayedTitleTextColor
Picture13.BackColor = Me.stdSS1.PlayedTimeTextColor



End Sub

Private Sub LoadGUIConstants()
'constants for the gui

stdSS1.GUI_SSOzadjeŠirina = 13
stdSS1.GUI_SSOzadjeVišina = 40
stdSS1.GUI_SSOzadjeX = 26
stdSS1.GUI_SSOzadjeY = 0
stdSS1.GUI_SSOzadjeXD = 26
stdSS1.GUI_SSOzadjeYD = 0

stdSS1.GUI_SSDrsnikOzadjeŠirina = 13
stdSS1.GUI_SSDrsnikOzadjeVišina = 40
stdSS1.GUI_SSDrsnikOzadjeX = 65
stdSS1.GUI_SSDrsnikOzadjeY = 0
stdSS1.GUI_SSDrsnikOzadjeXD = 78
stdSS1.GUI_SSDrsnikOzadjeYD = 0

stdSS1.GUI_SSDrsnikGorŠirina = 13
stdSS1.GUI_SSDrsnikGorVišina = 21
stdSS1.GUI_SSDrsnikGorX = 39
stdSS1.GUI_SSDrsnikGorY = 0
stdSS1.GUI_SSDrsnikGorXD = 52
stdSS1.GUI_SSDrsnikGorYD = 0

stdSS1.GUI_SSDrsnikDolŠirina = 13
stdSS1.GUI_SSDrsnikDolVišina = 20
stdSS1.GUI_SSDrsnikDolX = 39
stdSS1.GUI_SSDrsnikDolY = 21
stdSS1.GUI_SSDrsnikDolXD = 52
stdSS1.GUI_SSDrsnikDolYD = 21

stdSS1.GUI_SSGorŠirina = 13
stdSS1.GUI_SSGorVišina = 17
stdSS1.GUI_SSGorX = 0
stdSS1.GUI_SSGorY = 0
stdSS1.GUI_SSGorXD = 13
stdSS1.GUI_SSGorYD = 0

stdSS1.GUI_SSDolŠirina = 13
stdSS1.GUI_SSDolVišina = 17
stdSS1.GUI_SSDolX = 0
stdSS1.GUI_SSDolY = 17
stdSS1.GUI_SSDolXD = 13
stdSS1.GUI_SSDolYD = 17

stdSS1.GUI_SSDrsnikMiniVišina = 41
stdSS1.GUI_SSDrsnikScale = False
stdSS1.GUI_SSVednoKaži = True


End Sub

Private Sub Picture10_Click()
'Change Selected Time Text color
CD.ShowColor
Picture10.BackColor = CD.Color
Me.stdSS1.SelectedTimeTextColor = CD.Color

End Sub

Private Sub Picture11_Click()
'Change Played Border color
CD.ShowColor
Picture11.BackColor = CD.Color
Me.stdSS1.PlayedBorderColor = CD.Color

End Sub

Private Sub Picture12_Click()
'Change Played Title Text color
CD.ShowColor
Picture12.BackColor = CD.Color
Me.stdSS1.PlayedTitleTextColor = CD.Color

End Sub

Private Sub Picture13_Click()
'Change Played Time Text color
CD.ShowColor
Picture13.BackColor = CD.Color
Me.stdSS1.PlayedTimeTextColor = CD.Color

End Sub

Private Sub Picture2_Click()
'Change Background color
CD.ShowColor
Picture2.BackColor = CD.Color
Me.stdSS1.BackgroundColor = CD.Color

End Sub

Private Sub Picture3_Click()
'Change Time Background color
CD.ShowColor
Picture3.BackColor = CD.Color
Me.stdSS1.BackgroundTimeColor = CD.Color

End Sub

Private Sub Picture4_Click()
'Change Title Text color
CD.ShowColor
Picture4.BackColor = CD.Color
Me.stdSS1.TitleTextColor = CD.Color

End Sub

Private Sub Picture5_Click()
'Change Time Text color
CD.ShowColor
Picture5.BackColor = CD.Color
Me.stdSS1.TimeTextColor = CD.Color

End Sub

Private Sub Picture6_Click()
'Change SelectedBorder color
CD.ShowColor
Picture6.BackColor = CD.Color
Me.stdSS1.SelectedBorderColor = CD.Color

End Sub

Private Sub Picture7_Click()
'Change Selected Title Back color
CD.ShowColor
Picture7.BackColor = CD.Color
Me.stdSS1.SelectedTitleBackColor = CD.Color

End Sub

Private Sub Picture8_Click()
'Change Selected Time Back color
CD.ShowColor
Picture8.BackColor = CD.Color
Me.stdSS1.SelectedTimeBackColor = CD.Color

End Sub

Private Sub Picture9_Click()
'Change Selected Title Text color
CD.ShowColor
Picture9.BackColor = CD.Color
Me.stdSS1.SelectedTitleTextColor = CD.Color

End Sub

