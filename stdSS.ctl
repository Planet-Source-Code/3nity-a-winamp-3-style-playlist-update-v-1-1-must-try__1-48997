VERSION 5.00
Begin VB.UserControl stdSS 
   BackColor       =   &H00FEF5E9&
   ClientHeight    =   7050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7755
   ControlContainer=   -1  'True
   DataBindingBehavior=   1  'vbSimpleBound
   KeyPreview      =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   470
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   517
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picScroll 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5790
      Left            =   7425
      ScaleHeight     =   386
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   13
      Top             =   375
      Visible         =   0   'False
      Width           =   240
      Begin VB.PictureBox PicDol 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   0
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   16
         Top             =   4725
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox PicPoljeDrsnika 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF80FF&
         BorderStyle     =   0  'None
         Height          =   3315
         Left            =   0
         ScaleHeight     =   221
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   15
         Top             =   450
         Width           =   240
         Begin VB.PictureBox PicDrsnik 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C000C0&
            BorderStyle     =   0  'None
            Height          =   1440
            Left            =   0
            ScaleHeight     =   96
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   16
            TabIndex        =   17
            Top             =   0
            Width           =   240
            Begin VB.PictureBox PicDrsnikD 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00C000C0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   1440
               Left            =   0
               ScaleHeight     =   96
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   16
               TabIndex        =   18
               Top             =   600
               Width           =   240
               Begin VB.PictureBox PicDrsnikDDol 
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00FF00FF&
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   0
                  ScaleHeight     =   16
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   16
                  TabIndex        =   20
                  Top             =   900
                  Visible         =   0   'False
                  Width           =   240
               End
               Begin VB.PictureBox PicDrsnikDGor 
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00FF00FF&
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   0
                  ScaleHeight     =   16
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   16
                  TabIndex        =   19
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   240
               End
            End
            Begin VB.PictureBox PicDrsnikDol 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00FF00FF&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   240
               Left            =   0
               ScaleHeight     =   16
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   16
               TabIndex        =   22
               Top             =   900
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.PictureBox PicDrsnikGor 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00FF00FF&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   240
               Left            =   0
               ScaleHeight     =   16
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   16
               TabIndex        =   21
               Top             =   0
               Visible         =   0   'False
               Width           =   240
            End
         End
      End
      Begin VB.PictureBox PicGor 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   0
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.PictureBox PicOzadje 
      BackColor       =   &H00FEF5E9&
      BorderStyle     =   0  'None
      FillColor       =   &H00FEF5E9&
      Height          =   5640
      Left            =   0
      ScaleHeight     =   376
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   441
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.PictureBox picPremik 
         BackColor       =   &H007E511F&
         BorderStyle     =   0  'None
         Height          =   45
         Left            =   300
         ScaleHeight     =   3
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   221
         TabIndex        =   11
         Top             =   1950
         Visible         =   0   'False
         Width           =   3315
      End
      Begin VB.Line �rtaRob 
         BorderColor     =   &H00A76D46&
         Index           =   0
         Visible         =   0   'False
         X1              =   160
         X2              =   160
         Y1              =   160
         Y2              =   190
      End
      Begin VB.Line �rta�as 
         BorderColor     =   &H00A76D46&
         Index           =   0
         Visible         =   0   'False
         X1              =   225
         X2              =   400
         Y1              =   260
         Y2              =   260
      End
      Begin VB.Line �rta 
         BorderColor     =   &H00C89248&
         Index           =   0
         Visible         =   0   'False
         X1              =   110
         X2              =   285
         Y1              =   250
         Y2              =   250
      End
      Begin VB.Label lblSpot 
         AutoSize        =   -1  'True
         BackColor       =   &H00FEF5E9&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A76D46&
         Height          =   210
         Index           =   0
         Left            =   5475
         TabIndex        =   12
         Top             =   -1725
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape shp�asB 
         BorderColor     =   &H000080FF&
         FillColor       =   &H00A76D46&
         Height          =   540
         Left            =   3450
         Top             =   2700
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lbl�tevilkaB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lbl�tevilkaB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   210
         Left            =   3825
         TabIndex        =   10
         Top             =   2250
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label lblImeB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblImeB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   210
         Left            =   3825
         TabIndex        =   9
         Top             =   2475
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lbl�asB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2B"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   210
         Left            =   3825
         TabIndex        =   8
         Top             =   2700
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label lbl�tevilkaA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lbl�tevilkaA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FAE2B7&
         Height          =   210
         Left            =   4950
         TabIndex        =   7
         Top             =   2250
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label lblImeA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblImeA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FAE2B7&
         Height          =   210
         Left            =   4950
         TabIndex        =   6
         Top             =   2475
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lbl�asA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FEF5E9&
         Height          =   210
         Left            =   4950
         TabIndex        =   5
         Top             =   2700
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Shape shp�asA 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00A76D46&
         FillStyle       =   0  'Solid
         Height          =   540
         Left            =   4800
         Top             =   2700
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lbl�as 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A76D46&
         Height          =   210
         Index           =   0
         Left            =   4800
         TabIndex        =   3
         Top             =   1950
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblIme 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblIme"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A76D46&
         Height          =   210
         Index           =   0
         Left            =   4800
         TabIndex        =   2
         Top             =   1725
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label lbl�tevilka 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lbl�tevilka"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A76D46&
         Height          =   210
         Index           =   0
         Left            =   4800
         TabIndex        =   1
         Top             =   1500
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label lblIzbor 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   0
         Left            =   4650
         OLEDropMode     =   1  'Manual
         TabIndex        =   4
         Top             =   2250
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Shape shpOzadje 
         BorderColor     =   &H008A6544&
         FillColor       =   &H00C89248&
         FillStyle       =   0  'Solid
         Height          =   540
         Index           =   0
         Left            =   4800
         Top             =   2175
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Shape shp�as 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FAE2B7&
         FillStyle       =   0  'Solid
         Height          =   540
         Index           =   0
         Left            =   4725
         Top             =   1950
         Visible         =   0   'False
         Width           =   1215
      End
   End
End
Attribute VB_Name = "stdSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'had no time to coment
'just translated most of the public subs
'really had no time...
'sorry, hope it helps
'will translate it in the future if requested

Option Explicit
Dim MouseY As Integer

Dim xYz As Integer

Public GU As Boolean

Public VelikostVrstice As Integer
Public ListCount As Integer
Public Selected As Integer
Public Playing As Integer
Public Polo�ajZgoraj As Long
Public �irina�asa As Integer
Public AShowScroller As Boolean
Public bDrsnikMiniVi�ina As Integer
Public DrsnikScale As Boolean
Public MultiSelect As Boolean
Public Na�inMultiSelect As Integer
Public prvaMultiSelect As Integer
Public Za�etMS As Boolean
Public NePredvajaj As Boolean
Public Skupen�asSekund As Long

Public gFileName As String
Public gFileName2 As String
Public gTitle As String
Public gTime As String
Public gTimeInSeconds As Long

Public PictureData As PictureBox

Public Event Play(FileName As String)
Public Event RePlay()
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ShowMenu()
Public Event DurationChange(NewDuration As Long)

Public GUI_SSOzadje�irina As Integer
Public GUI_SSOzadjeVi�ina As Integer
Public GUI_SSOzadjeX As Integer
Public GUI_SSOzadjeY As Integer
Public GUI_SSOzadjeXD As Integer
Public GUI_SSOzadjeYD As Integer

Public GUI_SSDrsnikOzadje�irina As Integer
Public GUI_SSDrsnikOzadjeVi�ina As Integer
Public GUI_SSDrsnikOzadjeX As Integer
Public GUI_SSDrsnikOzadjeY As Integer
Public GUI_SSDrsnikOzadjeXD As Integer
Public GUI_SSDrsnikOzadjeYD As Integer

Public GUI_SSDrsnikGor�irina As Integer
Public GUI_SSDrsnikGorVi�ina As Integer
Public GUI_SSDrsnikGorX As Integer
Public GUI_SSDrsnikGorY As Integer
Public GUI_SSDrsnikGorXD As Integer
Public GUI_SSDrsnikGorYD As Integer

Public GUI_SSDrsnikDol�irina As Integer
Public GUI_SSDrsnikDolVi�ina As Integer
Public GUI_SSDrsnikDolX As Integer
Public GUI_SSDrsnikDolY As Integer
Public GUI_SSDrsnikDolXD As Integer
Public GUI_SSDrsnikDolYD As Integer

Public GUI_SSGor�irina As Integer
Public GUI_SSGorVi�ina As Integer
Public GUI_SSGorX As Integer
Public GUI_SSGorY As Integer
Public GUI_SSGorXD As Integer
Public GUI_SSGorYD As Integer

Public GUI_SSDol�irina As Integer
Public GUI_SSDolVi�ina As Integer
Public GUI_SSDolX As Integer
Public GUI_SSDolY As Integer
Public GUI_SSDolXD As Integer
Public GUI_SSDolYD As Integer

Public GUI_SSDrsnikMiniVi�ina As Integer
Public GUI_SSDrsnikScale As Boolean
Public GUI_SSVednoKa�i As Boolean

Public Sub AddItem(FileName As String, FileName2 As String, Title As String, Time As String, TimeInSeconds As Long)

Dim cc As Integer
cc = UserControl.lblIme.Count
ListCount = cc

Load UserControl.lbl�tevilka(cc)
Load UserControl.lbl�as(cc)
Load UserControl.lblIme(cc)
Load UserControl.shpOzadje(cc)
Load UserControl.shp�as(cc)
Load UserControl.lblIzbor(cc)
Load UserControl.lblSpot(cc)
Load UserControl.�rta(cc)
Load UserControl.�rta�as(cc)
Load UserControl.�rtaRob(cc)

Skupen�asSekund = Skupen�asSekund + TimeInSeconds
RaiseEvent DurationChange(Skupen�asSekund)

lbl�tevilka(cc).Caption = cc & ". "
lbl�tevilka(cc).Tag = FileName2
lbl�tevilka(cc).Left = 2
lbl�tevilka(cc).Top = (cc - 1) * (lbl�tevilka(cc).Height + 2) + 1

lblIme(cc).Caption = Title
lblIme(cc).Tag = FileName
lblIme(cc).Left = lbl�tevilka(cc).Left + lbl�tevilka(cc).Width
lblIme(cc).Top = (cc - 1) * (lbl�tevilka(cc).Height + 2) + 1

lbl�as(cc).Caption = Time
lbl�as(cc).Width = �irina�asa - 2
lbl�as(cc).Tag = TimeInSeconds
lbl�as(cc).Left = PicOzadje.Width - �irina�asa + 1
lbl�as(cc).Top = (cc - 1) * (lbl�tevilka(cc).Height + 2) + 1

lblSpot(cc).Left = lbl�as(cc).Left - lblSpot(cc).Width
lblSpot(cc).Top = (cc - 1) * (lbl�tevilka(cc).Height + 2) + 1

shpOzadje(cc).Left = 0
shpOzadje(cc).Top = (cc - 1) * (lbl�tevilka(cc).Height + 2)
shpOzadje(cc).Width = PicOzadje.Width
shpOzadje(cc).Height = lblIme(cc).Height + 3

shp�as(cc).Left = lbl�as(cc).Left - 1
shp�as(cc).Top = (cc - 1) * (lbl�tevilka(cc).Height + 2)
shp�as(cc).Width = PicOzadje.Width - shp�as(cc).Left + 1
shp�as(cc).Height = lblIme(cc).Height + 3

�rta(cc).x1 = 1
�rta(cc).x2 = shp�as(cc).Left
�rta(cc).Y1 = shpOzadje(cc).Top + shpOzadje(cc).Height - 1
�rta(cc).Y2 = �rta(cc).Y1

�rta�as(cc).x1 = shp�as(cc).Left
�rta�as(cc).x2 = PicOzadje.Width
�rta�as(cc).Y1 = shpOzadje(cc).Top + shpOzadje(cc).Height - 1
�rta�as(cc).Y2 = �rta(cc).Y1

�rtaRob(cc).x1 = PicOzadje.Width - 1
�rtaRob(cc).x2 = PicOzadje.Width - 1
�rtaRob(cc).Y1 = shpOzadje(cc).Top
�rtaRob(cc).Y2 = shpOzadje(cc).Top + shpOzadje(cc).Height - 1
�rtaRob(cc).Visible = False

lblIzbor(cc).Left = 0
lblIzbor(cc).Top = (cc - 1) * (lbl�tevilka(cc).Height + 2)
lblIzbor(cc).Width = PicOzadje.Width
lblIzbor(cc).Height = lblIme(cc).Height + 2

lblIme(cc).ZOrder 1
shpOzadje(cc).ZOrder 1

lblIme(cc).Visible = True
lbl�tevilka(cc).Visible = True
lbl�as(cc).Visible = True
shp�as(cc).Visible = True
lblIzbor(cc).Visible = True
lblIzbor(cc).ZOrder

If lblIme(cc).Width > lbl�as(cc).Left - lblIme(cc).Left Then
    lblSpot(cc).Visible = True
End If

PicOzadje.Height = (cc) * (shpOzadje(cc).Height - 1) + 1

If PicOzadje.Visible = False Then PicOzadje.Visible = True
LegaDrsnika

End Sub

Public Sub GUI()
Dim iCnt As Integer
On Error Resume Next

PicPoljeDrsnika.Height = 1700
PicPoljeDrsnika.Width = GUI_SSOzadje�irina
PicPoljeDrsnika.AutoRedraw = True
For iCnt = 0 To Int(1700 / GUI_SSOzadjeVi�ina)
    BitBlt PicPoljeDrsnika.hDC, 0, iCnt * GUI_SSOzadjeVi�ina, GUI_SSOzadje�irina, GUI_SSOzadjeVi�ina, PictureData.hDC, GUI_SSOzadjeX, GUI_SSOzadjeY, SRCCOPY
    PicPoljeDrsnika.Refresh
Next iCnt

PicDrsnik.Height = 1700
PicDrsnik.Width = GUI_SSDrsnikOzadje�irina
PicDrsnikD.Height = 1700
PicDrsnikD.Width = GUI_SSDrsnikOzadje�irina

For iCnt = 0 To Int(1700 / GUI_SSDrsnikOzadjeVi�ina)
    BitBlt PicDrsnik.hDC, 0, iCnt * GUI_SSDrsnikOzadjeVi�ina, GUI_SSDrsnikOzadje�irina, GUI_SSDrsnikOzadjeVi�ina, PictureData.hDC, GUI_SSDrsnikOzadjeX, GUI_SSDrsnikOzadjeY, SRCCOPY
    PicDrsnik.Refresh
    BitBlt PicDrsnikD.hDC, 0, iCnt * GUI_SSDrsnikOzadjeVi�ina, GUI_SSDrsnikOzadje�irina, GUI_SSDrsnikOzadjeVi�ina, PictureData.hDC, GUI_SSDrsnikOzadjeXD, GUI_SSDrsnikOzadjeYD, SRCCOPY
    PicDrsnikD.Refresh
    
Next iCnt
PicDrsnikD.Visible = False
PicDrsnikD.Top = 0
PicDrsnikD.Left = 0

PicDrsnikGor.Top = 0
PicDrsnikGor.Height = GUI_SSDrsnikGorVi�ina
PicDrsnikGor.Width = GUI_SSDrsnikGor�irina
PicDrsnikGor.Left = 0
BitBlt PicDrsnikGor.hDC, 0, 0, GUI_SSDrsnikGor�irina, GUI_SSDrsnikGorVi�ina, PictureData.hDC, GUI_SSDrsnikGorX, GUI_SSDrsnikGorY, SRCCOPY
PicDrsnikGor.Refresh
PicDrsnikGor.Visible = True

PicDrsnikDGor.Top = 0
PicDrsnikDGor.Height = GUI_SSDrsnikGorVi�ina
PicDrsnikDGor.Width = GUI_SSDrsnikGor�irina
PicDrsnikDGor.Left = 0
BitBlt PicDrsnikDGor.hDC, 0, 0, GUI_SSDrsnikGor�irina, GUI_SSDrsnikGorVi�ina, PictureData.hDC, GUI_SSDrsnikGorXD, GUI_SSDrsnikGorYD, SRCCOPY
PicDrsnikDGor.Refresh
PicDrsnikDGor.Visible = True

PicDrsnikDol.Height = GUI_SSDrsnikDolVi�ina
PicDrsnikDol.Width = GUI_SSDrsnikDol�irina
PicDrsnikDol.Left = 0
BitBlt PicDrsnikDol.hDC, 0, 0, GUI_SSDrsnikDol�irina, GUI_SSDrsnikDolVi�ina, PictureData.hDC, GUI_SSDrsnikDolX, GUI_SSDrsnikDolY, SRCCOPY
PicDrsnikDol.Refresh
PicDrsnikDol.Visible = True

PicDrsnikDDol.Height = GUI_SSDrsnikDolVi�ina
PicDrsnikDDol.Width = GUI_SSDrsnikDol�irina
PicDrsnikDDol.Left = 0
BitBlt PicDrsnikDDol.hDC, 0, 0, GUI_SSDrsnikDol�irina, GUI_SSDrsnikDolVi�ina, PictureData.hDC, GUI_SSDrsnikDolXD, GUI_SSDrsnikDolYD, SRCCOPY
PicDrsnikDDol.Refresh
PicDrsnikDDol.Visible = True

PicDol.Height = GUI_SSDolVi�ina
PicDol.Width = GUI_SSDol�irina

BitBlt PicDol.hDC, 0, 0, GUI_SSDol�irina, GUI_SSDolVi�ina, PictureData.hDC, GUI_SSDolX, GUI_SSDolY, SRCCOPY
PicDol.Refresh

PicGor.Height = GUI_SSGorVi�ina
PicGor.Width = GUI_SSGor�irina

BitBlt PicGor.hDC, 0, 0, GUI_SSGor�irina, GUI_SSGorVi�ina, PictureData.hDC, GUI_SSGorX, GUI_SSGorY, SRCCOPY
PicGor.Refresh
End Sub

Public Sub Clear()

On Error Resume Next
Dim iCnt As Integer
Skupen�asSekund = 0
RaiseEvent DurationChange(Skupen�asSekund)

For iCnt = 1 To lblIme.Count - 1
    Unload UserControl.lbl�tevilka(iCnt)
    Unload UserControl.lbl�as(iCnt)
    Unload UserControl.lblIme(iCnt)
    Unload UserControl.shpOzadje(iCnt)
    Unload UserControl.shp�as(iCnt)
    Unload UserControl.lblIzbor(iCnt)
    Unload UserControl.lblSpot(iCnt)
    Unload UserControl.�rta(iCnt)
    Unload UserControl.�rta�as(iCnt)
    Unload UserControl.�rtaRob(iCnt)
Next iCnt

If PicOzadje.Visible = True Then PicOzadje.Visible = False
PicOzadje.Height = 0
PicOzadje.Top = 0

If AShowScroller = False Then picScroll.Visible = False

shp�asB.Visible = False
ListCount = 0
Selected = 0
Playing = 0
LegaDrsnika

gFileName = ""
gFileName2 = ""
gTitle = ""
gTime = 0
gTimeInSeconds = 0

End Sub

Public Sub Remove(Index As Integer)

If Index > 0 Then

Skupen�asSekund = Skupen�asSekund - lbl�as(Index).Tag
RaiseEvent DurationChange(Skupen�asSekund)

    Me.NePredvajaj = True
    ListCount = ListCount - 1
    
    If ListCount > 0 Then
        PicOzadje.Height = (ListCount) * (shpOzadje(2).Height - 1) + 1
    Else
        PicOzadje.Visible = False
    End If
    
    shpOzadje(Selected).Visible = False
    
    shp�as(Selected).Left = lbl�as(Selected).Left - 1
    shp�as(Selected).Top = (Selected - 1) * (lbl�tevilka(Selected).Height + 2)
    shp�as(Selected).Width = PicOzadje.Width - shp�as(Selected).Left + 1
    shp�as(Selected).Height = lblIme(Selected).Height + 3
    shp�as(Selected).FillColor = shp�as(0).FillColor

    lblIme(Selected).Font = lblIme(0).Font
    lblIme(Selected).FontBold = lblIme(0).FontBold
    lblIme(Selected).FontItalic = lblIme(0).FontItalic
    lblIme(Selected).ForeColor = lblIme(0).ForeColor
    
    lbl�as(Selected).Font = lbl�as(0).Font
    lbl�as(Selected).FontBold = lbl�as(0).FontBold
    lbl�as(Selected).FontItalic = lbl�as(0).FontItalic
    lbl�as(Selected).ForeColor = lbl�as(0).ForeColor
    
    lbl�tevilka(Selected).Font = lbl�tevilka(0).Font
    lbl�tevilka(Selected).FontBold = lbl�tevilka(0).FontBold
    lbl�tevilka(Selected).FontItalic = lbl�tevilka(0).FontItalic
    lbl�tevilka(Selected).ForeColor = lbl�tevilka(0).ForeColor
    
    lblSpot(Selected).Font = lblIme(0).Font
    lblSpot(Selected).FontBold = lblIme(0).FontBold
    lblSpot(Selected).FontItalic = lblIme(0).FontItalic
    lblSpot(Selected).ForeColor = lblIme(0).ForeColor
    lblSpot(Selected).BackColor = PicOzadje.BackColor
    
    If Playing > Index Then
        Play Playing - 1
    ElseIf Playing = Index Then
        Playing = 0
        shp�asB.Visible = False
    End If
    
    
    lblIme(Selected).Left = lbl�tevilka(Selected).Width + lbl�tevilka(Selected).Left

    Dim iCnt As Integer
    
    For iCnt = Index To lblIme.Count - 2
        lbl�tevilka(iCnt).Tag = lbl�tevilka(iCnt + 1).Tag
        lblIme(iCnt).Caption = lblIme(iCnt + 1).Caption
        lblIme(iCnt).Tag = lblIme(iCnt + 1).Tag
        lbl�as(iCnt).Caption = lbl�as(iCnt + 1).Caption
        lbl�as(iCnt).Tag = lbl�as(iCnt + 1).Tag
        lblIzbor(iCnt).Tag = lblIzbor(iCnt + 1).Tag
        
        If lblIme(iCnt).Width > lbl�as(iCnt).Left - lblIme(iCnt).Left Then
            If Not lblSpot(iCnt).Left = lbl�as(iCnt).Left - lblSpot(iCnt).Width Then lblSpot(iCnt).Left = lbl�as(iCnt).Left - lblSpot(iCnt).Width
            lblSpot(iCnt).Visible = True
        Else
            lblSpot(iCnt).Visible = False
        End If
        
        
        
    Next iCnt
    
    iCnt = lblIme.Count - 1
    
    Unload UserControl.lbl�tevilka(iCnt)
    Unload UserControl.lbl�as(iCnt)
    Unload UserControl.lblIme(iCnt)
    Unload UserControl.shpOzadje(iCnt)
    Unload UserControl.shp�as(iCnt)
    Unload UserControl.lblIzbor(iCnt)
    Unload UserControl.lblSpot(iCnt)
    Unload UserControl.�rta(iCnt)
    Unload UserControl.�rta�as(iCnt)
    Unload UserControl.�rtaRob(iCnt)
    
    Selected = 0
    NePredvajaj = False
    If PicOzadje.Height > UserControl.Height / Screen.TwipsPerPixelY Then
        If PicOzadje.Top < UserControl.Height / Screen.TwipsPerPixelY - PicOzadje.Height Then
            PicOzadje.Top = UserControl.Height / Screen.TwipsPerPixelY - PicOzadje.Height
        End If
    Else
        If PicOzadje.Top < 0 Then PicOzadje.Top = 0
        
    End If

    LegaDrsnika
    
End If

End Sub

Public Sub MultiIzbris()
Dim iCnt As Integer
Dim a As Integer
Dim B As Integer
a = lblIme.Count

�rte True
Remove Selected
For iCnt = 1 To a - 1
    B = a - iCnt - 1
    
    If lblIzbor(B).Tag = "I" Then
        Selected = B
        Remove B
    End If
Next iCnt

End Sub

Public Sub RefreshTitle(Title As String, Index As Integer)
    lblIme(Index).Caption = Title
    
    If lblIme(Index).Width > lbl�as(Index).Left - lblIme(Index).Left Then
        If Not lblSpot(Index).Visible = True Then lblSpot(Index).Visible = True
    Else
        If Not lblSpot(Index).Visible = False Then lblSpot(Index).Visible = False
    End If
    lblSpot(Index).Refresh
    
End Sub

Public Sub Poravnaj()
On Error Resume Next
Dim iCnt As Integer
For iCnt = 1 To lblIme.Count
    If Not lbl�as(iCnt).Left = PicOzadje.Width - �irina�asa + 1 Then lbl�as(iCnt).Left = PicOzadje.Width - �irina�asa + 1
    If Not lblSpot(iCnt).Left = lbl�as(iCnt).Left - lblSpot(iCnt).Width Then lblSpot(iCnt).Left = lbl�as(iCnt).Left - lblSpot(iCnt).Width
    If lblIme(iCnt).Width > lbl�as(iCnt).Left - lblIme(iCnt).Left Then
        If Not lblSpot(iCnt).Visible = True Then lblSpot(iCnt).Visible = True
    Else
        If Not lblSpot(iCnt).Visible = False Then lblSpot(iCnt).Visible = False
    End If
    
    If Not shpOzadje(iCnt).Width = PicOzadje.Width Then shpOzadje(iCnt).Width = PicOzadje.Width
    If Not shp�as(iCnt).Left = lbl�as(iCnt).Left - 1 Then shp�as(iCnt).Left = lbl�as(iCnt).Left - 1
    If Not lblIzbor(iCnt).Width = PicOzadje.Width Then lblIzbor(iCnt).Width = PicOzadje.Width
    If Not shp�asB.Width = PicOzadje.Width Then shp�asB.Width = PicOzadje.Width
    
Next iCnt

End Sub

Public Sub SetScroller(sWidth As Integer, bUP As Boolean, bDown As Boolean, bScaleScroller As Boolean, Optional bUPWidth As Integer, Optional bUPHeight As Integer, Optional bDownWidth As Integer, Optional bDownHeight As Integer, Optional ScrollerHeight As Integer)
picScroll.Width = sWidth
Dim a1 As Integer
Dim b1 As Integer

If Not picScroll.Left = UserControl.Width / Screen.TwipsPerPixelX - picScroll.Width Then picScroll.Left = UserControl.Width / Screen.TwipsPerPixelX - picScroll.Width
picScroll.Height = UserControl.Height / Screen.TwipsPerPixelY
If Not picScroll.Top = 0 Then picScroll.Top = 0

If bUP = True Then
    PicGor.Visible = True
    PicGor.Width = bUPWidth
    PicGor.Height = bUPHeight
    PicGor.Top = 0
    PicGor.Left = 0
    a1 = PicGor.Height
    b1 = PicGor.Height
Else
    PicGor.Visible = False
    a1 = 0
End If

If bDown = True Then
    PicDol.Visible = True
    PicDol.Width = bDownWidth
    PicDol.Height = bDownHeight
    PicDol.Left = 0
    PicDol.Top = picScroll.Height - PicDol.Height
    b1 = b1 + a1
    
Else
    b1 = a1
    PicDol.Visible = False
End If

PicPoljeDrsnika.Height = picScroll.Height - b1
PicPoljeDrsnika.Top = a1
PicPoljeDrsnika.Width = sWidth
PicPoljeDrsnika.Left = 0
PicDrsnik.Width = sWidth
PicDrsnikD.Width = sWidth

bDrsnikMiniVi�ina = ScrollerHeight
DrsnikScale = bScaleScroller

LegaDrsnika

End Sub

Public Sub LegaDrsnika()
On Error Resume Next
Dim c1 As Long

c1 = PicPoljeDrsnika.Height * (UserControl.Height / Screen.TwipsPerPixelY) / PicOzadje.Height

If DrsnikScale = True Then
    If c1 >= PicPoljeDrsnika.Height Then
        PicDrsnik.Visible = False
    ElseIf c1 < bDrsnikMiniVi�ina Then
        PicDrsnik.Height = bDrsnikMiniVi�ina
        PicDrsnik.Visible = True
    Else
        PicDrsnik.Height = c1
        PicDrsnik.Visible = True
    End If
    

Else
    If c1 >= PicPoljeDrsnika.Height Then
        PicDrsnik.Visible = False
    Else
        PicDrsnik.Visible = True
    End If
    
    PicDrsnik.Height = bDrsnikMiniVi�ina

End If

If PicDrsnik.Visible = True Then
    If (UserControl.Height / Screen.TwipsPerPixelY - PicOzadje.Height) = 0 Then

    Else
        PicDrsnik.Top = ((PicOzadje.Top) * (PicPoljeDrsnika.Height - PicDrsnik.Height) / (UserControl.Height / Screen.TwipsPerPixelY - PicOzadje.Height))
    End If
End If

End Sub

Public Sub AlwaysShowScroller(vrednost As Boolean)
AShowScroller = vrednost

If PicOzadje.Height > UserControl.Height / Screen.TwipsPerPixelY Or AShowScroller = True Then
    If Not PicOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX - picScroll.Width Then PicOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX - picScroll.Width
    If Not picScroll.Left = PicOzadje.Width Then picScroll.Left = PicOzadje.Width
    If Not picScroll.Height = UserControl.Height / Screen.TwipsPerPixelY Then picScroll.Height = UserControl.Height / Screen.TwipsPerPixelY
    If Not picScroll.Top = 0 Then picScroll.Top = 0
    If Not picScroll.Visible = True Then picScroll.Visible = True
    Poravnaj
Else
    If Not PicOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX Then PicOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX
    If Not picScroll.Visible = False Then picScroll.Visible = False
    Poravnaj
End If
LegaDrsnika

End Sub

Private Sub lblIzbor_DblClick(Index As Integer)
If MultiSelect = False Then
    NePredvajaj = False
    If GU = False And Index = Playing Then
        RaiseEvent Play(lblIme(Playing).Tag)
        GU = True
    Else
        Play (Index)
    End If
End If

End Sub

Private Sub lblIzbor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim iCnt As Integer
If Button = vbLeftButton Then
    If MultiSelect = True Then
        MS Index
    Else
        SelectIndex (Index)
    End If
Else
    If lblIzbor(Index).Tag <> "I" Then
        If MultiSelect = True Then
            MS Index
        Else
            SelectIndex (Index)
        End If
    End If
    
    If Selected <> Index Then
        Selected = Index
    End If
        lblIme(Selected).Refresh
        
    RaiseEvent ShowMenu
End If

End Sub

Private Sub lblIzbor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    If MultiSelect = True Then
    
    Else
        If ListCount > 0 Then
            picPremik.Visible = True
            If Int((Y / Screen.TwipsPerPixelY + (Index - 1) * lblIzbor(Index).Height) / lblIzbor(Index).Height) * lblIzbor(Index).Height - 1 < -1 Then
                picPremik.Top = -1
            ElseIf Int((Y / Screen.TwipsPerPixelY + (Index - 1) * lblIzbor(Index).Height) / lblIzbor(Index).Height) * lblIzbor(Index).Height - 1 > (ListCount - 1) * lblIzbor(Index).Height Then
                picPremik.Top = (ListCount - 1) * lblIzbor(Index).Height - 1
            Else
                picPremik.Top = Int((Y / Screen.TwipsPerPixelY + (Index - 1) * lblIzbor(Index).Height) / lblIzbor(Index).Height) * lblIzbor(Index).Height - 1
    
            End If
        End If
    End If
ElseIf Button = vbRightButton Then

Else
    If lblSpot(Index).Visible = True Then
        If Not lblIzbor(Index).ToolTipText = lblIme(Index).Caption Then lblIzbor(Index).ToolTipText = lblIme(Index).Caption
    Else
        lblIzbor(Index).ToolTipText = ""
    End If
End If

End Sub

Private Sub lblIzbor_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    If picPremik.Visible = True Then
        NePredvajaj = True
        picPremik.Visible = False
        Dim QW As Integer
        Dim x1 As String
        Dim x2 As String
        Dim x3 As String
        Dim x4 As String
        Dim x5 As String
        Dim iCnt As Integer
        
        If Int((Y / Screen.TwipsPerPixelY + (Index - 1) * lblIzbor(Index).Height) / lblIzbor(Index).Height) * lblIzbor(Index).Height - 1 < -1 Then
            QW = 0
        ElseIf Int((Y / Screen.TwipsPerPixelY + (Index - 1) * lblIzbor(Index).Height) / lblIzbor(Index).Height) * lblIzbor(Index).Height - 1 > (ListCount - 1) * lblIzbor(Index).Height Then
            QW = (ListCount - 1)
        Else
            QW = Int((Y / Screen.TwipsPerPixelY + (Index - 1) * lblIzbor(Index).Height) / lblIzbor(Index).Height)
        End If
        
        If Index <> QW + 1 Then
        
            x1 = lbl�tevilka(Index).Tag
            x2 = lblIme(Index).Caption
            x3 = lblIme(Index).Tag
            x4 = lbl�as(Index).Caption
            x5 = lbl�as(Index).Tag
    
            If QW < Index Then
    
                Dim c As Integer
                For iCnt = QW To Index - 1
                    c = Index - iCnt + QW
                    lbl�tevilka(c).Tag = lbl�tevilka(c - 1).Tag
                    lblIme(c).Caption = lblIme(c - 1).Caption
                    lblIme(c).Tag = lblIme(c - 1).Tag
                    lbl�as(c).Caption = lbl�as(c - 1).Caption
                    lbl�as(c).Tag = lbl�as(c - 1).Tag
                    
                    If lblIme(c).Width > lbl�as(c).Left - lblIme(c).Left Then
                        If Not lblSpot(c).Left = lbl�as(c).Left - lblSpot(c).Width Then lblSpot(c).Left = lbl�as(c).Left - lblSpot(c).Width
                        lblSpot(c).Visible = True
                    Else
                        lblSpot(c).Visible = False
                    End If
                    
                Next iCnt
                
                lbl�tevilka(QW + 1).Tag = x1
                lblIme(QW + 1).Caption = x2
                lblIme(QW + 1).Tag = x3
                lbl�as(QW + 1).Caption = x4
                lbl�as(QW + 1).Tag = x5
    
                Dim ASD As Integer
                ASD = Selected
                SelectIndex QW + 1
    
                If Playing = ASD Then
                    Play QW + 1
                Else
                    If Playing >= QW + 1 And Playing <= Index - 1 Then
                        Play (Playing + 1)
                    End If
                End If
                
            Else
                For iCnt = Index To QW
                    lbl�tevilka(iCnt).Tag = lbl�tevilka(iCnt + 1).Tag
                    lblIme(iCnt).Caption = lblIme(iCnt + 1).Caption
                    lblIme(iCnt).Tag = lblIme(iCnt + 1).Tag
                    lbl�as(iCnt).Caption = lbl�as(iCnt + 1).Caption
                    lbl�as(iCnt).Tag = lbl�as(iCnt + 1).Tag
                    
                    If lblIme(iCnt).Width > lbl�as(iCnt).Left - lblIme(iCnt).Left Then
                        If Not lblSpot(iCnt).Left = lbl�as(iCnt).Left - lblSpot(iCnt).Width Then lblSpot(iCnt).Left = lbl�as(iCnt).Left - lblSpot(iCnt).Width
                        lblSpot(iCnt).Visible = True
                    Else
                        lblSpot(iCnt).Visible = False
                    End If
                    
                Next iCnt
                
                lbl�tevilka(QW + 1).Tag = x1
                lblIme(QW + 1).Caption = x2
                lblIme(QW + 1).Tag = x3
                lbl�as(QW + 1).Caption = x4
                lbl�as(QW + 1).Tag = x5
                ASD = Selected
                SelectIndex QW + 1
     
                If Playing = ASD Then
                    Play QW + 1
                Else
                    If Playing >= Index And Playing <= QW + 1 Then
                        Play (Playing - 1)
                    End If
                End If
            End If

        End If
    NePredvajaj = False
    End If
End If

End Sub

Private Sub lblIzbor_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)

End Sub

Private Sub PicDol_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    BitBlt PicDol.hDC, 0, 0, GUI_SSDol�irina, GUI_SSDolVi�ina, PictureData.hDC, GUI_SSDolXD, GUI_SSDolYD, SRCCOPY
    PicDol.Refresh
    
    tmrScroll.Tag = "DOL"
    tmrScroll.Enabled = True
End If

End Sub

Private Sub PicDol_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    If X >= 0 And X <= PicDol.Width And Y >= 0 And Y <= PicDol.Height Then
        BitBlt PicDol.hDC, 0, 0, GUI_SSDol�irina, GUI_SSDolVi�ina, PictureData.hDC, GUI_SSDolXD, GUI_SSDolYD, SRCCOPY
        PicDol.Refresh
        tmrScroll.Enabled = True
    Else
        BitBlt PicDol.hDC, 0, 0, GUI_SSDol�irina, GUI_SSDolVi�ina, PictureData.hDC, GUI_SSDolX, GUI_SSDolY, SRCCOPY
        PicDol.Refresh
        tmrScroll.Enabled = False
    End If
End If

End Sub

Private Sub PicDol_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    BitBlt PicDol.hDC, 0, 0, GUI_SSDol�irina, GUI_SSDolVi�ina, PictureData.hDC, GUI_SSDolX, GUI_SSDolY, SRCCOPY
    PicDol.Refresh
    
    tmrScroll.Enabled = False
End If

End Sub

Private Sub PicDrsnik_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    MouseY = Y
    PicDrsnikD.Visible = True
    
End If

End Sub

Private Sub PicDrsnik_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    If PicDrsnik.Top - (MouseY - Y) <= 0 Then
        PicDrsnik.Top = 0
        If Not PicOzadje.Top = 0 Then PicOzadje.Top = 0
        
    ElseIf PicDrsnik.Top - (MouseY - Y) > PicPoljeDrsnika.Height - PicDrsnik.Height Then
        PicDrsnik.Top = PicPoljeDrsnika.Height - PicDrsnik.Height
        If Not PicOzadje.Top = UserControl.Height / Screen.TwipsPerPixelY - PicOzadje.Height Then PicOzadje.Top = UserControl.Height / Screen.TwipsPerPixelY - PicOzadje.Height
    Else
        PicDrsnik.Top = PicDrsnik.Top - (MouseY - Y)
        PicOzadje.Top = PicDrsnik.Top * (UserControl.Height / Screen.TwipsPerPixelY - PicOzadje.Height) / (PicPoljeDrsnika.Height - PicDrsnik.Height)
    End If
End If

End Sub

Private Sub PicDrsnik_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PicDrsnikD.Visible = False

End Sub

Private Sub PicDrsnik_Resize()
PicDrsnikDDol.Top = PicDrsnik.Height - GUI_SSDrsnikDolVi�ina
PicDrsnikDol.Top = PicDrsnik.Height - GUI_SSDrsnikDolVi�ina
End Sub

Private Sub PicGor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    BitBlt PicGor.hDC, 0, 0, GUI_SSGor�irina, GUI_SSGorVi�ina, PictureData.hDC, GUI_SSGorXD, GUI_SSGorYD, SRCCOPY
    PicGor.Refresh
    
    tmrScroll.Tag = "GOR"
    tmrScroll.Enabled = True
End If

End Sub

Private Sub PicGor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    If X >= 0 And X <= PicGor.Width And Y >= 0 And Y <= PicGor.Height Then
        BitBlt PicGor.hDC, 0, 0, GUI_SSGor�irina, GUI_SSGorVi�ina, PictureData.hDC, GUI_SSGorXD, GUI_SSGorYD, SRCCOPY
        PicGor.Refresh
        tmrScroll.Enabled = True
    Else
        BitBlt PicGor.hDC, 0, 0, GUI_SSGor�irina, GUI_SSGorVi�ina, PictureData.hDC, GUI_SSGorX, GUI_SSGorY, SRCCOPY
        PicGor.Refresh
        tmrScroll.Enabled = False
    End If
End If

End Sub

Private Sub PicGor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    BitBlt PicGor.hDC, 0, 0, GUI_SSGor�irina, GUI_SSGorVi�ina, PictureData.hDC, GUI_SSGorX, GUI_SSGorY, SRCCOPY
    PicGor.Refresh
    
    tmrScroll.Enabled = False
End If

End Sub

Private Sub PicOzadje_Resize()

If Not shp�asA.Left = PicOzadje.Width - �irina�asa Then shp�asA.Left = PicOzadje.Width - �irina�asa
If Not shp�asA.Top = 0 Then shp�asA.Top = 0
If Not shp�asA.Height = PicOzadje.Height + 50 Then shp�asA.Height = PicOzadje.Height + 50
If Not picPremik.Left = 0 Then picPremik.Left = 0
If Not picPremik.Width = PicOzadje.Width Then picPremik.Width = PicOzadje.Width

If PicOzadje.Height > UserControl.Height / Screen.TwipsPerPixelY Or AShowScroller = True Then
    If Not PicOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX - picScroll.Width Then PicOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX - picScroll.Width
    If Not picScroll.Left = PicOzadje.Width Then picScroll.Left = PicOzadje.Width
    If Not picScroll.Height = UserControl.Height / Screen.TwipsPerPixelY Then picScroll.Height = UserControl.Height / Screen.TwipsPerPixelY
    If Not picScroll.Top = 0 Then picScroll.Top = 0
    If Not picScroll.Visible = True Then picScroll.Visible = True
    Poravnaj
Else
    If Not PicOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX Then PicOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX
    If Not picScroll.Visible = False Then picScroll.Visible = False
    Poravnaj
End If

Dim cc As Integer
For cc = 1 To lblIme.Count - 1
    If Not �rta(cc).x2 = shp�as(cc).Left Then �rta(cc).x2 = shp�as(cc).Left
    
    If Not �rta�as(cc).x1 = shp�as(cc).Left Then �rta�as(cc).x1 = shp�as(cc).Left
    If Not �rta�as(cc).x2 = PicOzadje.Width - 1 Then �rta�as(cc).x2 = PicOzadje.Width - 1

    If Not �rtaRob(cc).x1 = PicOzadje.Width - 1 Then �rtaRob(cc).x1 = PicOzadje.Width - 1
    If Not �rtaRob(cc).x2 = PicOzadje.Width - 1 Then �rtaRob(cc).x2 = PicOzadje.Width - 1
Next cc

End Sub

Private Sub PicPoljeDrsnika_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    If Y < PicDrsnik.Top Then
        tmrScroll.Tag = "MGOR"
    Else
        tmrScroll.Tag = "MDOL"
    End If
    
    tmrScroll.Enabled = True
    PicPoljeDrsnika.Tag = Y
    PicDrsnikD.Visible = True
    
End If

End Sub

Private Sub PicPoljeDrsnika_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
tmrScroll.Enabled = False
tmrScroll.Tag = ""
PicPoljeDrsnika.Tag = 0
PicDrsnikD.Visible = False
    
End Sub

Private Sub tmrScroll_Timer()
Dim cc As Integer
cc = (PicPoljeDrsnika.Height - PicDrsnik.Height) / ((-UserControl.Height / Screen.TwipsPerPixelY + PicOzadje.Height) / ListCount) / 2

    If tmrScroll.Tag = "DOL" Then
        If PicDrsnik.Top <= PicPoljeDrsnika.Height - PicDrsnik.Height - cc Then
            PicDrsnik.Top = PicDrsnik.Top + cc
        Else
            PicDrsnik.Top = PicPoljeDrsnika.Height - PicDrsnik.Height
        End If
    ElseIf tmrScroll.Tag = "GOR" Then
        If PicDrsnik.Top >= cc Then
            PicDrsnik.Top = PicDrsnik.Top - cc
        Else
            PicDrsnik.Top = 0
        End If
    ElseIf tmrScroll.Tag = "MGOR" Then
        
        If PicPoljeDrsnika.Tag < PicDrsnik.Top + PicDrsnik.Height / 2 Then
            If PicDrsnik.Top >= cc Then
                If PicDrsnik.Top - cc < PicPoljeDrsnika.Tag - PicDrsnik.Height / 2 Then
                    PicDrsnik.Top = PicPoljeDrsnika.Tag - PicDrsnik.Height / 2
                Else
                     PicDrsnik.Top = PicDrsnik.Top - cc
                End If
            Else
                PicDrsnik.Top = 0
            End If
        Else
            tmrScroll.Enabled = False
        End If
    ElseIf tmrScroll.Tag = "MDOL" Then
        If PicPoljeDrsnika.Tag > PicDrsnik.Top + PicDrsnik.Height / 2 Then
            If PicDrsnik.Top <= PicPoljeDrsnika.Height - PicDrsnik.Height - cc Then
                If PicDrsnik.Top + cc > PicPoljeDrsnika.Tag - PicDrsnik.Height / 2 Then
                    PicDrsnik.Top = PicPoljeDrsnika.Tag - PicDrsnik.Height / 2
                Else
                     PicDrsnik.Top = PicDrsnik.Top + cc
                End If
            Else
                PicDrsnik.Top = PicPoljeDrsnika.Height - PicDrsnik.Height
            End If
        Else
            tmrScroll.Enabled = False
        End If
    End If

If PicDrsnik.Top = 0 Then
    If Not PicOzadje.Top = 0 Then PicOzadje.Top = 0
ElseIf PicDrsnik.Top = PicPoljeDrsnika.Height - PicDrsnik.Height Then
    If Not PicOzadje.Top = UserControl.Height / Screen.TwipsPerPixelY - PicOzadje.Height Then PicOzadje.Top = UserControl.Height / Screen.TwipsPerPixelY - PicOzadje.Height
Else
    PicOzadje.Top = PicDrsnik.Top * (UserControl.Height / Screen.TwipsPerPixelY - PicOzadje.Height) / (PicPoljeDrsnika.Height - PicDrsnik.Height)
End If

End Sub

Private Sub UserControl_Initialize()
ListCount = 0
�irina�asa = 30
Selected = 0
Playing = 0
Na�inMultiSelect = 0
Skupen�asSekund = 0
PicOzadje.Height = 0

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
    If ListCount > 0 Then
        If Selected = 0 Then
            SelectIndex (1)
        ElseIf Selected <= ListCount - 1 Then
            SelectIndex (Selected + 1)
        End If
    End If
    
ElseIf KeyCode = vbKeyUp Then
    If ListCount > 0 Then
        If Selected = 0 Then
            SelectIndex (1)
        ElseIf Selected > 1 Then
            SelectIndex (Selected - 1)
        End If
    End If
    
ElseIf KeyCode = 13 Then
        If Za�etMS = True Then NoMS
        NePredvajaj = False
        If GU = False And Selected = Playing Then
            RaiseEvent Play(lblIme(Playing).Tag)
            GU = True
        Else
            Play (Selected)
        End If
ElseIf KeyCode = vbKeyDelete Then
    If Za�etMS = False Then
        Remove Selected
    Else
        MultiIzbris
    End If
    
End If

If Shift = 1 Or Shift = 2 Then
    MultiSelect = True
    Na�inMultiSelect = Shift
Else
    MultiSelect = False
    Na�inMultiSelect = 0

End If

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    MultiSelect = False
    Na�inMultiSelect = 0
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    RaiseEvent ShowMenu
End If

End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
BackgroundColor = PropBag.ReadProperty("BackgroundColor", &HFEF5E9)
SelectedBorderColor = PropBag.ReadProperty("SelectedBorderColor", &H8A6544)
SelectedTitleBackColor = PropBag.ReadProperty("SelectedTitleBackColor", &HC89248)
SelectedTimeBackColor = PropBag.ReadProperty("SelectedTimeBackColor", &HA76D46)
BackgroundTimeColor = PropBag.ReadProperty("BackgroundTimeColor", &HFAE2B7)
TimeTextColor = PropBag.ReadProperty("TimeTextColor", &HA76D46)
TitleTextColor = PropBag.ReadProperty("TitleTextColor", &HA76D46)
SelectedTitleTextColor = PropBag.ReadProperty("SelectedTitleTextColor", &HFAE2B7)
SelectedTimeTextColor = PropBag.ReadProperty("SelectedTimeTextColor", &HFEF5E9)
PlayedTimeTextColor = PropBag.ReadProperty("PlayedTimeTextColor", &H40C0&)
PlayedTitleTextColor = PropBag.ReadProperty("PlayedTitleTextColor", &H40C0&)
PlayedBorderColor = PropBag.ReadProperty("PlayedBorderColor", &H40C0&)

End Sub

Private Sub UserControl_Resize()

If PicOzadje.Left <> 0 Then PicOzadje.Left = 0
If PicOzadje.Width <> UserControl.Width / Screen.TwipsPerPixelX Then PicOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX
If PicOzadje.Top > 0 Then PicOzadje.Top = 0

If ListCount = 0 Then PicOzadje.Visible = False Else PicOzadje.Visible = True

If PicOzadje.Height > UserControl.Height / Screen.TwipsPerPixelY Or AShowScroller = True Then
    If Not PicOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX - picScroll.Width Then PicOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX - picScroll.Width
    If Not picScroll.Left = PicOzadje.Width Then picScroll.Left = PicOzadje.Width
    If Not picScroll.Height = UserControl.Height / Screen.TwipsPerPixelY Then picScroll.Height = UserControl.Height / Screen.TwipsPerPixelY
    If Not picScroll.Top = 0 Then picScroll.Top = 0
    If Not picScroll.Visible = True Then picScroll.Visible = True
    Poravnaj
Else
    If Not PicOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX Then PicOzadje.Width = UserControl.Width / Screen.TwipsPerPixelX
    If Not picScroll.Visible = False Then picScroll.Visible = False
    Poravnaj
End If

If PicOzadje.Height + PicOzadje.Top < UserControl.Height / Screen.TwipsPerPixelX Then
    If UserControl.Height / Screen.TwipsPerPixelY - PicOzadje.Height < 0 Then
        PicOzadje.Top = UserControl.Height / Screen.TwipsPerPixelY - PicOzadje.Height
    Else
        PicOzadje.Top = 0
    End If
End If

End Sub

Public Sub SelectIndex(Index As Integer)
If Index <> Selected Then
    shpOzadje(Selected).Visible = False
    shpOzadje(Index).Visible = True
    
    If Selected > 0 Then
        shp�as(Selected).Left = lbl�as(Selected).Left - 1
        shp�as(Selected).Top = (Selected - 1) * (lbl�tevilka(Selected).Height + 2)
        shp�as(Selected).Width = PicOzadje.Width - shp�as(Selected).Left + 1
        shp�as(Selected).Height = lblIme(Selected).Height + 3
        shp�as(Selected).FillColor = shp�as(0).FillColor
        lblIzbor(Selected).Tag = ""
    End If
    
    If Not Selected = Playing Then
        lblIme(Selected).Font = lblIme(0).Font
        lblIme(Selected).FontBold = lblIme(0).FontBold
        lblIme(Selected).FontItalic = lblIme(0).FontItalic
        lblIme(Selected).ForeColor = lblIme(0).ForeColor
        
        lbl�as(Selected).Font = lbl�as(0).Font
        lbl�as(Selected).FontBold = lbl�as(0).FontBold
        lbl�as(Selected).FontItalic = lbl�as(0).FontItalic
        lbl�as(Selected).ForeColor = lbl�as(0).ForeColor
        
        lbl�tevilka(Selected).Font = lbl�tevilka(0).Font
        lbl�tevilka(Selected).FontBold = lbl�tevilka(0).FontBold
        lbl�tevilka(Selected).FontItalic = lbl�tevilka(0).FontItalic
        lbl�tevilka(Selected).ForeColor = lbl�tevilka(0).ForeColor
        
        lblSpot(Selected).Font = lblIme(0).Font
        lblSpot(Selected).FontBold = lblIme(0).FontBold
        lblSpot(Selected).FontItalic = lblIme(0).FontItalic
        lblSpot(Selected).ForeColor = lblIme(0).ForeColor

        
        If lblIme(Selected).Width > lbl�as(Selected).Left - lblIme(Selected).Left Then
            If Not lblSpot(Selected).Left = lbl�as(Selected).Left - lblSpot(Selected).Width Then lblSpot(Selected).Left = lbl�as(Selected).Left - lblSpot(Selected).Width
            lblSpot(Selected).Visible = True
        Else
            lblSpot(Selected).Visible = False
        End If
        
    End If
        lblSpot(Selected).BackColor = PicOzadje.BackColor
    lblIme(Selected).Left = lbl�tevilka(Selected).Width + lbl�tevilka(Selected).Left

    
    Selected = Index
    prvaMultiSelect = Index
    
    shp�as(Selected).Top = (Selected - 1) * (lbl�tevilka(Selected).Height + 2) + 1
    shp�as(Selected).Width = PicOzadje.Width - shp�as(Selected).Left
    shp�as(Selected).Height = lblIme(Selected).Height + 2
    shp�as(Selected).FillColor = shp�asA.FillColor
   
    lblIzbor(Selected).Tag = "I"
    
    lbl�as(Selected).Refresh
    lblSpot(Selected).Refresh
    
    If Not Index = Playing Then
        lblIme(Selected).Font = lblImeA.Font
        lblIme(Selected).FontBold = lblImeA.FontBold
        lblIme(Selected).FontItalic = lblImeA.FontItalic
        lblIme(Selected).ForeColor = lblImeA.ForeColor
        
        lbl�as(Selected).Font = lbl�asA.Font
        lbl�as(Selected).FontBold = lbl�asA.FontBold
        lbl�as(Selected).FontItalic = lbl�asA.FontItalic
        lbl�as(Selected).ForeColor = lbl�asA.ForeColor
        
        lbl�tevilka(Selected).Font = lbl�tevilkaA.Font
        lbl�tevilka(Selected).FontBold = lbl�tevilkaA.FontBold
        lbl�tevilka(Selected).FontItalic = lbl�tevilkaA.FontItalic
        lbl�tevilka(Selected).ForeColor = lbl�tevilkaA.ForeColor
        
        lblSpot(Selected).Font = lblImeA.Font
        lblSpot(Selected).FontBold = lblImeA.FontBold
        lblSpot(Selected).FontItalic = lblImeA.FontItalic
        lblSpot(Selected).ForeColor = lblImeA.ForeColor

        
        If lblIme(Selected).Width > lbl�as(Selected).Left - lblIme(Selected).Left Then
            If Not lblSpot(Selected).Left = lbl�as(Selected).Left - lblSpot(Selected).Width Then lblSpot(Selected).Left = lbl�as(Selected).Left - lblSpot(Selected).Width
            lblSpot(Selected).Visible = True
        Else
            lblSpot(Selected).Visible = False
        End If
        
    End If
        lblSpot(Selected).BackColor = shpOzadje(0).FillColor
    lblIme(Selected).Left = lbl�tevilka(Selected).Width + lbl�tevilka(Selected).Left

    If PicOzadje.Height > UserControl.Height / Screen.TwipsPerPixelY Then
        If (Selected - 1) * lblIzbor(Selected).Height < -PicOzadje.Top Then
            PicOzadje.Top = -(Selected - 1) * lblIzbor(Selected).Height
            LegaDrsnika
        ElseIf (Selected) * lblIzbor(Selected).Height > UserControl.Height / Screen.TwipsPerPixelY - PicOzadje.Top Then
            PicOzadje.Top = UserControl.Height / Screen.TwipsPerPixelY - (Selected) * lblIzbor(Selected).Height - 1
            LegaDrsnika
        End If
    End If
    
    NoMS
End If

End Sub

Public Sub Play(Index As Integer)
On Error Resume Next

If Index = Playing And NePredvajaj = False Then
    RaiseEvent RePlay
    
    If PicOzadje.Height > UserControl.Height / Screen.TwipsPerPixelY Then
        If (Playing - 1) * lblIzbor(Playing).Height < -PicOzadje.Top Then
            PicOzadje.Top = -(Playing - 1) * lblIzbor(Playing).Height
            LegaDrsnika
        ElseIf (Playing) * lblIzbor(Playing).Height > UserControl.Height / Screen.TwipsPerPixelY - PicOzadje.Top Then
            PicOzadje.Top = UserControl.Height / Screen.TwipsPerPixelY - (Playing - 1) * lblIzbor(Playing).Height - 1
            LegaDrsnika
        End If
    End If
    
    Exit Sub

End If

If Index > 0 Then

    If Playing <> Selected Then
        lblIme(Playing).Font = lblIme(0).Font
        lblIme(Playing).FontBold = lblIme(0).FontBold
        lblIme(Playing).FontItalic = lblIme(0).FontItalic
        lblIme(Playing).ForeColor = lblIme(0).ForeColor
        
        lbl�as(Playing).Font = lbl�as(0).Font
        lbl�as(Playing).FontBold = lbl�as(0).FontBold
        lbl�as(Playing).FontItalic = lbl�as(0).FontItalic
        lbl�as(Playing).ForeColor = lbl�as(0).ForeColor
        
        lbl�tevilka(Playing).Font = lbl�tevilka(0).Font
        lbl�tevilka(Playing).FontBold = lbl�tevilka(0).FontBold
        lbl�tevilka(Playing).FontItalic = lbl�tevilka(0).FontItalic
        lbl�tevilka(Playing).ForeColor = lbl�tevilka(0).ForeColor
        
        lblSpot(Playing).Font = lblIme(0).Font
        lblSpot(Playing).FontBold = lblIme(0).FontBold
        lblSpot(Playing).FontItalic = lblIme(0).FontItalic
        lblSpot(Playing).ForeColor = lblIme(0).ForeColor

    Else
        If Not Playing = 0 Then
            lblIme(Playing).Font = lblImeA.Font
            lblIme(Playing).FontBold = lblImeA.FontBold
            lblIme(Playing).FontItalic = lblImeA.FontItalic
            lblIme(Playing).ForeColor = lblImeA.ForeColor
            
            lbl�as(Playing).Font = lbl�asA.Font
            lbl�as(Playing).FontBold = lbl�asA.FontBold
            lbl�as(Playing).FontItalic = lbl�asA.FontItalic
            lbl�as(Playing).ForeColor = lbl�asA.ForeColor
            
            lbl�tevilka(Playing).Font = lbl�tevilkaA.Font
            lbl�tevilka(Playing).FontBold = lbl�tevilkaA.FontBold
            lbl�tevilka(Playing).FontItalic = lbl�tevilkaA.FontItalic
            lbl�tevilka(Playing).ForeColor = lbl�tevilkaA.ForeColor
            
            lblSpot(Playing).Font = lblImeA.Font
            lblSpot(Playing).FontBold = lblImeA.FontBold
            lblSpot(Playing).FontItalic = lblImeA.FontItalic
            lblSpot(Playing).ForeColor = lblImeA.ForeColor
        End If
    End If
    
    If Playing <> 0 Then
        If lblIme(Playing).Width > lbl�as(Playing).Left - lblIme(Playing).Left Then
            If Not lblSpot(Playing).Left = lbl�as(Playing).Left - lblSpot(Playing).Width Then lblSpot(Playing).Left = lbl�as(Playing).Left - lblSpot(Playing).Width
            lblSpot(Playing).Visible = True
        Else
            lblSpot(Playing).Visible = False
        End If
    End If
    
    lblIme(Playing).Left = lbl�tevilka(Playing).Width + lbl�tevilka(Playing).Left
    Dim cc As Integer
    cc = Playing
    Playing = Index

    shp�asB.Top = (Playing - 1) * (lbl�tevilka(Playing).Height + 2)
    shp�asB.Width = PicOzadje.Width
    shp�asB.Height = lblIme(Playing).Height + 3
    If Not shp�asB.Left = 0 Then shp�asB.Left = 0
    
    shp�asB.Visible = True
    
    lblIme(Playing).Font = lblImeB.Font
    lblIme(Playing).FontBold = lblImeB.FontBold
    lblIme(Playing).FontItalic = lblImeB.FontItalic
    lblIme(Playing).ForeColor = lblImeB.ForeColor
    
    lbl�as(Playing).Font = lbl�asB.Font
    lbl�as(Playing).FontBold = lbl�asB.FontBold
    lbl�as(Playing).FontItalic = lbl�asB.FontItalic
    lbl�as(Playing).ForeColor = lbl�asB.ForeColor
    
    lbl�tevilka(Playing).Font = lbl�tevilkaB.Font
    lbl�tevilka(Playing).FontBold = lbl�tevilkaB.FontBold
    lbl�tevilka(Playing).FontItalic = lbl�tevilkaB.FontItalic
    lbl�tevilka(Playing).ForeColor = lbl�tevilkaB.ForeColor
    
    lblIme(Playing).Left = lbl�tevilka(Playing).Width + lbl�tevilka(Playing).Left

    
    lblSpot(Playing).Font = lblImeB.Font
    lblSpot(Playing).FontBold = lblImeB.FontBold
    lblSpot(Playing).FontItalic = lblImeB.FontItalic
    lblSpot(Playing).ForeColor = lblImeB.ForeColor
    
    If lblIme(Playing).Width > lbl�as(Playing).Left - lblIme(Playing).Left Then
        If Not lblSpot(Playing).Left = lbl�as(Playing).Left - lblSpot(Playing).Width Then lblSpot(Playing).Left = lbl�as(Playing).Left - lblSpot(Playing).Width
        lblSpot(Playing).Visible = True
    Else
        lblSpot(Playing).Visible = False
    End If

    If PicOzadje.Height > UserControl.Height / Screen.TwipsPerPixelY Then
        If (Playing - 1) * lblIzbor(Playing).Height < -PicOzadje.Top Then
            PicOzadje.Top = -(Playing - 1) * lblIzbor(Playing).Height
            LegaDrsnika
        ElseIf (Playing) * lblIzbor(Playing).Height > UserControl.Height / Screen.TwipsPerPixelY - PicOzadje.Top Then
            PicOzadje.Top = UserControl.Height / Screen.TwipsPerPixelY - (Playing) * lblIzbor(Playing).Height - 1
            LegaDrsnika
        End If
    End If
    
    UserControl.Refresh
    If NePredvajaj = False Then
        RaiseEvent Play(lblIme(Playing).Tag)
    End If
        
End If

End Sub

Public Sub MS(Index As Integer)
Dim iCnt As Integer
Dim a As Integer
Dim B As Integer
Za�etMS = True

If Na�inMultiSelect = 1 Then
    If Index >= prvaMultiSelect Then
        a = Index
        B = prvaMultiSelect
    Else
        B = Index
        a = prvaMultiSelect
        
    End If
        For iCnt = 1 To lblIme.Count - 1
            If iCnt >= B And iCnt <= a Then
                If Not shpOzadje(iCnt).Visible = True Then shpOzadje(iCnt).Visible = True
                If Not shp�as(iCnt).Top = (iCnt - 1) * (lbl�tevilka(iCnt).Height + 2) + 1 Then shp�as(iCnt).Top = (iCnt - 1) * (lbl�tevilka(iCnt).Height + 2) + 1
                If Not shp�as(iCnt).Width = PicOzadje.Width - shp�as(iCnt).Left Then shp�as(iCnt).Width = PicOzadje.Width - shp�as(iCnt).Left
                If Not shp�as(iCnt).Height = lblIme(iCnt).Height + 2 Then shp�as(iCnt).Height = lblIme(iCnt).Height + 2
                If Not shp�as(iCnt).FillColor = shp�asA.FillColor Then shp�as(iCnt).FillColor = shp�asA.FillColor

                
                If Not lblIzbor(iCnt).Tag = "I" Then lblIzbor(iCnt).Tag = "I"
                
                If iCnt < a Then
                    �rta(iCnt).Visible = True
                    �rta�as(iCnt).Visible = True
                Else
                    �rta(iCnt).Visible = False
                    �rta�as(iCnt).Visible = False
                End If
                
                If Not iCnt = Playing Then
                    If Not lblIme(iCnt).Font = lblImeA.Font Then lblIme(iCnt).Font = lblImeA.Font
                    If Not lblIme(iCnt).FontBold = lblImeA.FontBold Then lblIme(iCnt).FontBold = lblImeA.FontBold
                    If Not lblIme(iCnt).FontItalic = lblImeA.FontItalic Then lblIme(iCnt).FontItalic = lblImeA.FontItalic
                    If Not lblIme(iCnt).ForeColor = lblImeA.ForeColor Then lblIme(iCnt).ForeColor = lblImeA.ForeColor
                    
                    If Not lbl�as(iCnt).Font = lbl�asA.Font Then lbl�as(iCnt).Font = lbl�asA.Font
                    If Not lbl�as(iCnt).FontBold = lbl�asA.FontBold Then lbl�as(iCnt).FontBold = lbl�asA.FontBold
                    If Not lbl�as(iCnt).FontItalic = lbl�asA.FontItalic Then lbl�as(iCnt).FontItalic = lbl�asA.FontItalic
                    If Not lbl�as(iCnt).ForeColor = lbl�asA.ForeColor Then lbl�as(iCnt).ForeColor = lbl�asA.ForeColor
                    
                    If Not lbl�tevilka(iCnt).Font = lbl�tevilkaA.Font Then lbl�tevilka(iCnt).Font = lbl�tevilkaA.Font
                    If Not lbl�tevilka(iCnt).FontBold = lbl�tevilkaA.FontBold Then lbl�tevilka(iCnt).FontBold = lbl�tevilkaA.FontBold
                    If Not lbl�tevilka(iCnt).FontItalic = lbl�tevilkaA.FontItalic Then lbl�tevilka(iCnt).FontItalic = lbl�tevilkaA.FontItalic
                    If Not lbl�tevilka(iCnt).ForeColor = lbl�tevilkaA.ForeColor Then lbl�tevilka(iCnt).ForeColor = lbl�tevilkaA.ForeColor
                    
                    If Not lblSpot(iCnt).Font = lblImeA.Font Then lblSpot(iCnt).Font = lblImeA.Font
                    If Not lblSpot(iCnt).FontBold = lblImeA.FontBold Then lblSpot(iCnt).FontBold = lblImeA.FontBold
                    If Not lblSpot(iCnt).FontItalic = lblImeA.FontItalic Then lblSpot(iCnt).FontItalic = lblImeA.FontItalic
                    If Not lblSpot(iCnt).ForeColor = lblImeA.ForeColor Then lblSpot(iCnt).ForeColor = lblImeA.ForeColor
            
                    
                    If lblIme(iCnt).Width > lbl�as(iCnt).Left - lblIme(iCnt).Left Then
                        If Not lblSpot(iCnt).Left = lbl�as(iCnt).Left - lblSpot(iCnt).Width Then lblSpot(iCnt).Left = lbl�as(iCnt).Left - lblSpot(iCnt).Width
                        If Not lblSpot(iCnt).Visible = True Then lblSpot(iCnt).Visible = True
                    Else
                        lblSpot(iCnt).Visible = False
                    End If
                    
                End If
                
                If Not lblSpot(iCnt).BackColor = shpOzadje(0).FillColor Then lblSpot(iCnt).BackColor = shpOzadje(0).FillColor
                If Not lblIme(iCnt).Left = lbl�tevilka(iCnt).Width + lbl�tevilka(iCnt).Left Then lblIme(iCnt).Left = lbl�tevilka(iCnt).Width + lbl�tevilka(iCnt).Left

            Else
                �rta(iCnt).Visible = False
                �rta�as(iCnt).Visible = False
                If lblIzbor(iCnt).Tag = "I" Then
                    shpOzadje(iCnt).Visible = False
                   
                    shp�as(iCnt).Left = lbl�as(iCnt).Left - 1
                    shp�as(iCnt).Top = (iCnt - 1) * (lbl�tevilka(iCnt).Height + 2)
                    shp�as(iCnt).Width = PicOzadje.Width - shp�as(iCnt).Left + 1
                    shp�as(iCnt).Height = lblIme(iCnt).Height + 3
                    shp�as(iCnt).FillColor = shp�as(0).FillColor
                    lblIzbor(iCnt).Tag = ""
                    
                    If Not iCnt = Playing Then
                        lblIme(iCnt).Font = lblIme(0).Font
                        lblIme(iCnt).FontBold = lblIme(0).FontBold
                        lblIme(iCnt).FontItalic = lblIme(0).FontItalic
                        lblIme(iCnt).ForeColor = lblIme(0).ForeColor
                        
                        lbl�as(iCnt).Font = lbl�as(0).Font
                        lbl�as(iCnt).FontBold = lbl�as(0).FontBold
                        lbl�as(iCnt).FontItalic = lbl�as(0).FontItalic
                        lbl�as(iCnt).ForeColor = lbl�as(0).ForeColor
                        
                        lbl�tevilka(iCnt).Font = lbl�tevilka(0).Font
                        lbl�tevilka(iCnt).FontBold = lbl�tevilka(0).FontBold
                        lbl�tevilka(iCnt).FontItalic = lbl�tevilka(0).FontItalic
                        lbl�tevilka(iCnt).ForeColor = lbl�tevilka(0).ForeColor
                        
                        lblSpot(iCnt).Font = lblIme(0).Font
                        lblSpot(iCnt).FontBold = lblIme(0).FontBold
                        lblSpot(iCnt).FontItalic = lblIme(0).FontItalic
                        lblSpot(iCnt).ForeColor = lblIme(0).ForeColor
                
                        
                        If lblIme(iCnt).Width > lbl�as(iCnt).Left - lblIme(iCnt).Left Then
                            If Not lblSpot(iCnt).Left = lbl�as(iCnt).Left - lblSpot(iCnt).Width Then lblSpot(iCnt).Left = lbl�as(iCnt).Left - lblSpot(iCnt).Width
                            lblSpot(iCnt).Visible = True
                        Else
                            lblSpot(iCnt).Visible = False
                        End If
                        
                    End If
                        lblSpot(iCnt).BackColor = PicOzadje.BackColor
                    lblIme(iCnt).Left = lbl�tevilka(iCnt).Width + lbl�tevilka(iCnt).Left
                End If
            End If
        
        Next iCnt
Else
iCnt = Index
    If lblIzbor(iCnt).Tag = "" Then
        If Not shpOzadje(iCnt).Visible = True Then shpOzadje(iCnt).Visible = True
        If Not shp�as(iCnt).Top = (iCnt - 1) * (lbl�tevilka(iCnt).Height + 2) + 1 Then shp�as(iCnt).Top = (iCnt - 1) * (lbl�tevilka(iCnt).Height + 2) + 1
        If Not shp�as(iCnt).Width = PicOzadje.Width - shp�as(iCnt).Left Then shp�as(iCnt).Width = PicOzadje.Width - shp�as(iCnt).Left
        If Not shp�as(iCnt).Height = lblIme(iCnt).Height + 2 Then shp�as(iCnt).Height = lblIme(iCnt).Height + 2
        If Not shp�as(iCnt).FillColor = shp�asA.FillColor Then shp�as(iCnt).FillColor = shp�asA.FillColor
        
        If Not lblIzbor(iCnt).Tag = "I" Then lblIzbor(iCnt).Tag = "I"
        
        If Not iCnt = Playing Then
            If Not lblIme(iCnt).Font = lblImeA.Font Then lblIme(iCnt).Font = lblImeA.Font
            If Not lblIme(iCnt).FontBold = lblImeA.FontBold Then lblIme(iCnt).FontBold = lblImeA.FontBold
            If Not lblIme(iCnt).FontItalic = lblImeA.FontItalic Then lblIme(iCnt).FontItalic = lblImeA.FontItalic
            If Not lblIme(iCnt).ForeColor = lblImeA.ForeColor Then lblIme(iCnt).ForeColor = lblImeA.ForeColor
            
            If Not lbl�as(iCnt).Font = lbl�asA.Font Then lbl�as(iCnt).Font = lbl�asA.Font
            If Not lbl�as(iCnt).FontBold = lbl�asA.FontBold Then lbl�as(iCnt).FontBold = lbl�asA.FontBold
            If Not lbl�as(iCnt).FontItalic = lbl�asA.FontItalic Then lbl�as(iCnt).FontItalic = lbl�asA.FontItalic
            If Not lbl�as(iCnt).ForeColor = lbl�asA.ForeColor Then lbl�as(iCnt).ForeColor = lbl�asA.ForeColor
            
            If Not lbl�tevilka(iCnt).Font = lbl�tevilkaA.Font Then lbl�tevilka(iCnt).Font = lbl�tevilkaA.Font
            If Not lbl�tevilka(iCnt).FontBold = lbl�tevilkaA.FontBold Then lbl�tevilka(iCnt).FontBold = lbl�tevilkaA.FontBold
            If Not lbl�tevilka(iCnt).FontItalic = lbl�tevilkaA.FontItalic Then lbl�tevilka(iCnt).FontItalic = lbl�tevilkaA.FontItalic
            If Not lbl�tevilka(iCnt).ForeColor = lbl�tevilkaA.ForeColor Then lbl�tevilka(iCnt).ForeColor = lbl�tevilkaA.ForeColor
            
            If Not lblSpot(iCnt).Font = lblImeA.Font Then lblSpot(iCnt).Font = lblImeA.Font
            If Not lblSpot(iCnt).FontBold = lblImeA.FontBold Then lblSpot(iCnt).FontBold = lblImeA.FontBold
            If Not lblSpot(iCnt).FontItalic = lblImeA.FontItalic Then lblSpot(iCnt).FontItalic = lblImeA.FontItalic
            If Not lblSpot(iCnt).ForeColor = lblImeA.ForeColor Then lblSpot(iCnt).ForeColor = lblImeA.ForeColor
    
            
            If lblIme(iCnt).Width > lbl�as(iCnt).Left - lblIme(iCnt).Left Then
                If Not lblSpot(iCnt).Left = lbl�as(iCnt).Left - lblSpot(iCnt).Width Then lblSpot(iCnt).Left = lbl�as(iCnt).Left - lblSpot(iCnt).Width
                If Not lblSpot(iCnt).Visible = True Then lblSpot(iCnt).Visible = True
            Else
                lblSpot(iCnt).Visible = False
            End If
            
        End If
        
        If Not lblSpot(iCnt).BackColor = shpOzadje(0).FillColor Then lblSpot(iCnt).BackColor = shpOzadje(0).FillColor
        If Not lblIme(iCnt).Left = lbl�tevilka(iCnt).Width + lbl�tevilka(iCnt).Left Then lblIme(iCnt).Left = lbl�tevilka(iCnt).Width + lbl�tevilka(iCnt).Left
   Else
        If lblIzbor(iCnt).Tag = "I" Then
        shpOzadje(iCnt).Visible = False
       
        shp�as(iCnt).Left = lbl�as(iCnt).Left - 1
        shp�as(iCnt).Top = (iCnt - 1) * (lbl�tevilka(iCnt).Height + 2)
        shp�as(iCnt).Width = PicOzadje.Width - shp�as(iCnt).Left + 1
        shp�as(iCnt).Height = lblIme(iCnt).Height + 3
        shp�as(iCnt).FillColor = shp�as(0).FillColor
        lblIzbor(iCnt).Tag = ""
        
        If Not iCnt = Playing Then
            lblIme(iCnt).Font = lblIme(0).Font
            lblIme(iCnt).FontBold = lblIme(0).FontBold
            lblIme(iCnt).FontItalic = lblIme(0).FontItalic
            lblIme(iCnt).ForeColor = lblIme(0).ForeColor
            
            lbl�as(iCnt).Font = lbl�as(0).Font
            lbl�as(iCnt).FontBold = lbl�as(0).FontBold
            lbl�as(iCnt).FontItalic = lbl�as(0).FontItalic
            lbl�as(iCnt).ForeColor = lbl�as(0).ForeColor
            
            lbl�tevilka(iCnt).Font = lbl�tevilka(0).Font
            lbl�tevilka(iCnt).FontBold = lbl�tevilka(0).FontBold
            lbl�tevilka(iCnt).FontItalic = lbl�tevilka(0).FontItalic
            lbl�tevilka(iCnt).ForeColor = lbl�tevilka(0).ForeColor
            
            lblSpot(iCnt).Font = lblIme(0).Font
            lblSpot(iCnt).FontBold = lblIme(0).FontBold
            lblSpot(iCnt).FontItalic = lblIme(0).FontItalic
            lblSpot(iCnt).ForeColor = lblIme(0).ForeColor
    
            
            If lblIme(iCnt).Width > lbl�as(iCnt).Left - lblIme(iCnt).Left Then
                If Not lblSpot(iCnt).Left = lbl�as(iCnt).Left - lblSpot(iCnt).Width Then lblSpot(iCnt).Left = lbl�as(iCnt).Left - lblSpot(iCnt).Width
                lblSpot(iCnt).Visible = True
            Else
                lblSpot(iCnt).Visible = False
            End If
            
        End If
            lblSpot(iCnt).BackColor = PicOzadje.BackColor
        lblIme(iCnt).Left = lbl�tevilka(iCnt).Width + lbl�tevilka(iCnt).Left
    End If
   End If
   �rte

End If

End Sub

Public Sub NoMS()
Za�etMS = False
Dim iCnt As Integer

For iCnt = 1 To lblIme.Count - 1
    �rta(iCnt).Visible = False
    �rta�as(iCnt).Visible = False
    If Not iCnt = Selected Then
        If lblIzbor(iCnt).Tag = "I" Then
            shpOzadje(iCnt).Visible = False
           
            shp�as(iCnt).Left = lbl�as(iCnt).Left - 1
            shp�as(iCnt).Top = (iCnt - 1) * (lbl�tevilka(iCnt).Height + 2)
            shp�as(iCnt).Width = PicOzadje.Width - shp�as(iCnt).Left + 1
            shp�as(iCnt).Height = lblIme(iCnt).Height + 3
            shp�as(iCnt).FillColor = shp�as(0).FillColor
            lblIzbor(iCnt).Tag = ""
            
            If Not iCnt = Playing Then
                lblIme(iCnt).Font = lblIme(0).Font
                lblIme(iCnt).FontBold = lblIme(0).FontBold
                lblIme(iCnt).FontItalic = lblIme(0).FontItalic
                lblIme(iCnt).ForeColor = lblIme(0).ForeColor
                
                lbl�as(iCnt).Font = lbl�as(0).Font
                lbl�as(iCnt).FontBold = lbl�as(0).FontBold
                lbl�as(iCnt).FontItalic = lbl�as(0).FontItalic
                lbl�as(iCnt).ForeColor = lbl�as(0).ForeColor
                
                lbl�tevilka(iCnt).Font = lbl�tevilka(0).Font
                lbl�tevilka(iCnt).FontBold = lbl�tevilka(0).FontBold
                lbl�tevilka(iCnt).FontItalic = lbl�tevilka(0).FontItalic
                lbl�tevilka(iCnt).ForeColor = lbl�tevilka(0).ForeColor
                
                lblSpot(iCnt).Font = lblIme(0).Font
                lblSpot(iCnt).FontBold = lblIme(0).FontBold
                lblSpot(iCnt).FontItalic = lblIme(0).FontItalic
                lblSpot(iCnt).ForeColor = lblIme(0).ForeColor
        
                
                If lblIme(iCnt).Width > lbl�as(iCnt).Left - lblIme(iCnt).Left Then
                    If Not lblSpot(iCnt).Left = lbl�as(iCnt).Left - lblSpot(iCnt).Width Then lblSpot(iCnt).Left = lbl�as(iCnt).Left - lblSpot(iCnt).Width
                    lblSpot(iCnt).Visible = True
                Else
                    lblSpot(iCnt).Visible = False
                End If
                
            End If
                lblSpot(iCnt).BackColor = PicOzadje.BackColor
            lblIme(iCnt).Left = lbl�tevilka(iCnt).Width + lbl�tevilka(iCnt).Left
        End If
    End If
Next iCnt

End Sub

Private Sub �rte(Optional SkrijVse As Boolean)
On Error Resume Next
Dim iCnt As Integer

If SkrijVse = True Then
    For iCnt = 1 To �rta.Count - 1
        �rta(iCnt).Visible = False
        �rta�as(iCnt).Visible = False
    Next iCnt
Else
    For iCnt = 0 To �rta.Count - 2
        If lblIzbor(iCnt).Tag = "I" And lblIzbor(iCnt + 1).Tag = "I" Then
            If Not �rta(iCnt).Visible = True Then �rta(iCnt).Visible = True
            If Not �rta�as(iCnt).Visible = True Then �rta�as(iCnt).Visible = True
        Else
            If Not �rta(iCnt).Visible = False Then �rta(iCnt).Visible = False
            If Not �rta�as(iCnt).Visible = False Then �rta�as(iCnt).Visible = False
        End If
    Next iCnt
End If

End Sub

Public Sub RefreshTime(NewTimeInSeconds As Long, NewTime As String, Index As Integer)
If lbl�as(Index).Tag <> NewTimeInSeconds Then
    Skupen�asSekund = Skupen�asSekund - lbl�as(Index).Tag + NewTimeInSeconds
    RaiseEvent DurationChange(Skupen�asSekund)
    
    lbl�as(Index).Tag = NewTimeInSeconds
    lbl�as(Index).Caption = NewTime
End If

End Sub

Public Sub GetData(Index As Integer)
If Index > 0 Then
    gFileName = lblIme(Index).Tag
    gFileName2 = lbl�tevilka(Index).Tag
    gTitle = lblIme(Index).Caption
    gTime = lbl�as(Index).Caption
    gTimeInSeconds = lbl�as(Index).Tag
Else
    gFileName = ""
    gFileName2 = ""
    gTitle = ""
    gTime = ""
    gTimeInSeconds = 0
End If

End Sub

Public Sub AddFileName2(FileName As String, Index As Integer)
lbl�tevilka(Index).Tag = FileName

End Sub

Public Property Get BackgroundColor() As OLE_COLOR
BackgroundColor = PicOzadje.BackColor

End Property

Public Property Get BackgroundTimeColor() As OLE_COLOR
BackgroundTimeColor = shp�as(0).FillColor

End Property

Public Property Get SelectedTimeTextColor() As OLE_COLOR
SelectedTimeTextColor = lbl�asA.ForeColor

End Property

Public Property Get PlayedTimeTextColor() As OLE_COLOR
PlayedTimeTextColor = lbl�asB.ForeColor

End Property

Public Property Get TimeTextColor() As OLE_COLOR
TimeTextColor = lbl�as(0).ForeColor

End Property

Public Property Get PlayedTitleTextColor() As OLE_COLOR
PlayedTitleTextColor = lblImeB.ForeColor

End Property

Public Property Get TitleTextColor() As OLE_COLOR
TitleTextColor = lblIme(0).ForeColor

End Property

Public Property Get SelectedTitleTextColor() As OLE_COLOR
SelectedTitleTextColor = lblImeA.ForeColor

End Property

Public Property Get SelectedTitleBackColor() As OLE_COLOR
SelectedTitleBackColor = shpOzadje(0).FillColor

End Property

Public Property Get SelectedTimeBackColor() As OLE_COLOR
SelectedTimeBackColor = shp�asA.FillColor

End Property

Public Property Get SelectedBorderColor() As OLE_COLOR
SelectedBorderColor = UserControl.shpOzadje(0).BorderColor

End Property

Public Property Get PlayedBorderColor() As OLE_COLOR
PlayedBorderColor = UserControl.shp�asB.BorderColor

End Property

Public Property Let SelectedTitleBackColor(Color As OLE_COLOR)
For xYz = 0 To shpOzadje.Count - 1
    UserControl.shpOzadje(xYz).FillColor = Color
    �rta(xYz).BorderColor = Color
Next xYz

End Property

Public Property Let SelectedTimeBackColor(Color As OLE_COLOR)
UserControl.shp�asA.FillColor = Color
For xYz = 0 To shpOzadje.Count - 1
    �rta�as(xYz).BorderColor = Color
    If lblIzbor(xYz).Tag = "I" Then
        shp�as(xYz).FillColor = Color
    End If
Next xYz

End Property

Public Property Let SelectedBorderColor(Color As OLE_COLOR)
For xYz = 0 To shpOzadje.Count - 1
    UserControl.shpOzadje(xYz).BorderColor = Color
Next xYz

End Property

Public Property Let PlayedBorderColor(Color As OLE_COLOR)
UserControl.shp�asB.BorderColor = Color

End Property

Public Property Let BackgroundColor(Color As OLE_COLOR)
PicOzadje.BackColor = Color
UserControl.BackColor = Color

End Property

Public Property Let BackgroundTimeColor(Color As OLE_COLOR)
For xYz = 0 To shpOzadje.Count - 1
    If Not lblIzbor(xYz).Tag = "I" Then
        shp�as(xYz).FillColor = Color
    End If
Next xYz

End Property

Public Property Let SelectedTimeTextColor(Color As OLE_COLOR)
lbl�asA.ForeColor = Color
For xYz = 1 To lblIme.Count - 1
    If lblIzbor(xYz).Tag = "I" And xYz <> Playing Then
        lbl�as(xYz).ForeColor = Color
    End If
Next xYz

End Property

Public Property Let TimeTextColor(Color As OLE_COLOR)
For xYz = 0 To shpOzadje.Count - 1
    If Not lblIzbor(xYz).Tag = "I" Then
        If Playing = 0 Or Playing <> xYz Then
            lbl�as(xYz).ForeColor = Color
        End If
    End If
Next xYz

End Property

Public Property Let PlayedTimeTextColor(Color As OLE_COLOR)
lbl�asB.ForeColor = Color

If Playing > 0 Then
    lbl�as(Playing).ForeColor = Color
End If

End Property

Public Property Let TitleTextColor(Color As OLE_COLOR)
For xYz = 0 To shpOzadje.Count - 1
    If Not lblIzbor(xYz).Tag = "I" Then
        If Playing = 0 Or Playing <> xYz Then
            lblIme(xYz).ForeColor = Color
            lbl�tevilka(xYz).ForeColor = Color
        End If
    End If
Next xYz

End Property

Public Property Let PlayedTitleTextColor(Color As OLE_COLOR)
lblImeB.ForeColor = Color
lbl�tevilkaB.ForeColor = Color

If Playing > 0 Then
    lblIme(Playing).ForeColor = Color
    lbl�tevilka(Playing).ForeColor = Color
End If

End Property

Public Property Let SelectedTitleTextColor(Color As OLE_COLOR)
lblImeA.ForeColor = Color
lbl�tevilkaA.ForeColor = Color
For xYz = 1 To lblIme.Count - 1
    If lblIzbor(xYz).Tag = "I" Then
        If Playing = 0 Or Playing <> xYz Then
            lblIme(xYz).ForeColor = Color
            lbl�tevilka(xYz).ForeColor = Color
        End If
    End If
Next xYz

End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "BackgroundColor", PicOzadje.BackColor, &HFEF5E9
PropBag.WriteProperty "SelectedBorderColor", UserControl.shpOzadje(0).BorderColor, &H8A6544
PropBag.WriteProperty "SelectedTitleBackColor", UserControl.shpOzadje(0).FillColor, &HC89248
PropBag.WriteProperty "SelectedTimeBackColor", UserControl.shp�asA.FillColor, &HA76D46
PropBag.WriteProperty "BackgroundTimeColor", UserControl.shp�as(0).FillColor, &HFAE2B7
PropBag.WriteProperty "TimeTextColor", UserControl.lbl�as(0).ForeColor, &HA76D46
PropBag.WriteProperty "TitleTextColor", UserControl.lblIme(0).ForeColor, &HA76D46
PropBag.WriteProperty "SelectedTitleTextColor", UserControl.lblImeA.ForeColor, &HFAE2B7
PropBag.WriteProperty "SelectedTimeTextColor", UserControl.lbl�asA.ForeColor, &HFEF5E9
PropBag.WriteProperty "PlayedTimeTextColor", UserControl.lbl�asB.ForeColor, &H40C0&
PropBag.WriteProperty "PlayedTitleTextColor", UserControl.lblImeB.ForeColor, &H40C0&
PropBag.WriteProperty "PlayedBorderColor", UserControl.shp�asB.BorderColor, &H40C0&

End Sub
