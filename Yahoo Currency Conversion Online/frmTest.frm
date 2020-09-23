VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "YCCO (Yahoo Currency Conversion Online) Demo"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9660
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   9660
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Converting currencies using the YCCO module"
      Height          =   2055
      Left            =   90
      TabIndex        =   5
      Top             =   1740
      Width           =   9480
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2010
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   405
         Width           =   3405
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   5820
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   405
         Width           =   3405
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   840
         TabIndex        =   8
         Top             =   390
         Width           =   1065
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   8190
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1515
         Width           =   1065
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Go!"
         Height          =   465
         Left            =   7965
         TabIndex        =   6
         Top             =   870
         Width           =   1260
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   210
         Top             =   1245
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Caption         =   "Convert"
         Height          =   450
         Left            =   180
         TabIndex        =   13
         Top             =   465
         Width           =   690
      End
      Begin VB.Label Label2 
         Caption         =   "into"
         Height          =   300
         Left            =   5475
         TabIndex        =   12
         Top             =   465
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "The result is"
         Height          =   255
         Left            =   7215
         TabIndex        =   11
         Top             =   1575
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Initialising the YCCO module"
      Height          =   1545
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   9480
      Begin VB.CheckBox Check2 
         Caption         =   "Update internal record of each conversion rate...."
         Height          =   360
         Left            =   2205
         TabIndex        =   4
         Top             =   630
         Value           =   1  'Checked
         Width           =   5085
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Enable registry cache - improves performance immensly!"
         Height          =   360
         Left            =   2205
         TabIndex        =   3
         Top             =   270
         Width           =   5085
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   2445
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   975
         Width           =   4950
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Populate the currency ComboBoxes"
         Height          =   615
         Left            =   135
         TabIndex        =   1
         ToolTipText     =   $"frmTest.frx":0000
         Top             =   345
         Width           =   1875
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
   If Check1.Value = vbChecked Then
      YCCO_EnableRegistryCache True
   Else
      YCCO_EnableRegistryCache False
   End If
End Sub

Private Sub Check2_Click()
   If Check1.Value = vbChecked Then
      Combo3.Enabled = True
      YCCO_SetInternetRefreshInterval Combo3.ItemData(Combo3.ListIndex)
   Else
      Combo3.Enabled = False
      YCCO_SetInternetRefreshInterval YCCO_EVERY_TIME
   End If
   
End Sub


Private Sub Combo3_Click()
   YCCO_SetInternetRefreshInterval Combo3.ItemData(Combo3.ListIndex)
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   Command2.Enabled = False
   Screen.MousePointer = vbArrowHourglass
   
   YCCO_PopulateCurrencySelectionControls
   
   Command1.Enabled = True
   Command2.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
   Dim conversion_value!
   
   Command1.Enabled = False
   Command2.Enabled = False
   Screen.MousePointer = vbArrowHourglass
   
   If Text1.Text = "" Then Text1.Text = "1"
   
'   Select Case YCCO_Convert(CSng(Text1.Text), conversion_value)
   Select Case YCCO_ConvertGeneric(Text1, conversion_value)
      Case YCCO_SUCCEEDED
         Text2.Text = CStr(conversion_value)
      Case YCCO_NO_DATA_AVAILABLE
         Text2.Text = "?"
         MsgBox "No conversion rate available!"
      Case YCCO_FAILED
         Text2.Text = ""
         MsgBox "Failed!"
   End Select

   Command1.Enabled = True
   Command2.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   
   YCCO_RegisterInternetTransferControl Inet1
   YCCO_RegisterCurrencySelectionControls Combo1, Combo2
   
   Combo3.AddItem "Every Day - often enough for most requirements (default)."
   Combo3.ItemData(Combo3.NewIndex) = YCCO_EVERY_DAY
   Combo3.AddItem "Every Hour - for a serious business application."
   Combo3.ItemData(Combo3.NewIndex) = YCCO_EVERY_HOUR
   Combo3.AddItem "Every Minute - are you writing software for the Stock Exchange?!"
   Combo3.ItemData(Combo3.NewIndex) = YCCO_EVERY_MINUTE
   Combo3.AddItem "Every Month - only useful for monthly accounting, etc."
   Combo3.ItemData(Combo3.NewIndex) = YCCO_EVERY_MONTH
   Combo3.ListIndex = 0
End Sub

Private Sub TabStrip1_Click()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii > Asc(" ") And KeyAscii <> Asc(".") Then
      If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
         KeyAscii = 0
      End If
   End If
End Sub

