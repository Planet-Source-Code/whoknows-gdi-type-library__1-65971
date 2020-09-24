VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmGDIPlus 
   Caption         =   "GDI+ Demo"
   ClientHeight    =   3168
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5148
   LinkTopic       =   "Form1"
   ScaleHeight     =   3168
   ScaleWidth      =   5148
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   372
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   972
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   1800
      Top             =   0
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.PictureBox picDisplay 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2676
      Left            =   0
      ScaleHeight     =   2676
      ScaleWidth      =   5148
      TabIndex        =   1
      Top             =   480
      Width           =   5148
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   372
      Left            =   4180
      TabIndex        =   0
      Top             =   0
      Width           =   972
   End
End
Attribute VB_Name = "frmGDIPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' http://www.syix.com/wpsjr1/index.html

' Example of using GDI+ with VB5/6
' This will work on all 32bit OS except Win95.
' Of course, gdiplus.dll must be installed.
' This dll is installed with the .net framework,
' or may be installed seperately.

' gdiplus.dll is available at:
' http://www.microsoft.com/downloads/release.asp?releaseid=32738

Private WithEvents gdip As cGdiPlus

Private Sub Form_Load()
  Set gdip = New cGdiPlus
End Sub

Private Sub cmdSave_Click()
  On Error GoTo errorhandler
  
  cdl.CancelError = True
  cdl.filename = "foo.gif"
  cdl.Filter = "Images Files (*.gif;*.jpg;*.png;*.tif)|*.gif;;*.jpg;*.png;*.tif"
  cdl.ShowSave
  
  If gdip.PictureBoxToFile(picDisplay, cdl.filename) Then
    Debug.Print "Sucessfully saved file: " & cdl.filename
  Else
    Debug.Print "Could not save file"
  End If
  
  Exit Sub
errorhandler:
  ' handle cancel silently
End Sub

Private Sub cmdOpen_Click()
  On Error GoTo errorhandler
  
  cdl.Filter = "Images Files (*.bmp;*.gif;*.jpg)|*.bmp;*.ico;*.jpg"
  cdl.CancelError = True
  cdl.ShowOpen
  
  Set picDisplay.Picture = LoadPicture(cdl.filename)
  Caption = "Displaying: " & cdl.filename
  
  Exit Sub
errorhandler:
  ' handle cancel silently
End Sub

Private Sub gdip_Error(ByVal lGdiError As Long, ByVal sErrorDesc As String)
  Debug.Print "A GDI+ Error has occured, Error Number: " & lGdiError & "   Error Description: " & sErrorDesc
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set gdip = Nothing
End Sub
