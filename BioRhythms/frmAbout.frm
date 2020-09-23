VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAbout 
   Caption         =   "Intepreting Your Chart"
   ClientHeight    =   4956
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6576
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4956
   ScaleWidth      =   6576
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   372
      Left            =   2400
      TabIndex        =   1
      Top             =   4440
      Width           =   972
   End
   Begin RichTextLib.RichTextBox rtbAbout 
      Height          =   4092
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6372
      _ExtentX        =   11240
      _ExtentY        =   7218
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmAbout.frx":030A
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Plagiarize by Louis Boldt
' --------------------------------------------------------------------
Private Sub cmdOK_Click()
' --------------------------------------------------------------------
  Unload Me
  
End Sub

' --------------------------------------------------------------------
Private Sub Form_Load()
' --------------------------------------------------------------------
' Text is stored in the resource file
' Load rtf file into a temporary file
' then rtf box reads the file in
' done with file so kill it

' Thanks to Clint Lafever for the custom resource
' file code.
  
  Dim resByte() As Byte
  Dim intTempFile As Integer
  Dim strTempFile As String
  
  On Error GoTo Handle_Error
  
  strTempFile = fnGetTempPath("rtf")
  intTempFile = FreeFile
  resByte = LoadResData(101, "RTF")
  Open strTempFile For Binary Access Write As #intTempFile
  Put #1, , resByte
  Close #intTempFile
  rtbAbout.LoadFile strTempFile
  Kill strTempFile
Exit_Point:
  Exit Sub

Handle_Error:

  MsgBox "Unexpected Error Returned" & vbCrLf _
       & "Nmbr " & Err.Number & vbCrLf _
       & "Desc " & Err.Description & vbCrLf _
       & "Srce " & Err.Source & vbCrLf _
       & "Time " & Now & vbCrLf _
       & "Path " & App.Path, _
     16, "You have a Problem!"

  Resume Exit_Point
End Sub


