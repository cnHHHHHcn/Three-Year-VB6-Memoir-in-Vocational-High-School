VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   8400
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   7695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   3840
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long

Private Declare Function Process32First Lib "kernel32.dll" (ByVal hSnapshot As Long, leep As PROCESSENTRY32) As Long

Private Declare Function Process32Next Lib "kernel32.dll" (ByVal hSnapshot As Long, leep As PROCESSENTRY32) As Long

Private Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long

Const TH32CS_INHERIT = 2147483648#
Const TH32CS_SNPALL As Long = 19
Const TH32CS_SNAPHEAPLIST As Long = 1
Const TH32CS_SNAPMODULE As Long = 8
Const TH32CS_SNAPPROCESS As Long = 2
Const TH32CS_SNAPTHREAD As Long = 4
Const INVALID_HANDLE_VALUE = -1

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 256
End Type

Private Sub Command1_Click()
    Dim hSnapshot As Long, th32ProcessID  As Long
    Dim ProcessEntry As PROCESSENTRY32
    Dim Name As String
    hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, th32ProcessID)
    If hSnapshot > 0 Then
        ProcessEntry.dwSize = 296
        If Process32First(hSnapshot, ProcessEntry) Then
            Do
                Name = Space(256)
                Name = Left(ProcessEntry.szExeFile, InStr(1, ProcessEntry.szExeFile, Chr(0)) - 1)
                List1.AddItem Name & "  " & ProcessEntry.th32ProcessID & "  " & ProcessEntry.th32ParentProcessID & "  " & ProcessEntry.th32ModuleID & "  " & ProcessEntry.cntThreads
            Loop While Process32Next(hSnapshot, ProcessEntry)
        End If
    End If
End Sub
