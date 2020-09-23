VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAsm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "VBEnsamblador MASM 5.10 - Sin titulo"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7920
   Icon            =   "frmAssembler.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdgOpen 
      Left            =   7245
      Top             =   2625
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FraCode 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   5160
      Left            =   105
      TabIndex        =   2
      Top             =   2940
      Width           =   7680
      Begin VB.TextBox TxtCode 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   4605
         Left            =   315
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "frmAssembler.frx":030A
         Top             =   315
         Width           =   7170
      End
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Header"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   2640
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   7680
      Begin VB.TextBox TxtHeader 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   2085
         Left            =   315
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "frmAssembler.frx":031E
         Top             =   315
         Width           =   7170
      End
   End
   Begin VB.Menu mnuPrincipal 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuArchivo 
         Caption         =   "New"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "Open"
         Index           =   1
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "Save"
         Index           =   2
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "Save as..."
         Index           =   3
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuArchivo 
         Caption         =   "Exit"
         Index           =   5
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuPrincipal 
      Caption         =   "B&uild"
      Index           =   1
      Begin VB.Menu mnuConstruir 
         Caption         =   "Compile"
         Index           =   0
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuConstruir 
         Caption         =   "Make EXE"
         Index           =   1
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuConstruir 
         Caption         =   "Comp + Make"
         Index           =   2
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuPrincipal 
      Caption         =   "&Execute"
      Index           =   2
      Begin VB.Menu mnuEjecuta 
         Caption         =   "Execute"
         Index           =   0
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuEjecuta 
         Caption         =   "MS-DOS"
         Index           =   1
      End
   End
   Begin VB.Menu mnuPrincipal 
      Caption         =   "Help"
      Index           =   3
      Begin VB.Menu mnuAyuda 
         Caption         =   "Help"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAyuda 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuAyuda 
         Caption         =   "About..."
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmAsm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'Objetive: Validates Structure
'Author: Héctor Raúl González Juárez
   
   frmAsm.Caption = TITULO & SINTITULO
   Call ASMValidaEstructura

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Objetive: Asks if you wish to save before exit
'Author: Héctor Raúl González Juárez

   If Not frmAsm.Caption = TITULO & SINTITULO Then
      ASMSalir
   Else
      End
   End If

End Sub

Private Sub Form_Resize()
'Objetive: Resize controls
'Author: Héctor Raúl González Juárez

   With frmAsm
      If Not .WindowState = vbMinimized Then
         If .Height <= 6360 Then .Height = 6360
         If .Width <= 5680 Then .Width = 5680
         .FraCode.Height = .Height - 3720
         .FraHeader.Width = .Width - 360
         .FraCode.Width = .Width - 360
         .TxtCode.Height = .FraCode.Height - 555
         .TxtCode.Width = .FraCode.Width - 510
         .TxtHeader.Width = .FraHeader.Width - 510
      End If
   End With

End Sub

Private Sub mnuArchivo_Click(Index As Integer)
'Objetive: Depends of the menu option selected
'Author: Héctor Raúl González Juárez

   Dim sFilename As String

   Select Case Index
      Case 0   'New
         ASMGuardaArchivo
         ASMNuevoArchivo
      Case 1   'Open
         ASMAbreArchivo
      Case 2   'Save
         ASMGuardaArchivo
      Case 3   'Save As
         sFilename = ASMGetFileName(sFilename)
         ASMGuardaArchivo
      Case 5   'Exit
         If Not frmAsm.Caption = TITULO & SINTITULO Then
            ASMSalir
         Else
            End
         End If
   End Select
   
End Sub

Private Sub mnuAyuda_Click(Index As Integer)
'Objetive: Depends of the menu option selected
'Author: Héctor Raúl González Juárez

   Select Case Index
      Case 0
         FrmAyuda.Show
      Case 2
         frmAbout.Show 1
   End Select

End Sub

Private Sub mnuConstruir_Click(Index As Integer)
'Objetive: Depends of the menu option selected (on Build Option)
'Author: Héctor Raúl González Juárez

   Select Case Index
      Case 0
         ASMCompila
      Case 1
         ASMEjecutable
      Case 2
         ASMCompilaEjec
   End Select
End Sub

Private Sub mnuEjecuta_Click(Index As Integer)
'Objetive: Opens an MSDOS Session
'Author: Héctor Raúl González Juárez
   
   Dim lVar As Long
   
   Select Case Index
      Case 0
         ASMEjecuta
      Case 1
         On Error Resume Next
         lVar = Shell("COMMAND.COM", vbNormalFocus)   'Windows 9X / ME
         lVar = Shell("CMD.COM", vbNormalFocus)       'Windows NT
         On Error GoTo 0
   End Select
   
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
'Objetive: Put cursor to 'Header' TextBox
'Author: Héctor Raúl González Juárez

   If KeyCode = vbKeyUp And TxtCode.SelStart = 0 Then
      KeyCode = 0
      TxtHeader.SetFocus
   End If

End Sub

Private Sub TxtHeader_KeyDown(KeyCode As Integer, Shift As Integer)
'Objetive: Put cursor to 'Code' TextBox
'Author: Héctor Raúl González Juárez

   If KeyCode = vbKeyDown And TxtHeader.SelStart = Len(TxtHeader.Text) Then
      KeyCode = 0
      TxtCode.SetFocus
   End If

End Sub

