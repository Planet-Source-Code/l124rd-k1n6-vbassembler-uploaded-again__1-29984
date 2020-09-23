Attribute VB_Name = "basAsm"
Option Explicit

'API's used for create a MSDOS process and wait for the endingd of this process
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Type FormState
    Deleted As Integer
    Dirty As Integer
    Color As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

Public Const TITULO$ = "VBEnsamblador 8086 MASM 5.10 - "
Public Const SINTITULO$ = "Sin titulo"
Public Const KEYEVENTF_KEYUP = &H2
Public FState As FormState

Const RUTA_COMP$ = "\COMP\MASM.EXE"
Const RUTA_LINK$ = "\COMP\LINK.EXE"
Const RUTA_CODE$ = "\ASMCodes"

Sub ASMValidaEstructura()
'Objetive: Validates structure of directories and files needed
'Author: Héctor Raúl González Juárez

   If Dir(App.Path & "\COMP", vbDirectory) = "" Then
      MkDir App.Path & "\COMP"
   End If
   
   If Dir(App.Path & "\ASMCodes", vbDirectory) = "" Then
      MkDir App.Path & "\ASMCodes"
   End If
   
   If Dir(App.Path & RUTA_COMP, vbArchive) = "" Then
      MsgBox "File MASM.EXE not found!", vbCritical + vbOKOnly, "VB-Assembler"
   End If

   If Dir(App.Path & RUTA_LINK, vbArchive) = "" Then
      MsgBox "File LINK.EXE not found!", vbCritical + vbOKOnly, "VB-Assembler"
   End If

End Sub

Sub ASMAbreArchivo()
'Objetive: Opens file
'Author: Héctor Raúl González Juárez
   
   Dim sCad As String

   frmAsm.cdgOpen.Filter = "*.asm"
   frmAsm.cdgOpen.DialogTitle = "Open *.ASM"
   frmAsm.cdgOpen.InitDir = App.Path & "\" & RUTA_CODE
   frmAsm.cdgOpen.FileName = "*.asm"
   frmAsm.cdgOpen.Action = 1
   
   If Not InStr(frmAsm.cdgOpen.FileName, "*") > 0 Then
      Call ASMOpenFile(frmAsm.cdgOpen.FileName)
   End If
   
End Sub

Sub ASMOpenFile(sNomArchivo As String)
'Objetive: Opens file and shows it on TextBox
'Author: Héctor Raúl González Juárez

   Dim fIndex     As Integer
   Dim iPuntero   As Integer
   Dim sArchivo   As String

   On Error Resume Next
    
   Open sNomArchivo For Input As #1
      If Err Then
         MsgBox "File no exists: " + sNomArchivo, vbCritical, "VB-Assembler"
         Exit Sub
      End If
   
      Screen.MousePointer = 11
      
      sArchivo = Input(LOF(1), 1)
      FState.Dirty = False
   Close #1
   
   iPuntero = InStr(1, sArchivo, ".CODE")
   
   If iPuntero > 0 Then
      frmAsm.TxtHeader.Text = Trim$(Mid$(sArchivo, 1, iPuntero + 6))
      frmAsm.TxtCode.Text = Trim$(Mid$(sArchivo, iPuntero + 7, Len(sArchivo) - (iPuntero + 8)))
      frmAsm.Caption = TITULO$ & frmAsm.cdgOpen.FileTitle
   Else
      frmAsm.TxtCode.Text = Trim$(sArchivo)
   End If
   Screen.MousePointer = 0
End Sub

Sub ASMNuevoArchivo()
'Objetive: New File
'Author: Héctor Raúl González Juárez
   
   Dim sHeaderDefault   As String
   Dim sCodeDefault     As String
   
   sHeaderDefault = "DOSSEG" & vbNewLine & ".MODEL SMALL" & vbNewLine & _
                    ".STACK 100h" & vbNewLine & ".Data" & vbNewLine & _
                    "temp DB 10 DUP(?) ; This is an internal variable" & vbNewLine & _
                    "Msg  DB 10,13,'MENSAJE PERSONALIZADO',10,13,'$'" & vbNewLine & _
                    ".CODE" & vbNewLine
   
   sCodeDefault = ";/Guarda datos" & vbNewLine & "MOV AX, @DATA" & vbNewLine & "MOV DS, AX" & _
                  vbNewLine & vbNewLine & ";/Código" & vbNewLine & vbNewLine & vbNewLine & _
                  vbNewLine & ";/Fin del Código" & vbNewLine & "EndProgram:" & vbNewLine & _
                  "MOV AH, 4CH" & vbNewLine & "INT 21H" & vbNewLine & "End" & vbNewLine & _
                  ";/Fin del Programa"
   
   frmAsm.Caption = TITULO & SINTITULO
   frmAsm.TxtHeader = sHeaderDefault
   frmAsm.TxtCode = sCodeDefault
   frmAsm.cdgOpen.FileName = ""
   
End Sub

Sub ASMSalir()
'Objetive: Asks if you wish to save before exit
'Author: Héctor Raúl González Juárez

   Dim iRes As Integer
   
   iRes = MsgBox("¿Save?", vbYesNoCancel, "VB-Assembler")
   Select Case iRes
      Case vbYes
         ASMGuardaArchivo
         End
      Case vbNo
         End
      Case vbCancel
         Exit Sub
   End Select
   
End Sub

Function ASMGuardaArchivo() As Integer
'Objetive: Saves file
'Author: Héctor Raúl González Juárez

    Dim sFilename As String

   If frmAsm.Caption = TITULO & SINTITULO Then
      sFilename = ASMGetFileName(sFilename)
   Else
      If InStr(frmAsm.cdgOpen.FileTitle, ".") > 0 Then
         sFilename = frmAsm.cdgOpen.FileTitle
      Else
         sFilename = frmAsm.cdgOpen.FileTitle & ".ASM"
      End If
   End If
   
   If sFilename <> "" Then
      ASMSaveFileAs sFilename
      ASMGuardaArchivo = True
   Else
      ASMGuardaArchivo = False
   End If
    
End Function

Function ASMGetFileName(FileName As Variant)
'Objetive: Gets file name
'Author: Héctor Raúl González Juárez
   On Error Resume Next
   
   Dim byLenNom As Byte
   
   frmAsm.cdgOpen.FileName = FileName
   frmAsm.cdgOpen.Filter = "*.asm"
   
   Do
      MsgBox "Name of file 8 characters maximum", vbInformation + vbOKOnly, "VB-Assembler"
      frmAsm.cdgOpen.ShowSave
      If Not InStr(frmAsm.cdgOpen.FileTitle, ".") > 0 Then byLenNom = Len(frmAsm.cdgOpen.FileTitle)
   Loop While InStr(frmAsm.cdgOpen.FileTitle, ".") > 9 Or byLenNom > 8
      
   If Err <> 32755 Then
      If InStr(frmAsm.cdgOpen.FileName, ".") > 0 Then
         ASMGetFileName = frmAsm.cdgOpen.FileName
      Else
         ASMGetFileName = frmAsm.cdgOpen.FileName & ".ASM"
         frmAsm.cdgOpen.FileName = ASMGetFileName
         frmAsm.cdgOpen.FileTitle = frmAsm.cdgOpen.FileTitle & ".ASM"
      End If
   Else
      ASMGetFileName = ""
   End If
   
End Function

Sub ASMSaveFileAs(FileName)
'Objetive: Shows Dialogbox Save As...
'Author: Héctor Raúl González Juárez

   On Error Resume Next
   Dim sContents As String
   
   Open FileName For Output As #1
      sContents = frmAsm.TxtHeader.Text & frmAsm.TxtCode.Text
      Screen.MousePointer = 11
   Print #1, sContents
   Close #1
   
   Screen.MousePointer = 0
   
   If Err Then
      MsgBox Error, 48, App.Title
   Else
      If InStr(frmAsm.cdgOpen.FileTitle, ".") > 0 Then
         frmAsm.Caption = TITULO$ & frmAsm.cdgOpen.FileTitle
      Else
         frmAsm.Caption = TITULO$ & frmAsm.cdgOpen.FileTitle & ".ASM"
      End If
      FState.Dirty = False
   End If
   
End Sub

Sub ASMCompila()
'Objetive: Compiles .ASM and generates .OBG Files
'Author: Héctor Raúl González Juárez
   
   Dim sNomArchivo   As String
   
   If frmAsm.Caption = TITULO & SINTITULO Then Exit Sub
   MsgBox "Press ENTER 3 times to generate OBJ on Shell session and close MSDOS Window", vbInformation + vbOKOnly, "VB-Assembler"
   sNomArchivo = sASMObtenNomArc
   ASMGuardaArchivo
   Call ASMBorra(sNomArchivo, "OBJ")
   Call ASMExecCmd(App.Path & RUTA_COMP & " " & sNomArchivo)
   Call ASMCambiaExtension(sNomArchivo, "OBJ")
   If Not bASMExisteArchivo(sNomArchivo) Then
      MsgBox "Error on code", vbCritical + vbOKOnly, "VB-Assembler"
   Else
      MsgBox "File OBJ generated", vbExclamation + vbOKOnly, "VB-Assembler"
   End If
   
End Sub

Sub ASMEjecutable()
'Objetive: Generates .EXE File from .OBJ generated in rutine ASMCompila()
'Author: Héctor Raúl González Juárez
   
   Dim sNomArchivo   As String
   
   If frmAsm.Caption = TITULO & SINTITULO Then Exit Sub
   MsgBox "Press ENTER 3 times to generate EXE on Shell session and close MSDOS Window", vbInformation + vbOKOnly, "VB-Assembler"
   sNomArchivo = sASMObtenNomArc
   ASMGuardaArchivo
   Call ASMBorra(sNomArchivo, "EXE")
   Call ASMCambiaExtension(sNomArchivo, "OBJ")
   Call ASMExecCmd(App.Path & RUTA_LINK & " " & sNomArchivo)
   Call ASMCambiaExtension(sNomArchivo, "EXE")
   If Not bASMExisteArchivo(sNomArchivo) Then
      MsgBox "Error on code", vbCritical + vbOKOnly, "VB-Assembler"
   Else
      MsgBox "File EXE generated", vbExclamation + vbOKOnly, "VB-Assembler"
   End If
   
End Sub

Sub ASMCompilaEjec()
'Objetive: Generates .OBJ and .EXE in one step
'Author: Héctor Raúl González Juárez

   Dim sNomArchivo   As String
   
   If frmAsm.Caption = TITULO & SINTITULO Then Exit Sub
   
   sNomArchivo = sASMObtenNomArc
   ASMGuardaArchivo
   Call ASMBorra(sNomArchivo, "OBJ")
   Call ASMExecCmd(App.Path & RUTA_COMP & " " & sNomArchivo)
   Call ASMCambiaExtension(sNomArchivo, "OBJ")
   
   If Not bASMExisteArchivo(sNomArchivo) Then
      MsgBox "Error on code", vbCritical + vbOKOnly, "VB-Assembler"
   Else
      sNomArchivo = sASMObtenNomArc
      ASMGuardaArchivo
      Call ASMBorra(sNomArchivo, "EXE")
      Call ASMCambiaExtension(sNomArchivo, "OBJ")
      Call ASMExecCmd(App.Path & RUTA_LINK & " " & sNomArchivo)
      Call ASMCambiaExtension(sNomArchivo, "EXE")
      If Not bASMExisteArchivo(sNomArchivo) Then
         MsgBox "Error on code", vbCritical + vbOKOnly, "VB-Assembler"
      Else
         MsgBox "Archivo EXE generado", vbExclamation + vbOKOnly, "VB-Assembler"
      End If
   End If

End Sub

Sub ASMEjecuta()
'Objetive: Executes EXE file previously generated
'Author: Héctor Raúl González Juárez
         
   Dim sNomArchivo   As String
         
   If frmAsm.Caption = TITULO & SINTITULO Then Exit Sub
   
   sNomArchivo = sASMObtenNomArc
   Call ASMCambiaExtension(sNomArchivo, "EXE")
   If Not bASMExisteArchivo(sNomArchivo) Then
      MsgBox "EXE has not been generated", vbCritical + vbOKOnly, "VB-Assembler"
   Else
      Call ASMExecCmd(App.Path & RUTA_CODE & "\" & sNomArchivo)
   End If
   
End Sub

Sub ASMCambiaExtension(ByRef sNomArchivo As String, sNvaExt As String)
'Objetive: Changes the file extension
'Author: Héctor Raúl González Juárez

   If InStr(sNomArchivo, ".") > 0 Then
      sNomArchivo = Trim$(Mid$(sNomArchivo, 1, InStr(sNomArchivo, ".")) & sNvaExt)
   Else
      sNomArchivo = sNomArchivo & "." & sNvaExt
   End If

End Sub

Function sASMObtenNomArc() As String
'Objetive: Gets the name of the actual file
'Author: Héctor Raúl González Juárez
   
   sASMObtenNomArc = ""
   
   If Not Trim$(frmAsm.cdgOpen.FileTitle) = "" Then
      sASMObtenNomArc = Trim$(frmAsm.cdgOpen.FileTitle)
   Else
      sASMObtenNomArc = Trim$(Mid$(frmAsm.Caption, InStr(sASMObtenNomArc, "-") + 1))
   End If
   
End Function

Sub ASMBorra(ByVal sNomObj As String, sExtension As String)
'Objetive: Erases file .OBJ
'Author: Héctor Raúl González Juárez
   
   sNomObj = Trim$(Mid$(sNomObj, 1, InStr(sNomObj, ".")) & sExtension)
   
   If bASMExisteArchivo(sNomObj) Then
      Kill App.Path & RUTA_CODE & "\" & sNomObj
   End If

End Sub

Function bASMExisteArchivo(sArchivo As String) As Boolean
'Objetive: Validate if file exists
'Author: Héctor Raúl González Juárez

   bASMExisteArchivo = False
   
   If Dir(App.Path & RUTA_CODE & "\" & sArchivo, vbArchive) <> "" Then
      bASMExisteArchivo = True
   End If

End Function

Public Sub ASMExecCmd(sCmdLine As String)
'Objetive: Executes a line on Shell
'Author: Héctor Raúl González Juárez
'Note: To create .OBJ the shell have to say -> MASM SOURCEFILE.ASM  -> so you get SOURCEFILE.OBJ
'      To create .EXE the shell have to say -> LINK SOURCEFILE.OBJ  -> (You just press ENTER key 3 times) so you get SOURCEFILE.EXE
  
   Dim tProc   As PROCESS_INFORMATION
   Dim tStart  As STARTUPINFO
   Dim lRes    As Long

   tStart.cb = Len(tStart)
      
   lRes = CreateProcessA(0&, sCmdLine, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, tStart, tProc)
   lRes = WaitForSingleObject(tProc.hProcess, INFINITE)
   lRes = CloseHandle(tProc.hProcess)
      
End Sub

Sub ASMLeeAyuda()
'Objetive: Reads help file
'Author Héctor Raúl González Juárez
   
   Dim iFreeFile As Integer
      
   If Dir(App.Path & "\VBAHelp.TXT", vbArchive) <> "" Then
      iFreeFile = FreeFile
      
      Open App.Path & "\VBAHelp.TXT" For Input As #iFreeFile
         If Err Then
            MsgBox "The file: VBAHelp.TXT is corrupted", vbCritical + vbOKOnly, "VB-Assembler"
            Exit Sub
         End If
         Screen.MousePointer = 11
         FrmAyuda.TxtHelp.Text = Input(LOF(1), 1)
      Close #iFreeFile
      
      Screen.MousePointer = 0
   Else
      MsgBox "The file: VBAHelp.TXT doesn't exist.", vbCritical + vbOKOnly, "VB-Assembler"
   End If
   
End Sub
