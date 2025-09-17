' =========================================================
' PSR Automation - QA Evidence Standardization
' Author: Rodrigo Bonifacio Chagas dos Santos
' Created: 2025-09-13
' License: MIT (see LICENSE file)
' =========================================================

Option Explicit

' >>>>> CONFIGURAÇÃO <<<<<
Dim watchFolder : watchFolder = "C:\PSR_Automatico"     ' Pasta monitorada (ZIP/MHT entram aqui)
Dim processorVbs : processorVbs = watchFolder & "\PadronizarArraste_v2.vbs" ' Seu VBS atual (intocado)

' Destino final do DOCX: Desktop\Homologacao (mantém na Área de Trabalho)
Dim destDocFolder

' Limpeza / Arquivamento dos originais (após SUCESSO/FALHA):
' "move" -> move ZIP e MHT para subpastas \Processados (com histórico)
' "delete"-> apaga ZIP e MHT (sem histórico)
Const POST_ACTION_FOR_SOURCES = "move"   ' "move" ou "delete"

' Notificações:
Const ENABLE_BEEP = True                 ' Beep do sistema
Const ENABLE_TTS = False                 ' Fala "Evidência pronta"
Const ENABLE_POPUP = False               ' Popup 3s
Const POPUP_TIMEOUT_SEC = 3
Const TTS_TEXT = "Evidência pronta"

' Controles:
Const PROCESS_ZIP = True                 ' Processar ZIP automaticamente
Const PROCESS_MHT_DIRECT = True          ' Processar MHT colocado diretamente
Const RECENT_WINDOW_SEC = 15             ' Cooldown p/ mesmo MHT
Const EXTRACT_TIME_SLACK = 5             ' Tolerância seleção do MHT pós-extração (s)
Const LOCK_STALE_MIN = 10                ' Lock obsoleto após X minutos

' Fallback recursivo (para MHTs em subpastas após extração do ZIP):
Const ENABLE_RECURSIVE_FALLBACK = True   ' Ativado por padrão
Const RECURSIVE_MAX_DEPTH = 2            ' Profundidade máxima ao procurar MHT

' Log:
Const MAX_LOG_KB = 2048                  ' ~2MB antes de rotacionar

' >>>>> FIM CONFIG <<<<<

Dim fso, shell, wsh
Set fso   = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("Shell.Application")
Set wsh   = CreateObject("WScript.Shell")

' ========= Instância Única (via WMI) =========
If IsAnotherInstanceRunning() Then
  WScript.Quit 0
End If

' Pasta de destino: Desktop\Homologacao (independe de idioma/redirecionamento)
destDocFolder = wsh.SpecialFolders("Desktop") & "\Homologacao"

' Subpastas de arquivamento (se POST_ACTION_FOR_SOURCES="move")
Dim processedRoot : processedRoot      = watchFolder & "\Processados"
Dim processedZipFolder : processedZipFolder = processedRoot & "\ZIP"
Dim processedMhtFolder : processedMhtFolder = processedRoot & "\MHT"

Dim logFile, loopDelaySec
logFile = watchFolder & "\AutomacaoPSR.log"
loopDelaySec = 2 ' segundos

Dim recent : Set recent = CreateObject("Scripting.Dictionary") ' cooldown
Dim lockFile    : lockFile    = watchFolder & "\processing.lock"
Dim watcherLock : watcherLock = watchFolder & "\watcher.lock"  ' heartbeat do watcher

' ========= Logging =========
Sub Log(msg)
  On Error Resume Next
  RotateLogIfNeeded
  Dim ts : Set ts = fso.OpenTextFile(logFile, 8, True) ' Append
  ts.WriteLine "[" & Now & "] " & msg
  ts.Close
  On Error GoTo 0
End Sub

Sub RotateLogIfNeeded()
  On Error Resume Next
  If fso.FileExists(logFile) Then
    Dim kb : kb = CLng(fso.GetFile(logFile).Size \ 1024)
    If kb >= MAX_LOG_KB Then
      Dim rotated : rotated = watchFolder & "\AutomacaoPSR_" & Timestamp() & ".log"
      fso.MoveFile logFile, rotated
    End If
  End If
  On Error GoTo 0
End Sub

' ========= Pré-checagens =========
If Not fso.FolderExists(watchFolder) Then
  MsgBox "Pasta não existe: " & watchFolder, vbCritical, "WatchPSR"
  WScript.Quit 1
End If
If Not fso.FileExists(processorVbs) Then
  MsgBox "VBS não encontrado: " & processorVbs, vbCritical, "WatchPSR"
  WScript.Quit 1
End If
Call EnsureFolderExists(destDocFolder)
If LCase(POST_ACTION_FOR_SOURCES) = "move" Then
  Call EnsureFolderExists(processedZipFolder)
  Call EnsureFolderExists(processedMhtFolder)
End If

Log "Iniciando monitoramento em: " & watchFolder
Log "Usando VBS: " & processorVbs
Log "Pasta de destino (DOCX): " & destDocFolder
Log "Pós-ação p/ originais: " & POST_ACTION_FOR_SOURCES

' ========= Utilidades =========
Function ExtLower(fn)
  ExtLower = LCase(fso.GetExtensionName(fn))
End Function

Function DonePath(filePath)
  DonePath = filePath & ".done"
End Function

Function NeedsProcess(filePath)
  On Error Resume Next
  Dim dp : dp = DonePath(filePath)
  If Not fso.FileExists(dp) Then NeedsProcess = True : Exit Function
  Dim f : Set f = fso.GetFile(filePath)
  Dim d : Set d = fso.GetFile(dp)
  NeedsProcess = (f.DateLastModified > d.DateLastModified)
  On Error GoTo 0
End Function

Sub TouchDone(filePath)
  On Error Resume Next
  Dim dp : dp = DonePath(filePath)
  Dim ts : Set ts = fso.OpenTextFile(dp, 2, True) ' ForWriting (cria)
  ts.WriteLine "done: " & Now
  ts.Close
  On Error GoTo 0
End Sub

Sub DeleteDoneIfExists(filePath)
  On Error Resume Next
  Dim dp : dp = DonePath(filePath)
  If fso.FileExists(dp) Then fso.DeleteFile dp, True
  On Error GoTo 0
End Sub

Sub EnsureFolderExists(path)
  On Error Resume Next
  If fso.FolderExists(path) Then Exit Sub
  Dim parts, i, cur
  parts = Split(path, "\")
  cur = parts(0) ' "C:"
  For i = 1 To UBound(parts)
    cur = cur & "\" & parts(i)
    If Not fso.FolderExists(cur) Then fso.CreateFolder cur
  Next
  On Error GoTo 0
End Sub

Function BaseNameNoExt(p)
  Dim nm, dotPos
  nm = fso.GetFileName(p)
  dotPos = InStrRev(nm, ".")
  If dotPos > 0 Then BaseNameNoExt = Left(nm, dotPos - 1) Else BaseNameNoExt = nm
End Function

Function Timestamp()
  Dim yr, mo, dy, hr, mn, sc
  yr = CStr(Year(Now))
  mo = Right("0" & CStr(Month(Now)), 2)
  dy = Right("0" & CStr(Day(Now)), 2)
  hr = Right("0" & CStr(Hour(Now)), 2)
  mn = Right("0" & CStr(Minute(Now)), 2)
  sc = Right("0" & CStr(Second(Now)), 2)
  Timestamp = yr & mo & dy & "_" & hr & mn & sc
End Function

' ---- Cooldown ----
Function IsRecentlyProcessed(mhtPath)
  On Error Resume Next
  If recent.Exists(mhtPath) Then
    If DateDiff("s", recent(mhtPath), Now) < RECENT_WINDOW_SEC Then
      IsRecentlyProcessed = True
      Exit Function
    Else
      recent.Remove mhtPath
    End If
  End If
  IsRecentlyProcessed = False
  On Error GoTo 0
End Function

Sub MarkRecentlyProcessed(mhtPath)
  On Error Resume Next
  recent(mhtPath) = Now
  On Error GoTo 0
End Sub

' ---- Lock (arquivo) para processamento unitário de itens ----
Function LockIsStale()
  On Error Resume Next
  LockIsStale = False
  If fso.FileExists(lockFile) Then
    Dim dt : dt = fso.GetFile(lockFile).DateLastModified
    If DateDiff("n", dt, Now) > LOCK_STALE_MIN Then LockIsStale = True
  End If
  On Error GoTo 0
End Function

Function AcquireLock(tag)
  On Error Resume Next
  AcquireLock = False
  If fso.FileExists(lockFile) Then
    If LockIsStale() Then
      fso.DeleteFile lockFile, True
      Log "Lock obsoleto removido."
    Else
      Log "Lock ativo, ocupando (" & tag & ")."
      Exit Function
    End If
  End If
  Dim ts : Set ts = fso.CreateTextFile(lockFile, True)
  ts.WriteLine "lock: " & Now & " " & vbCrLf & " " & tag
  ts.Close
  AcquireLock = True
  On Error GoTo 0
End Function

Sub ReleaseLock()
  On Error Resume Next
  If fso.FileExists(lockFile) Then fso.DeleteFile lockFile, True
  On Error GoTo 0
End Sub

' ---- Avisos ----
Sub NotifyOk(filePath)
  On Error Resume Next
  If ENABLE_BEEP Then wsh.Run "rundll32 user32.dll,MessageBeep", 0, False
  If ENABLE_TTS Then
    Dim v : Set v = CreateObject("SAPI.SpVoice")
    v.Speak TTS_TEXT, 1 ' async
  End If
  If ENABLE_POPUP Then wsh.Popup "Processado: " & fso.GetFileName(filePath), POPUP_TIMEOUT_SEC, "WatchPSR", 64
  On Error GoTo 0
End Sub

' ---- ZIP helpers ----
Function FileStable(zipPath, maxWaitSeconds)
  On Error Resume Next
  Dim stable, prevSize, start, currSize
  stable = False : prevSize = -1 : start = Timer
  Do While (Timer - start) < maxWaitSeconds
    If fso.FileExists(zipPath) Then
      currSize = fso.GetFile(zipPath).Size
      If currSize > 0 And currSize = prevSize Then
        WScript.Sleep 3000
        If fso.FileExists(zipPath) Then
          If fso.GetFile(zipPath).Size = currSize Then stable = True : Exit Do
        End If
      End If
      prevSize = currSize
    End If
    WScript.Sleep 500
  Loop
  FileStable = stable
  On Error GoTo 0
End Function

Sub ExtractZip(zipPath, destFolder)
  On Error Resume Next
  Dim destNS, srcNS
  Set destNS = shell.NameSpace(destFolder)
  Set srcNS  = shell.NameSpace(zipPath)
  If (destNS Is Nothing) Or (srcNS Is Nothing) Then
    Log "Falha ao preparar extração (Shell.NameSpace): " & zipPath
  Else
    destNS.CopyHere srcNS.Items, 16 ' 16 = sem UI
    WScript.Sleep 2000
  End If
  On Error GoTo 0
End Sub

' ---- Seleciona DOCX gerado a partir do nome base do MHT ----
Function FindLatestRedimensionadoDoc(mhtPath)
  On Error Resume Next
  Dim mhtDir, base, folder, f, bestPath, bestDt
  mhtDir = fso.GetParentFolderName(mhtPath)
  base = BaseNameNoExt(mhtPath)
  bestPath = ""
  bestDt = DateSerial(1900,1,1)
  If fso.FolderExists(mhtDir) Then
    Set folder = fso.GetFolder(mhtDir)
    For Each f In folder.Files
      If LCase(ExtLower(f.Name)) = "docx" Then
        If Left(LCase(fso.GetBaseName(f.Name)), Len(LCase(base & "_redimensionado"))) = LCase(base & "_redimensionado") Then
          If bestPath = "" Or f.DateLastModified > bestDt Then bestPath = f.Path : bestDt = f.DateLastModified
        End If
      End If
    Next
  End If
  If bestPath = "" Then
    Dim tmp, tmpFolder
    tmp = wsh.ExpandEnvironmentStrings("%TEMP%")
    If fso.FolderExists(tmp) Then
      Set tmpFolder = fso.GetFolder(tmp)
      For Each f In tmpFolder.Files
        If LCase(ExtLower(f.Name)) = "docx" Then
          If Left(LCase(fso.GetBaseName(f.Name)), Len(LCase(base & "_redimensionado"))) = LCase(base & "_redimensionado") Then
            If bestPath = "" Or f.DateLastModified > bestDt Then bestPath = f.Path : bestDt = f.DateLastModified
          End If
        End If
      Next
    End If
  End If
  FindLatestRedimensionadoDoc = bestPath
  On Error GoTo 0
End Function

' ---- Move/Copy/Delete seguros ----
Function MoveFileSafe(srcPath, destFolder)
  On Error Resume Next
  Call EnsureFolderExists(destFolder)
  Dim fileName, destPath
  fileName = fso.GetFileName(srcPath)
  destPath = destFolder & "\" & fileName
  If fso.FileExists(destPath) Then
    Dim base, ext
    base = BaseNameNoExt(fileName)
    ext  = fso.GetExtensionName(fileName)
    destPath = destFolder & "\" & base & "_" & Timestamp() & "." & ext
  End If
  Err.Clear
  fso.MoveFile srcPath, destPath
  If Err.Number <> 0 Then
    Err.Clear
    fso.CopyFile srcPath, destPath, True
    If Err.Number = 0 Then
      On Error Resume Next
      fso.DeleteFile srcPath, True
      On Error GoTo 0
    Else
      Log "Falha ao mover/copiar '" & srcPath & "' para '" & destPath & "': " & Err.Description
    End If
  End If
  MoveFileSafe = destPath
End Function

Sub DeleteFileSafe(path)
  On Error Resume Next
  If fso.FileExists(path) Then fso.DeleteFile path, True
  On Error GoTo 0
End Sub

' ---- Procura MHT apenas na RAIZ (sem recursão) ----
Function FindNewestMhtInRootSince(startFolder, minDt)
  On Error Resume Next
  Dim folder, file, bestPath, bestDt
  Set folder = fso.GetFolder(startFolder)
  bestPath = "" : bestDt = DateSerial(1900,1,1)
  For Each file In folder.Files
    Dim e : e = ExtLower(file.Name)
    If (e = "mht" Or e = "mhtml") Then
      If DateDiff("s", minDt, file.DateLastModified) >= 0 Then
        If file.DateLastModified > bestDt Then
          bestDt = file.DateLastModified
          bestPath = file.Path
        End If
      End If
    End If
  Next
  FindNewestMhtInRootSince = bestPath
  On Error GoTo 0
End Function

' ---- Fallback: busca MHT de forma recursiva (profundidade limitada)
Function FindNewestMhtRecursiveSince(startFolder, minDt, maxDepth)
  On Error Resume Next
  Dim bestPath, bestDt
  bestPath = "" : bestDt = DateSerial(1900,1,1)
  RecurseMhtSearch startFolder, 0, maxDepth, minDt, bestPath, bestDt
  FindNewestMhtRecursiveSince = bestPath
  On Error GoTo 0
End Function

' (CORREÇÃO) Sub separado: sem aninhar em função (VBScript não permite)
Sub RecurseMhtSearch(curFolder, depth, maxDepth, minDt, ByRef bestPath, ByRef bestDt)
  On Error Resume Next
  If depth > maxDepth Then Exit Sub
  Dim fld, f, subf, e
  If Not fso.FolderExists(curFolder) Then Exit Sub
  Set fld = fso.GetFolder(curFolder)

  For Each f In fld.Files
    e = LCase(fso.GetExtensionName(f.Name))
    If (e = "mht" Or e = "mhtml") Then
      If DateDiff("s", minDt, f.DateLastModified) >= 0 Then
        If f.DateLastModified > bestDt Then
          bestDt = f.DateLastModified
          bestPath = f.Path
        End If
      End If
    End If
  Next
  For Each subf In fld.SubFolders
    RecurseMhtSearch subf.Path, depth + 1, maxDepth, minDt, bestPath, bestDt
  Next
  On Error GoTo 0
End Sub

' ---- Subpasta por data (YYYY-MM-DD) com base na data do arquivo ----
Function DateFolderFromFile(fp)
  On Error Resume Next
  Dim d : d = fso.GetFile(fp).DateLastModified
  DateFolderFromFile = Year(d) & "-" & Right("0" & Month(d), 2) & "-" & Right("0" & Day(d), 2)
  On Error GoTo 0
End Function

' ---- Heartbeat do watcher ----
Sub UpdateWatcherHeartbeat()
  On Error Resume Next
  Dim ts : Set ts = fso.OpenTextFile(watcherLock, 2, True) ' ForWriting, cria se não existir
  ts.WriteLine "watcher: " & Now & " (" & WScript.FullName & " " & WScript.ScriptName & ")"
  ts.Close
  On Error GoTo 0
End Sub

' ---- Checagem de instância única via WMI ----
Function IsAnotherInstanceRunning()
  On Error Resume Next
  Dim mePath : mePath = LCase(WScript.ScriptFullName)
  Dim svc : Set svc = GetObject("winmgmts:\\.\root\cimv2")
  If Err.Number <> 0 Or (svc Is Nothing) Then
    Err.Clear
    IsAnotherInstanceRunning = False
    Exit Function
  End If
  Dim procs, p, cnt : cnt = 0
  Set procs = svc.ExecQuery("SELECT ProcessId,Name,CommandLine FROM Win32_Process WHERE Name='wscript.exe' OR Name='cscript.exe'")
  For Each p In procs
    If InStr(LCase("" & p.CommandLine), mePath) > 0 Then
      cnt = cnt + 1
      If cnt > 1 Then Exit For
    End If
  Next
  If cnt > 1 Then
    On Error Resume Next
    Dim lf : lf = logFile ' usa o mesmo log configurado
    Dim ts
    Set ts = CreateObject("Scripting.FileSystemObject").OpenTextFile(lf, 8, True)
    ts.WriteLine "[" & Now & "] " & "Outra instância do WatchPSR já está em execução. Encerrando esta."
    ts.Close
    On Error GoTo 0
    IsAnotherInstanceRunning = True
  Else
    IsAnotherInstanceRunning = False
  End If
End Function

' ========= Núcleo =========
Sub ProcessAndMove(mhtPath, originTag, successOut)
  On Error Resume Next
  successOut = False

  ' ---- LOCK por item ----
  If Not AcquireLock(originTag & " -> " & mhtPath) Then Exit Sub
  Log originTag & ": " & mhtPath

  If IsRecentlyProcessed(mhtPath) Then
    Log "Ignorado (recentemente processado): " & mhtPath
    ReleaseLock
    Exit Sub
  End If

  Dim rc : rc = RunProcessor(mhtPath)
  Log "VBS finalizado (ExitCode=" & rc & ") para: " & mhtPath

  ' Marca cooldown e .done do MHT
  MarkRecentlyProcessed mhtPath
  TouchDone mhtPath

  ' Procura o DOCX resultante
  Dim docxPath : docxPath = FindLatestRedimensionadoDoc(mhtPath)
  If docxPath = "" Then
    Log "ATENÇÃO: DOCX não encontrado para '" & mhtPath & "'."
    ' Mantém MHT e .done para permitir futura análise
  Else
    Dim newPath : newPath = MoveFileSafe(docxPath, destDocFolder)
    Log "DOCX movido para: " & newPath
    NotifyOk newPath
    successOut = True

    ' Limpeza da origem (MHT) após sucesso
    If LCase(POST_ACTION_FOR_SOURCES) = "move" Then
      Dim mhtDateFolder, moved
      mhtDateFolder = processedMhtFolder & "\" & DateFolderFromFile(mhtPath)
      EnsureFolderExists mhtDateFolder
      moved = MoveFileSafe(mhtPath, mhtDateFolder)
      Log "MHT arquivado em: " & moved
    ElseIf LCase(POST_ACTION_FOR_SOURCES) = "delete" Then
      DeleteFileSafe mhtPath
      Log "MHT apagado: " & mhtPath
    End If

    ' Remove o .done do MHT (não é mais necessário)
    DeleteDoneIfExists mhtPath
  End If

  ReleaseLock
  On Error GoTo 0
End Sub

Function RunProcessor(mhtPath)
  On Error Resume Next
  Dim sh, cmd, rc
  Set sh = CreateObject("WScript.Shell")
  cmd = "cscript //nologo """ & processorVbs & """ """ & mhtPath & """"
  rc = sh.Run(cmd, 0, True) ' 0=oculto, True=aguarda
  RunProcessor = rc
  On Error GoTo 0
End Function
' ==== Loop principal ====
Do
  On Error Resume Next
  ' Heartbeat do watcher (diagnóstico)
  UpdateWatcherHeartbeat()

  Dim folder, file, ext, zipPath, mhtPath
  Set folder = fso.GetFolder(watchFolder)

  For Each file In folder.Files
    ext = ExtLower(file.Name)

    ' -------- ZIP --------
    If PROCESS_ZIP And ext = "zip" Then
      zipPath = file.Path
      If NeedsProcess(zipPath) Then
        If FileStable(zipPath, 60) Then
          Log "ZIP estável: " & zipPath
          Dim t0 : t0 = Now
          ExtractZip zipPath, watchFolder
          Log "ZIP extraído em: " & watchFolder
          Dim minDt : minDt = DateAdd("s", -EXTRACT_TIME_SLACK, t0)

          ' Preferência: busca de MHT na raiz
          Dim mhtFromZip : mhtFromZip = FindNewestMhtInRootSince(watchFolder, minDt)

          ' Fallback recursivo opcional (subpastas)
          If mhtFromZip = "" And ENABLE_RECURSIVE_FALLBACK Then
            mhtFromZip = FindNewestMhtRecursiveSince(watchFolder, minDt, RECURSIVE_MAX_DEPTH)
          End If

          Dim ok : ok = False
          If mhtFromZip <> "" Then
            ProcessAndMove mhtFromZip, "MHT extraído", ok
          Else
            Log "Nenhum .MHT/.MHTML novo encontrado após extrair: " & zipPath
          End If

          ' Marca o ZIP como processado (para evitar loop)
          TouchDone zipPath

          ' NOVO: arquiva ZIP SEMPRE (separa OK/Falha)
          Dim zipBaseFolder, movedZip, statusTag
          If ok = True Then
            zipBaseFolder = processedZipFolder & "\" & DateFolderFromFile(zipPath)
            statusTag = " (OK)"
          Else
            zipBaseFolder = processedZipFolder & "\Falha\" & DateFolderFromFile(zipPath)
            statusTag = " (FALHA)"
          End If
          EnsureFolderExists zipBaseFolder

          If LCase(POST_ACTION_FOR_SOURCES) = "move" Then
            movedZip = MoveFileSafe(zipPath, zipBaseFolder)
            Log "ZIP arquivado em: " & movedZip & statusTag
          ElseIf LCase(POST_ACTION_FOR_SOURCES) = "delete" Then
            DeleteFileSafe zipPath
            Log "ZIP apagado: " & zipPath & statusTag
          End If
          DeleteDoneIfExists zipPath

        Else
          Log "Arquivo ZIP não estabilizou a tempo: " & zipPath
          ' Mantém ZIP e .done (se criado) para análise/manual
        End If
      End If

    ' -------- MHT/MHTML DIRETO --------
    ElseIf PROCESS_MHT_DIRECT And (ext = "mht" Or ext = "mhtml") Then
      mhtPath = file.Path
      If NeedsProcess(mhtPath) Then
        If Not IsRecentlyProcessed(mhtPath) Then
          Dim ok2 : ok2 = False
          ProcessAndMove mhtPath, "MHT direto", ok2
          ' Se ok2=True, já moveu/apagou MHT e limpou o .done
        Else
          Log "Ignorado (recentemente processado): " & mhtPath
        End If
      End If
    End If
  Next
  On Error GoTo 0
  WScript.Sleep loopDelaySec * 1000
Loop
