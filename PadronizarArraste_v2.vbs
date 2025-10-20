
' PadronizarArraste_v3_prompt_L18A12_CENTER_FIX11_ASCII_Header_Spacing_P2.vbs
' - ASCII-only (accents via ChrW)
' - Header textual fixo na 1a pagina; insere quebra de pagina apos o cabecalho (imagens so a partir da 2a pagina)
' - Centralizacao de imagens (inline: paragrafo; flutuantes: wdShapeCenter relativo as margens)
' - PACK: Normaliza espaco antes/depois de cada bloco (imagem + 1 paragrafo de texto)
' - Preserva a secao de instrucoes no final (procura por "Instru")
' - Remove todo texto literal "Previous Next " em qualquer ponto do documento

Option Explicit

' ================= Config =================
Const FORCE_INLINE = True
Const ISOLAR_PARAGRAFO = True
Const MIN_SPACE_BEFORE_CM = 0.0
Const MIN_SPACE_AFTER_CM  = 0.25

Const APPLY_WRAP_FOR_FLOATING = True
Const FLOAT_DIST_TOP_CM    = 0.2
Const FLOAT_DIST_BOTTOM_CM = 0.2
Const FLOAT_DIST_LEFT_CM   = 0.2
Const FLOAT_DIST_RIGHT_CM  = 0.2

Const PROMPT_DEFAULT = "L 18 A 12"

' PACK de espacamento
Const ENABLE_SPACING_PACK     = True
Const MAX_EMPTY_BEFORE_IMAGE  = 0   ' linhas vazias permitidas antes da imagem
Const MAX_EMPTY_AFTER_BLOCK   = 1   ' linhas vazias permitidas apos o bloco imagem+texto
Const CAPTION_SPACE_AFTER_PT  = 6   ' espaco depois do paragrafo de texto (pt)
Const CAPTION_SPACE_BEFORE_PT = 0

' ======= Word/Office constants =======
Const MsoTypePicture       = 13
Const MsoTypeLinkedPicture = 11
Const MsoTypeGroup         = 6

Const wdWrapTopBottom = 4

' WdInformation
Const wdWithInTable = 12
Const wdActiveEndPageNumber = 3

' GoTo / Alignment
Const wdGoToPage      = 1
Const wdGoToAbsolute  = 1
Const wdAlignParagraphCenter = 1

' Relative positions
Const wdRelHorizPosMargin = 0

' Shape position
Const wdShapeCenter = -999995

' WdBreakType (para InsertBreak)
Const wdPageBreak = 7

' ================= Main =================
Dim filePath, inputAll, hasL, Lcm, hasA, Acm
Dim wordApp, doc, ok

If WScript.Arguments.Count = 0 Then
  MsgBox "Arraste um arquivo DOC/DOCX/MHT sobre este .VBS.", vbInformation, "Padronizacao"
  WScript.Quit 0
End If

filePath = WScript.Arguments(0)
If WScript.Arguments.Count >= 2 Then
  inputAll = WScript.Arguments(1)
  If WScript.Arguments.Count >= 3 Then inputAll = inputAll & " " & WScript.Arguments(2)
End If

ok = ParseInput(Trim(inputAll), hasL, Lcm, hasA, Acm)
Do While Not ok
  Dim promptMsg, promptTitle
  promptMsg = "Digite as medidas (em cm) na mesma linha." & vbCrLf & _
              "Exemplos:" & vbCrLf & _
              " L 12 (somente largura)" & vbCrLf & _
              " A 7 (somente altura)" & vbCrLf & _
              " L 18 A 12 (encaixar em 18 x 12)" & vbCrLf & _
              " 18x12 (encaixar em 18 x 12)"
  promptTitle = "Base(s) + Tamanho(s)"
  inputAll = InputBox(promptMsg, promptTitle, PROMPT_DEFAULT)
  If Len(inputAll) = 0 Then WScript.Quit 0
  ok = ParseInput(inputAll, hasL, Lcm, hasA, Acm)
Loop

On Error Resume Next
Set wordApp = CreateObject("Word.Application")
If Err.Number <> 0 Then
  MsgBox "Nao foi possivel iniciar o Word. Erro: " & Err.Description, vbCritical, "Erro"
  WScript.Quit 1
End If
On Error GoTo 0

wordApp.Visible = False
wordApp.DisplayAlerts = 0

Dim Lpt, Apt
If hasL Then Lpt = wordApp.CentimetersToPoints(Lcm)
If hasA Then Apt = wordApp.CentimetersToPoints(Acm)

Set doc = OpenDocSmart(wordApp, filePath)
If doc Is Nothing Then
  wordApp.Quit
  MsgBox "Nao foi possivel abrir o arquivo.", vbCritical, "Abertura bloqueada"
  WScript.Quit 1
End If

' 1) Cabecalho na 1a pagina (limpa e insere) + quebra de pagina apos cabecalho
ReplaceFirstPageIntroWithHeader doc, wordApp

' 2) Centralizacao / redimensionamento de imagens
Dim pageWpt, pageHpt, maxWpt, maxHpt
pageWpt = doc.PageSetup.PageWidth
pageHpt = doc.PageSetup.PageHeight
maxWpt = pageWpt - (doc.PageSetup.LeftMargin + doc.PageSetup.RightMargin)
maxHpt = pageHpt - (doc.PageSetup.TopMargin + doc.PageSetup.BottomMargin)

Dim i, shp, ils, ilsNew

If FORCE_INLINE Then
  For i = doc.Shapes.Count To 1 Step -1
    Set shp = doc.Shapes(i)
    If shp.Type = MsoTypePicture Or shp.Type = MsoTypeLinkedPicture Then
      On Error Resume Next
      Set ilsNew = shp.ConvertToInlineShape
      Err.Clear
      On Error GoTo 0
    ElseIf shp.Type = MsoTypeGroup Then
      If APPLY_WRAP_FOR_FLOATING Then
        SafeWrapForShape shp, wordApp
        CenterFloatingShapeToMargins shp
      End If
    End If
  Next
Else
  If APPLY_WRAP_FOR_FLOATING Then
    For i = 1 To doc.Shapes.Count
      Set shp = doc.Shapes(i)
      If shp.Type = MsoTypePicture Or shp.Type = MsoTypeLinkedPicture Then
        SafeWrapForShape shp, wordApp
        CenterFloatingShapeToMargins shp
      ElseIf shp.Type = MsoTypeGroup Then
        SafeWrapForShape shp, wordApp
        CenterFloatingShapeToMargins shp
      End If
    Next
  End If
End If

If hasL And hasA Then
  For Each ils In doc.InlineShapes
    FitInline ils, Lpt, Apt
    ClampInline ils, maxWpt, maxHpt
    ApplySpacingAndIsolation ils, wordApp
    CenterInlineParagraph ils
  Next
Else
  For Each ils In doc.InlineShapes
    ScaleInline ils, hasL, Lpt, hasA, Apt
    ClampInline ils, maxWpt, maxHpt
    ApplySpacingAndIsolation ils, wordApp
    CenterInlineParagraph ils
  Next
End If

If hasL And hasA Then
  For Each shp In doc.Shapes
    If shp.Type = MsoTypePicture Or shp.Type = MsoTypeLinkedPicture Then
      FitShape shp, Lpt, Apt
      CenterFloatingShapeToMargins shp
    ElseIf shp.Type = MsoTypeGroup Then
      FitShape shp, Lpt, Apt
      CenterFloatingShapeToMargins shp
    End If
  Next
Else
  For Each shp In doc.Shapes
    If shp.Type = MsoTypePicture Or shp.Type = MsoTypeLinkedPicture Then
      ScaleShape shp, hasL, Lpt, hasA, Apt
      CenterFloatingShapeToMargins shp
    ElseIf shp.Type = MsoTypeGroup Then
      ScaleShape shp, hasL, Lpt, hasA, Apt
      CenterFloatingShapeToMargins shp
    End If
  Next
End If

' 3) PACK de espacamento, preservando o final (Instru...)
If ENABLE_SPACING_PACK Then
  Dim protectStart
  protectStart = GetProtectedTailStart(doc)
  If protectStart <= 0 Then protectStart = doc.Content.End
  PackSpacingAroundInlineImages doc, wordApp, protectStart
End If

' 4) Remover qualquer ocorrencia literal de "Previous Next "
RemovePreviousNextText doc

' 5) Salvar copia
Dim fso, fName, fDir, baseName, newPath, fmt
Dim yr, mo, dy, hr, mn, sc, stamp
Set fso = CreateObject("Scripting.FileSystemObject")
fName = fso.GetFileName(filePath)
fDir = fso.GetParentFolderName(filePath)
If InStrRev(fName, ".") > 0 Then
  baseName = Left(fName, InStrRev(fName, ".") - 1)
Else
  baseName = fName
End If
fmt = 16
newPath = fDir & "\" & baseName & "_redimensionado.docx"
If fso.FileExists(newPath) Then
  yr = CStr(Year(Now))
  mo = Right("0" & CStr(Month(Now)), 2)
  dy = Right("0" & CStr(Day(Now)), 2)
  hr = Right("0" & CStr(Hour(Now)), 2)
  mn = Right("0" & CStr(Minute(Now)), 2)
  sc = Right("0" & CStr(Second(Now)), 2)
  stamp = yr & mo & dy & "_" & hr & mn & sc
  newPath = fDir & "\" & baseName & "_redimensionado_" & stamp & ".docx"
End If

On Error Resume Next
doc.SaveAs2 newPath, fmt
If Err.Number <> 0 Then
  Err.Clear
  Dim shl, tmp, safePath
  Set shl = CreateObject("WScript.Shell")
  tmp = shl.ExpandEnvironmentStrings("%TEMP%")
  safePath = tmp & "\" & baseName & "_redimensionado.docx"
  doc.SaveAs2 safePath, fmt
  If Err.Number <> 0 Then
    doc.Close False
    wordApp.Quit
    MsgBox "Falha ao salvar a copia (pastas originais e TEMP). Verifique permissoes.", vbCritical, "Erro ao salvar"
    WScript.Quit 1
  Else
    newPath = safePath
  End If
End If
On Error GoTo 0

doc.Close False
wordApp.Quit
MsgBox "Pronto! Copia salva em:" & vbCrLf & newPath, vbInformation, "Concluido"

' ================= Helpers =================
Sub ReplaceFirstPageIntroWithHeader(ByRef d, ByRef app)
  On Error Resume Next
  app.Selection.GoTo wdGoToPage, wdGoToAbsolute, 1

  Dim pg
  Set pg = d.Bookmarks("\Page").Range
  If Not pg Is Nothing Then
    pg.Text = vbCr
  End If

  Dim i2, shp2, anchPg
  For i2 = d.Shapes.Count To 1 Step -1
    Set shp2 = d.Shapes(i2)
    On Error Resume Next
    anchPg = shp2.Anchor.Information(wdActiveEndPageNumber)
    On Error GoTo 0
    If anchPg = 1 Then
      If shp2.TextFrame.HasText Then shp2.Delete
    End If
  Next

  Dim a_tilde, c_ced, a_acute, i_acute
  a_tilde = ChrW(&HE3)
  c_ced   = ChrW(&HE7)
  a_acute = ChrW(&HE1)
  i_acute = ChrW(&HED)

  Dim lines(6)
  lines(0) = "Projeto: Conex" & a_tilde & "o Digital"
  lines(1) = "Frente:"
  lines(2) = "Respons" & a_acute & "vel:"
  lines(3) = "Produto/Serv" & i_acute & c_ced & "o:"
  lines(4) = "Vers" & a_tilde & "o:"
  lines(5) = "Dispositivo:"
  lines(6) = "Resultado Esperado:"

  Dim headerText
  headerText = lines(0) & vbCr & lines(1) & vbCr & lines(2) & vbCr & lines(3) & vbCr & lines(4) & vbCr & lines(5) & vbCr & lines(6) & vbCr

  ' Insere header no inicio e em seguida uma quebra de pagina (garante imagens apenas a partir da pagina 2)
  Dim rIns, rAfter
  Set rIns = d.Range(0,0)
  rIns.Text = headerText & vbCr

  Set rAfter = d.Range(Len(headerText) + 1, Len(headerText) + 1)
  rAfter.InsertBreak wdPageBreak

  ' Formata fonte do bloco inicial
  Dim rHdr
  Set rHdr = d.Range(0, Len(headerText))
  With rHdr.Font
    .Name = "Calibri"
    .Size = 11
  End With
  rHdr.ParagraphFormat.SpaceAfter = app.CentimetersToPoints(0.1)

  BoldLabelAscii d, "Projeto:"
  BoldLabelAscii d, "Frente:"
  BoldLabelAscii d, "Respons" & a_acute & "vel:"
  BoldLabelAscii d, "Produto/Serv" & i_acute & c_ced & "o:"
  BoldLabelAscii d, "Vers" & a_tilde & "o:"
  BoldLabelAscii d, "Dispositivo:"
  BoldLabelAscii d, "Resultado Esperado:"

  On Error GoTo 0
End Sub

Sub BoldLabelAscii(ByRef d, ByVal label)
  On Error Resume Next
  Dim rng
  Set rng = d.Range(0, 0)
  With rng.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = label
    .Replacement.Text = ""
    .Forward = True
    .Wrap = 0
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
  End With
  Do While rng.Find.Execute
    rng.Font.Bold = True
    rng.Collapse 0
    rng.MoveStart 1
    rng.MoveEnd 1
  Loop
  On Error GoTo 0
End Sub

Sub CenterInlineParagraph(ByRef ils)
  On Error Resume Next
  With ils.Range.ParagraphFormat
    .Alignment = wdAlignParagraphCenter
    .LeftIndent = 0
    .RightIndent = 0
  End With
  On Error GoTo 0
End Sub

Sub CenterFloatingShapeToMargins(ByRef shp)
  On Error Resume Next
  shp.RelativeHorizontalPosition = wdRelHorizPosMargin
  shp.Left = wdShapeCenter
  On Error GoTo 0
End Sub

Function OpenDocSmart(ByRef app, ByVal path)
  On Error Resume Next
  Dim doc2, ext, pv
  ext = LCase(GetExt(path))
  If (ext = "mht") Or (ext = "mhtml") Then
    Set pv = app.ProtectedViewWindows.Open(path, False, "", False, True)
    If Err.Number = 0 And Not pv Is Nothing Then
      Set doc2 = pv.Edit
    Else
      Err.Clear
    End If
  End If
  If doc2 Is Nothing Then
    Set doc2 = app.Documents.Open(path, False, False, False, "", "", False, "", "", False, False, False, False, True)
  End If
  Set OpenDocSmart = doc2
End Function

Function GetExt(ByVal p)
  Dim dotPos
  dotPos = InStrRev(p, ".")
  If dotPos > 0 Then
    GetExt = Mid(p, dotPos + 1)
  Else
    GetExt = ""
  End If
End Function

Sub ApplySpacingAndIsolation(ByRef ils, ByRef app)
  On Error Resume Next
  Dim spaceBeforePt, spaceAfterPt
  spaceBeforePt = app.CentimetersToPoints(MIN_SPACE_BEFORE_CM)
  spaceAfterPt  = app.CentimetersToPoints(MIN_SPACE_AFTER_CM)
  Dim inTable
  inTable = False
  If Not ils Is Nothing Then inTable = CBool(ils.Range.Information(wdWithInTable))
  If ISOLAR_PARAGRAFO And (Not inTable) Then
    Dim r, pRange
    Set r = ils.Range
    Set pRange = r.Paragraphs(1).Range
    If r.Start > pRange.Start Then r.InsertBefore vbCr
    Set r = ils.Range
    Set pRange = r.Paragraphs(1).Range
    If r.End < pRange.End Then r.InsertAfter vbCr
  End If
  Dim pf
  Set pf = ils.Range.ParagraphFormat
  If pf.SpaceBefore < spaceBeforePt Then pf.SpaceBefore = spaceBeforePt
  If pf.SpaceAfter  < spaceAfterPt  Then pf.SpaceAfter  = spaceAfterPt
  On Error GoTo 0
End Sub

Sub SafeWrapForShape(ByRef shp, ByRef app)
  On Error Resume Next
  shp.WrapFormat.Type = wdWrapTopBottom
  shp.WrapFormat.AllowOverlap = False
  shp.WrapFormat.DistanceTop    = app.CentimetersToPoints(FLOAT_DIST_TOP_CM)
  shp.WrapFormat.DistanceBottom = app.CentimetersToPoints(FLOAT_DIST_BOTTOM_CM)
  shp.WrapFormat.DistanceLeft   = app.CentimetersToPoints(FLOAT_DIST_LEFT_CM)
  shp.WrapFormat.DistanceRight  = app.CentimetersToPoints(FLOAT_DIST_RIGHT_CM)
  On Error GoTo 0
End Sub

Sub ScaleInline(ByRef ils, ByVal hasL, ByVal Lpt, ByVal hasA, ByVal Apt)
  On Error Resume Next
  ils.LockAspectRatio = True
  If hasL Then
    ils.Width = Lpt
  ElseIf hasA Then
    ils.Height = Apt
  End If
  On Error GoTo 0
End Sub

Sub ClampInline(ByRef ils, ByVal maxWpt, ByVal maxHpt)
  On Error Resume Next
  ils.LockAspectRatio = True
  If ils.Width > maxWpt Then
    ils.Width = maxWpt
  End If
  If ils.Height > maxHpt Then
    ils.Height = maxHpt
  End If
  On Error GoTo 0
End Sub

Sub ScaleShape(ByRef shp, ByVal hasL, ByVal Lpt, ByVal hasA, ByVal Apt)
  Dim i3
  If shp.Type = MsoTypePicture Or shp.Type = MsoTypeLinkedPicture Then
    On Error Resume Next
    shp.LockAspectRatio = True
    If hasL Then
      shp.Width = Lpt
    ElseIf hasA Then
      shp.Height = Apt
    End If
    On Error GoTo 0
  ElseIf shp.Type = MsoTypeGroup Then
    For i3 = 1 To shp.GroupItems.Count
      ScaleShape shp.GroupItems(i3), hasL, Lpt, hasA, Apt
    Next
  End If
End Sub

Sub FitInline(ByRef ils, ByVal Lpt, ByVal Apt)
  Dim cw, ch, s
  On Error Resume Next
  cw = ils.Width
  ch = ils.Height
  If cw > 0 And ch > 0 Then
    s = Min2(Lpt / cw, Apt / ch)
    If s > 0 Then
      ils.LockAspectRatio = True
      ils.Width = cw * s
    End If
  End If
  On Error GoTo 0
End Sub

Sub FitShape(ByRef shp, ByVal Lpt, ByVal Apt)
  Dim cw2, ch2, s2, j
  If shp.Type = MsoTypePicture Or shp.Type = MsoTypeLinkedPicture Then
    On Error Resume Next
    cw2 = shp.Width
    ch2 = shp.Height
    If cw2 > 0 And ch2 > 0 Then
      s2 = Min2(Lpt / cw2, Apt / ch2)
      If s2 > 0 Then
        shp.LockAspectRatio = True
        shp.Width = cw2 * s2
      End If
    End If
    On Error GoTo 0
  ElseIf shp.Type = MsoTypeGroup Then
    For j = 1 To shp.GroupItems.Count
      FitShape shp.GroupItems(j), Lpt, Apt
    Next
  End If
End Sub

Function Min2(a, b)
  If a < b Then Min2 = a Else Min2 = b
End Function

' ===== PACK helpers =====
Function GetProtectedTailStart(ByRef d)
  On Error Resume Next
  Dim rng
  Set rng = d.Content
  With rng.Find
    .ClearFormatting
    .Text = "Instru"
    .Forward = True
    .Wrap = 0
    .MatchCase = False
    .Format = False
  End With
  If rng.Find.Execute Then
    GetProtectedTailStart = rng.Start
  Else
    GetProtectedTailStart = 0
  End If
  On Error GoTo 0
End Function

Sub PackSpacingAroundInlineImages(ByRef d, ByRef app, ByVal protectStart)
  On Error Resume Next
  Dim ils, par, capPar
  Dim capAfter, capBefore
  capAfter  = CAPTION_SPACE_AFTER_PT
  capBefore = CAPTION_SPACE_BEFORE_PT

  For Each ils In d.InlineShapes
    If ils.Range.Start < protectStart Then
      Set par = ils.Range.Paragraphs(1)
      ReduceEmptyBefore par, MAX_EMPTY_BEFORE_IMAGE
      Set capPar = GetNextNonEmptyParagraph(d, par)
      If Not capPar Is Nothing Then
        With capPar.Format
          .SpaceBefore = capBefore
          .SpaceAfter  = capAfter
          .LeftIndent  = 0
          .RightIndent = 0
          .KeepTogether = True
        End With
      End If
      If Not capPar Is Nothing Then
        ReduceEmptyAfter capPar, MAX_EMPTY_AFTER_BLOCK
      Else
        ReduceEmptyAfter par, MAX_EMPTY_AFTER_BLOCK
      End If
    End If
  Next

  On Error GoTo 0
End Sub

Sub ReduceEmptyBefore(ByRef anyPar, ByVal maxEmpty)
  On Error Resume Next
  Dim p, countEmpty
  countEmpty = 0
  Set p = anyPar.Previous
  Do While Not p Is Nothing
    If IsEmptyParagraph(p) Then
      countEmpty = countEmpty + 1
      If countEmpty > maxEmpty Then p.Range.Delete
      Set p = p.Previous
    Else
      Exit Do
    End If
  Loop
  On Error GoTo 0
End Sub

Sub ReduceEmptyAfter(ByRef anyPar, ByVal maxEmpty)
  On Error Resume Next
  Dim p, countEmpty
  countEmpty = 0
  Set p = anyPar.Next
  Do While Not p Is Nothing
    If IsEmptyParagraph(p) Then
      countEmpty = countEmpty + 1
      If countEmpty > maxEmpty Then
        p.Range.Delete
      Else
        Set p = p.Next
      End If
    Else
      Exit Do
    End If
  Loop
  On Error GoTo 0
End Sub

Function IsEmptyParagraph(ByRef paragraph)
  Dim t
  t = paragraph.Range.Text
  t = Replace(t, Chr(13), "")
  t = Trim(t)
  IsEmptyParagraph = (Len(t) = 0)
End Function

Function GetNextNonEmptyParagraph(ByRef d, ByRef basePar)
  On Error Resume Next
  Dim p
  Set p = basePar.Next
  Do While Not p Is Nothing
    If Not IsEmptyParagraph(p) Then
      Set GetNextNonEmptyParagraph = p
      Exit Function
    End If
    Set p = p.Next
  Loop
  Set GetNextNonEmptyParagraph = Nothing
  On Error GoTo 0
End Function

' ===== Cleaners =====
Sub RemovePreviousNextText(ByRef d)
  On Error Resume Next
  Dim rng
  Set rng = d.Content
  With rng.Find
    .ClearFormatting
    .Text = "Previous Next "
    .Replacement.ClearFormatting
    .Replacement.Text = ""
    .Forward = True
    .Wrap = 1 ' wdFindContinue
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
  End With
  Do While rng.Find.Execute
    rng.Text = ""
    rng.Collapse 0
    rng.MoveStart 1
  Loop
  On Error GoTo 0
End Sub

' ===== Parse / Number helpers =====
Function ParseInput(ByVal s, ByRef hasL, ByRef Lcm, ByRef hasA, ByRef Acm)
  Dim t, parts, i4, tok, first, rest, n1, n2
  Dim timesSym
  timesSym = ChrW(&HD7)

  ParseInput = False
  hasL = False
  hasA = False
  Lcm = 0
  Acm = 0

  t = UCase(Trim(s))
  If t = "" Then Exit Function

  t = Replace(t, vbTab, " ")
  t = Replace(t, "=", " ")
  t = Replace(t, ":", " ")
  t = Replace(t, ";", " ")
  t = Replace(t, "POR", " X ")
  t = Replace(t, timesSym, " X ")
  t = Replace(t, "/", " X ")

  Do While InStr(t, "  ") > 0
    t = Replace(t, "  ", " ")
  Loop

  If InStr(t, "X") > 0 Then
    n1 = ToNumber(Left(t, InStr(t, "X") - 1))
    n2 = ToNumber(Mid(t, InStr(t, "X") + 1))
    If n1 > 0 And n2 > 0 Then
      hasL = True
      Lcm = n1
      hasA = True
      Acm = n2
      ParseInput = True
      Exit Function
    End If
  End If

  parts = Split(t, " ")
  For i4 = 0 To UBound(parts)
    tok = Trim(parts(i4))
    If tok <> "" Then
      first = Left(tok, 1)
      If (first = "L" Or first = "A") And Len(tok) > 1 Then
        rest = Mid(tok, 2)
        rest = Replace(rest, "CM", "")
        n1 = ToNumber(rest)
        If n1 > 0 Then
          If first = "L" Then
            hasL = True
            Lcm = n1
          Else
            hasA = True
            Acm = n1
          End If
        End If
      ElseIf tok = "L" Or tok = "A" Then
        If i4 < UBound(parts) Then
          rest = Replace(Trim(parts(i4 + 1)), "CM", "")
          n1 = ToNumber(rest)
          If n1 > 0 Then
            If tok = "L" Then
              hasL = True
              Lcm = n1
            Else
              hasA = True
              Acm = n1
            End If
          End If
        End If
      Else
        n1 = ToNumber(Replace(tok, "CM", ""))
        If n1 > 0 Then
          If Not hasL Then
            hasL = True
            Lcm = n1
          ElseIf Not hasA Then
            hasA = True
            Acm = n1
          End If
        End If
      End If
    End If
  Next

  ParseInput = (hasL Or hasA)
End Function

Function ToNumber(ByVal s)
  Dim i5, ch, buf, dotCount
  s = Trim(UCase(s))
  s = Replace(s, ",", ".")
  s = Replace(s, "CM", "")
  buf = ""
  dotCount = 0
  For i5 = 1 To Len(s)
    ch = Mid(s, i5, 1)
    If ch >= "0" And ch <= "9" Then
      buf = buf & ch
    ElseIf ch = "." Then
      If dotCount = 0 Then
        buf = buf & "."
        dotCount = 1
      Else
        Exit For
      End If
    ElseIf ch = "-" And buf = "" Then
      buf = "-"
    ElseIf Len(buf) > 0 Then
      Exit For
    End If
  Next
  If buf = "" Or buf = "-" Or buf = "." Or buf = "-." Then
    ToNumber = 0
  Else
    On Error Resume Next
    ToNumber = CDbl(buf)
    If Err.Number <> 0 Then ToNumber = 0
    On Error GoTo 0
  End If
End Function
