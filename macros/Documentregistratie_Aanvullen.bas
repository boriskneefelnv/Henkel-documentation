Attribute VB_Name = "DocumentregistratieAanvullen"
Option Explicit

'=====================================================================================
' Documentregistratie aanvullen vanuit de Henkel-planning
'
' Werkt op het ACTIEVE werkblad (de documentregistratie). Kop op rij 1 met o.a.:
'   Master | BP File | Reference | IC/EC | Dest | Loading date | Document code | ...
' Per regel wordt het masternummer (kolom Master) opgezocht in blad "Uit" van
' New_planning_Henkel.xlsm (tabel HenkelPlanning) en worden aangevuld:
'   BP File      <- BP FILE
'   Reference    <- SHIPMENT
'   IC/EC        <- EC/IC
'   Dest         <- DEST
'   Loading date <- LAAD DATUM (of LD DATE, als de kolom zo heet)
'
' - Alleen LEGE cellen worden gevuld; wat er al staat blijft staan
'   (zet OVERSCHRIJVEN op True om alles opnieuw te vullen).
' - Staat de planning niet open, dan vraagt de macro om het bestand en opent hij
'   het alleen-lezen.
' - Masters die niet in de planning staan worden lichtrood gemarkeerd.
' - Komt een master meerdere keren voor in de planning, dan wint de regel die nog
'   niet geladen is (LOADED <> Y); daarna de regel met de laatste laaddatum.
' - Placeholderdatums (jaar < 2000, zoals 01-01-1900) worden niet overgenomen.
'=====================================================================================

Private Const BRON_BLAD As String = "Uit"
Private Const BRON_TABEL As String = "HenkelPlanning"
Private Const OVERSCHRIJVEN As Boolean = False
Private Const KLEUR_NIET_GEVONDEN As Long = 13551615    ' RGB(255,199,206) lichtrood

Public Sub Documentregistratie_Aanvullen()
    Dim wsDoel As Worksheet, wsBron As Worksheet
    Dim bron As Variant
    Dim cMaster As Long, cBp As Long, cShip As Long, cEcIc As Long, cDest As Long, cLaad As Long, cLoaded As Long
    Dim dMaster As Long, dBp As Long, dRef As Long, dEcIc As Long, dDest As Long, dLaad As Long
    Dim idx As Object, dubbel As Object
    Dim r As Long, laatste As Long, br As Long, oud As Long
    Dim sleutelWaarde As String
    Dim nGevuld As Long, nNietGevonden As Long, nRegels As Long
    Dim nietGevonden As String, dubbelLijst As String, k As Variant

    Set wsDoel = ActiveSheet
    If LCase$(Trim$(wsDoel.Name)) = LCase$(BRON_BLAD) Then
        MsgBox "Start de macro vanaf het blad van de documentregistratie, niet vanaf blad '" & BRON_BLAD & "'.", vbExclamation
        Exit Sub
    End If

    '--- kolommen op het doelblad (kop op rij 1) ---
    dMaster = DoelKolom(wsDoel, "Master")
    dBp = DoelKolom(wsDoel, "BP File")
    dRef = DoelKolom(wsDoel, "Reference")
    dEcIc = DoelKolom(wsDoel, "IC/EC", "EC/IC")
    dDest = DoelKolom(wsDoel, "Dest")
    dLaad = DoelKolom(wsDoel, "Loading date")
    If dMaster = 0 Then
        MsgBox "Kolom 'Master' niet gevonden op rij 1 van blad '" & wsDoel.Name & "'.", vbCritical
        Exit Sub
    End If
    If dBp + dRef + dEcIc + dDest + dLaad = 0 Then
        MsgBox "Geen van de kolommen BP File, Reference, IC/EC, Dest of Loading date gevonden op rij 1.", vbCritical
        Exit Sub
    End If

    '--- planning zoeken (open werkmap, anders laten kiezen) ---
    Set wsBron = ZoekBronBlad()
    If wsBron Is Nothing Then Exit Sub
    If Not LeesBron(wsBron, bron) Then
        MsgBox "In blad '" & wsBron.Name & "' van " & wsBron.Parent.Name & " is geen kop met MASTER en BP FILE gevonden.", vbCritical
        Exit Sub
    End If

    cMaster = BronKolom(bron, "MASTER")
    cBp = BronKolom(bron, "BP FILE")
    cShip = BronKolom(bron, "SHIPMENT")
    cEcIc = BronKolom(bron, "EC/IC", "IC/EC")
    cDest = BronKolom(bron, "DEST")
    cLaad = BronKolom(bron, "LAAD DATUM", "LD DATE", "LD. DATE")
    cLoaded = BronKolom(bron, "LOADED")
    If cMaster = 0 Then
        MsgBox "Kolom MASTER niet gevonden in blad '" & wsBron.Name & "'.", vbCritical
        Exit Sub
    End If

    '--- idx master -> beste bronregel ---
    Set idx = CreateObject("Scripting.Dictionary")
    Set dubbel = CreateObject("Scripting.Dictionary")
    For br = 2 To UBound(bron, 1)
        sleutelWaarde = Sleutel(bron(br, cMaster))
        If Len(sleutelWaarde) > 0 Then
            If idx.Exists(sleutelWaarde) Then
                dubbel(sleutelWaarde) = True
                oud = idx(sleutelWaarde)
                If Beter(bron, br, oud, cLoaded, cLaad) Then idx(sleutelWaarde) = br
            Else
                idx.Add sleutelWaarde, br
            End If
        End If
    Next br

    '--- doelregels aanvullen ---
    laatste = wsDoel.Cells(wsDoel.Rows.Count, dMaster).End(xlUp).Row
    If laatste < 2 Then
        MsgBox "Er staan nog geen masternummers in kolom 'Master'.", vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    For r = 2 To laatste
        sleutelWaarde = Sleutel(wsDoel.Cells(r, dMaster).Value)
        If Len(sleutelWaarde) > 0 Then
            nRegels = nRegels + 1
            If idx.Exists(sleutelWaarde) Then
                br = idx(sleutelWaarde)
                If wsDoel.Cells(r, dMaster).Interior.Color = KLEUR_NIET_GEVONDEN Then wsDoel.Cells(r, dMaster).Interior.Pattern = xlNone
                nGevuld = nGevuld + Zet(wsDoel, r, dBp, Waarde(bron, br, cBp), False)
                nGevuld = nGevuld + Zet(wsDoel, r, dRef, Waarde(bron, br, cShip), False)
                nGevuld = nGevuld + Zet(wsDoel, r, dEcIc, UCase$(Trim$(CStr(Waarde(bron, br, cEcIc)))), False)
                nGevuld = nGevuld + Zet(wsDoel, r, dDest, UCase$(Trim$(CStr(Waarde(bron, br, cDest)))), False)
                nGevuld = nGevuld + Zet(wsDoel, r, dLaad, Datum(Waarde(bron, br, cLaad)), True)
            Else
                wsDoel.Cells(r, dMaster).Interior.Color = KLEUR_NIET_GEVONDEN
                nNietGevonden = nNietGevonden + 1
                If InStr(1, "," & nietGevonden & ",", "," & sleutelWaarde & ",") = 0 Then
                    nietGevonden = nietGevonden & IIf(Len(nietGevonden) > 0, ",", "") & sleutelWaarde
                End If
            End If
        End If
    Next r
    Application.ScreenUpdating = True

    '--- melding ---
    Dim opBlad As String
    opBlad = "|" & Join(MastersOpBlad(wsDoel, dMaster, laatste), "|") & "|"
    For Each k In dubbel.Keys
        If InStr(1, opBlad, "|" & k & "|") > 0 Then
            dubbelLijst = dubbelLijst & IIf(Len(dubbelLijst) > 0, ", ", "") & k
        End If
    Next k

    Dim tekst As String
    tekst = nRegels & " regel(s) met een masternummer verwerkt; " & nGevuld & " cel(len) ingevuld." & vbCrLf & _
            "Bron: " & wsBron.Parent.Name & " / blad " & wsBron.Name & "."
    If nNietGevonden > 0 Then
        tekst = tekst & vbCrLf & vbCrLf & nNietGevonden & " regel(s) niet gevonden in de planning (lichtrood gemarkeerd):" & _
                vbCrLf & Replace(nietGevonden, ",", ", ")
    End If
    If Len(dubbelLijst) > 0 Then
        tekst = tekst & vbCrLf & vbCrLf & "Let op: deze masters staan meerdere keren in de planning (open regel / laatste laaddatum gebruikt):" & _
                vbCrLf & dubbelLijst
    End If
    If cLaad = 0 Then tekst = tekst & vbCrLf & vbCrLf & "Kolom LAAD DATUM / LD DATE niet gevonden in de planning: Loading date is niet ingevuld."
    MsgBox tekst, IIf(nNietGevonden > 0, vbExclamation, vbInformation), "Documentregistratie aanvullen"
End Sub

'--- schrijft een waarde in een lege cel (of altijd bij OVERSCHRIJVEN); geeft 1 terug als er geschreven is
Private Function Zet(ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal v As Variant, ByVal isDatum As Boolean) As Long
    If c = 0 Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then If Len(Trim$(v)) = 0 Then Exit Function
    If Not OVERSCHRIJVEN Then
        If Len(Trim$(CStr(ws.Cells(r, c).Value))) > 0 Then Exit Function
    End If
    ws.Cells(r, c).Value = v
    If isDatum Then ws.Cells(r, c).NumberFormat = "dd-mm-yyyy"
    Zet = 1
End Function

'--- masternummer -> alleen cijfers (vangt tekst/getal, spaties en regeleindes af)
Private Function Sleutel(ByVal v As Variant) As String
    Dim s As String, i As Long, ch As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        s = v
    ElseIf IsNumeric(v) Then
        s = Format$(v, "0")
    Else
        s = CStr(v)
    End If
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Then Sleutel = Sleutel & ch
    Next i
End Function

Private Function Waarde(bron As Variant, ByVal r As Long, ByVal c As Long) As Variant
    If c = 0 Then Waarde = Empty: Exit Function
    If IsError(bron(r, c)) Then Waarde = Empty: Exit Function
    Waarde = bron(r, c)
End Function

'--- laaddatum: alleen echte datums (jaar >= 2000), anders leeg
Private Function Datum(ByVal v As Variant) As Variant
    Datum = Empty
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then
        If Year(v) >= 2000 Then Datum = CDate(Int(CDbl(v)))
    ElseIf IsNumeric(v) Then
        If CDbl(v) >= 36526 Then Datum = CDate(Int(CDbl(v)))
    ElseIf IsDate(v) Then
        If Year(CDate(v)) >= 2000 Then Datum = CDate(Int(CDbl(CDate(v))))
    End If
End Function

'--- bij dubbele masters: open regel wint, daarna de laatste laaddatum
Private Function Beter(bron As Variant, ByVal nieuw As Long, ByVal oud As Long, ByVal cLoaded As Long, ByVal cLaad As Long) As Boolean
    Dim openN As Boolean, openO As Boolean, dN As Double, dO As Double, v As Variant
    If cLoaded > 0 Then
        openN = UCase$(Trim$(CStr(Waarde(bron, nieuw, cLoaded)))) <> "Y"
        openO = UCase$(Trim$(CStr(Waarde(bron, oud, cLoaded)))) <> "Y"
        If openN <> openO Then Beter = openN: Exit Function
    End If
    v = Datum(Waarde(bron, nieuw, cLaad)): If Not IsEmpty(v) Then dN = CDbl(v)
    v = Datum(Waarde(bron, oud, cLaad)): If Not IsEmpty(v) Then dO = CDbl(v)
    Beter = dN > dO
End Function

'--- kolomnummer op rij 1 van het doelblad (hoofdletterongevoelig, spaties genegeerd)
Private Function DoelKolom(ws As Worksheet, ParamArray namen() As Variant) As Long
    Dim c As Long, laatsteKol As Long, kop As String, n As Variant
    laatsteKol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To laatsteKol
        kop = Norm(ws.Cells(1, c).Value)
        For Each n In namen
            If kop = Norm(n) Then DoelKolom = c: Exit Function
        Next n
    Next c
End Function

'--- kolomnummer in de koprij (rij 1 van de bronarray)
Private Function BronKolom(bron As Variant, ParamArray namen() As Variant) As Long
    Dim c As Long, kop As String, n As Variant
    For c = 1 To UBound(bron, 2)
        kop = Norm(bron(1, c))
        For Each n In namen
            If kop = Norm(n) Then BronKolom = c: Exit Function
        Next n
    Next c
End Function

Private Function Norm(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    Norm = UCase$(Replace(Replace(Replace(CStr(v), " ", ""), vbLf, ""), Chr(160), ""))
End Function

'--- leest de planning als array met de koprij als rij 1
Private Function LeesBron(ws As Worksheet, ByRef bron As Variant) As Boolean
    Dim lo As ListObject, r As Long, c As Long, kopRij As Long, laatsteRij As Long, laatsteKol As Long
    On Error Resume Next
    Set lo = ws.ListObjects(BRON_TABEL)
    On Error GoTo 0
    If Not lo Is Nothing Then
        bron = lo.Range.Value
        LeesBron = True
        Exit Function
    End If
    'geen tabel: zoek de koprij (MASTER + BP FILE) in de eerste 10 rijen
    For r = 1 To 10
        laatsteKol = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        For c = 1 To laatsteKol
            If Norm(ws.Cells(r, c).Value) = "MASTER" Then
                If Application.WorksheetFunction.CountIf(ws.Rows(r), "BP FILE") > 0 Then kopRij = r
            End If
        Next c
        If kopRij > 0 Then Exit For
    Next r
    If kopRij = 0 Then Exit Function
    laatsteRij = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    laatsteKol = ws.Cells(kopRij, ws.Columns.Count).End(xlToLeft).Column
    If laatsteRij <= kopRij Then laatsteRij = kopRij + 1
    bron = ws.Range(ws.Cells(kopRij, 1), ws.Cells(laatsteRij, laatsteKol)).Value
    LeesBron = True
End Function

'--- blad "Uit" in een open werkmap; anders het planningbestand laten kiezen
Private Function ZoekBronBlad() As Worksheet
    Dim wb As Workbook, ws As Worksheet, f As Variant, bron As Variant
    For Each wb In Application.Workbooks
        For Each ws In wb.Worksheets
            If LCase$(Trim$(ws.Name)) = LCase$(BRON_BLAD) Then
                If LeesBron(ws, bron) Then Set ZoekBronBlad = ws: Exit Function
            End If
        Next ws
    Next wb
    f = Application.GetOpenFilename("Excel-bestanden (*.xlsm;*.xlsx;*.xls),*.xlsm;*.xlsx;*.xls", , _
        "Kies de planning (New_planning_Henkel.xlsm) met blad '" & BRON_BLAD & "'")
    If VarType(f) = vbBoolean Then Exit Function
    Dim actief As Worksheet
    Set actief = ActiveSheet
    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(Filename:=f, ReadOnly:=True, UpdateLinks:=0)
    actief.Parent.Activate
    actief.Activate
    Application.ScreenUpdating = True
    On Error Resume Next
    Set ZoekBronBlad = wb.Worksheets(BRON_BLAD)
    On Error GoTo 0
    If ZoekBronBlad Is Nothing Then
        MsgBox "In '" & wb.Name & "' staat geen blad '" & BRON_BLAD & "'.", vbCritical
    End If
End Function

Private Function MastersOpBlad(ws As Worksheet, ByVal c As Long, ByVal laatste As Long) As Variant
    Dim r As Long, uit() As String, n As Long
    ReDim uit(0 To laatste)
    For r = 2 To laatste
        uit(n) = Sleutel(ws.Cells(r, c).Value)
        n = n + 1
    Next r
    MastersOpBlad = uit
End Function
