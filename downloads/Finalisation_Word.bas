Attribute VB_Name = "Typo"
Sub Lancer_le_traitement()
    Call Supprimer_Caracteres_Speciaux(Display_Message:=False)
    Call Traits(Display_Message:=False)
    Call Petites_Majuscules(Display_Message:=False)
    Call Renvois(Display_Message:=False)
    MsgBox "Le traitement est terminé.", vbInformation
End Sub
Sub Lancer_le_nettoyage()
    Call Nettoyage
End Sub

Sub Renvois(Optional Display_Message As Boolean = True)
    Dim doc As Document
    Dim fNote As footnote
    Dim eNote As endnote
    Dim rng As Range
    Dim charBefore As String
    Dim punctuations As String
    Dim punctSeq As String

    ' Définition des signes de ponctuation de fin de phrase
    punctuations = " .!?;:»”" & ChrW(&HA0) & ChrW(&H202F) ' Ponctuation de fin de phrase et guillemets fermants et espaces insécables

    Set doc = ActiveDocument

    Application.ScreenUpdating = False

    ' Loop through each footnote in the document
    For Each fNote In doc.Footnotes
        ' Set rng to the footnote reference marker in the main text
        Set rng = fNote.Reference
        punctSeq = ""
        
        ' Move backwards from the footnote reference to collect any trailing punctuation
        Do While rng.Start > 1
            charBefore = Mid(doc.Range(rng.Start - 1, rng.Start).Text, 1, 1)
            If InStr(punctuations, charBefore) > 0 Then
                punctSeq = charBefore & punctSeq
                rng.Start = rng.Start - 1
            Else
                Exit Do
            End If
        Loop
        
        ' If punctuation was found preceding the footnote reference, move it after the reference
        If punctSeq <> "" Then
            Dim Foot_Note_Ponctuation_Range As Range
            Set Foot_Note_Ponctuation_Range = doc.Range(rng.Start, rng.Start + Len(punctSeq))
            Foot_Note_Ponctuation_Range.Cut
            rng.Collapse Direction:=wdCollapseEnd
            rng.Paste
        End If
        
        ' Optional: highlight the footnote reference marker for visual confirmation
        fNote.Reference.HighlightColorIndex = wdBrightGreen
    Next fNote


    ' Loop through each endnote in the document
    For Each End_Note In doc.Endnotes
        ' Set rng to the footnote reference marker in the main text
        Set rng = End_Note.Reference
        punctSeq = ""
        
        ' Move backwards from the endnote reference to collect any trailing punctuation
        Do While rng.Start > 1
            charBefore = Mid(doc.Range(rng.Start - 1, rng.Start).Text, 1, 1)
            If InStr(punctuations, charBefore) > 0 Then
                punctSeq = charBefore & punctSeq
                rng.Start = rng.Start - 1
            Else
                Exit Do
            End If
        Loop
        
        ' If punctuation was found preceding the endnote reference, move it after the reference
        If punctSeq <> "" Then
            Dim End_Note_Ponctuation_Range As Range
            Set End_Note_Ponctuation_Range = doc.Range(rng.Start, rng.Start + Len(punctSeq))
            End_Note_Ponctuation_Range.Cut
            rng.Collapse Direction:=wdCollapseEnd
            rng.Paste
        End If
        
        ' Optional: highlight the endnote reference marker for visual confirmation
        End_Note.Reference.HighlightColorIndex = wdBrightGreen
    Next End_Note

    doc.Repaginate

    Application.ScreenUpdating = True

    ' Display message box with error count
    If Display_Message Then
        If errorCount > 0 Then
            MsgBox errorCount & " erreurs sont survenues pendant le replacement des renvois.", vbExclamation
        Else
            MsgBox "Le replacement des renvois s'est déroulé sans erreur.", vbInformation
        End If
    End If

End Sub


Sub Traits(Optional Display_Message As Boolean = True)
    Dim doc As Document: Set doc = ActiveDocument
    Dim rngStory As Range, rngSub As Range
    Dim motifs As Variant, remplacements As Variant
    Dim totalRepl As Long, thisRepl As Long, i As Long
    
    Application.ScreenUpdating = False
    Options.DefaultHighlightColorIndex = wdTurquoise
    
    '--- caractère “trait d’union insécable” ----------------------
    Dim NBH As String: NBH = ChrW(&H2011)
    
    '--- tableaux motifs / remplacements -------------------------
    motifs = Array("-([A-ZÀ-ÖØ-Þ0-9])", "([A-Za-zÀ-ÖØ-Þ])\.-", "-([A-Za-zÀ-ÖØ-Þ])\.")
    
    remplacements = Array(NBH & "\1", "\1." & NBH, NBH & "\1.")
    
    '--- on balaie chacune des StoryRanges du document -------------
    For Each rngStory In doc.StoryRanges
        For i = LBound(motifs) To UBound(motifs)
            RemplacerDansPlage rngStory, CStr(motifs(i)), CStr(remplacements(i))
        Next i
    Next rngStory
    
    Application.ScreenUpdating = True
    
    If Display_Message Then
        MsgBox totalRepl & " remplacement(s) effectué(s).", vbInformation
    End If
End Sub


Sub Petites_Majuscules(Optional Display_Message As Boolean = True)
    Dim rng As Range
    Dim doc As Document
    Dim footnote As footnote
    Dim endnote As endnote
    Dim searchPattern As String
    Dim mentionList As Variant
    Dim mention As Variant
    Dim errorCount As Long
    errorCount = 0
    
    On Error GoTo ErrorHandler ' Set error handling

    Set doc = ActiveDocument
    searchPattern = "\[*\]"
    
    mentionList = Array("en anglais seulement", "anglais seulement", "traduction")
    
    ' Process main document content
    Set rng = doc.Content
    If Not rng Is Nothing Then
        Call Process_Small_Caps_Range(rng, mentionList, searchPattern, errorCount)
    End If
    
    ' Process footnotes
    If doc.Footnotes.Count > 0 Then
        For Each footnote In doc.Footnotes
            Set rng = footnote.Range
            If Not rng Is Nothing Then
                Call Process_Small_Caps_Range(rng, mentionList, searchPattern, errorCount)
            End If
        Next footnote
    End If

    ' Process endnotes
    If doc.Endnotes.Count > 0 Then
        For Each endnote In doc.Endnotes
            Set rng = endnote.Range
            If Not rng Is Nothing Then
                Call Process_Small_Caps_Range(rng, mentionList, searchPattern, errorCount)
            End If
        Next endnote
    End If

    ' Display message box with error count
    If Display_Message Then
        If errorCount > 0 Then
            MsgBox errorCount & " erreurs sont survenues pendant l'application des petites majuscules.", vbExclamation
        Else
            MsgBox "L'application des petites majuscules s'est déroulée sans erreur.", vbInformation
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    errorCount = errorCount + 1
    Resume Next
End Sub

Sub Nettoyage(Optional Display_Message As Boolean = True)
    Dim rng As Range
    Dim sec As Section
    Dim hf As HeaderFooter
    Dim shp As Shape
    errorCount = 0

    ' Unhide text in the main document body
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .Font.Hidden = True
        With .Replacement
            .ClearFormatting
            .Font.Hidden = False
        End With
        .Execute Replace:=wdReplaceAll
    End With

    ' Remove highlighting in the main document body
    With rng.Find
        .ClearFormatting
        .Highlight = True
        With .Replacement
            .ClearFormatting
            .Highlight = False
        End With
        .Execute Replace:=wdReplaceAll
    End With

    ' Unhide text and remove highlighting in headers and footers
    For Each sec In ActiveDocument.Sections
        For Each hf In sec.Headers
            Set rng = hf.Range
            With rng.Find
                .ClearFormatting
                .Font.Hidden = True
                With .Replacement
                    .ClearFormatting
                    .Font.Hidden = False
                End With
                .Execute Replace:=wdReplaceAll
            End With

            With rng.Find
                .ClearFormatting
                .Highlight = True
                With .Replacement
                    .ClearFormatting
                    .Highlight = False
                End With
                .Execute Replace:=wdReplaceAll
            End With
        Next hf

        For Each hf In sec.Footers
            Set rng = hf.Range
            With rng.Find
                .ClearFormatting
                .Font.Hidden = True
                With .Replacement
                    .ClearFormatting
                    .Font.Hidden = False
                End With
                .Execute Replace:=wdReplaceAll
            End With

            With rng.Find
                .ClearFormatting
                .Highlight = True
                With .Replacement
                    .ClearFormatting
                    .Highlight = False
                End With
                .Execute Replace:=wdReplaceAll
            End With
        Next hf
    Next sec

    ' Unhide text and remove highlighting in text boxes and other shapes
    For Each shp In ActiveDocument.Shapes
        If shp.Type = msoTextBox Then
            Set rng = shp.TextFrame.TextRange
            With rng.Find
                .ClearFormatting
                .Font.Hidden = True
                With .Replacement
                    .ClearFormatting
                    .Font.Hidden = False
                End With
                .Execute Replace:=wdReplaceAll
            End With

            With rng.Find
                .ClearFormatting
                .Highlight = True
                With .Replacement
                    .ClearFormatting
                    .Highlight = False
                End With
                .Execute Replace:=wdReplaceAll
            End With
        End If
    Next shp

    ' Unhide text and remove highlighting in footnotes
    For Each ft In ActiveDocument.Footnotes
        Set rng = ft.Range
        With rng.Find
            .ClearFormatting
            .Font.Hidden = True
            With .Replacement
                .ClearFormatting
                .Font.Hidden = False
            End With
            .Execute Replace:=wdReplaceAll
        End With

        With rng.Find
            .ClearFormatting
            .Highlight = True
            With .Replacement
                .ClearFormatting
                .Highlight = False
            End With
            .Execute Replace:=wdReplaceAll
        End With
    Next ft

    ' Unhide text and remove highlighting in endnotes
    For Each en In ActiveDocument.Endnotes
        Set rng = en.Range
        With rng.Find
            .ClearFormatting
            .Font.Hidden = True
            With .Replacement
                .ClearFormatting
                .Font.Hidden = False
            End With
            .Execute Replace:=wdReplaceAll
        End With

        With rng.Find
            .ClearFormatting
            .Highlight = True
            With .Replacement
                .ClearFormatting
                .Highlight = False
            End With
            .Execute Replace:=wdReplaceAll
        End With
    Next en

    ' Display message box with error count
    If Display_Message Then
        If errorCount > 0 Then
            MsgBox errorCount & " erreurs sont survenues pendant le nettoyage.", vbExclamation
        Else
            MsgBox "Le nettoyage s'est déroulé sans erreur.", vbInformation
        End If
    End If

End Sub

Sub Process_Small_Caps_Range(rng As Range, mentionList As Variant, searchPattern As String, ByRef errorCount As Long)
    On Error GoTo ErrorHandler ' Set error handling

    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = searchPattern
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = True
        
        Do While .Execute
            Set rng = .Parent
            If Not rng.Paragraphs.First Is Nothing Then
                Dim textWithinBrackets As String
                textWithinBrackets = Trim(Mid(rng.Text, 2, Len(rng.Text) - 2))
                If InStr(textWithinBrackets, vbCr) = 0 And InStr(textWithinBrackets, vbLf) = 0 Then
                    Dim isSpecialMention As Boolean
                    isSpecialMention = False

                    For Each mention In mentionList
                        If textWithinBrackets = mention Then
                            isSpecialMention = True
                            Exit For
                        End If
                    Next mention
                    
                    If isSpecialMention Then
                        rng.Font.SmallCaps = True
                        rng.HighlightColorIndex = wdBrightGreen
                    Else
                        rng.HighlightColorIndex = wdRed
                    End If
                End If
            End If
        Loop
    End With
    
    Exit Sub
    
ErrorHandler:
    errorCount = errorCount + 1
    Resume Next
End Sub

Sub Supprimer_Caracteres_Speciaux(Optional Display_Message As Boolean = True)
    
    Options.DefaultHighlightColorIndex = wdBlue
    
    '--- 1) Déclaration des points de code à supprimer ----------
    Dim chars As Variant
    chars = Array(&H2995, &H2996, &H22D8, &H22D9, &H272D, &H2729)


    '--- 2) a) Nettoyage du contenu visible --------------------
    Dim i As Long
    For i = LBound(chars) To UBound(chars)
        With ActiveDocument.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = ChrW(chars(i))
            .Replacement.Text = ""
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    Next i

    '--- 2) b) Nettoyage des hyperliens -----------------------
    Dim hl As Hyperlink, charToRemove As String
    For Each hl In ActiveDocument.Hyperlinks
        For i = LBound(chars) To UBound(chars)
            charToRemove = ChrW(chars(i))
            hl.Address = Replace(hl.Address, charToRemove, "")
            hl.TextToDisplay = Replace(hl.TextToDisplay, charToRemove, "")
        Next i
    Next hl

    If Display_Message Then
        MsgBox "Élimination des symboles TAGmax terminée"
    End If
End Sub

Private Sub RemplacerDansPlage(rng As Range, _
                               ByVal motif As String, _
                               ByVal remplacement As String)
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = motif
        .Replacement.Text = remplacement
        .MatchWildcards = True
        .Wrap = wdFindStop
        .Forward = True
        .Replacement.Highlight = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub

