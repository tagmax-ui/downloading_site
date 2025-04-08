Attribute VB_Name = "Typo"
Sub Lancer_le_traitement()
    Call Traits(Display_Message:=False)
    Call Petites_Majuscules(Display_Message:=False)
    Call Renvois(Display_Message:=False)
    MsgBox "Le traitement est termin�.", vbInformation
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

    ' D�finition des signes de ponctuation de fin de phrase
    punctuations = " .!?;:��" & ChrW(&HA0) & ChrW(&H202F) ' Ponctuation de fin de phrase et guillemets fermants et espaces ins�cables

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
            MsgBox "Le replacement des renvois s'est d�roul� sans erreur.", vbInformation
        End If
    End If

End Sub


Sub Traits(Optional Display_Message As Boolean = True)
    Dim doc As Document
    Dim rng As Range
    Dim findString As String
    Dim replaceString As String
    errorCount = 0
    
    Options.DefaultHighlightColorIndex = wdRed
    Application.ScreenUpdating = False
    
    Set doc = ActiveDocument
    
    ' D�finir la cha�ne � rechercher et la cha�ne de remplacement
    findString = "-([A-Z�-��-�0-9])"
    replaceString = ChrW(&H2011) & "\1" ' Utilisation du trait d'union ins�cable
    
    ' Parcourir toutes les instances dans le document
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findString
        .Replacement.Text = replaceString
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = True
        .Replacement.Highlight = True
        .Execute Replace:=wdReplaceAll
    End With

    Application.ScreenUpdating = True
    
    ' Display message box with error count
    If Display_Message Then
        If errorCount > 0 Then
            MsgBox errorCount & " erreurs sont survenues pendant l'application des traits d'union ins�cables.", vbExclamation
        Else
            MsgBox "L'application des traits d'union ins�cables s'est d�roul�e sans erreur.", vbInformation
        End If
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
            MsgBox "L'application des petites majuscules s'est d�roul�e sans erreur.", vbInformation
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
            MsgBox "Le nettoyage s'est d�roul� sans erreur.", vbInformation
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

