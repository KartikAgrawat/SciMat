Attribute VB_Name = "SciMatEngine"
Sub ConvertMarkdownMath()
    Dim rng As Range
    Dim i As Integer
    Dim monoIons() As String
    Dim diatomicGases() As String
    
    ' ==============================================================================
    ' ⚙️ USER CONFIGURATION: THE EXCEPTION WHITELISTS
    ' ==============================================================================
    monoIons = Split("O2-,S2-,P3-,N3-,C4-,H+,H-,K+,F-,I-,V2+,V3+,V4+,V5+,Y3+,U4+,U6+", ",")
    diatomicGases = Split("H2,N2,O2,O3,P4,S8,C60", ",")

    ' ==============================================================================
    ' PART 1: MARKDOWN MATH CONVERTER
    ' ==============================================================================
    Application.Dialogs(2844).Execute
    DoEvents
    
    Set rng = ActiveDocument.Content
    With rng.Find
        .Text = "$$[!$]@$$"
        .MatchWildcards = True
        Do While .Execute
            rng.Text = Mid(rng.Text, 3, Len(rng.Text) - 4)
            ActiveDocument.OMaths.Add rng
            rng.OMaths(1).BuildUp
            rng.Collapse 0
        Loop
    End With
    
    Set rng = ActiveDocument.Content
    With rng.Find
        .Text = "$[!$]@$"
        .MatchWildcards = True
        Do While .Execute
            rng.Text = Mid(rng.Text, 2, Len(rng.Text) - 2)
            ActiveDocument.OMaths.Add rng
            rng.OMaths(1).BuildUp
            rng.Collapse 0
        Loop
    End With
    
    ' ==============================================================================
    ' PART 2: THE CHEMISTRY ENGINE (V1.5.6)
    ' ==============================================================================
    
    ' ---------------------------------------------------------
    ' PASS 0: TEXT SANITIZER (Normalizes notation before formatting)
    ' ---------------------------------------------------------
    
    ' 0A. Caret Sweeper (Word requires ^^ to find a literal caret!)
    Dim caretCharges() As String
    caretCharges = Split("^^1+,^^1-,^^2+,^^2-,^^3+,^^3-,^^4+,^^4-,^^+,^^-", ",")
    
    For i = LBound(caretCharges) To UBound(caretCharges)
        Set rng = ActiveDocument.Content
        With rng.Find
            .Text = caretCharges(i)
            .MatchWildcards = False
            .MatchWholeWord = False
            Do While .Execute
                rng.Text = Replace(rng.Text, "^", "")
                rng.Collapse 0
            Loop
        End With
    Next i

    ' 0B. Notation Normalizer (Converts -- to 2-, ++ to 2+)
    Set rng = ActiveDocument.Content
    With rng.Find
        .Text = "([A-Za-z0-9])\-\-\-"
        .MatchWildcards = True
        Do While .Execute
            rng.Text = Left(rng.Text, 1) & "3-"
            rng.Collapse 0
        Loop
    End With
    Set rng = ActiveDocument.Content
    With rng.Find
        .Text = "([A-Za-z0-9])\-\-"
        .MatchWildcards = True
        Do While .Execute
            rng.Text = Left(rng.Text, 1) & "2-"
            rng.Collapse 0
        Loop
    End With
    Set rng = ActiveDocument.Content
    With rng.Find
        .Text = "([A-Za-z0-9])\+\+"
        .MatchWildcards = True
        Do While .Execute
            rng.Text = Left(rng.Text, 1) & "2+"
            rng.Collapse 0
        Loop
    End With

    ' ---------------------------------------------------------
    ' PASS 1: WHITELISTS & EDGE CASES
    ' ---------------------------------------------------------
    
    ' 1A. Polyatomic Edge Cases (Fixes O3-, OH-, CN- without hitting AB-)
    Dim polyExceptions() As String
    polyExceptions = Split("O3-,I3-,OH-,CN-", ",")
    For i = LBound(polyExceptions) To UBound(polyExceptions)
        Set rng = ActiveDocument.Content
        With rng.Find
            .Text = polyExceptions(i)
            .MatchWildcards = False
            .MatchWholeWord = False
            .MatchCase = True
            Do While .Execute
                If polyExceptions(i) = "O3-" Or polyExceptions(i) = "I3-" Then
                    ActiveDocument.Range(rng.Start + 1, rng.Start + 2).Font.Subscript = True
                    ActiveDocument.Range(rng.Start + 2, rng.End).Font.Superscript = True
                Else
                    ActiveDocument.Range(rng.Start + 2, rng.End).Font.Superscript = True
                End If
                rng.Collapse 0
            Loop
        End With
    Next i

    ' 1B. Mono-Ions
    For i = LBound(monoIons) To UBound(monoIons)
        Set rng = ActiveDocument.Content
        With rng.Find
            .Text = monoIons(i)
            .MatchWildcards = False
            .MatchWholeWord = False
            .MatchCase = True
            Do While .Execute
                ActiveDocument.Range(rng.Start + 1, rng.End).Font.Superscript = True
                rng.Collapse 0
            Loop
        End With
    Next i

    ' 1C. Gases
    For i = LBound(diatomicGases) To UBound(diatomicGases)
        Set rng = ActiveDocument.Content
        With rng.Find
            .Text = diatomicGases(i)
            .MatchWildcards = False
            .MatchWholeWord = False
            .MatchCase = True
            Do While .Execute
                ActiveDocument.Range(rng.Start + 1, rng.End).Font.Subscript = True
                rng.Collapse 0
            Loop
        End With
    Next i

    ' ---------------------------------------------------------
    ' PASS 1.5: Polyatomic Charge Splitters
    ' ---------------------------------------------------------
    Set rng = ActiveDocument.Content
    With rng.Find
        .Text = "[A-Za-z][0-9][0-9][-+]"
        .MatchWildcards = True
        Do While .Execute
            If rng.Font.Superscript = False And rng.Font.Subscript = False Then
                ActiveDocument.Range(rng.Start + 1, rng.Start + 2).Font.Subscript = True
                ActiveDocument.Range(rng.Start + 2, rng.End).Font.Superscript = True
            End If
            rng.Collapse 0
        Loop
    End With

    Set rng = ActiveDocument.Content
    With rng.Find
        .Text = "[A-Z][0-9][-+]"
        .MatchWildcards = True
        Do While .Execute
            If rng.Font.Superscript = False And rng.Font.Subscript = False Then
                ActiveDocument.Range(rng.Start + 1, rng.Start + 2).Font.Subscript = True
                ActiveDocument.Range(rng.Start + 2, rng.End).Font.Superscript = True
            End If
            rng.Collapse 0
        Loop
    End With

    ' ---------------------------------------------------------
    ' PASS 2: MONO-IONS FIRST
    ' ---------------------------------------------------------
    Set rng = ActiveDocument.Content
    With rng.Find
        .Text = "[A-Z][a-z][0-9]{1,2}[-+]"
        .MatchWildcards = True
        Do While .Execute
            If rng.Font.Subscript = False Then
                rng.MoveStart wdCharacter, 2
                rng.Font.Superscript = True
            End If
            rng.Collapse 0
        Loop
    End With

    Set rng = ActiveDocument.Content
    With rng.Find
        .Text = "[A-Z][a-z][-+]"
        .MatchWildcards = True
        Do While .Execute
            If rng.Font.Subscript = False Then
                rng.MoveStart wdCharacter, 2
                rng.Font.Superscript = True
            End If
            rng.Collapse 0
        Loop
    End With
    
    Set rng = ActiveDocument.Content
    With rng.Find
        .Text = "[\)\]][0-9]{1,2}[-+]"
        .MatchWildcards = True
        Do While .Execute
            If rng.Font.Subscript = False Then
                rng.MoveStart wdCharacter, 1
                rng.Font.Superscript = True
            End If
            rng.Collapse 0
        Loop
    End With

    Set rng = ActiveDocument.Content
    With rng.Find
        .Text = "[\)\]][-+]"
        .MatchWildcards = True
        Do While .Execute
            If rng.Font.Subscript = False Then
                rng.MoveStart wdCharacter, 1
                rng.Font.Superscript = True
            End If
            rng.Collapse 0
        Loop
    End With

    ' ---------------------------------------------------------
    ' PASS 3: BASE ELEMENTS LAST
    ' ---------------------------------------------------------
    Set rng = ActiveDocument.Content
    With rng.Find
        .Text = "[a-z][A-Z][0-9]{1,2}"
        .MatchWildcards = True
        Do While .Execute
            If rng.Font.Superscript = False And rng.Font.Subscript = False Then
                rng.MoveStart wdCharacter, 2
                rng.Font.Subscript = True
            End If
            rng.Collapse 0
        Loop
    End With

    Set rng = ActiveDocument.Content
    With rng.Find
        .Text = "[A-Z][A-Z][0-9]{1,2}" 
        .MatchWildcards = True
        Do While .Execute
            If rng.Font.Superscript = False And rng.Font.Subscript = False Then
                rng.MoveStart wdCharacter, 2
                rng.Font.Subscript = True
            End If
            rng.Collapse 0
        Loop
    End With
    
    ' THE ORGANIC PATCH (With Standalone Gatekeeper)
    Set rng = ActiveDocument.Content
    With rng.Find
        .Text = "[CHONPSK][0-9]{1,2}"
        .MatchWildcards = True
        Do While .Execute
            Dim isSafe As Boolean
            isSafe = True
            
            ' If isolated by spaces or punctuation, ignore it
            If rng.Start > 0 And rng.End < ActiveDocument.Content.End - 1 Then
                Dim pChar As String, nChar As String
                pChar = ActiveDocument.Range(rng.Start - 1, rng.Start).Text
                nChar = ActiveDocument.Range(rng.End, rng.End + 1).Text
                
                If (pChar = " " Or pChar = Chr(13) Or pChar = Chr(9)) And _
                   (nChar = " " Or nChar = "," Or nChar = "." Or nChar = Chr(13) Or nChar = Chr(9)) Then
                    isSafe = False
                End If
            End If
            
            If isSafe And rng.Font.Superscript = False And rng.Font.Subscript = False Then
                rng.MoveStart wdCharacter, 1
                rng.Font.Subscript = True
            End If
            rng.Collapse 0
        Loop
    End With

    Set rng = ActiveDocument.Content
    With rng.Find
        .Text = "[A-Z][a-z][0-9]{1,2}"
        .MatchWildcards = True
        Do While .Execute
            If rng.Font.Superscript = False And rng.Font.Subscript = False Then
                rng.MoveStart wdCharacter, 2
                rng.Font.Subscript = True
            End If
            rng.Collapse 0
        Loop
    End With
    
    Set rng = ActiveDocument.Content
    With rng.Find
        .Text = "[\)\]][0-9]{1,2}"
        .MatchWildcards = True
        Do While .Execute
            If rng.Font.Superscript = False And rng.Font.Subscript = False Then
                rng.MoveStart wdCharacter, 1
                rng.Font.Subscript = True
            End If
            rng.Collapse 0
        Loop
    End With

    ' ---------------------------------------------------------
    ' PASS 4: THE ISOTOPE SWEEPER
    ' ---------------------------------------------------------
    Set rng = ActiveDocument.Content
    With rng.Find
        .Text = "<[0-9]{1,3}[A-Z]"
        .MatchWildcards = True
        Do While .Execute
            Dim isSafeIso As Boolean
            isSafeIso = True
            If rng.Start > 0 Then
                Dim prevChar As String
                prevChar = ActiveDocument.Range(rng.Start - 1, rng.Start).Text
                If prevChar = ")" Or prevChar = "]" Or prevChar = "}" Then
                    isSafeIso = False
                End If
            End If
            If isSafeIso Then
                rng.MoveEnd wdCharacter, -1
                rng.Font.Superscript = True
            End If
            rng.Collapse 0
        Loop
    End With

End Sub