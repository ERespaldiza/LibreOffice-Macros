' ============================================================
' FormatCodeBlocks
' Finds lines with exact code-language names, selects the
' following "Preformatted Text" paragraphs and applies the
' matching dk_Code_* paragraph style.
' ============================================================

' Module-level variable used by the dialog button handler
Dim g_sDialogResult As String
Dim g_oDialog       As Object


' ── Language definitions ─────────────────────────────────────
' Add or remove entries here to support more languages.
' Format: Array("Label in dialog", "Exact heading text", "Style to apply")
' ─────────────────────────────────────────────────────────────
Function GetLanguages() As Variant
    GetLanguages = Array( _
        Array("C++",        "C++",        "dk_Code_Cpp"),        _
        Array("CSS",        "CSS",        "dk_Code_CSS"),        _
        Array("Dart",       "Dart",       "dk_Code_Dart"),       _
        Array("HTML",       "HTML",       "dk_Code_HTML"),       _
        Array("Java",       "Java",       "dk_Code_Java"),       _
        Array("JavaScript", "JavaScript", "dk_Code_JavaScript"), _
        Array("Python",     "Python",     "dk_Code_Python"),     _
        Array("SQL",        "SQL",        "dk_Code_SQL")         _
    )
End Function


' ── Main entry point ─────────────────────────────────────────
Sub FormatCodeBlocks()
    Dim oLangs       As Variant
    Dim oSelected()  As Integer
    Dim nSelected    As Integer
    Dim oDoc         As Object
    Dim oText        As Object
    Dim frame        As Object
    Dim dispatcher   As Object
    Dim nTotalParas  As Integer
    Dim nTotalBlocks As Integer
    Dim nParas       As Integer
    Dim nBlocks      As Integer
    Dim k            As Integer
    Dim idx          As Integer
    Dim sFind        As String
    Dim sStyle       As String

    oLangs = GetLanguages()

    If Not ShowLanguageDialog(oLangs, oSelected(), nSelected) Then
        Exit Sub
    End If

    oDoc       = ThisComponent
    oText      = oDoc.getText()
    frame      = oDoc.getCurrentController().getFrame()
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

    nTotalParas  = 0
    nTotalBlocks = 0

    For k = 0 To nSelected - 1
        idx    = oSelected(k)
        sFind  = oLangs(idx)(1)
        sStyle = oLangs(idx)(2)
        Call ProcessLanguage(oDoc, oText, frame, dispatcher, _
                             sFind, sStyle, nParas, nBlocks)
        nTotalParas  = nTotalParas  + nParas
        nTotalBlocks = nTotalBlocks + nBlocks
    Next k

    If nTotalBlocks = 0 Then
        MsgBox "No matching blocks found for the selected language(s).", _
               MB_ICONWARNING, "FormatCodeBlocks"
    Else
        MsgBox "Done! Formatted " & nTotalParas & " paragraph(s) across " & _
               nTotalBlocks & " block(s).", _
               MB_ICONINFORMATION, "FormatCodeBlocks"
    End If
End Sub


' ── Dialog button handlers ────────────────────────────────────
Sub OnDialogOK()
    g_sDialogResult = "OK"
    g_oDialog.endExecute()
End Sub

Sub OnDialogCancel()
    g_sDialogResult = "CANCEL"
    g_oDialog.endExecute()
End Sub


' ── Multi-select language dialog ──────────────────────────────
Function ShowLanguageDialog(oLangs As Variant, _
                            ByRef oSelected() As Integer, _
                            ByRef nSelected As Integer) As Boolean
    Dim oDialogModel  As Object
    Dim oListModel    As Object
    Dim oOKModel      As Object
    Dim oCancelModel  As Object
    Dim oLabelModel   As Object
    Dim oList         As Object
    Dim aItems()      As String
    Dim aSelPos()     As Integer
    Dim nLangs        As Integer
    Dim i             As Integer

    nLangs = UBound(oLangs) + 1

    ' Build dialog model
    oDialogModel = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
    oDialogModel.Width  = 200
    oDialogModel.Height = 185
    oDialogModel.Title  = "Format Code Blocks"

    ' Label
    oLabelModel = oDialogModel.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    oLabelModel.PositionX = 8
    oLabelModel.PositionY = 8
    oLabelModel.Width     = 184
    oLabelModel.Height    = 12
    oLabelModel.Label     = "Select one or more languages (Ctrl+click):"
    oDialogModel.insertByName("lbl", oLabelModel)

    ' Listbox
    oListModel = oDialogModel.createInstance("com.sun.star.awt.UnoControlListBoxModel")
    oListModel.PositionX      = 8
    oListModel.PositionY      = 24
    oListModel.Width          = 184
    oListModel.Height         = 116
    oListModel.MultiSelection = True
    oListModel.Border         = 1
    ReDim aItems(nLangs - 1)
    For i = 0 To nLangs - 1
        aItems(i) = oLangs(i)(0)
    Next i
    oListModel.StringItemList = aItems
    oDialogModel.insertByName("lst", oListModel)

    ' OK button — wired to macro name via ActionCommand + ButtonType PUSH
    oOKModel = oDialogModel.createInstance("com.sun.star.awt.UnoControlButtonModel")
    oOKModel.PositionX       = 56
    oOKModel.PositionY       = 152
    oOKModel.Width           = 50
    oOKModel.Height          = 16
    oOKModel.Label           = "OK"
    oOKModel.DefaultButton   = True
    oOKModel.PushButtonType  = 0
    oDialogModel.insertByName("btnOK", oOKModel)

    ' Cancel button
    oCancelModel = oDialogModel.createInstance("com.sun.star.awt.UnoControlButtonModel")
    oCancelModel.PositionX      = 114
    oCancelModel.PositionY      = 152
    oCancelModel.Width          = 50
    oCancelModel.Height         = 16
    oCancelModel.Label          = "Cancel"
    oCancelModel.PushButtonType = 0
    oDialogModel.insertByName("btnCancel", oCancelModel)

    ' Instantiate dialog
    g_sDialogResult = "CANCEL"
    g_oDialog = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    g_oDialog.setModel(oDialogModel)
    g_oDialog.createPeer(CreateUnoService("com.sun.star.awt.Toolkit"), Nothing)

    ' Wire buttons using createActionListener workaround via Basic macro strings
    Dim oOKBtn     As Object
    Dim oCancelBtn As Object
    oOKBtn     = g_oDialog.getControl("btnOK")
    oCancelBtn = g_oDialog.getControl("btnCancel")

    ' Use a simple UNO ActionListener via createUnoListener
    Dim oOKListener As Object
    oOKListener = CreateUnoListener("OK_", "com.sun.star.awt.XActionListener")
    oOKBtn.addActionListener(oOKListener)

    Dim oCancelListener As Object
    oCancelListener = CreateUnoListener("Cancel_", "com.sun.star.awt.XActionListener")
    oCancelBtn.addActionListener(oCancelListener)

    g_oDialog.execute()

    ' Read selection before disposing
    If g_sDialogResult = "OK" Then
        oList    = g_oDialog.getControl("lst")
        aSelPos  = oList.getSelectedItemsPos()
        nSelected = UBound(aSelPos) + 1
        If nSelected = 0 Then
            MsgBox "Please select at least one language.", _
                   MB_ICONWARNING, "FormatCodeBlocks"
            g_oDialog.dispose()
            ShowLanguageDialog = False
            Exit Function
        End If
        ReDim oSelected(nSelected - 1)
        For i = 0 To nSelected - 1
            oSelected(i) = aSelPos(i)
        Next i
        ShowLanguageDialog = True
    Else
        nSelected = 0
        ShowLanguageDialog = False
    End If

    g_oDialog.dispose()
End Function


' ── UNO listener stubs for OK button ─────────────────────────
Sub OK_actionPerformed(oEvent As Object)
    g_sDialogResult = "OK"
    g_oDialog.endExecute()
End Sub

Sub OK_disposing(oEvent As Object)
End Sub


' ── UNO listener stubs for Cancel button ─────────────────────
Sub Cancel_actionPerformed(oEvent As Object)
    g_sDialogResult = "CANCEL"
    g_oDialog.endExecute()
End Sub

Sub Cancel_disposing(oEvent As Object)
End Sub


' ── Process one language through the whole document ──────────
Sub ProcessLanguage(oDoc As Object, oText As Object, _
                    frame As Object, dispatcher As Object, _
                    sFind As String, sStyle As String, _
                    ByRef nTotalParas As Integer, _
                    ByRef nTotalBlocks As Integer)
    Dim oEnum       As Object
    Dim oPar        As Object
    Dim oStartPar   As Object
    Dim oEndPar     As Object
    Dim bFoundTag   As Boolean
    Dim bInBlock    As Boolean
    Dim nExtraParas As Integer
    Dim nParas      As Integer
    Dim nBlocks     As Integer

    bFoundTag   = False
    bInBlock    = False
    nExtraParas = 0
    nParas      = 0
    nBlocks     = 0

    oEnum = oText.createEnumeration()
    Do While oEnum.hasMoreElements()
        oPar = oEnum.nextElement()
        If Not oPar.supportsService("com.sun.star.text.Paragraph") Then
            GoTo NextElement
        End If

        If Trim(oPar.getString()) = sFind Then
            If bInBlock And Not (IsNull(oStartPar) Or IsEmpty(oStartPar)) Then
                Call ProcessBlock(oDoc, oText, frame, dispatcher, _
                                  oStartPar, oEndPar, sStyle, nExtraParas)
                nBlocks = nBlocks + 1
                nParas  = nParas + nExtraParas + 1
            End If
            bFoundTag   = True
            bInBlock    = False
            nExtraParas = 0
            Set oStartPar = Nothing
            Set oEndPar   = Nothing

        ElseIf bFoundTag Then
            If oPar.ParaStyleName = "Preformatted Text" Then
                bInBlock = True
                If IsNull(oStartPar) Or IsEmpty(oStartPar) Then
                    Set oStartPar = oPar
                Else
                    nExtraParas = nExtraParas + 1
                End If
                Set oEndPar = oPar
            Else
                If bInBlock And Not (IsNull(oStartPar) Or IsEmpty(oStartPar)) Then
                    Call ProcessBlock(oDoc, oText, frame, dispatcher, _
                                      oStartPar, oEndPar, sStyle, nExtraParas)
                    nBlocks = nBlocks + 1
                    nParas  = nParas + nExtraParas + 1
                End If
                bFoundTag   = False
                bInBlock    = False
                nExtraParas = 0
                Set oStartPar = Nothing
                Set oEndPar   = Nothing
            End If
        End If
        NextElement:
    Loop

    ' Flush any block still open at end of document
    If bInBlock And Not (IsNull(oStartPar) Or IsEmpty(oStartPar)) Then
        Call ProcessBlock(oDoc, oText, frame, dispatcher, _
                          oStartPar, oEndPar, sStyle, nExtraParas)
        nBlocks = nBlocks + 1
        nParas  = nParas + nExtraParas + 1
    End If

    nTotalParas  = nParas
    nTotalBlocks = nBlocks
End Sub


' ── Apply style + clear formatting on one block ───────────────
Sub ProcessBlock(oDoc As Object, oText As Object, _
                 frame As Object, dispatcher As Object, _
                 oStartPar As Object, oEndPar As Object, _
                 sStyle As String, nExtraParas As Integer)
    Dim oEnum      As Object
    Dim oPar       As Object
    Dim oParCursor As Object
    Dim oSelCursor As Object
    Dim bInRange   As Boolean
    Dim nDone      As Integer
    Dim nTarget    As Integer
    Dim i          As Integer

    nTarget  = nExtraParas + 1
    bInRange = False
    nDone    = 0

    oEnum = oText.createEnumeration()
    Do While oEnum.hasMoreElements()
        oPar = oEnum.nextElement()
        If Not oPar.supportsService("com.sun.star.text.Paragraph") Then
            GoTo NextPar
        End If

        If Not bInRange Then
            If oText.compareRegionStarts(oPar.getStart(), oStartPar.getStart()) = 0 Then
                bInRange = True
            End If
        End If

        If bInRange Then
            oParCursor = oText.createTextCursorByRange(oPar.getStart())
            oParCursor.gotoEndOfParagraph(True)
            oParCursor.setPropertyToDefault("CharWeight")
            oParCursor.setPropertyToDefault("CharColor")
            oParCursor.setPropertyToDefault("CharHeight")
            oParCursor.setPropertyToDefault("CharPosture")
            oParCursor.setPropertyToDefault("CharUnderline")
            oParCursor.setPropertyToDefault("CharBackColor")
            oParCursor.setPropertyToDefault("CharBackTransparent")
            oPar.ParaStyleName = sStyle
            nDone = nDone + 1
            If nDone = nTarget Then Exit Do
        End If
        NextPar:
    Loop

    ' Select block and dispatch ResetAttributes
    oSelCursor = oText.createTextCursorByRange(oStartPar.getStart())
    oSelCursor.gotoStartOfParagraph(False)
    For i = 1 To nExtraParas
        oSelCursor.gotoNextParagraph(True)
    Next i
    oSelCursor.gotoEndOfParagraph(True)
    oDoc.getCurrentController().select(oSelCursor)
    dispatcher.executeDispatch(frame, ".uno:ResetAttributes", "", 0, Array())
End Sub
