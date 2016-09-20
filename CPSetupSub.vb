' CPFullSetup
'Compass Plan Setup

Sub CPAllSetup()
    '
    ' CPAllSetup Macro
    '
    Dim StartTime As Double
    Dim SecondsElapsed As Double

    'Remember time when macro starts
     StartTime = Timer
  
    Application.ScreenUpdating = False

    Dim wholedoc As Range
    Dim MyRange As Range
    Dim rng As Range

'I was the worst for coming up with variable names, omg
    Dim aTbl As Table
    Dim aCol1 As Column
    Dim aRows As Integer
    Dim arng As Range
    Dim p As Paragraph
    Dim EOC As String
    Dim drng As Range
    Dim k As Integer
    Dim l As Integer
    
    Dim Sec As Section
    Dim tbl As Table
    Dim oRow As Row
    Dim ocel As Cell
    Dim i As Integer
    Dim j As Integer
    
    Dim pathName As String
    Dim o As Document
    
    Set o = ActiveDocument
    Dim newname As String

'turns off track changse for final version
    o.TrackRevisions = Not o.TrackRevisions
'search document name for ".", if it's there, remove it
    If InStrRev(o.Name, ".") <> 0 Then
        newname = Left(o.Name, InStrRev(o.Name, ".") - 1)
    End If
'save new version to specific directory with new name
    ChangeFileOpenDirectory "C:\Users\171856123\Desktop\"
    o.SaveAs2 FileName:=newname & " - Final.doc" _
        , FileFormat:=wdFormatDocument, LockComments:=False, Password:="", _
        AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
        EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
        :=False, SaveAsAOCELetter:=False, CompatibilityMode:=0

'now that a new document has been created, let's get to business.

'find table of contents--the stuff I want to change is from that point onward.
    Set wholedoc = o.Range
    Set MyRange = o.Content
    MyRange.Find.Execute FindText:="Table of Contents", _
        Forward:=True
    If MyRange.Find.Found = True Then
        MyRange.SetRange (MyRange.Start), o.Content.End
    End If

    'remove extra spaces
    'for some reason there were often extra line breaks at the beginning of some pages.
    wholedoc.Find.ClearFormatting
    wholedoc.Find.Font.Size = 1
    wholedoc.Find.Replacement.ClearFormatting
    With wholedoc.Find
        .Text = "^p"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    wholedoc.Find.Execute Replace:=wdReplaceAll

    'Replace manual page breaks with section page breaks (it threw off the formatting if they were inconsistent)
    Selection.HomeKey wdStory
    Selection.Find.ClearFormatting
    With Selection.Find
        Do While .Execute(FindText:="^m", Forward:=True, _
        MatchWildcards:=False, Wrap:=wdFindStop) = True
            Set rng = Selection.Range.Duplicate
            Selection.InsertBreak Type:=wdSectionBreakNextPage
            rng.MoveStart wdCharacter, 1
            rng.Delete
        Loop
    End With
    
    'Set page margins
    For Each Sec In MyRange.Sections
        With Sec.PageSetup
            .LineNumbering.Active = False
            .Orientation = wdOrientPortrait
            .TopMargin = CentimetersToPoints(2)
            .BottomMargin = CentimetersToPoints(2)
            .LeftMargin = CentimetersToPoints(2.3)
            .RightMargin = CentimetersToPoints(2.3)
            .Gutter = CentimetersToPoints(0)
            .HeaderDistance = CentimetersToPoints(1.27)
            .FooterDistance = CentimetersToPoints(1.27)
            .PageWidth = CentimetersToPoints(21.59)
            .PageHeight = CentimetersToPoints(27.94)
            .FirstPageTray = wdPrinterDefaultBin
            .OtherPagesTray = wdPrinterDefaultBin
            .SectionStart = wdSectionNewPage
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .VerticalAlignment = wdAlignVerticalTop
            .SuppressEndnotes = False
            .MirrorMargins = False
            .TwoPagesOnOne = False
            .BookFoldPrinting = False
            .BookFoldRevPrinting = False
            .BookFoldPrintingSheets = 1
            .GutterPos = wdGutterPosLeft
        End With
    Next Sec

'go through the tables in the document and apply consisent formatting

'Ok, I'm sure I had a great reason to for replacing spaces with a character that wouldn't appear anywhere else in the document, but why this order?
'This would be to remove extra spaces that shouldn't be here

    For Each tbl In MyRange.Tables
        On Error Resume Next
        i = tbl.Rows.Count
        If Err.Number <> 5991 Then
            tbl.Range.Find.ClearFormatting
            With tbl.Range.Find
                .Text = " "
                .Replacement.Text = "~"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            tbl.Range.Find.Execute Replace:=wdReplaceAll
            tbl.Rows.Alignment = wdAlignRowCenter
            'set text allignment for tables
            For j = 1 To i
                tbl.Rows(j).Cells.VerticalAlignment = wdCellAlignVerticalCenter
            Next j
        End If
    Next tbl

    MyRange.Find.ClearFormatting
    MyRange.Find.Replacement.ClearFormatting
    With MyRange.Find
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    MyRange.Find.Execute Replace:=wdReplaceAll
    With MyRange.Find
        .Text = "~"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    MyRange.Find.Execute Replace:=wdReplaceAll

    'Format Recommended Actions Table
    
    'update TOC
    If o.TablesOfContents.Count = 1 Then
        o.TablesOfContents(1).Update
    End If

Application.ScreenUpdating = True

'Determine how many seconds code took to run
  SecondsElapsed = Round(Timer - StartTime, 2)

'Notify user in seconds
  MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation

End Sub
