'与SAS程序中Patient Profile结合使用
Sub MCompare()

Dim Orig As String
Dim Revi As String
Dim Comp As String
Dim name As String

'Orig: The path of original patientprofiles
'Revi: The path of revised patientprofiles
'Comp: The path of compared outputs


Orig = "C:\users\yuanyuanp\Desktop\patient profile\Prd\Output\RTF\20211117_Oncology\"
Revi = "C:\users\yuanyuanp\Desktop\patient profile\Prd\Output\RTF\20220815_Oncology\"
Comp = "C:\users\yuanyuanp\Desktop\patient profile\Prd\Output\Compare\20211117-20220815_Oncology\"

myfile = Dir(Orig, vbdictionary)

Do While myfile <> ""

'To compare documents
Documents.Open FileName:=Orig & myfile
Documents.Open FileName:=Revi & myfile
Application.CompareDocuments OriginalDocument:=Documents(Orig & myfile), RevisedDocument:=Documents(Revi & myfile), Destination:= _
        wdCompareDestinationNew, Granularity:=wdGranularityWordLevel, _
        CompareFormatting:=True, CompareCaseChanges:=True, CompareWhitespace:= _
        False, CompareTables:=True, CompareHeaders:=False, CompareFootnotes:= _
        False, CompareTextboxes:=True, CompareFields:=True, CompareComments:=True _
        , CompareMoves:=True, RevisedAuthor:="Pei Yuanyuan", IgnoreAllComparisonWarnings:=False
ActiveWindow.ShowSourceDocuments = wdShowSourceDocumentsBoth
ActiveWindow.View.MarkupMode = wdInLineRevisions
ActiveDocument.SaveAs2 FileName:=Comp & myfile
Documents.Close

'To Save as .pdf
name = Left(myfile, Len(myfile) - 4)
Documents.Open FileName:=Comp & myfile
ActiveDocument.ExportAsFixedFormat OutputFileName:=Comp & name & ".pdf", ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
        wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, To:=1, _
        Item:=wdExportDocumentWithMarkup, IncludeDocProps:=True, KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False
Documents.Close

myfile = Dir

Loop


End Sub



