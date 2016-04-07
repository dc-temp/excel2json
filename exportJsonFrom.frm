VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} exportJsonFrom 
   Caption         =   "功能窗口"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3660
   OleObjectBlob   =   "exportJsonFrom.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "exportJsonFrom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ExportBtn_Click()
    myrange = Worksheets("Sheet1").UsedRange  '通过有效数据区来选择数据,表格名字是Sheet1如果不同需要修改
    Total = UBound(myrange, 1) '获取行数
    Fields = UBound(myrange, 2) '获取列数
     Dim objStream As Object
     Set objStream = CreateObject("ADODB.Stream")
      
     With objStream
            .Type = 2
            .Charset = "UTF-8"
            .Open
            .WriteText "["
     
            For i = 2 To Total
                .WriteText "{"
                For j = 1 To Fields
                    .WriteText """" & myrange(1, j) & """:""" & Replace(myrange(i, j), """", "\""") & """"
                     If j <> Fields Then
                        .WriteText ","
                     End If
                Next
                If i = Total Then
                        .WriteText "}"
                Else
                        .WriteText "},"
                End If
            Next
            .WriteText "]"
            .SaveToFile "c:\" & ActiveWorkbook.Name & ".json", 2
     End With
     Set objStream = Nothing
End Sub
