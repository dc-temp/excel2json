VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} exportJsonFrom 
   Caption         =   "���ܴ���"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3660
   OleObjectBlob   =   "exportJsonFrom.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "exportJsonFrom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ExportBtn_Click()
    myrange = Worksheets("Sheet1").UsedRange  'ͨ����Ч��������ѡ������,���������Sheet1�����ͬ��Ҫ�޸�
    Total = UBound(myrange, 1) '��ȡ����
    Fields = UBound(myrange, 2) '��ȡ����
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
