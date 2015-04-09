' @(h) clsMSAccessCSV.vb          ver 01.00.00
'
' @(s)
' 
'
'
Option Strict Off
Option Explicit On 

Imports System
Imports System.IO

Public Class MSAccessCSV

    ' @(f)
    '
    ' �@�\�@�@ :�w���ް����߰ď���
    '
    ' �Ԃ�l�@ :����I�� - 0    �װ - 1
    '
    ' �������@ :ARG1 - mdb̧�ٖ�(���߽�t)
    ' �@�@�@    ARG2 - csv̧�ٖ�(���߽�t)
    ' �@�@�@    ARG3 - ð��ٖ�
    ' �@�@�@    ARG4 - ���߰Ē�`��
    '
    ' �@�\���� :
    '
    ' ���l�@�@ :
    '
    Public Function lngExportCSV(ByRef mdbPath As String, _
                                 ByRef csvPath As String, _
                                 ByRef tableName As String, _
                                 ByRef schemaFile As String, _
                                 ByRef hasFieldNames As Boolean) As Long

        lngExportCSV = 0

        ''---- mdb�������� ----
        If File.Exists(mdbPath) = False Then
            lngExportCSV = -1
            Exit Function
        End If

        Dim oAccess As Access.ApplicationClass
        oAccess = New Access.ApplicationClass

        ''---- DB�ڑ� ----
        Try
            oAccess.OpenCurrentDatabase(mdbPath)
        Catch ex As Exception
            lngExportCSV = Err.Number
            Exit Function
        End Try

        ''---- csv����߰� ----
        Try
            oAccess.DoCmd.TransferText(Access.AcTextTransferType.acExportDelim, _
                                       schemaFile, _
                                       tableName, _
                                       csvPath, _
                                       hasFieldNames)
        Catch ex As Exception
            lngExportCSV = Err.Number
            Exit Function
        End Try

        oAccess.Quit()

    End Function

    ' @(f)
    '
    ' �@�\�@�@ :�w���ް����߰ď���
    '
    ' �Ԃ�l�@ :����I�� - 0    �װ - 0�ȊO
    '
    ' �������@ :ARG1 - mdb̧�ٖ�(���߽�t)
    ' �@�@�@    ARG2 - csv̧�ٖ�(���߽�t)
    ' �@�@�@    ARG3 - ð��ٖ�
    ' �@�@�@    ARG4 - ���߰Ē�`��
    ' �@�@�@    ARG5 - �擪�s̨���ޔF���׸�
    '
    ' �@�\���� :
    '
    ' ���l�@�@ :
    '
    Public Function lngImportCSV(ByRef mdbPath As String, _
                                 ByRef csvPath As String, _
                                 ByRef tableName As String, _
                                 ByRef schemaFile As String, _
                                 ByRef hasFieldNames As Boolean) As Long

        lngImportCSV = 0

        ''---- mdb�������� ----
        If File.Exists(mdbPath) = False Then
            lngImportCSV = -1
            Exit Function
        End If

        ''---- csv�������� ----
        If File.Exists(csvPath) = False Then
            lngImportCSV = -2
            Exit Function
        End If

        Dim oAccess As Access.ApplicationClass
        oAccess = New Access.ApplicationClass

        ''---- DB�ڑ� ----
        Try
            oAccess.OpenCurrentDatabase(mdbPath)
        Catch ex As Exception
            lngImportCSV = Err.Number
            Exit Function
        End Try

        ''---- csv���߰� ----
        Try
            oAccess.DoCmd.TransferText(Access.AcTextTransferType.acImportDelim, _
                                       schemaFile, _
                                       tableName, _
                                       csvPath, _
                                       hasFieldNames)
        Catch ex As Exception
            lngImportCSV = Err.Number
            Exit Function
        End Try

        oAccess.Quit()

    End Function

End Class
