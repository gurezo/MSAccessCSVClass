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
    ' 機能　　 :指定ﾃﾞｰﾀｲﾝﾎﾟｰﾄ処理
    '
    ' 返り値　 :正常終了 - 0    ｴﾗｰ - 1
    '
    ' 引き数　 :ARG1 - mdbﾌｧｲﾙ名(ﾌﾙﾊﾟｽ付)
    ' 　　　    ARG2 - csvﾌｧｲﾙ名(ﾌﾙﾊﾟｽ付)
    ' 　　　    ARG3 - ﾃｰﾌﾞﾙ名
    ' 　　　    ARG4 - ｲﾝﾎﾟｰﾄ定義名
    '
    ' 機能説明 :
    '
    ' 備考　　 :
    '
    Public Function lngExportCSV(ByRef mdbPath As String, _
                                 ByRef csvPath As String, _
                                 ByRef tableName As String, _
                                 ByRef schemaFile As String, _
                                 ByRef hasFieldNames As Boolean) As Long

        lngExportCSV = 0

        ''---- mdb存在ﾁｪｯｸ ----
        If File.Exists(mdbPath) = False Then
            lngExportCSV = -1
            Exit Function
        End If

        Dim oAccess As Access.ApplicationClass
        oAccess = New Access.ApplicationClass

        ''---- DB接続 ----
        Try
            oAccess.OpenCurrentDatabase(mdbPath)
        Catch ex As Exception
            lngExportCSV = Err.Number
            Exit Function
        End Try

        ''---- csvｴｸｽﾎﾟｰﾄ ----
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
    ' 機能　　 :指定ﾃﾞｰﾀｲﾝﾎﾟｰﾄ処理
    '
    ' 返り値　 :正常終了 - 0    ｴﾗｰ - 0以外
    '
    ' 引き数　 :ARG1 - mdbﾌｧｲﾙ名(ﾌﾙﾊﾟｽ付)
    ' 　　　    ARG2 - csvﾌｧｲﾙ名(ﾌﾙﾊﾟｽ付)
    ' 　　　    ARG3 - ﾃｰﾌﾞﾙ名
    ' 　　　    ARG4 - ｲﾝﾎﾟｰﾄ定義名
    ' 　　　    ARG5 - 先頭行ﾌｨｰﾙﾄﾞ認識ﾌﾗｸﾞ
    '
    ' 機能説明 :
    '
    ' 備考　　 :
    '
    Public Function lngImportCSV(ByRef mdbPath As String, _
                                 ByRef csvPath As String, _
                                 ByRef tableName As String, _
                                 ByRef schemaFile As String, _
                                 ByRef hasFieldNames As Boolean) As Long

        lngImportCSV = 0

        ''---- mdb存在ﾁｪｯｸ ----
        If File.Exists(mdbPath) = False Then
            lngImportCSV = -1
            Exit Function
        End If

        ''---- csv存在ﾁｪｯｸ ----
        If File.Exists(csvPath) = False Then
            lngImportCSV = -2
            Exit Function
        End If

        Dim oAccess As Access.ApplicationClass
        oAccess = New Access.ApplicationClass

        ''---- DB接続 ----
        Try
            oAccess.OpenCurrentDatabase(mdbPath)
        Catch ex As Exception
            lngImportCSV = Err.Number
            Exit Function
        End Try

        ''---- csvｲﾝﾎﾟｰﾄ ----
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
