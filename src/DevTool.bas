Attribute VB_Name = "DevTool"
Option Explicit

'<summary>
' 機能: 全VBAソースをエクスポートする
' 引数: なし
' </summary>
' <remarks>
' 1. VBEにおいてMicrosoft Visual Basic for Applications Extensibilityへの参照を追加する。
' 2. 「VBAプロジェクト オブジェクトモデルへのアクセスを信頼する」オプションを指定する。
' </remarks>
Public Sub ExportAllSource()
    Dim module                  As VBComponent      ' モジュール
    Dim moduleList              As VBComponents     ' VBAプロジェクトの全モジュール
    Dim extension               As String           ' モジュールの拡張子
    Dim sFilePath               As String           ' エクスポートファイルパス
    Dim sSaveFolder             As String           ' 保存先フォルダ
    
    sSaveFolder = ActiveWorkbook.Path & "\" & "src"

    If Dir(sSaveFolder, vbDirectory) = "" Then
        MkDir sSaveFolder
    End If
    
    ' 処理対象ブックのモジュール一覧を取得
    Set moduleList = ActiveWorkbook.VBProject.VBComponents
    
    ' VBAプロジェクトに含まれる全てのモジュールをループ
    For Each module In moduleList
        ' クラス
        If (module.Type = vbext_ct_ClassModule) Then
            extension = "cls"
        ' フォーム
        ElseIf (module.Type = vbext_ct_MSForm) Then
            ' .frxも一緒にエクスポートされる
            extension = "frm"
        ' 標準モジュール
        ElseIf (module.Type = vbext_ct_StdModule) Then
            extension = "bas"
        ' その他
        Else
            ' エクスポート対象外のため次ループへ
            GoTo CONTINUE
        End If
        
        ' エクスポート実施
        sFilePath = sSaveFolder & "\" & module.Name & "." & extension
        
        Call module.Export(sFilePath)
        
        ' 出力先確認用ログ出力
        Debug.Print sFilePath
CONTINUE:
    Next
    
End Sub
