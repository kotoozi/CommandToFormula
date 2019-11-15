VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub InsertEquation(str)
    ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, 0, 0, 0, 0).Select
    Selection.Font.Size = 15
    Application.CommandBars.ExecuteMso "InsertBuildingBlocksEquationsGallery"
    Selection.Text = str
    Selection.ShapeRange.TextFrame2.WordWrap = msoFalse
    Application.CommandBars.ExecuteMso "EquationProfessional"
End Sub
Sub Sample()
    InsertEquation "F(t)=\sum_(x=0)^(n-1) \of \begin f(x) e^(-\ii  2\pi tx/N) \end"
End Sub
Sub OutputMatrix()
'<summary>
' OutputMatrix : ()
'</summary>
'<remarks>
' セルから文字列情報を取得して行列化して出力する
' 空白個所は0として処理する
'</remarks>
    '#region 変数
    Dim matrixCommand As String          ' 行列コマンド用の文字列(String型)
    Dim i, j As Integer                  ' カウント用の変数(int型)
    Dim outputPermission As Integer      ' 行列出力許可の確認
    '#endregion (変数)
    matrixCommand = "(\matrix("          ' 行列コマンド作成開始
    If (TypeName(Selection) <> "Range") Then    ' セルを選択していない場合は異常終了
        MsgBox "セルを選択してください"
        End                              ' 終了
    End If
    For i = Selection.Row To (Selection.Row + Selection.Rows.Count - 1)
        For j = Selection.Column To (Selection.Column + Selection.Columns.Count - 1)
            matrixCommand = matrixCommand & nullToZero(Cells(i, j).Value)   ' セル内のテキスト情報を取得
            If (j <> Selection.Column + Selection.Columns.Count - 1) Then   ' 選択範囲の終端の場合の処理
                matrixCommand = matrixCommand & "&"
            End If
        Next j
        If (i <> Selection.Row + Selection.Rows.Count - 1) Then
            matrixCommand = matrixCommand + "@"                             ' 選択範囲の終端の場合の処理
        End If
    Next i
    matrixCommand = matrixCommand + "))" ' 行列取得コマンド作成終了
    outputPermission = MsgBox("行列を出力しますがよろしいですか？", vbYesNo, "行列出力")
    If (outputPermission = vbYes) Then
        InsertEquation matrixCommand
    End If
End Sub
Function nullToZero(ByVal target As String)
'<summary>
' nullToZero : (引数 : String)
'</summary>
'<remarks>
' 文字列に何も無ければ0を返す関数
' 関数化してメソッド(サブルーチン)の行数を減らす
'</remarks>
    If target = "" Then
        nullToZero = "0"
    Else
        nullToZero = target
    End If
End Function

