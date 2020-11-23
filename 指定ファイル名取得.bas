Attribute VB_Name = "Module1"
Function FileName() As String
      
    '=====================
    '   ファイル指定
    '=====================
    With Application.FileDialog(msoFileDialogOpen)
        .Title = "ファイルの選択"
        'ファイルの種類を設定
        .Filters.Clear
        .Filters.Add "すべてのファイル", "*.*"
        '複数ファイル選択を許可しない
        .AllowMultiSelect = False
          
        'ダイアログを表示
        If .Show = -1 Then
            'ファイルが選択されたとき
            'そのフルバスを返り値に設定
            FileName = Trim(.SelectedItems.Item(1))
        Else
            'ファイルが選択されなければ長さゼロの文字列を返す
            FileName = ""
        End If
    End With
           
End Function
