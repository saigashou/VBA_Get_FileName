Attribute VB_Name = "Module1"
Function FileName() As String
      
    '=====================
    '   �t�@�C���w��
    '=====================
    With Application.FileDialog(msoFileDialogOpen)
        .Title = "�t�@�C���̑I��"
        '�t�@�C���̎�ނ�ݒ�
        .Filters.Clear
        .Filters.Add "���ׂẴt�@�C��", "*.*"
        '�����t�@�C���I���������Ȃ�
        .AllowMultiSelect = False
          
        '�_�C�A���O��\��
        If .Show = -1 Then
            '�t�@�C�����I�����ꂽ�Ƃ�
            '���̃t���o�X��Ԃ�l�ɐݒ�
            FileName = Trim(.SelectedItems.Item(1))
        Else
            '�t�@�C�����I������Ȃ���Β����[���̕������Ԃ�
            FileName = ""
        End If
    End With
           
End Function
