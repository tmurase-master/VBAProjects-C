Attribute VB_Name = "kitazima"
'---------------------------------------------------------------
'�A�w��t�H���_���̃G�N�Z���V�[�g�����ԂɊJ���ĕ���}�N��
'�i�P�j��ƑΏۃt�H���_�p�X���w��
'�i�Q�j��ƑΏۃt�H���_�p�X���̃G�N�Z���V�[�g���擾
'�i�R�j�i�Q�j�̃G�N�Z���V�[�g�����ԂɁu�J���˕���v
'Owner kitazima
'---------------------------------------------------------------

Sub OpenAndCloseBooks()

'-----�ϐ��錾-----
    Dim foPath As String '�Ώۂ̃t�H���_�p�X
    Dim fiName As String 'foPath���̃G�N�Z���t�@�C����
    
'-----�t�H���_�p�X�E�t�@�C���p�X�擾�i3�p�^�[���j-----
    '�t�H���_�p�X�I��1�F�_�C�A���O����I��
    With Application.FileDialog(msoFileDialogFolderPicker) '�t�H���_�I����ʂ�\��
        If .Show = 0 Then '���I���̏ꍇ
            Exit Sub '�}�N�����I��
        Else '�I�������ꍇ
            foPath = .SelectedItems(1) '�I�������t�H���_�p�X���擾
        End If
    End With
    
    '�t�H���_�p�X�I��2�F�}�N����u�����t�H���_�Ƃ���
    foPath = ThisWorkbook.Path
    
    '�t�H���_�p�X�I��3�F�p�X�Œ�i���ړ��́j
    foPath = "C:\Users\Kitajima\Desktop"
    
    fiName = Dir(foPath & "\*.xls*") '�Ώۃt�H���_�̍ŏ��̃t�@�C����
    
'-----�t�@�C�����J���ĕ���-----
    Do While fiName <> "" '�t�H���_�ɃG�N�Z���t�@�C��������ꍇ
        Workbooks.Open foPath & "\" & fiName '�J��
        
        '�����ɊJ�������Ƃ̏������L��
        
        Workbooks(fiName).Close SaveChanges:=False '�㏑�������t�@�C�������
        fiName = Dir '���̃t�@�C���̌���
    Loop
    
End Sub

