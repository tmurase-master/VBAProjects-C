Attribute VB_Name = "Main"
Sub �G�N�Z���V�[�g�̐��l���W�v����}�N��()

'-----�ϐ��錾�J�n-----
    Dim targetFolderPath As String  '��ƑΏۃt�H���_
    Dim targetFileNames() As String '��ƑΏۃG�N�Z���V�[�g���i�z��j
    
    Dim resultFileName As String  '���ʏo�̓G�N�Z���t�@�C����
    Dim resultFile As Workbook    '���ʏo�̓G�N�Z���u�b�N
    Dim resultSheet As Worksheet  '���ʏo�̓G�N�Z���V�[�g
'-----�ϐ��錾�I��-----

'-----���������J�n-----
    ' �t�H���_���w�肵�ĕϐ��Ɋi�[�iOwner kitazima�j
    Call SelectBooks(targetFolderPath, targetFileNames)
    ' ���ʏo�͗p�V�[�g��ݒ�iOwner suzuki�j
    Call OpenResultSheet(resultFile, resultSheet)
'-----���������I��-----
    
'-----���C�������J�n-----
    ' ���t�@�C�����J���A���������s���擾�iOwner kitazima�j
    Call ProcessBooks(targetFolderPath, targetFileNames, resultSheet)
'-----���C�������I��-----
    
'-----�ŏI�����J�n-----
    ' ���ʏo�̓t�@�C���̕ۊǁiOwner suzuki�j
    Call OutputResultFile(resultFile)
'-----�ŏI�����I��-----

End Sub
