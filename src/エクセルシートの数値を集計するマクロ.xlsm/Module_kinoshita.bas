Attribute VB_Name = "Module_kinoshita"
'---------------------------------------------------------------
'�Z���ʒu�E�V�[�g�����w�肷��@�\
'�i�P�j�G�N�Z���V�[�g�w�肵�ĊJ��
'�i�Q�j�J�����G�N�Z���V�[�g�̂����u�Z���ʒu�v�u�V�[�g���v���w��
'�i�R�j�w��ʒu�̃Z�����A�N�e�B�u�ɂ��Ēl��C�ӂ̕ϐ��ɑ������
'Owner kinoshita
'---------------------------------------------------------------


Function Kagebunshin(openSheet As String, rangeSelect As String) As Long

' kagebunshin Macro

    'Dim sheetname As String
    'Dim rangeselect As String
    'Dim opensheet As String
    Dim targetvalue As Long

    Worksheets(openSheet).Select
    'MsgBox "openSheet ���J���܂���"
    
    Worksheets(openSheet).Activate
    Range(rangeSelect).Activate

    targetvalue = Range(rangeSelect).Value
    
    Kagebunshin = targetvalue

    'MsgBox "�ϐ��̒l��" & targetvalue & "�ł�"
    '    Worksheets(sheetname).Range(rangeselect) = "test2"
    '    Sheets.Add After:=ActiveSheet
    '    ActiveCell.Formula = "a"
    '
    '      MsgBox "���� "
      
End Function
