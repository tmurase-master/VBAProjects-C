Attribute VB_Name = "Module1"

'---------------------------------------------------------------
'�@�Z���ʒu�E�V�[�g�����w�肷��}�N��
'�i�P�j�G�N�Z���V�[�g�w�肵�ĊJ��
'�i�Q�j�J�����G�N�Z���V�[�g�̂����u�Z���ʒu�v�u�V�[�g���v���w��
'�i�R�j�w��ʒu�̃Z�����A�N�e�B�u�ɂ��Ēl��C�ӂ̕ϐ��ɑ������
'Owner kinoshita
'---------------------------------------------------------------


Sub kagebunshin()
Attribute kagebunshin.VB_ProcData.VB_Invoke_Func = " \n14"
'
' kagebunshin Macro

Dim sheetname As String
Dim rangeselect As String
Dim opensheet As String
Dim targetvalue As Long

sheetname = "Sheet1"

rangeselect = "A1"

opensheet = "test01"



Worksheets(opensheet).Select
    MsgBox "opensheet ���J���܂���"
    
Worksheets(opensheet).Activate
Range("A3").Activate

targetvalue = Range("C3").Value

MsgBox "�ϐ��̒l��" & targetvalue & "�ł�"
'    Worksheets(sheetname).Range(rangeselect) = "test2"
'    Sheets.Add After:=ActiveSheet
'    ActiveCell.Formula = "a"
'
'      MsgBox "���� "
      
End Sub
