var App = new ActiveXObject("Excel.Application");
var WshShell = new ActiveXObject( "WScript.Shell" );

// Excel ��\��( ����������R�����g�� )
App.Visible = true;
// �x�����o���Ȃ�
App.DisplayAlerts = false;

// Minimize : GetSaveAsFilename ��O�ʂɏo����
// �{���A-4140 �ł��� WScript.Shell �� Run �Ɠ��� 2 ���g����
App.WindowState = 2

// ****************************
// ��̃t�@�C�����J��
// ****************************
// https://docs.microsoft.com/ja-jp/office/vba/api/excel.application.getopenfilename
var FilePath = App.GetOpenFilename("Excel �t�@�C�� (*.xlsx), *.xlsx,�S��,*.*", 1,"�t�@�C���̑I��",null, false );
// ���I���̏ꍇ
if( FilePath === false ) {
    WshShell.Popup("�t�@�C���̎Q�ƑI�����L�����Z������܂���")
    App.Quit();
    App = null;
    WScript.Quit();
}

// ****************************
// �u�b�N���J��
// ****************************
var Book = App.Workbooks.Open(FilePath);

// ****************************
// �ŏI�V�[�g��O�ɃR�s�[
// ****************************
Book.Sheets(Book.Sheets.Count).Copy( Book.Sheets(Book.Sheets.Count) );

// �R�s�[�����̂� �A�N�e�B�u�ɂȂ�܂�
var Target = Book.ActiveSheet;

// �擪��𕶎���ɐݒ�
Target.Range("A:A").NumberFormatLocal = "@";

// �Z���ɒl���Z�b�g
Target.Cells(1, 1).Value = "�Ј��R�[�h";
Target.Cells(2, 1).Value = "0001"
Target.Cells(3, 1).Value = "0002";
Target.Cells(4, 1).Value = "0003";
Target.Cells(1, 2).Value = "�Ј���";
Target.Cells(2, 2).Value = "�R�c�@���Y�r���q��";
Target.Cells(3, 2).Value = "��؁@��Y";
Target.Cells(4, 2).Value = "�����@�m�q";

// �񕝎�������
Target.Columns("B:B").EntireColumn.AutoFit();

// ****************************
// �ۑ�
// ****************************
try {
    Book.SaveAs( FilePath );
    Book.Close();
}
catch (error) {
    WshShell.Popup( "Book.SaveAs �ŃG���[���������܂��� : " + error.description );
}

// �I��
App.Quit();

App = null;

WshShell.Run( "RunDLL32.EXE shell32.dll,ShellExec_RunDLL " + FilePath );
