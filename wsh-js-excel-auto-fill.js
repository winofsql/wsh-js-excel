// ****************************
// ��������
// ****************************
WScript.Echo( "�������J�n���܂�" );
var WshShell = new ActiveXObject("WScript.Shell");
var ExcelApp = new ActiveXObject( "Excel.Application" );

// �f�o�b�O���́AExcel �̖{�̂�\�������ď󋵂�����悤�ɂ���
ExcelApp.Visible = true;
// UI �Ń`�F�b�N������悤�ȃ_�C�A���O��\�������Ɏ��s����
ExcelApp.DisplayAlerts = false;

try {

    // ****************************
    // �u�b�N�ǉ�
    // ****************************
    var Book = ExcelApp.Workbooks.Add();

    // �ʏ��̃V�[�g���쐬����Ă��܂�
    var Sheet = Book.Worksheets( 1 );

    // ****************************
    // �V�[�g���ύX
    // ****************************
    Sheet.Name = "JScript�̏���";

    // ****************************
    // �Z���ɒl�𒼐ڃZ�b�g
    // ****************************
    for( var i = 1; i <= 10; i++ )
    {
        Sheet.Cells(i, 1) = "���� : " + i;
    }

    // ****************************
    // 1�̃Z������
    // AutoFill �Œl���Z�b�g
    // ****************************
    Sheet.Cells(1, 2) = "�q";
    // ��ƂȂ�Z���͈�
    var SourceRange = Sheet.Range(Sheet.Cells(1, 2), Sheet.Cells(1,2));
    // �I�[�g�t�B���͈̔�(��ƂȂ�Z���͈͂��܂� )
    var FillRange = Sheet.Range(Sheet.Cells(1, 2), Sheet.Cells(10,2));
    SourceRange.AutoFill(FillRange);

    // ****************************
    // �ۑ�
    // ****************************
    Book.SaveAs( WshShell.CurrentDirectory + "\\sample.xlsx" );

} catch (error) {
    ExcelApp.Quit();
    ExcelApp = null;
    WshShell.Popup(error.description);
    WScript.Quit();	
}

ExcelApp.Quit();
ExcelApp = null;

// ****************************
// �t�@�C���̍Ō�
// ****************************
WshShell.Popup("�������I�����܂�");
