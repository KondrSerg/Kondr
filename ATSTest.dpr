////////////////////////////////////////////////////////////////////////////////
//             ATSTest - ����������� ����������� ��������� ����������
//                             (c) ������� ��������
//    http://mazurkin.virtualave.net, mazurkin@mailru.com, mazurkin@chat.ru
// _____________________________________________________________________________
//                                 ������� ������
////////////////////////////////////////////////////////////////////////////////

program ATSTest;

uses
  Forms,
  ATSDialogMain in 'ATSDialogMain.pas' {DialogMain},
  ATSAssert in 'ATSAssert.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TDialogMain, DialogMain);
  Application.Run;
end.
