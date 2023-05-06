////////////////////////////////////////////////////////////////////////////////
//             ATSTest - иллюстрация комплексной обработки исключений
//                             (c) Николай Мазуркин
//    http://mazurkin.virtualave.net, mazurkin@mailru.com, mazurkin@chat.ru
// _____________________________________________________________________________
//                                 Главный модуль
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
