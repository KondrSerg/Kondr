////////////////////////////////////////////////////////////////////////////////
//             ATSTest - иллюстрация комплексной обработки исключений
//                             (c) Николай Мазуркин
//    http://mazurkin.virtualave.net, mazurkin@mailru.com, mazurkin@chat.ru
// _____________________________________________________________________________
//                      Процедуры обработки фатальных ошибок
////////////////////////////////////////////////////////////////////////////////

unit ATSAssert;

interface

uses
  Windows, SysUtils, Forms;

////////////////////////////////////////////////////////////////////////////////
// Процедуры обработки фатальных ситуаций
////////////////////////////////////////////////////////////////////////////////
procedure AssertMsg(Condition: Boolean; const Mark: string);
procedure AssertWin32(Condition: Boolean; const Mark: string);
function  AssertInternal(const Mark: string): TObject;

implementation

uses Main;

////////////////////////////////////////////////////////////////////////////////
// Обработка ошибочной ситуации
////////////////////////////////////////////////////////////////////////////////
procedure HandleError(const Msg: string);
var
  DeadFile : TextFile;
  DeadName : string;
  ErrorMsg : string;
begin
  // Формируем сообщение
  ErrorMsg := Msg + #13#13;
  // Если это исключение, то добавим его тип и адрес
  if ExceptObject <> nil then
    ErrorMsg := ErrorMsg + Format('Исключение типа %s, адрес 0x%8.8x'#13#13,
      [ExceptObject.ClassName, DWord(ExceptAddr)]
    );

  ErrorMsg := ErrorMsg + 'Выполнение программы будет прекращено.';
  // Показываем диалог
  MessageBox(0, PChar(ErrorMsg),
    'Фатальная ошибка', MB_OK+MB_ICONWARNING+MB_APPLMODAL+MB_SETFOREGROUND+MB_TOPMOST);
  // Генерируем уникальное имя файла
  SetLength(DeadName, MAX_PATH+10);
  GetTempFileName(PChar(ExtractFilePath(ParamStr(0))), 'err', 0, @DeadName[1]);
  // Записываем ошибку в посмертный лог - на случай если пользователь не догадается
  // записать код ошибки и сообщение.
  AssignFile(DeadFile, DeadName);
  ReWrite(DeadFile);
  WriteLn(DeadFile, ErrorMsg);
  CloseFile(DeadFile);
  // Немедленно закрываем процесс
  //ExitProcess(1);
end;

////////////////////////////////////////////////////////////////////////////////
// TExceptHandler
// _____________________________________________________________________________
// Глобальный обработчик необработанных локально исключений
////////////////////////////////////////////////////////////////////////////////
type
  TExceptHandler = class (TObject)
    procedure   OnException(Sender: TObject; E: Exception);
  end;

var
  ExceptHandler : TExceptHandler;

////////////////////////////////////////////////////////////////////////////////
// Обработка исключений
////////////////////////////////////////////////////////////////////////////////
procedure TExceptHandler.OnException(Sender: TObject; E: Exception);
begin
  HandleError(PChar(E.Message));
end;

////////////////////////////////////////////////////////////////////////////////
// ATSAssert
// _____________________________________________________________________________
// Процедуры обработки фатальных ситуаций
////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////
// Проверка выполнимости условия
////////////////////////////////////////////////////////////////////////////////
procedure AssertMsg(Condition: Boolean; const Mark: string);
begin
  // Если условие выполняется, то выходим сразу
  if Condition then Exit;
  // Обрабатываем возникшую ошибку
  HandleError('Фатальная ошибка '+Mark);
end;

////////////////////////////////////////////////////////////////////////////////
// Проверка выполнимости API-функции
////////////////////////////////////////////////////////////////////////////////
procedure AssertWin32(Condition: Boolean; const Mark: string);
var
  ErrorMsg : string;
begin
  // Если условие выполняется, то выходим
  if Condition then Exit;
  // Получаем сообщение об ошибке
  ErrorMsg := SysErrorMessage(GetLastError);
  // Обрабатываем возникшую ошибку
  HandleError('Фатальная ошибка '+Mark+#13#13+ErrorMsg);
end;

////////////////////////////////////////////////////////////////////////////////
// Генерация фатального исключения
////////////////////////////////////////////////////////////////////////////////
function AssertInternal(const Mark: string): TObject;
begin
  // Проверяем тип сгенерированного исключения и обрабатываем возникшую ошибку
  if ExceptObject is Exception then
    HandleError('Фатальная ошибка '+Mark+#13#13+PChar(Exception(ExceptObject).Message))
  else
    HandleError('Фатальная ошибка '+Mark);
  // Следующая строчка никогда не выполнится, она нужна только для успокоения компилятора
  Result := nil;
end;

////////////////////////////////////////////////////////////////////////////////
// Автоматическое создание глобального обработчика исключений
////////////////////////////////////////////////////////////////////////////////
initialization
  ExceptHandler := TExceptHandler.Create;
  Application.OnException := ExceptHandler.OnException;

end.
