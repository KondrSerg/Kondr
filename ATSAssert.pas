////////////////////////////////////////////////////////////////////////////////
//             ATSTest - ����������� ����������� ��������� ����������
//                             (c) ������� ��������
//    http://mazurkin.virtualave.net, mazurkin@mailru.com, mazurkin@chat.ru
// _____________________________________________________________________________
//                      ��������� ��������� ��������� ������
////////////////////////////////////////////////////////////////////////////////

unit ATSAssert;

interface

uses
  Windows, SysUtils, Forms;

////////////////////////////////////////////////////////////////////////////////
// ��������� ��������� ��������� ��������
////////////////////////////////////////////////////////////////////////////////
procedure AssertMsg(Condition: Boolean; const Mark: string);
procedure AssertWin32(Condition: Boolean; const Mark: string);
function  AssertInternal(const Mark: string): TObject;

implementation

uses Main;

////////////////////////////////////////////////////////////////////////////////
// ��������� ��������� ��������
////////////////////////////////////////////////////////////////////////////////
procedure HandleError(const Msg: string);
var
  DeadFile : TextFile;
  DeadName : string;
  ErrorMsg : string;
begin
  // ��������� ���������
  ErrorMsg := Msg + #13#13;
  // ���� ��� ����������, �� ������� ��� ��� � �����
  if ExceptObject <> nil then
    ErrorMsg := ErrorMsg + Format('���������� ���� %s, ����� 0x%8.8x'#13#13,
      [ExceptObject.ClassName, DWord(ExceptAddr)]
    );

  ErrorMsg := ErrorMsg + '���������� ��������� ����� ����������.';
  // ���������� ������
  MessageBox(0, PChar(ErrorMsg),
    '��������� ������', MB_OK+MB_ICONWARNING+MB_APPLMODAL+MB_SETFOREGROUND+MB_TOPMOST);
  // ���������� ���������� ��� �����
  SetLength(DeadName, MAX_PATH+10);
  GetTempFileName(PChar(ExtractFilePath(ParamStr(0))), 'err', 0, @DeadName[1]);
  // ���������� ������ � ���������� ��� - �� ������ ���� ������������ �� ����������
  // �������� ��� ������ � ���������.
  AssignFile(DeadFile, DeadName);
  ReWrite(DeadFile);
  WriteLn(DeadFile, ErrorMsg);
  CloseFile(DeadFile);
  // ���������� ��������� �������
  //ExitProcess(1);
end;

////////////////////////////////////////////////////////////////////////////////
// TExceptHandler
// _____________________________________________________________________________
// ���������� ���������� �������������� �������� ����������
////////////////////////////////////////////////////////////////////////////////
type
  TExceptHandler = class (TObject)
    procedure   OnException(Sender: TObject; E: Exception);
  end;

var
  ExceptHandler : TExceptHandler;

////////////////////////////////////////////////////////////////////////////////
// ��������� ����������
////////////////////////////////////////////////////////////////////////////////
procedure TExceptHandler.OnException(Sender: TObject; E: Exception);
begin
  HandleError(PChar(E.Message));
end;

////////////////////////////////////////////////////////////////////////////////
// ATSAssert
// _____________________________________________________________________________
// ��������� ��������� ��������� ��������
////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////
// �������� ������������ �������
////////////////////////////////////////////////////////////////////////////////
procedure AssertMsg(Condition: Boolean; const Mark: string);
begin
  // ���� ������� �����������, �� ������� �����
  if Condition then Exit;
  // ������������ ��������� ������
  HandleError('��������� ������ '+Mark);
end;

////////////////////////////////////////////////////////////////////////////////
// �������� ������������ API-�������
////////////////////////////////////////////////////////////////////////////////
procedure AssertWin32(Condition: Boolean; const Mark: string);
var
  ErrorMsg : string;
begin
  // ���� ������� �����������, �� �������
  if Condition then Exit;
  // �������� ��������� �� ������
  ErrorMsg := SysErrorMessage(GetLastError);
  // ������������ ��������� ������
  HandleError('��������� ������ '+Mark+#13#13+ErrorMsg);
end;

////////////////////////////////////////////////////////////////////////////////
// ��������� ���������� ����������
////////////////////////////////////////////////////////////////////////////////
function AssertInternal(const Mark: string): TObject;
begin
  // ��������� ��� ���������������� ���������� � ������������ ��������� ������
  if ExceptObject is Exception then
    HandleError('��������� ������ '+Mark+#13#13+PChar(Exception(ExceptObject).Message))
  else
    HandleError('��������� ������ '+Mark);
  // ��������� ������� ������� �� ����������, ��� ����� ������ ��� ���������� �����������
  Result := nil;
end;

////////////////////////////////////////////////////////////////////////////////
// �������������� �������� ����������� ����������� ����������
////////////////////////////////////////////////////////////////////////////////
initialization
  ExceptHandler := TExceptHandler.Create;
  Application.OnException := ExceptHandler.OnException;

end.
