////////////////////////////////////////////////////////////////////////////////
//             ATSTest - ����������� ����������� ��������� ����������
//                             (c) ������� ��������
//    http://mazurkin.virtualave.net, mazurkin@mailru.com, mazurkin@chat.ru
// _____________________________________________________________________________
//                                 ������� �����
////////////////////////////////////////////////////////////////////////////////

unit ATSDialogMain;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls;

type
  TDialogMain = class(TForm)
    btnErrorRaise: TButton;
    rgrpErrorType: TRadioGroup;
    procedure OnRaiseError(Sender: TObject);
  protected
    procedure ErrorPointer;
    procedure ErrorExceptionUnhandled;  
    procedure ErrorExceptionHandled;
    procedure ErrorAPI;
    procedure ErrorUnhandled;
    procedure ErrorHandled;
    procedure ErrorAssert;
    function  ErrorFunction: Integer;
  end;

var
  DialogMain: TDialogMain;

implementation

uses ATSAssert;

{$R *.DFM}

////////////////////////////////////////////////////////////////////////////////
// ������������� ������
////////////////////////////////////////////////////////////////////////////////
procedure TDialogMain.OnRaiseError(Sender: TObject);
begin
  try
    // ���������� ������, ������� ������� ������������
    case rgrpErrorType.ItemIndex of
      0 : ErrorPointer;
      1 : ErrorExceptionUnhandled;
      2 : ErrorExceptionHandled;
      3 : ErrorAPI;
      4 : ErrorUnhandled;
      5 : ErrorHandled;
      6 : ErrorAssert;
      7 : ErrorFunction;
    end;
    // � ������ ��������� �� ����� ��������� ������� ������ � ����������. ������
    // �������������� ���������� ����� ��������� � �������� ���������� ����������.
    // ���������� � ���� ��������� ������������ ��� ������ ���� �������������� ����������,
    // ������� ����� ��������� "������" ���������� �������������� ��� ������������.
    // � ������ ������ ����������� ������ ������ ��������� ����� �������, ��� ���
    // � ������������ ������ ���� ������� ��� ���������� ���������.
  except
    raise AssertInternal('{97D33BA0-2A45-11D4-ACD0-009027350D25}');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
// ��������� ������ �� "��������" ���������
////////////////////////////////////////////////////////////////////////////////
procedure TDialogMain.ErrorPointer;
var
  WrongPointer : PInteger;
begin
  try
    // ����������� ��������������� ��������� �������� ����� (��������� ��������)
    WrongPointer  := nil;
    // �������� �������� �� ����� ������ �����-���� ��������
    // ������������ ����������, ��������� ��������������� ����� ������������
    // ����� ����� ������������ ������.
    WrongPointer^ := 100;
  except
    raise AssertInternal('{97D33BA1-2A45-11D4-ACD0-009027350D25}');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
// ���������������� �������� ����� (�������� �������)
////////////////////////////////////////////////////////////////////////////////
procedure TDialogMain.ErrorExceptionUnhandled;
var
  FS : TFileStream;
begin
  try
    // ��������� �������������� ���� (��������� ��������), ������ �� �������, ���
    // �� ������ ��������� ��������� (�������� �������) � �� ����� ������������� �������
    // �������� �� ��������� ������.
    // ��������� ����������, ������� ��������������� ����� ������������.
    FS := TFileStream.Create('qgfasajasasasa.das', fmOpenRead);
    // ��� ������� �� �����������, �.�. ���� �� ����� ���� ������
    FS.Seek(0, 0);
    FS.Free;
    // �����������, � ����� ���������� ������������ �����������
    //    FS := TFileStream.Create('qgfasajasasasa.das', fmOpenRead);
    //    try
    //      FS.Seek(0, 0);
    //    finally
    //      FS.Free;
    //    end;
    // ������, ��� ��� �� ���������� ������� "���������" ������ � ������� FS.Seek(0, 0)
    // ���������� ������ ����� ������ ���������� � ���������� ���������, �������
    // �������� �� ������������ ������� ������ ������������� �� ����� - ��������� ��� 
    // ����� ����� �������.
  except
    raise AssertInternal('{97D33BA3-2A45-11D4-ACD0-009027350D25}');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
// �������������� �������� ����� (���������� �������)
////////////////////////////////////////////////////////////////////////////////
procedure TDialogMain.ErrorExceptionHandled;
var
  FS : TFileStream;
begin
  try
    // ��������� �������������� ���� (��������� ��������), ������ �� �� �������, ���
    // �� ������ ����� ������ (���������� �������). �� ������������� ��������������
    // ����������, ������� �������������, ������������ ��������� ������ � ������ ���������.
    try
      FS := TFileStream.Create('qgfasajasasasa.das', fmOpenRead);
    except
      ShowMessage('���������� ������� ���� �����-��, ��������� ��-��.');
      Exit;
    end;
    // ��� ������� �� �����������, �.�. ���� �� ����� ���� ������
    FS.Seek(0, 0);
    FS.Free;
  except
    raise AssertInternal('{97D33BA2-2A45-11D4-ACD0-009027350D25}');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
// ��������� ������ API
////////////////////////////////////////////////////////////////////////////////
procedure TDialogMain.ErrorAPI;
var
  SomeHandle : THandle;
begin
  try
    // ����������� ��������������� ����������� �������� �������� (��������� ��������)
    SomeHandle := INVALID_HANDLE_VALUE;
    // ��������� ��������� ���������� API-������� �, ����-���, ������� ��������� ������ 
    AssertWin32(CloseHandle(SomeHandle), '{97D33BA5-2A45-11D4-ACD0-009027350D25}');
    // � ������ ��������� �� ����� ���������� ������� ���������� �, �� ������ ������,
    // �� ���������� ������� � ������� ���������� ����� ���������. ������, ����������
    // ��� ���� ����� ���������� � � API-������� (�������� ��� ��������� � ������), �,
    // ��-������, ������� ������ ����� �������� ����������� �� ��� ���������, �.�. ��������
    // �� �����������, ����������� � ��., � ���� ������� ����� ��������� ������� ������� ����.
  except
    raise AssertInternal('{97D33BA4-2A45-11D4-ACD0-009027350D25}');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
// �������� ���������� � ���������� ����������
////////////////////////////////////////////////////////////////////////////////
procedure TDialogMain.ErrorAssert;
begin
  try
    // ��������� �������� ������������ ������� �������� ��������� ��������.
    AssertMsg(Self = nil, '{CFCC0842-2A5C-11D4-ACD0-009027350D25}');
    // �� �� �����. ���� � ���� ��������� �� ����� ���������� ����������, ����������
    // ��������� � ������ ����������� ���������� ���������� ���������.
  except
    raise AssertInternal('{97D33BA7-2A45-11D4-ACD0-009027350D25}');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
// ���������������� ������
////////////////////////////////////////////////////////////////////////////////
procedure TDialogMain.ErrorUnhandled;
var
  Strs : TStringList;
begin
  // � ������ ������� �� "������" � ����������� ��������� ������ � ����� ���������������
  // � ������������ �����.
  // ����������� � ���� ��������� ���������� �� �������������� ��������, ������ ���
  // �������������� � ���������, ������� ������� ������ ���������. ����������� �����
  // ����� �������.
  // ���� �� ��������� ���������, ��� �� ���� ���������� �� ����� �� ������ �����������
  // ���������� ����� ���������� ���������� ������������ ���������� �� �������
  // TApplication.OnException - ������ � ����������� ������ ��� �� ����� ���� � ����.
  Strs := TStringList.Create;
  try
    Strs.Strings[10];
  finally
    Strs.Free;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
// �������������� ������
////////////////////////////////////////////////////////////////////////////////
procedure TDialogMain.ErrorHandled;
var
  Strs : TStringList;
begin
  // ���������� ������ - ������, � ����� ����������� ��������� ������.
  try
    Strs := TStringList.Create;
    Strs.Strings[10];
    Strs.Free;
  except
    raise AssertInternal('{97D33BA6-2A45-11D4-ACD0-009027350D25}');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
// ������� � �������
////////////////////////////////////////////////////////////////////////////////
function TDialogMain.ErrorFunction: Integer;
begin
  try
    // ��������� �������� ������ �����������
    AssertMsg(1 = 2, '{CFCC0844-2A5C-11D4-ACD0-009027350D25}');
    // ����������� �������� ���������������� ����������
    Result := 0;
    // ������ ������ ������������ ������������� ���������� ����������
    //   raise AssertInternal(..)
    // ������ ������
    //   AssertInternal(..);
    // �� ����� ���� ��������� �������� raise ��� ������� �������������,
    // ��� ��� �� ��� ���� �� �����������, � ��������� ����������� � ����������
    // ������� AssertInternal. ������, ��� �������� ��������� raise, ����������
    // ����� �������� �������������� ��������������, ��� ��� �� ��� ������, ��� ������
    // ������� ���������, ������� ���������� ������� �������. � � ���� ������ ���������
    // ������� ������������� ��� �� �����������.  
  except
    raise AssertInternal('{CFCC0843-2A5C-11D4-ACD0-009027350D25}');
  end;
end;

end.