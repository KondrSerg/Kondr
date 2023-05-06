unit UpolsRed;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Data.DB, Data.Win.ADODB,ActiveDs_TLB,ActiveX,ComObj;

type
  TFPolsRed = class(TForm)
    Button1: TButton;
    ComboBox1: TComboBox;
    Button2: TButton;
    Label1: TLabel;
    Label2: TLabel;
    ComboBox2: TComboBox;
    Label3: TLabel;
    ComboBox3: TComboBox;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    Memo1: TMemo;
    Label4: TLabel;
    Memo2: TMemo;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    function GetObject(const name: string): IDispatch;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FPolsRed: TFPolsRed;
         const
 LOGON_WITH_PROFILE  = $00000001;
  LOGON_NETCREDENTIALS_ONLY = $00000002;
  LOGON_ZERO_PASSWORD_BUFFER = $80000000;
implementation

uses
  Main;

{$R *.dfm}
function CreateProcessWithLogonW(
lpUsername: PWideChar;
lpDomain: PWideChar;
lpPassword: PWideChar;
dwLogonFlags: DWORD;
lpApplicationName: PWideChar;
lpCommandLine: PWideChar;
dwCreationFlags: DWORD;
lpEnvironment: Pointer;
lpCurrentDirectory: PWideChar;
const lpStartupInfo: TStartupInfo;
var lpProcessInfo: TPROCESSINFORMATION): BOOL; stdcall; external 'advapi32.dll' name 'CreateProcessWithLogonW';

procedure TFPolsRed.Button1Click(Sender: TObject);
var
{StartupInfo : TStartupInfow;
ProcessInfo : TProcessInformation;

S,SS:PWideChar;
  sss,Pass,Department,title,Manager:string;
  SessionID: DWORD;
  UserToken: THandle;
  CmdLine: PChar;
 // Usr: IADsUser;
   Department:=ParamStr(3);
  title:=ParamStr(4);
  Manager:=ParamStr(5);
      Usr.Department:=Department; //Отдел  КТО ОТК ПДО  ....
    Usr.Title:=title;//Должность - Участок - Клапана, КЦКП
    Usr.Manager:=Manager;//Начальник участка}
 Str,Nam,ss,mail:string;
I,y:Integer;
  F_Otk,F_Kud,Fil,Nam_en,Pass,
  Department,//Отдел
  title,//Должность
  Manager//Ночальник
  ,s:string;
  buffer: array[0..MAX_COMPUTERNAME_LENGTH + 1] of Char;
  Size: Cardinal;
  Usr: IADsUser;
begin
  {nam :=ParamStr(1);
  nam_En :=ParamStr(2);
  //Pass:=ParamStr(3);
  Department:=ParamStr(3);
  title:=ParamStr(4);
  Manager:=ParamStr(5);//=1- Начальник
  Memo1.Lines.Add(nam);
  Memo1.Lines.Add(nam_En);
  Memo1.Lines.Add(Department);
  Memo1.Lines.Add(title);
  Memo1.Lines.Add(Manager); }
  if ComboBox1.Text='' then
     Department:='0'
    Else
    Department:= ComboBox1.Text;

    if ComboBox2.Text='' then
     Title:='0'
    Else
    Title:= ComboBox2.Text;
    if ComboBox3.Text='' then
     Manager:='0'
    Else
    Manager:= ComboBox3.Text;
    nam :=FPolsRed.Caption;
  nam_En :=FPolsRed.Label4.Caption;
  try      //CN=Кондратенко С.А.,CN=Users,DC=veza-miass,DC=local sAMAccountName
  //CN=Кондратенко С.А.,OU=Workstations,DC=veza-miass,DC=local
    Str:= 'LDAP://CN='+Nam+',OU=Workstations,DC=veza-miass,DC=local';
    Usr := GetObject(Str) as IADsUser;
    if Manager='0' then
    Usr.PutEx(ADS_PROPERTY_CLEAR,'FaxNumber',0)
    else
    Usr.FaxNumber:=Manager;//Начальник участка
    //
    if Department='0' then
    Usr.PutEx(ADS_PROPERTY_CLEAR,'Department',0)
    else
    Usr.Department:=Department; //Отдел  КТО ОТК ПДО  ....
    //
    if title='0' then
    Usr.PutEx(ADS_PROPERTY_CLEAR,'title',0)
    else
    Usr.Title:=title;//Должность - Участок - Клапана, КЦКП
    //Usr.TelephonePager:=

     Usr.SetInfo; //Сохранить
  except
    on E: EOleException do
    begin
    ShowMessage(E.message);
    end;
  end;

  Close;

   { S := PWideChar(ExtractFileDir(ParamStr(0))+'\Redaktor.exe');
    if ComboBox1.Text='' then
     Department:='0'
    Else
    Department:= ComboBox1.Text;

    if ComboBox2.Text='' then
     Title:='0'
    Else
    Title:= ComboBox2.Text;
    if ComboBox3.Text='' then
     Manager:='0'
    Else
    Manager:= ComboBox3.Text;

    FillChar(StartupInfo, SizeOf(StartupInfo), 0);
    StartupInfo.cb := SizeOf(StartupInfo);

if CreateProcessWithLogonW('ldap_user',
 'veza-miass',
 'vez1ld1p',
LOGON_WITH_PROFILE
, (s),
PWideChar('0 "'+FPolsRed.Caption+'" "'+Label4.Caption+'" "'+Department+'" "'+Title+'" "'+Manager+'"'),
CREATE_DEFAULT_ERROR_MODE, nil, nil,
StartupInfo, ProcessInfo) then
begin
CloseHandle(ProcessInfo.hProcess);
CloseHandle(ProcessInfo.hThread)
end else
Win32Check(False);
 Close;   }
end;

function TFPolsRed.GetObject(const name: string): IDispatch;
var
Moniker: IMoniker;
Eaten: integer;
BindContext: IBindCtx;
Dispatch: IDispatch;
begin
OleCheck(CreateBindCtx(0, BindContext));
OleCheck(MkParseDisplayName(BindContext, PWideChar(WideString(name)), Eaten, Moniker));
OleCheck(Moniker.BindToObject(BindContext, nil, IDispatch, Dispatch));
Result := Dispatch;
end;
Const ADS_PROPERTY_CLEAR = 1 ;
procedure TFPolsRed.Button2Click(Sender: TObject);
var Str,Nam,ss,mail:string;
I,y:Integer;
  F_Otk,F_Kud,Fil,Nam_en,Pass,
  Department,//Отдел
  title,//Должность
  Manager//Ночальник
  ,s:string;
  buffer: array[0..MAX_COMPUTERNAME_LENGTH + 1] of Char;
  Size: Cardinal;
  Usr: IADsUser;
begin
 { nam :=FPolsRed.Caption;
  nam_En :=FPolsRed.Label4.Caption;
  //Pass:=ParamStr(3);
  Department:=ComboBox1.Text;
  title:=ComboBox2.Text;
  Manager:=ComboBox3.Text;//=1- Начальник

  try      //CN=Кондратенко С.А.,CN=Users,DC=veza-miass,DC=local sAMAccountName
    Str:= 'LDAP://CN='+Nam+',CN=Users,DC=veza-miass,DC=local';
    Usr := GetObject(Str) as IADsUser;
    if Manager='' then
    Usr.PutEx(ADS_PROPERTY_CLEAR,'FaxNumber',0)
    else
    Usr.FaxNumber:=Manager;//Начальник участка
    //
    if Department='' then
    Usr.PutEx(ADS_PROPERTY_CLEAR,'Department',0)
    else
    Usr.Department:=Department; //Отдел  КТО ОТК ПДО  ....
    //
    if title='' then
    Usr.PutEx(ADS_PROPERTY_CLEAR,'title',0)
    else
    Usr.Title:=title;//Должность - Участок - Клапана, КЦКП
    //Usr.Description:=Manager;//ОПИСАНИЕ 1=Начальник участка
    Usr.SetInfo;
  except
    on E: EOleException do
    begin
    ShowMessage(E.message);
    end;
  end; }
  Close;
end;

procedure TFPolsRed.FormShow(Sender: TObject);
var I,y:Integer;
S,Serv:string;
begin
        ComboBox3.Clear;
            S := ExtractFileDir(ParamStr(0));
    Memo2.Lines.Clear;
    Memo2.Lines.LoadFromFile(S + '\ConnKlap.ini');
    Serv:= Memo2.Lines[7] ;
    for i :=0 to 1 do
    Begin
    try
      S:=Memo1.Lines.Strings[i];
      ADOQuery1.Close;
      ADOQuery1.ConnectionString := 'Provider=ADsDSOObject;Encrypt Password=False;Data Source='+Serv+';Mode=Read;Bind Flags=0;ADSI Flag=-2147483648';
      ADOQuery1.SQL.Clear;
      ADOQuery1.SQL.Add('select name,CN,MAIL from ');
      ADOQuery1.SQL.Add(#39+'LDAP://'+Serv+#39);
      ADOQuery1.SQL.Add(' WHERE objectcategory = '+#39+'User'+#39+' AND memberof=''CN='+S+',CN=Users,DC=veza-miass,DC=local''' );
      ADOQuery1.SQL.Add(' order by CN');

      ADOQuery1.open;

      //StringGrid1.RowCount:=StringGrid1.RowCount+ADOQuery1.RecordCount;
      for y:=0 to ADOQuery1.RecordCount-1 do
      begin
            ComboBox3.Items.Add(ADOQuery1.FieldByName('CN').AsString);
            ADOQuery1.Next;
      end;
      ComboBox3.Items.Add('Начальник');
      ComboBox3.ItemIndex:=1;
    except on E: Exception do
      showmessage('Ошибка: '+e.Message);
    end;
    End;
end;

end.
