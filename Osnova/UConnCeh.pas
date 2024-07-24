unit UConnCeh;

interface
    uses SysUtils, Classes, Variants,
       ADODB, Controls,Dialogs;
    Type
    Connect_Miass_Ceh = class
    public
    AConnect : TADOConnection;
    AQuery: TADOQuery;
    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    AConnect_Del : TADOConnection;
    AQuery_Del: TADOQuery;
    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    procedure Check(Module: string; Str: variant; Col: Integer);

    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    function SetConnect(var AQuery: TADOQuery;
        var AConnect: TAdoConnection): Boolean;
    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    function SetConnect_Del(var AQuery_Del: TADOQuery;
        var AConnect_Del: TAdoConnection): Boolean;

    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    function mkQuerySelect( const sQ: WideString; Args:
        array of const): boolean;
    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    function mkQuerySelect2( const sQ: WideString; Args:
        array of const): boolean;
    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    function mkQueryInsert( const sQ: WideString; Args:
        array of const): boolean;
    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    function mkQueryDelete( const sQ: WideString; Args:
        array of const): boolean;
    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    function mkQueryUpdate( const sQ: WideString; Args:
        array of const): boolean;
    end;

implementation

uses   UConnKlap;
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++ Check
procedure Connect_Miass_Ceh.Check(Module: string; Str: variant;
  Col: Integer);
var
        myFile: TextFile;
        S, Str1,Str2: string;
begin
        S := ExtractFileDir(ParamStr(0)) + '\Log.ini';
        AssignFile(myFile, S);
        Str2:=FormatDateTime('dd.mm.YYYY hh:mm:ss',Now) ;
        Str1 := str2 +' : '+Module + '  ' + Str + '  ' + IntToStr(Col);

        try
                Append(myFile);
                WriteLn(myFile, Str1);
                CloseFile(myFile);
        except
                ReWrite(myFile);
        end;

end;

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++Коннект MIASS_CEH
function Connect_Miass_Ceh.SetConnect(var AQuery: TADOQuery;
  var AConnect: TAdoConnection): Boolean;
var
        MSSQL_CONN_STR,DEF_CatalogName, S: string;
        Memo9:TStringList;
begin
        Result := false;
        S := ExtractFileDir(ParamStr(0));
        Memo9:=TStringList.Create;
        Memo9.Clear;
        Memo9.LoadFromFile(S + '\ConnCeh.ini');
        DEF_CatalogName := Memo9.Strings[0];
        DEF_CatalogName := DEF_CatalogName;
        MSSQL_CONN_STR :=  Memo9.Strings[4];
               { 'Provider=SQLOLEDB.1;Packet Size = 4096;Password=111;' +
                'Persist Security Info=True;User Id=testuser;' +
                'Initial Catalog=' + DEF_CatalogName +
                ';Data Source=DINAMO\SQLEXPRESS';   }

        if AConnect = nil then
                AConnect := TADOConnection.Create(Nil);

        if AQuery = nil then
                AQuery := TADOQuery.Create(Nil);

        AConnect.LoginPrompt := False;
        AConnect.Connected := False;
        AConnect.ConnectionString := MSSQL_CONN_STR;
        try
                AConnect.Connected := True;
                AQuery.Connection:=AConnect;
        except
                MessageDlg('Не удалось подключиться к базе данных!', mtError,
                        [mbOk], 0);
                Exit;
        end;
        Result := true;
end;
//+++++++++++++++++++++++++++++++++++++++++++++++++коннект для ins, del, update
function Connect_Miass_Ceh.SetConnect_Del(var AQuery_Del: TADOQuery;
  var AConnect_Del: TAdoConnection): Boolean;
var
        MSSQL_CONN_STR,DEF_CatalogName, S: string; //DEF_CatalogName ,
                Memo9:TStringList;
begin
        Result := false;
        S := ExtractFileDir(ParamStr(0));
        Memo9:=TStringList.Create;
        Memo9.Clear;
        Memo9.LoadFromFile(S + '\ConnCeh.ini');
        DEF_CatalogName := Memo9.Strings[0];
        DEF_CatalogName := DEF_CatalogName;

        MSSQL_CONN_STR :=Memo9.Strings[4];
               { 'Provider=SQLOLEDB.1;Packet Size = 4096;Password=111;' +
                'Persist Security Info=True;User Id=testuser;' +
                'Initial Catalog=' + DEF_CatalogName +
                ';Data Source=DINAMO\SQLEXPRESS'; }

        if AConnect_Del = nil then
                AConnect_Del := TADOConnection.Create(Nil);

        if AQuery_Del = nil then
                AQuery_Del := TADOQuery.Create(Nil);

        AConnect_Del.LoginPrompt := False;
        AConnect_Del.Connected := False;
        AConnect_Del.ConnectionString := MSSQL_CONN_STR;
        try
                AConnect_Del.Connected := True;
                AQuery_Del.Connection:=AConnect_Del;
        except
                MessageDlg('Не удалось подключиться к базе данных!', mtError,
                        [mbOk], 0);
                Exit;
        end;
        Result := true;
end;
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Select
function Connect_Miass_Ceh.mkQuerySelect(
  const sQ: WideString; Args: array of const): boolean;

begin
        Result := false;
        if not  SetConnect(AQuery,AConnect) then
                exit;
        try
                AQuery.Close;
                AQuery.SQL.Clear;
                AQuery.SQL.Text := Format(sQ, Args);
                AQuery.Active := true;
                Check('query', AQuery.SQL.Text, AQuery.RecordCount);
        except
                Check('Error in Query', Format(sQ, Args), 0);
                AQuery.Close;
                exit;
        end;
        Result := true;
end;
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Select
function Connect_Miass_Ceh.mkQuerySelect2(
  const sQ: WideString; Args: array of const): boolean;

begin
        Result := false;
        if not  SetConnect_Del(AQuery_Del,AConnect_Del) then
                exit;
        try
                AQuery_Del.Close;
                AQuery_Del.SQL.Clear;
                AQuery_Del.SQL.Text := Format(sQ, Args);
                AQuery_Del.Active := true;
                Check('query', AQuery_Del.SQL.Text, AQuery_Del.RecordCount);
        except
                Check('Error in Query', Format(sQ, Args), 0);
                AQuery_Del.Close;
                exit;
        end;
        Result := true;
end;
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Delete
function Connect_Miass_Ceh.mkQueryDelete(const sQ: WideString;
  Args: array of const): boolean;
var
        Str: string;
begin
        Result := false;
        if not SetConnect_Del(AQuery_Del,AConnect_Del) then
                exit;
        try
                AQuery_Del.Close;
                AQuery_Del.SQL.Clear;
                Str := Format(sQ, Args);
                AQuery_Del.SQL.Text := Format(sQ, Args);
                AQuery_Del.ExecSQL;
                check('query', AQuery_Del.SQL.Text, 1);
        except
                Check('Error in Query', Format(sQ, Args), 0);
                AQuery_Del.Close;
                MessageDlg('Не удалось Удалить запись!', mtError, [mbOk], 0);
                exit;
        end;
        AQuery_Del.Close;
        Result := true;
end;
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++Insert
function Connect_Miass_Ceh.mkQueryInsert(const sQ: WideString;
  Args: array of const): boolean;
var
        Str: string;
begin
        Result := false;
        if not SetConnect_Del(AQuery_Del,AConnect_Del) then
                exit;
        try
                AQuery_Del.Close;
                AQuery_Del.SQL.Clear;
                Str := Format(sQ, Args);
                AQuery_Del.SQL.Text := Format(sQ, Args);
                AQuery_Del.ExecSQL;
                check('query', AQuery_Del.SQL.Text, 1);
        except
                Check('Error in Query', Format(sQ, Args), 0);
                AQuery_Del.Close;
                MessageDlg('Не удалось вставить запись!', mtError, [mbOk], 0);
                exit;
        end;
        AQuery_Del.Close;
        Result := true;
end;
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Updat
function Connect_Miass_Ceh.mkQueryUpdate(const sQ: WideString;
  Args: array of const): boolean;
var
        Str: string;
begin
        Result := false;
        if not SetConnect_Del(AQuery_Del,AConnect_Del) then
                exit;
        try
                AQuery_Del.Close;
                AQuery_Del.SQL.Clear;
                Str := Format(sQ, Args);
                AQuery_Del.SQL.Text := Format(sQ, Args);
                AQuery_Del.ExecSQL;
                check('query', AQuery_Del.SQL.Text, 1);
        except
                Check('Error in Query', Format(sQ, Args), 0);
                AQuery_Del.Close;
                MessageDlg('Не удалось Обновить запись!', mtError, [mbOk], 0);
                exit;
        end;
        AQuery_Del.Close;
        Result := true;
end;
end.
