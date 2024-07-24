unit UConnKlap;
interface
    uses SysUtils, Classes, Variants,
       ADODB, Controls,Dialogs;
    Type
    Connect_Miass_Klap = class
    public
    DEF_CatalogName1:string;
    DEF_CatalogName3:string;
    AConnect : TADOConnection;
    AQuery: TADOQuery;
    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    AConnect_Del : TADOConnection;
    AQuery_Del: TADOQuery;
    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    AConnect3 : TADOConnection;
    AQuery3: TADOQuery;
    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    AConnect4 : TADOConnection;
    AQuery4: TADOQuery;
    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    AConnect5 : TADOConnection;
    AQuery5: TADOQuery;
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
    //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Select2
    function mkQuerySelect2(
    const sQ: WideString; Args: array of const): boolean;
    //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Select3
        function mkQuerySelect3(
        const sQ: WideString; Args: array of const): boolean;
    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    function SetConnect3(var AQuery3: TADOQuery;
        var AConnect3: TAdoConnection): Boolean;
    //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Select4
        function mkQuerySelect4(
        const sQ: WideString; Args: array of const): boolean;
    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    function SetConnect4(var AQuery4: TADOQuery;
        var AConnect4: TAdoConnection): Boolean;
    //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Select5
        function mkQuerySelect5(
        const sQ: WideString; Args: array of const): boolean;
    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    function SetConnect5(var AQuery5: TADOQuery;
        var AConnect5: TAdoConnection): Boolean;
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

uses
  Main, UConnCeh;


//+++++++++++++++++++++++++++++++++++++++++++++++++++++++ Check
procedure Connect_Miass_Klap.Check(Module: string; Str: variant;
  Col: Integer);
var
        myFile: TextFile;
        S, Str1,Str2: string;
        Memo2 : TStringList;
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
        //Memo2:=TStringList.Create;
        //Form1.Memo2.Clear;
        Form1.Memo2.Lines.Add(Module + '  ' + Str + '  ' + IntToStr(Col));
end;

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++Коннект MIASS_CEH
function Connect_Miass_Klap.SetConnect(var AQuery: TADOQuery;
  var AConnect: TAdoConnection): Boolean;
var
        MSSQL_CONN_STR, S,Text,SSS: string;
        Memo9:TStringList;
begin
        Result := false;
        S := ExtractFileDir(ParamStr(0));
        Memo9:=TStringList.Create;
        Memo9.Clear;
        Memo9.LoadFromFile(S + '\ConnKlap.ini');
        //Text:=Memo9.Text;
        SSS  := Memo9.Strings[0];
        //DEF_CatalogName3 := SSS;
        DEF_CatalogName1 := SSS;
        MSSQL_CONN_STR := Memo9.Strings[4];
               { 'Provider=SQLOLEDB.1;Packet Size = 4096;Password=111;' +
                'Persist Security Info=True;User Id=testuser;' +
                'Initial Catalog=' + DEF_CatalogName1 +
                ';Data Source=DINAMO\SQLEXPRESS'; }

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
function Connect_Miass_Klap.SetConnect_Del(var AQuery_Del: TADOQuery;
  var AConnect_Del: TAdoConnection): Boolean;
var
        MSSQL_CONN_STR, S: string; //DEF_CatalogName ,
        Memo9:TStringList;
begin
        Result := false;
        S := ExtractFileDir(ParamStr(0));
        Memo9:=TStringList.Create;
        Memo9.Clear;
        Memo9.LoadFromFile(S + '\ConnKlap.ini');
        DEF_CatalogName1 := Memo9.Strings[0];

        MSSQL_CONN_STR := Memo9.Strings[4];
                {'Provider=SQLOLEDB.1;Packet Size = 4096;Password=111;' +
                'Persist Security Info=True;User Id=testuser;' +
                'Initial Catalog=' + DEF_CatalogName1 +
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
function Connect_Miass_Klap.mkQuerySelect(
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
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Select2
function Connect_Miass_Klap.mkQuerySelect2(
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
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Select3
function Connect_Miass_Klap.mkQuerySelect3(
  const sQ: WideString; Args: array of const): boolean;

begin
        Result := false;
        if not  SetConnect3(AQuery3,AConnect3) then
                exit;
        try
                AQuery3.Close;
                AQuery3.SQL.Clear;
                AQuery3.SQL.Text := Format(sQ, Args);
                AQuery3.Active := true;
                Check('query', AQuery3.SQL.Text, AQuery3.RecordCount);
        except
                Check('Error in Query', Format(sQ, Args), 0);
                AQuery3.Close;
                exit;
        end;
        Result := true;
end;
//+++++++++++++++++++++++++++++++++++++++++++++++++коннект Select3
function Connect_Miass_Klap.SetConnect3(var AQuery3: TADOQuery;
  var AConnect3: TAdoConnection): Boolean;
var
        MSSQL_CONN_STR, S: string; //DEF_CatalogName ,
        Memo9:TStringList;
begin
        Result := false;
        S := ExtractFileDir(ParamStr(0));
        Memo9:=TStringList.Create;
        Memo9.Clear;
        Memo9.LoadFromFile(S + '\ConnKlap.ini');
        DEF_CatalogName1 := Memo9.Strings[0];

        MSSQL_CONN_STR := Memo9.Strings[4];
                {'Provider=SQLOLEDB.1;Packet Size = 4096;Password=111;' +
                'Persist Security Info=True;User Id=testuser;' +
                'Initial Catalog=' + DEF_CatalogName1 +
                ';Data Source=DINAMO\SQLEXPRESS'; }

        if AConnect3 = nil then
                AConnect3 := TADOConnection.Create(Nil);

        if AQuery3 = nil then
                AQuery3 := TADOQuery.Create(Nil);

        AConnect3.LoginPrompt := False;
        AConnect3.Connected := False;
        AConnect3.ConnectionString := MSSQL_CONN_STR;
        try
                AConnect3.Connected := True;
                AQuery3.Connection:=AConnect3;
        except
                MessageDlg('Не удалось подключиться к базе данных!', mtError,
                        [mbOk], 0);
                Exit;
        end;
        Result := true;
end;
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Select4
function Connect_Miass_Klap.mkQuerySelect4(
  const sQ: WideString; Args: array of const): boolean;

begin
        Result := false;
        if not  SetConnect4(AQuery4,AConnect4) then
                exit;
        try
                AQuery4.Close;
                AQuery4.SQL.Clear;
                AQuery4.SQL.Text := Format(sQ, Args);
                AQuery4.Active := true;
                Check('query', AQuery4.SQL.Text, AQuery4.RecordCount);
        except
                Check('Error in Query', Format(sQ, Args), 0);
                AQuery4.Close;
                exit;
        end;
        Result := true;
end;
//+++++++++++++++++++++++++++++++++++++++++++++++++коннект Select3
function Connect_Miass_Klap.SetConnect4(var AQuery4: TADOQuery;
  var AConnect4: TAdoConnection): Boolean;
var
        MSSQL_CONN_STR, S: string; //DEF_CatalogName ,
        Memo9:TStringList;
begin
        Result := false;
        S := ExtractFileDir(ParamStr(0));
        Memo9:=TStringList.Create;
        Memo9.Clear;
        Memo9.LoadFromFile(S + '\ConnKlap.ini');
        DEF_CatalogName1 := Memo9.Strings[0];

        MSSQL_CONN_STR :=Memo9.Strings[4];
                {'Provider=SQLOLEDB.1;Packet Size = 4096;Password=111;' +
                'Persist Security Info=True;User Id=testuser;' +
                'Initial Catalog=' + DEF_CatalogName1 +
                ';Data Source=DINAMO\SQLEXPRESS';  }

        if AConnect4 = nil then
                AConnect4 := TADOConnection.Create(Nil);

        if AQuery4 = nil then
                AQuery4 := TADOQuery.Create(Nil);

        AConnect4.LoginPrompt := False;
        AConnect4.Connected := False;
        AConnect4.ConnectionString := MSSQL_CONN_STR;
        try
                AConnect4.Connected := True;
                AQuery4.Connection:=AConnect4;
        except
                MessageDlg('Не удалось подключиться к базе данных!', mtError,
                        [mbOk], 0);
                Exit;
        end;
        Result := true;
end;
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Select5
function Connect_Miass_Klap.mkQuerySelect5(
  const sQ: WideString; Args: array of const): boolean;

begin
        Result := false;
        if not  SetConnect5(AQuery5,AConnect5) then
                exit;
        try
                AQuery5.Close;
                AQuery5.SQL.Clear;
                AQuery5.SQL.Text := Format(sQ, Args);
                AQuery5.Active := true;
                Check('query', AQuery5.SQL.Text, AQuery5.RecordCount);
        except
                Check('Error in Query5', Format(sQ, Args), 0);
                AQuery5.Close;
                exit;
        end;
        Result := true;
end;
//+++++++++++++++++++++++++++++++++++++++++++++++++коннект Select5
function Connect_Miass_Klap.SetConnect5(var AQuery5: TADOQuery;
  var AConnect5: TAdoConnection): Boolean;
var
        MSSQL_CONN_STR, S: string; //DEF_CatalogName ,
        Memo9:TStringList;
begin
        Result := false;
        S := ExtractFileDir(ParamStr(0));
        Memo9:=TStringList.Create;
        Memo9.Clear;
        Memo9.LoadFromFile(S + '\ConnKlap.ini');
        DEF_CatalogName1 := Memo9.Strings[0];

        MSSQL_CONN_STR :=  Memo9.Strings[4];
                {'Provider=SQLOLEDB.1;Packet Size = 4096;Password=111;' +
                'Persist Security Info=True;User Id=testuser;' +
                'Initial Catalog=' + DEF_CatalogName1 +
                ';Data Source=DINAMO\SQLEXPRESS'; }

        if AConnect5 = nil then
                AConnect5 := TADOConnection.Create(Nil);

        if AQuery5 = nil then
                AQuery5 := TADOQuery.Create(Nil);

        AConnect5.LoginPrompt := False;
        AConnect5.Connected := False;
        AConnect5.ConnectionString := MSSQL_CONN_STR;
        try
                AConnect5.Connected := True;
                AQuery5.Connection:=AConnect5;
        except
                MessageDlg('Не удалось подключиться к базе данных!', mtError,
                        [mbOk], 0);
                Exit;
        end;
        Result := true;
end;
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ Delete
function Connect_Miass_Klap.mkQueryDelete(const sQ: WideString;
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
function Connect_Miass_Klap.mkQueryInsert(const sQ: WideString;
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
function Connect_Miass_Klap.mkQueryUpdate(const sQ: WideString;
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
