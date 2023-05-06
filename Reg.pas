unit Reg;

interface

uses
        Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls,
        Forms,
        Dialogs, StdCtrls,ActiveX, // используется для COM Moniker stuff...
ActiveDs_TLB, // созданная библиотека типов
ComObj,ShellApi;

type
        TRegist = class(TForm)
                Button1: TButton;
                Button2: TButton;
                ComboBox1: TComboBox;
                Edit1: TEdit;
                Memo1: TMemo;
    Memo2: TMemo;
    Button3: TButton;
    Edit2: TEdit;
    Label1: TLabel;
    Label20: TLabel;
                procedure Button2Click(Sender: TObject);
                procedure FormShow(Sender: TObject);
                procedure Button1Click(Sender: TObject);
                procedure Edit1KeyPress(Sender: TObject; var Key: Char);
                procedure FormClose(Sender: TObject; var Action: TCloseAction);
                function GetObject(const name: string): IDispatch;
    procedure Button3Click(Sender: TObject);
        private
                { Private declarations }
        public
                { Public declarations }
                Doljnost: Integer;
                Famili:String;
        end;

var
        Regist: TRegist;
         const
 LOGON_WITH_PROFILE  = $00000001;
  LOGON_NETCREDENTIALS_ONLY = $00000002;
  LOGON_ZERO_PASSWORD_BUFFER = $80000000;

implementation
uses Main, UKokZap;
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


{$R *.dfm}

procedure TRegist.Button2Click(Sender: TObject);
begin

        Close;
        Form1.Close;

      {          Flag_Razb_Klapana := 0;
        if not Form1.mkQuerySelect(Form1.ADOQuery2, 'Select * from %s Where (ВидЭлемента=' + #39 + 'Сборочные единицы' + #39 + ') And (IdГП=' + #39 + IDGP + #39 + ') ',
        ['Специф']) then
          exit;
        Flag_Razb_Klapana := 0;
        for j := 0 to Form1.ADOQuery2.RecordCount - 1 do  //Проверка на Клапан разбитый на 2
        begin
          Elem := Form1.ADOQuery2.FieldByName('Элемент').AsString;
          Kol_Ed_Str := StringReplace(Form1.ADOQuery2.FieldByName('Количество').AsString, '.', ',', [rfReplaceAll]);
          Kol_Ed := StrToFloat(Kol_Ed_Str);
          Klap := AnsiCompareStr('Клапан', Elem);
          if (Klap = 0) and (Kol_Ed > 1) then
          begin
            Flag_Razb_Klapana := 1; //Клапан разбит на 2
            Izdel := Form1.ADOQuery2.FieldByName('Обозначение').AsString;
            Nam := Form1.ADOQuery2.FieldByName('Обозначение').AsString;
            Break;
          end;
          Form1.ADOQuery2.Next;                                                                    // (Тип=' + #39 + 'Лопатка' + #39 + ')
        end;                 //Лопатка-заготовка
        if not mkQuerySelect(ADOQuery1, 'Select * from %s Where (IdГП=' + #39 + IDGP + #39 +
        ') AND (ВидЭлемента=' + #39 + 'Сборочные единицы' + #39 +
        ') AND (Элемент=' + #39 + 'Лопатка' + #39 + ')', ['Специф']) then
          exit;
        if ADOQuery1.RecordCount<>0 then
        Lop := ADOQuery1.FieldByName('Количество').AsString;

        if not mkQuerySelect(ADOQuery1, 'Select * from %s Where (IdГП=' + #39 + IDGP + #39 +
        ') AND (ВидЭлемента=' + #39 + 'Сборочные единицы' + #39 +
        ') AND (Элемент=' + #39 + 'Лопатка-заготовка' + #39 + ')', ['Специф']) then
          exit;
        if ADOQuery1.RecordCount<>0 then
        Lop := ADOQuery1.FieldByName('Количество').AsString;

        Privod := SG6.Cells[I_FN_BRIKET, I + 1];
        if Privod = '0' then
        begin
          MessageDlg('Необходимо поставить кол-во приводов!', mtError, [mbOk], 0);
          Continue;
        end;
        Privod_Nam := StringGrid10.Cells[I_FN_MOD_PRIVOD, I + 1];
                                        //ADOQuery1.FieldByName('Количество').AsString;
        NC_Sbor := 0;
        Res := Pos(' ', Izdel);
        Delete(Izdel, 1, Res);
                                //======================================== KPU
        Res := Pos('-', Izdel);
        KPu_S := Copy(Izdel, 1, Res - 1);
        Delete(Izdel, 1, Res);
                                //======================================== 1 H
        Res := Pos('-', Izdel);
        N1 := Copy(Izdel, 1, Res - 1);
                                //Pos3 := Copy(Izdel, 2, 1);
        Delete(Izdel, 1, Res);
                                //========================================
                                //Клапан КПУ-1Н-О-Н-600*400-2*ф-MB220-сн-0-0-0-0-0-0
                                //========================================
        Res := Pos('-', Izdel);
        Nazn := Copy(Izdel, 1, Res - 1);
        Delete(Izdel, 1, Res);
                                //========================================
        Res := Pos('-', Izdel);
        Ispol := Copy(Izdel, 1, Res - 1);
        Delete(Izdel, 1, Res);
                                //========================================
                                //Клапан КПУ-1Н-О-Н-100-2*ф-MB220-сн-0-0-0-0-0-0
        Res := Pos('*', Izdel);
        if Res > 5 then
          Continue;
        a := StrToInt(Copy(Izdel, 1, Res - 1));
        Delete(Izdel, 1, Res);
                                //========================================
        Res := Pos('-', Izdel);
        b := StrToInt(Copy(Izdel, 1, Res - 1));
        Delete(Izdel, 1, Res);
                                //========================================
        Res := Pos('-', Izdel);
        F := Copy(Izdel, 1, 1);
        Delete(Izdel, 1, Res);
        if F <> '' then
          Kol_F := StrToInt(F);
                                //========================================
       Term:='';                                                            // T ?
                                //Клапан КПУ-1Н-О-Н-700*250-2*ф-MV220-сн-0-0-0-0-0-0
        Res := Pos('-', Izdel);
        Priv := Copy(Izdel, 1, Res - 1);
        Delete(Izdel, 1, Res);

                                //========================================
        Res := Pos('T', Izdel);
        if Res <> 0 then
        begin
          Term := Copy(Izdel, 1, Res);
          Delete(Izdel, 1, Res + 1);

        end;
                                //========================================
        Res := Pos('-', Izdel);
        Razmesh := Copy(Izdel, 1, Res - 1);
        Delete(Izdel, 1, Res);
                                //======================================== Klem,Resh,luk,Pereh,Ruk,Rama
        Res := Pos('-', Izdel);
        Klem := Copy(Izdel, 1, Res - 1);
        Delete(Izdel, 1, Res);
                                //========================================
        Res := Pos('-', Izdel);
        Resh := Copy(Izdel, 1, Res - 1);
        Delete(Izdel, 1, Res);
                                //========================================
        Res := Pos('-', Izdel);
        luk := Copy(Izdel, 1, Res - 1);
        Delete(Izdel, 1, Res);
                                //========================================
        Res := Pos('-', Izdel);
        Pereh := Copy(Izdel, 1, Res - 1);
        Delete(Izdel, 1, Res);
                                //========================================
                                //Клапан КПУ-1Н-О-Н-700*250-2*ф-MV220-сн-0-0-0-0-0-0
        Res := Pos('-', Izdel);
        Ruk := Copy(Izdel, 1, Res - 1);
        Delete(Izdel, 1, Res);
                                //========================================
                                //Клапан КПУ-1Н-О-Н-200*200-2*ф-MB220-сн-0-р-0-0-0-мрп
        Rama := Izdel;
                                //-----------------------------------------------------

        if Privod = '' then
          Kol_Priv := 0
        else
          Kol_Priv := StrToInt(Privod);
                                //-----------------------------------------------------

        if Lop = '' then
          Kol_lop := 0
        else
          Kol_lop := StrToInt(Lop);
        Memo13.Lines.Add(Lop);
        if Kol_lop = 0 then
          Lopatka := 'Лопатка1';
        if Kol_lop = 1 then
          Lopatka := 'Лопатка1';

        if Kol_lop = 2 then
          Lopatka := 'Лопатка2';

        if Kol_lop = 3 then
          Lopatka := 'Лопатка3';

        if Kol_lop = 4 then
          Lopatka := 'Лопатка4';
        //------------------------------------------------------
        Perim := a + b;
        if Perim <= 600 then
          Perim1 := 600;

        if (Perim > 600) and (Perim <= 850) then
          Perim1 := 850;  }
end;

function TRegist.GetObject(const name: string): IDispatch;
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

procedure TRegist.Button3Click(Sender: TObject);
var
StartupInfo : TStartupInfow;
ProcessInfo : TProcessInformation;

S,SS:PWideChar;
  sss,Pass:string;
  SessionID: DWORD;
  UserToken: THandle;
  CmdLine: PChar;
  Usr: IADsUser;
begin

{try
Usr := GetObject('LDAP://CN=Кондратенко С.А.,CN=Users,DC=veza-miass,DC=local') as IADsUser;
Usr.TelephonePager:='123';//('UserFlags', Usr.Get('UserFlags') or 65536) ;
Usr.SetInfo;
except
on E: EOleException do
begin
ShowMessage(E.message);
end;
end;  }

                   // 'D:\Source\Копии\Kopy.exe'
S := PWideChar(ExtractFileDir(ParamStr(0))+'\Kopy.exe');
Pass:= Edit2.Text;
Form1.Pass:=Pass;
FillChar(StartupInfo, SizeOf(StartupInfo), 0);
StartupInfo.cb := SizeOf(StartupInfo);  //gapova-ai  7777777

if CreateProcessWithLogonW('ldap_user',
 'veza-miass',
 'vez1ld1p',
LOGON_WITH_PROFILE
, (s),
PWideChar('0 "'+Form1.UserName1+'" "'+Form1.UserName_Eng+'" "'+Pass+'"'),
CREATE_DEFAULT_ERROR_MODE, nil, nil,
StartupInfo, ProcessInfo) then
begin
CloseHandle(ProcessInfo.hProcess);
CloseHandle(ProcessInfo.hThread)
end else
Win32Check(False);
        Edit2.Visible:=False;
        label1.Visible:=False;
        Button1.Enabled:=True;
        Button3.Visible:=False;
        Edit1.SetFocus;
Button1.Click;
end;

procedure TRegist.FormShow(Sender: TObject);
var
        i,res1,Grupp_I: Integer;
        S,Nam,ss,Grupp,name,Serv:String;
            buffer: array[0..MAX_COMPUTERNAME_LENGTH + 1] of Char;
  Size: Cardinal;
begin
        Grupp_I:=0;
        Edit2.Visible:=False;
        label1.Visible:=False;
        Button1.Enabled:=True;
        Button3.Visible:=False;
        S := ExtractFileDir(ParamStr(0));
        Memo2.Lines.Clear;
        Memo2.Lines.LoadFromFile(S + '\ConnKlap.ini');
        Form1.ADOConnection1.Connected:=False;
        Form1.ADOConnection1.ConnectionString:=Memo2.Lines[4] ;
        Serv:= Memo2.Lines[7] ;
        ComboBox1.DropDownCount := 30;
        ComboBox1.Style := csDropDownList;
        {Form1.ADOQuery1.Close;
        Form1.ADOQuery1.SQL.Clear;
        Form1.ADOQuery1.SQL.Text :=
                'Select * from Сотрудники Order By Сотрудник';
        Form1.ADOQuery1.Open;
        Form1.ADOQuery1.First;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
                ComboBox1.Items.Add(Form1.ADOQuery1.FieldByName('Сотрудник').AsString);
                Form1.ADOQuery1.Next;
        end;
        Edit2.SetFocus; }
        Memo1.Lines.Clear;
        Memo1.Lines.LoadFromFile('Avto.ini');
        if Memo1.Lines[0] = '' then
                ComboBox1.ItemIndex := 0
        else
                ComboBox1.ItemIndex := StrToInt(Memo1.Lines[0]);
  Size := MAX_COMPUTERNAME_LENGTH + 1;
  Windows.GetUserName(@buffer, Size);
  nam := StrPas(buffer);
  Form1.UserName_Eng:=Nam;
        Form1.ADOQuery1.Close;
      Form1.ADOQuery1.ConnectionString := 'Provider=ADsDSOObject;Encrypt Password=False;Data Source='+Serv+';Mode=Read;Bind Flags=0;ADSI Flag=-2147483648';
      Form1.ADOQuery1.SQL.Clear;
      Form1.ADOQuery1.SQL.Add('select name,CN,MAIL,Pager,Department,Title,facsimileTelephoneNumber from ');
      Form1.ADOQuery1.SQL.Add(#39+'LDAP://'+Serv+#39);
      Form1.ADOQuery1.SQL.Add(' WHERE objectcategory = '+#39+'User'+#39+' AND sAMAccountName  ='+#39+nam+#39 );
      SS:=Form1.ADOQuery1.SQL.Text;
      Form1.ADOQuery1.open;

      if Form1.ADOQuery1.RecordCount=0 then
      Begin
           MessageDlg('Не нашел пользователя!'+#10#13+Nam, mtError, [mbOk], 0);
      end
      else
      Begin
           Form1.UserName1:=Form1.ADOQuery1.FieldByName('name').AsString;
           Form1.Departament:=Form1.ADOQuery1.FieldByName('Department').AsString; //Отдел
           Form1.Title:=Form1.ADOQuery1.FieldByName('Title').AsString; //Участок
           Form1.Division:=Form1.ADOQuery1.FieldByName('facsimileTelephoneNumber').AsString; //Н- Начальник
           Form1.mail:=Form1.ADOQuery1.FieldByName('MAIL').AsString;
           if Form1.mail='' then
           MessageDlg('К вашей учетной записи не привязан адрес электоронной почты!'+#10#13+
           'Вы не сможете получать важные сообщения!'+#10#13+
           'Обратитесь к системному администратору!', mtError, [mbOk], 0);

           Form1.Pass:=Form1.ADOQuery1.FieldByName('Pager').AsString;
           if Form1.Pass='' then
           begin
              //
              Edit2.Visible:=True;
              label1.Visible:=True;
              Button1.Enabled:=False;
              Button3.Visible:=True;
              MessageDlg('К вашей учетной записи привязан адрес электоронной почты!'+#10#13+
              Form1.mail+#10#13+
              'Введите пожалуйста пароль вашей электронной почты!', mtError, [mbOk], 0);
              Edit2.SetFocus;
             // Button3.Click;
              //FNewBRAK.UserName1:= Form1.UserName1;

           end;
      End;
     Form1.UserName1:=Form1.ADOQuery1.FieldByName('name').AsString;
     res1:=Pos('Кондратенко', Form1.UserName1);
     if res1<>0 then
     Form1.tmr2.Enabled:=False;
     Form1.mail:=Form1.ADOQuery1.FieldByName('MAIL').AsString;
     Form1.UserName_Eng:=Nam;
     Famili:=Form1.ADOQuery1.FieldByName('name').AsString;

{Grupp_I
ПДО           0   7
КТО           1   2
Производство  2   4
ОТК           3   10
ОИТ           4   1}
Name:= Form1.UserName1;
      for I:=0 to Form1.ADOQuery1.RecordCount-1 do
      begin
          Grupp:= Form1.ADOQuery1.FieldByName('Department').AsString;
          res1:=Pos('ПДО',Grupp);
          if Res1<>0 then
          Begin
             Grupp_I:=7;
             Label20.Caption:= Name+'('+Grupp+')';
          End;
          //
          res1:=Pos('КТО',Grupp);
          if Res1<>0 then
          Begin
             Grupp_I:=2;
             Label20.Caption:= Name+'('+Grupp+')';
          End;
          //
         res1:=Pos('Производство',Grupp);
          if Res1<>0 then
          Begin
             Grupp_I:=4;
             Label20.Caption:= Name+'('+Grupp+')';
          End;
          //
         res1:=Pos('ОТК',Grupp);
          if Res1<>0 then
          Begin
             Grupp_I:=10;
             Label20.Caption:= Name+'('+Grupp+')';
          End;
          //
          res1:=Pos('ОИТ',Grupp);
          if Res1<>0 then
          Begin
             Grupp_I:=1;
             Label20.Caption:= Name+'('+Grupp+')';
          End;
          //
          Form1.ADOQuery1.Next;
      end;
      Form1.FlagDolg :=Grupp_I ;
      Doljnost :=Grupp_I ;
        if Regist.Button1.Enabled then
        begin

          Regist.Button1.Click;
        end;

     // Button1.Click;
end;
{Grupp_I
ПДО          0  7
КТО           1  2
Производство  2   4
ОТК           3   10
ОИТ           4   1}
procedure TRegist.Button1Click(Sender: TObject);
var
        Res, Dol: Integer;
        Pass1: string;
begin
      {  if not Form1.mkQuerySelect(Form1.ADOQuery1,'Select * from Сотрудники Where Сотрудник = '
                + #39+ ComboBox1.Text + #39, []) then
      exit;
        Pass1 := Form1.ADOQuery1.FieldByName('Пароль').AsString;
        Dol := Form1.ADOQuery1.FieldByName('Должность').AsInteger;
        Res := AnsiCompareStr(Pass1, Edit1.Text);
        if Res = 0 then
        begin
                Form1.FlagDolg := Dol;
                Doljnost := Form1.ADOQuery1.FieldByName('Должность').AsInteger;
               if (Doljnost = 1)  then
                begin
                        Form1.tmr2.Enabled := False; //Почта
                End;  }
                if (Doljnost = 1) OR (Doljnost = 2)  then
                begin
               //         Form1.Button56.Visible := True;
                End;
                if (Doljnost = 1) OR (Doljnost = 2)  then
                begin
                        Form1.Del.Visible := True;
                End;
                if (Doljnost = 1) OR (Doljnost = 11)  then
                begin
                        Form1.Button17.Visible := True;
                        Form1.PopupKPU_L.Items.Items[11].Visible:=True;
                        Form1.PopupVOZD_L.Items.Items[6].Visible:=True;
                End;
                if (Doljnost = 1) OR (Doljnost = 2) then
                begin
                        Form1.MenuItem1.Visible := True;
                        //Делете для воздуш клапанов
                        Form1.Del.Visible := True;
                        Form1.MenuItem1.Visible := True;
                        Form1.N5.Visible := True;
                        Form1.Button18.Visible := True;
                        Form1.Button14.Visible := True;

                end
                else
                begin
                        Form1.MenuItem1.Visible := False;
                        //Делете для воздуш клапанов
                        Form1.Del.Visible := False;
                        Form1.MenuItem1.Visible := False;
                        Form1.N5.Visible := False;
                end;
                if (Doljnost = 1) or (Doljnost = 2) or (Doljnost = 3) or
                        (Doljnost = 0) then
                begin
                       // Form1.Button2.Visible := True;
                        Form1.N62.Visible:=True;
                        Form1.N63.Visible:=True;
                        Form1.N64.Visible:=True;
                        Form1.N65.Visible:=True;
                end;
                if (Doljnost = 1) or (Doljnost = 7) or (Doljnost = 8) OR (Doljnost = 4) then
                begin
                        Form1.Button18.Visible := True;
                        Form1.Button14.Visible := True;
                        //Form1.Button2.Visible := True;
                        Form1.Button42.Visible := True;
                end;
                if (Doljnost = 1) or  (Doljnost = 8) OR (Doljnost = 4) OR (Doljnost = 7) then
                begin
                        Form1.Button41.Visible := True;
                        Form1.DateTimePicker3.Visible := True;
                        Form1.Label67.Visible := True;
                end;
                if (Doljnost = 1) or (Doljnost = 10) OR (Doljnost = 4) then //OTK
                begin
                        FKolZap.DateTimePicker1.Visible := True;
                        FKolZap.Button1.Visible := True;
                        FKolZap.Label7.Visible := True;
                end;
                if (Doljnost = 1) OR (Doljnost = 4) then //Trapez
                begin
                        //Form1.btn4.Visible := True;
                        Form1.N9.Visible := True;
                       // Form1.N15.Visible := True;
                end;
                if (Doljnost = 1) OR (Doljnost = 10) OR (Doljnost = 4) then //
                begin
                        //Form1.MenuItem1.Visible := True;
                        //Form1.N9.Visible := True;
                         Form1.N15.Visible := True;
                end;
                Form1.check('Регистрация ', ComboBox1.Text + '  ' + Pass1, 1);
               // Famili:=ComboBox1.Text;
               PostMessage(Handle, WM_CLOSE, 0, 0);
               Regist.ModalResult:=100;
               Regist.Close;
      {  end
        else
        begin
                MessageDlg('Не верный пароль!', mtError, [mbOk], 0);
                Form1.check('Регистрация ', ComboBox1.Text + '  ' + Edit1.Text,
                        -1);
                Edit1.Text := '';
        end;  }

end;

procedure TRegist.Edit1KeyPress(Sender: TObject; var Key: Char);
begin
        if (Key = #13) and (Button1.Enabled) then
                Button1Click(nil);
end;

procedure TRegist.FormClose(Sender: TObject; var Action: TCloseAction);
var S,FileName_Server,FileName:string;
 FileDate_Server ,FileDate:Integer;

begin
        Memo1.Lines.Clear;
        Memo1.Lines.Add(IntToStr(ComboBox1.ItemIndex));
        Memo1.Lines.SaveToFile('Avto.ini');
        Form1.Caption := Form1.Memo1.Lines.Strings[0] + '  001'+ '   '+Regist.Famili ;
        //--------------------------------------------------------------------
        fileName := ExtractFileDir(ParamStr(0))+'\Klapan.EXE';
        fileDate := FileAge(fileName);
        S:= DateToStr(FileDateToDateTime(fileDate)) ;
        //--------------------
        FileName_Server :=Form1.Put_Avto+'\CKlapana2\Klapan.exe';
        fileDate_Server := FileAge(fileName_server);
        if (fileDate<>fileDate_Server) and (Doljnost<>1 ) then
        begin
                MessageDlg('Внимание! Версии файлов не совпадают! Вызовите Кондратенко, тел. 115.', mtError,
                        [mbOk], 0);
                        Close;
                        Form1.Close;
        end;
        //
end;

end.
