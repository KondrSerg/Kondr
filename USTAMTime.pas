unit USTAMTime;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls,
  UConnCeh, UConnKlap,UOsnova;

type
  TFStamTime = class(TForm)
    dtp1: TDateTimePicker;
    btn1: TButton;
    btn2: TButton;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    procedure btn1Click(Sender: TObject);
    procedure btn2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    UOsnova_Main:Osnova_Main ;
  end;

var
  FStamTime: TFStamTime;

implementation

uses Main;

{$R *.dfm}

procedure TFStamTime.btn1Click(Sender: TObject);
begin
        Close;
end;

procedure TFStamTime.btn2Click(Sender: TObject);
var
str,Str1,Vn_Dat,fileName,Nom:string;
Res,Res1,Res2:Integer;
begin
        Str := FormatDateTime('mm.dd.yyyy', DTP1.Date);
        Str1 := FormatDateTime('dd.mm.yyyy', DTP1.Date);
        Res:=AnsiCompareStr(FStamTime.Caption,'Дата запуска');
        Res1:=AnsiCompareStr(FStamTime.Caption,'TRUMPF');
  if Form1.Luk=0 then
  begin
        If  (Res1=0) Then
        Begin
                {UOsnova_Main.Flag_Error :=0;
                Vn_Dat := FormatDateTime('dd.mm.yyyy', DTP1.Date);
                fileName := ExtractFileDir(ParamStr(0));//+'\Klapan.EXE';
                if not UOsnova_Main.Osnova2( Vn_Dat,'ЗапускСТАМ','СпецифСТАМ',#39+Label2.Caption+#39,fileName,1 )
                then
                Exit;   }
               if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 FStamTime.Caption + ']=' + #39 + Str + #39 +
                // ',[TRUMPF]=' + #39 + Str + #39 +
                ' WHERE ([IdГП]=' + #39 + Label1.Caption + #39 + ') AND (Номер=' +
                 #39 + Label2.Caption + #39 + ') ', ['ЗапускСТАМ']) then
                Exit;
                Form1.SG215.Cells[StrToInt(Label3.Caption),StrToInt(Label4.Caption)]:=Str1;
        End;
        Res:=AnsiCompareStr(FStamTime.Caption,'СВАРКА');
        If Res=0 Then
        begin
          if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 FStamTime.Caption + ']=' + #39 + Str + #39 +
                ' WHERE ([IdГП]=' + #39 + Label1.Caption + #39 + ') AND (Номер=' +
                 #39 + Label2.Caption + #39 + ') ', ['ЗапускСТАМ']) then
                Exit;
                Form1.SG215.Cells[StrToInt(Label3.Caption),StrToInt(Label4.Caption)]:=Str1;
        end;
        Res:=AnsiCompareStr(FStamTime.Caption,'ПОКРАСКА');
        If Res=0 Then
        begin
          if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 FStamTime.Caption + ']=' + #39 + Str + #39 +
                ',[Заготовка Готовность]=' + #39 + Str + #39 +
                ' WHERE ([IdГП]=' + #39 + Label1.Caption + #39 + ') AND (Номер=' +
                 #39 + Label2.Caption + #39 + ') ', ['ЗапускСТАМ']) then
                Exit;
                Form1.SG215.Cells[StrToInt(Label3.Caption),StrToInt(Label4.Caption)]:=Str1;
        end;
        Res:=AnsiCompareStr(FStamTime.Caption,'ГИБКА');
        if Res=0 then
        Begin
          if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 FStamTime.Caption+ ']=' + #39 + Str + #39 +
                ' WHERE ([IdГП]=' + #39 + Label1.Caption + #39 + ') AND (Номер=' +
                 #39 + Label2.Caption + #39 + ') ', ['ЗапускСТАМ']) then
                Exit;
                Form1.SG215.Cells[StrToInt(Label3.Caption),StrToInt(Label4.Caption)]:=Str1;
        End;

        Res1:=AnsiCompareStr(Label6.Caption,'1,5');
        if (Res=0) and (Res1=0) then
        begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 FStamTime.Caption + ']=' + #39 + Str + #39 +
                ',[СВАРКА]=' + #39 + Str + #39 +
                ',[ПОКРАСКА]=' + #39 + Str + #39 +
                ',[Заготовка Готовность]=' + #39 + Str + #39 +
                ' WHERE ([IdГП]=' + #39 + Label1.Caption + #39 + ') AND (Номер=' +
                 #39 + Label2.Caption + #39 + ') ', ['ЗапускСТАМ']) then
                Exit;
                Form1.SG215.Cells[StrToInt(Label3.Caption),StrToInt(Label4.Caption)]:=Str1;
                Form1.SG215.Cells[14,StrToInt(Label4.Caption)]:=Str1;
                Form1.SG215.Cells[15,StrToInt(Label4.Caption)]:=Str1;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 FStamTime.Caption+ ']=' + #39 + Str + #39 +
                ',[СВАРКА]=' + #39 + Str + #39 +
                ',[ПОКРАСКА]=' + #39 + Str + #39 +
                ' WHERE ([IdГП]=' + #39 + Label1.Caption + #39 + ') AND (Номер=' +
                 #39 + Label2.Caption + #39 + ') ', ['СТАМ']) then
                Exit;
        end;

       Res:=Pos('Планирование',FStamTime.Caption);
       if Res=0 then
            if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 FStamTime.Caption+ ']=' + #39 + Str + #39 +
                ' WHERE ([IdГП]=' + #39 + Label1.Caption + #39 + ') AND (Номер=' +
                 #39 + Label2.Caption + #39 + ')', ['ЗапускСТАМ']) then
                Exit;

       if Res<>0 then
            if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 FStamTime.Caption+ ']=' + #39 + Str + #39 +
                ' WHERE ([IdГП]=' + #39 + Label1.Caption + #39 + ')', ['ЗапускСТАМ']) then
                Exit;
       Form1.SG215.Cells[StrToInt(Label3.Caption),StrToInt(Label4.Caption)]:=Str1;
  end;

  Res:=AnsiCompareStr(FStamTime.Caption,'Дата запуска');
  Res1:=AnsiCompareStr(FStamTime.Caption,'TRUMPF');
  if Form1.Luk=1 then
  begin
          If  (Res1=0) Then
        Begin
               { UOsnova_Main.Flag_Error :=0;
                Vn_Dat := FormatDateTime('dd.mm.yyyy', DTP1.Date);
                fileName := ExtractFileDir(ParamStr(0));//+'\Klapan.EXE';
                if not UOsnova_Main.Osnova2( Vn_Dat,'ЗапускЛЮК','СпецифЛЮК',#39+Label2.Caption+#39,fileName,1 )
                then
                Exit; }
               if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 FStamTime.Caption + ']=' + #39 + Str + #39 +
                // ',[TRUMPF]=' + #39 + Str + #39 +
                ' WHERE ([IdГП]=' + #39 + Label1.Caption + #39 + ') AND (Номер=' +
                 #39 + Label2.Caption + #39 + ') ', ['ЗапускЛЮК']) then
                Exit;
                Form1.SG5.Cells[StrToInt(Label3.Caption),StrToInt(Label4.Caption)]:=Str1;
        End;
        Res:=AnsiCompareStr(FStamTime.Caption,'СВАРКА');
        If Res=0 Then
        begin
          if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 FStamTime.Caption + ']=' + #39 + Str + #39 +
                ' WHERE ([IdГП]=' + #39 + Label1.Caption + #39 + ') AND (Номер=' +
                 #39 + Label2.Caption + #39 + ') ', ['ЗапускЛЮК']) then
                Exit;
                Form1.SG5.Cells[StrToInt(Label3.Caption),StrToInt(Label4.Caption)]:=Str1;
        end;
        Res:=AnsiCompareStr(FStamTime.Caption,'ПОКРАСКА');
        If Res=0 Then
        begin
          if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 FStamTime.Caption + ']=' + #39 + Str + #39 +
                ',[Заготовка Готовность]=' + #39 + Str + #39 +
                ' WHERE ([IdГП]=' + #39 + Label1.Caption + #39 + ') AND (Номер=' +
                 #39 + Label2.Caption + #39 + ') ', ['ЗапускЛЮК']) then
                Exit;
                Form1.SG5.Cells[StrToInt(Label3.Caption),StrToInt(Label4.Caption)]:=Str1;
        end;
        Res:=AnsiCompareStr(FStamTime.Caption,'ГИБКА');
        if Res=0 then
        Begin
          if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 FStamTime.Caption+ ']=' + #39 + Str + #39 +
                ' WHERE ([IdГП]=' + #39 + Label1.Caption + #39 + ') AND (Номер=' +
                 #39 + Label2.Caption + #39 + ')', ['ЗапускЛЮК']) then
                Exit;
                Form1.SG5.Cells[StrToInt(Label3.Caption),StrToInt(Label4.Caption)]:=Str1;
        End;
        Res:=AnsiCompareStr(FStamTime.Caption,'Сборка Запуск');
        if Res=0 then
        Begin
          if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 FStamTime.Caption+ ']=' + #39 + Str + #39 +
                ' WHERE ([IdГП]=' + #39 + Label1.Caption + #39 + ') AND (Номер=' +
                 #39 + Label2.Caption + #39 + ')', ['ЗапускЛЮК']) then
                Exit;
                Form1.SG5.Cells[StrToInt(Label3.Caption),StrToInt(Label4.Caption)]:=Str1;
        End;
        Res1:=AnsiCompareStr(Label6.Caption,'1,5');
        if (Res=0) and (Res1=0) then
        begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 FStamTime.Caption + ']=' + #39 + Str + #39 +
                ',[СВАРКА]=' + #39 + Str + #39 +
                ',[ПОКРАСКА]=' + #39 + Str + #39 +
                ',[Заготовка Готовность]=' + #39 + Str + #39 +
                ' WHERE ([IdГП]=' + #39 + Label1.Caption + #39 + ') AND (Номер=' +
                 #39 + Label2.Caption + #39 + ') ', ['ЗапускЛЮК']) then
                Exit;
                Form1.SG5.Cells[StrToInt(Label3.Caption),StrToInt(Label4.Caption)]:=Str1;
                Form1.SG5.Cells[14,StrToInt(Label4.Caption)]:=Str1;
                Form1.SG5.Cells[15,StrToInt(Label4.Caption)]:=Str1;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 FStamTime.Caption+ ']=' + #39 + Str + #39 +
                ',[СВАРКА]=' + #39 + Str + #39 +
                ',[ПОКРАСКА]=' + #39 + Str + #39 +
                ' WHERE ([IdГП]=' + #39 + Label1.Caption + #39 + ')', ['ЛЮК']) then
                Exit;
        end;

      { Res:=Pos('Планирование',FStamTime.Caption);
       if Res<>0 then
            if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 FStamTime.Caption+ ']=' + #39 + Str + #39 +
                ' WHERE ([IdГП]=' + #39 + Label1.Caption + #39 + ') AND (Номер=' +
                 #39 + Label2.Caption + #39 + ')', ['ЗапускЛЮК']) then
                Exit;  }
       Res:=Pos('Упаковка',FStamTime.Caption);
       if Res<>0 then
            if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 FStamTime.Caption+ ']=' + #39 + Str + #39 +
                ' WHERE ([IdГП]=' + #39 + Label1.Caption + #39 + ')  AND (Номер=' +
                 #39 + Label2.Caption + #39 + ')', ['ЗапускЛЮК']) then
                Exit;
                Form1.SG5.Cells[StrToInt(Label3.Caption),StrToInt(Label4.Caption)]:=Str1;
  end;
        Close;
end;

procedure TFStamTime.FormShow(Sender: TObject);
begin
        //dtp1.DateTime:=Now;
end;

end.
