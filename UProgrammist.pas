unit UProgrammist;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, AdvObj, BaseGrid, AdvGrid, ExtCtrls, Menus,Clipbrd;

type
  TFProgramm = class(TForm)
    SG21: TAdvStringGrid;
    pm1: TPopupMenu;
    Paste1: TMenuItem;
    pnl1: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    btn1: TButton;
    SG: TAdvStringGrid;
    btn2: TButton;
    Memo111: TMemo;
    Label3: TLabel;
    Memo: TMemo;
    Memo11: TMemo;
    SG211: TAdvStringGrid;
    procedure Paste1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure btn2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FProgramm: TFProgramm;

implementation

uses Main;

{$R *.dfm}

procedure TFProgramm.Paste1Click(Sender: TObject);
begin
        SG21.PasteSelectionFromClipboard;
end;

procedure TFProgramm.FormShow(Sender: TObject);
var Str,Nam,Trumph,Elem,Oboz,Dlina,DlinRaz,Shir,ShirRaz,Mat:string;
res,I,Res_KPD,Flan1,Flan2,Res_KPU,A,B,Sten,Trump1,Trump2,
Dlina_I1,Dlina_I2,hh,S,y:Integer;
Kol_Ed,Kol_Oboz,Kol_Zap:Double;
begin
        Memo.Lines.Clear;
        Memo11.Lines.Clear;
        Memo111.Lines.Clear;
        Nam:=Label2.Caption;
        Kol_Zap:=StrToFloat(Label3.Caption);
        Res_KPU:=Pos('КПУ',Nam);
        Res_KPD:=Pos('КПД',Nam);
        Flan2:=Pos('2*ф',Nam);
        Flan1:=Pos('1*ф',Nam);
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
            'Select IDГП,Путь,Изделие from %s  Where  ([IDГП]='+#39+Label1.Caption+#39+
            ')  ', ['Запуск']) then
            exit;
            Clipboard.AsText:=Form1.ADOQuery1.FieldByName('Путь').AsString;
            SG21.SelectRows(1,1);
            SG21.PasteSelectionFromClipboard;
        //++++++++++++++++++++++++++++++++++++++++++++++++++

        Memo111.Lines.Clear;
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
            'Select * from %s  Where  ([IDГП]='+#39+Label1.Caption+#39+
            ') AND (ВидЭлемента='+#39+'Детали'+#39+') ', ['Специф']) then
        exit;
        S:=0;
        Y:=0;
        for I:=0 To Form1.ADOQuery1.RecordCount-1 do
        Begin
                Str:=Form1.ADOQuery1.FieldByName('Обозначение').AsString;
                Elem:= Form1.ADOQuery1.FieldByName('Элемент').AsString;
                Kol_Ed:= StrToFloat(Form1.ADOQuery1.FieldByName('Количество').AsString);
                Kol_Oboz:= StrToFloat(Form1.ADOQuery1.FieldByName('Количество').AsString);
                Oboz:= Form1.ADOQuery1.FieldByName('Обозначение').AsString;
                Dlina:= Form1.ADOQuery1.FieldByName('Длина').AsString;
                Shir:= Form1.ADOQuery1.FieldByName('Ширина').AsString;
                DlinRaz:= Form1.ADOQuery1.FieldByName('ДлинаРазв').AsString;
                ShirRaz:= Form1.ADOQuery1.FieldByName('ШиринаРазв').AsString;
                Mat:= Form1.ADOQuery1.FieldByName('Материал').AsString;
                Sten:=Pos('Стенка',Elem);
                if Sten<>0 then
                begin
                        SG211.Cells[0,y+1]:=IntToStr(y+1);
                        SG211.Cells[1,y+1]:=Elem;
                        SG211.Cells[2,y+1]:=FloatToStr(Kol_Oboz*Kol_Zap);
                        SG211.Cells[3,y+1]:=Oboz;
                        SG211.Cells[4,y+1]:=Dlina;
                        SG211.Cells[5,y+1]:=Shir;
                        SG211.Cells[6,y+1]:=DlinRaz;
                        SG211.Cells[7,y+1]:=ShirRaz;
                        SG211.Cells[8,y+1]:=Mat;
                        Memo.Lines.Add(Str);
                        Inc(y);
                        Form1.ADOQuery1.Next;
                        Continue;

                end;
                SG.Cells[0,I+1]:=IntToStr(I+1);
                SG.Cells[1,I+1]:=Elem;
                SG.Cells[2,I+1]:=FloatToStr(Kol_Oboz*Kol_Zap);
                SG.Cells[3,I+1]:=Oboz;
                SG.Cells[4,I+1]:=Dlina;
                SG.Cells[5,I+1]:=Shir;
                SG.Cells[6,I+1]:=DlinRaz;
                SG.Cells[7,I+1]:=ShirRaz;
                SG.Cells[8,I+1]:=Mat;
                Memo11.Lines.Add(Str);
                Res:=Pos('-',Str);
                If Res<>0 then
                 Delete(Str,1,Res+1);
                Memo111.Lines.Add(Str);
                Form1.ADOQuery1.Next;
        end;
        //+++++++++++++++++++++++++++++++++++++++++++++
                if not Form1.mkQuerySelect(Form1.ADOQuery1,
            'Select * from %s  Where  ([IDГП]='+#39+Label1.Caption+#39+
            ')  ', ['Специф']) then
        exit;
        S:=0;
        Y:=0;
        SG21.RowCount:=Form1.ADOQuery1.RecordCount+2;
        for I:=0 To Form1.ADOQuery1.RecordCount-1 do
        Begin
                Str:=Form1.ADOQuery1.FieldByName('Обозначение').AsString;
                Elem:= Form1.ADOQuery1.FieldByName('Элемент').AsString;
                Kol_Ed:= StrToFloat(Form1.ADOQuery1.FieldByName('Количество').AsString);
                Kol_Oboz:= StrToFloat(Form1.ADOQuery1.FieldByName('Количество').AsString);
                Oboz:= Form1.ADOQuery1.FieldByName('Обозначение').AsString;
                Dlina:= Form1.ADOQuery1.FieldByName('Длина').AsString;
                Shir:= Form1.ADOQuery1.FieldByName('Ширина').AsString;
                DlinRaz:= Form1.ADOQuery1.FieldByName('ДлинаРазв').AsString;
                ShirRaz:= Form1.ADOQuery1.FieldByName('ШиринаРазв').AsString;
                Mat:= Form1.ADOQuery1.FieldByName('Материал').AsString;

                        SG21.Cells[0,y+1]:=IntToStr(y+1);
                        SG21.Cells[1,y+1]:=Elem;
                        SG21.Cells[2,y+1]:=FloatToStr(Kol_Oboz*Kol_Zap);
                        SG21.Cells[3,y+1]:=Oboz;
                        SG21.Cells[4,y+1]:=Dlina;
                        SG21.Cells[5,y+1]:=Shir;
                        SG21.Cells[6,y+1]:=DlinRaz;
                        SG21.Cells[7,y+1]:=ShirRaz;
                        SG21.Cells[8,y+1]:=Mat;
                        Memo.Lines.Add(Str);
                        Inc(y);

                Form1.ADOQuery1.Next;
        end;
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        S:=1;
        for I:=0 to Memo111.Lines.Count-1 do
        begin
             if not Form1.mkQuerySelect(Form1.ADOQuery1,
            'Select * from %s  Where  ([ОбозначениеШ]='+#39+Memo111.Lines.Strings[i]+#39+
            ') AND (Trumph='+#39+'False'+#39+
            ')', ['Шаблон']) then
            exit;
            if Form1.ADOQuery1.RecordCount<>0 then
            Memo111.Lines.Delete(i);
        end;
        {S:=I+1;
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++STENKI
        //for I:=0 to Memo.Lines.Count-1 do
        //begin
        if Memo.Lines.Count=2 then
        Begin
                Str:=Memo11.Lines.Strings[0];
                Res:=Pos('-',Str);
                Dlin_I1:=StrToInt(Copy(Str,res,8));

                Str:=Memo11.Lines.Strings[1];
                Res:=Pos('-',Str);
                Dlin_I2:=StrToInt(Copy(Str,res,8));

            if (Res_KPU<>0) and (Flan2<>0) and (Sten<>0) Then
            Begin
                hh:=DlinaI1 MOD 50;
                if  ((DlinaI1<150)  OR (DlinaI1>1000)) OR (hh<>0) Then
                Begin
                        SG.Cells[0,s+1]:=IntToStr(s+1);
                        SG.Cells[1,s+1]:=Elem;
                        SG.Cells[2,s+1]:=FloatToStr(Kol_Oboz*Kol_Zap);
                        SG.Cells[3,s+1]:=Oboz;
                        SG.Cells[4,s+1]:=Dlina;
                        SG.Cells[5,s+1]:=Shir;
                        SG.Cells[6,s+1]:=DlinRaz;
                        SG.Cells[7,s+1]:=ShirRaz;
                        SG.Cells[8,s+1]:=Mat;
                End;
                hh:=DlinaI2 MOD 50;
                if  ((DlinaI2<150)  OR (DlinaI2>1000)) OR (hh<>0) Then
                Begin
                        SG.Cells[0,s+2]:=IntToStr(s+2);
                        SG.Cells[1,s+2]:=Elem;
                        SG.Cells[2,s+2]:=FloatToStr(Kol_Oboz*Kol_Zap);
                        SG.Cells[3,s+2]:=Oboz;
                        SG.Cells[4,s+2]:=Dlina;
                        SG.Cells[5,s+2]:=Shir;
                        SG.Cells[6,s+2]:=DlinRaz;
                        SG.Cells[7,s+2]:=ShirRaz;
                        SG.Cells[8,s+2]:=Mat;
                End;
            end;

        end; }
end;

procedure TFProgramm.btn1Click(Sender: TObject);
begin
        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                'UPDATE %s SET [Путь]=' + #39
                + Memo11.Text + #39 + ' WHERE ([IdГП]=' + #39 +
                        Label1.Caption
                + #39 + ')',
                ['Klapana']) then
                Exit;
        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                'UPDATE %s SET [Путь]=' + #39
                + Memo11.Text  + #39 + ' WHERE ([IdГП]=' + #39 +
                        Label1.Caption
                + #39 + ')',
                ['Запуск']) then
                Exit;
end;

procedure TFProgramm.btn2Click(Sender: TObject);
begin
        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                'UPDATE %s SET [Путь]=' + #39
                +  #39 + ' WHERE ([IdГП]=' + #39+
                        Label1.Caption + #39 + ')',
                ['Klapana']) then
                Exit;
        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                'UPDATE %s SET [Путь]=' + #39+ #39 + ' WHERE ([IdГП]=' + #39 +
                        Label1.Caption
                + #39 + ')',
                ['Запуск']) then
                Exit;
end;

end.

