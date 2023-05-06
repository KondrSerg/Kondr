unit UPDO;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, AdvObj, BaseGrid, AdvGrid, StdCtrls, Buttons, AdvMemo, AdvmSQLS, ExtCtrls;

type
  TFPDO = class(TForm)
    SG1: TAdvStringGrid;
    advsqlmstylr1: TAdvSQLMemoStyler;
    Memo1: TMemo;
    pnl1: TPanel;
    btn2: TSpeedButton;
    btn1: TButton;
    procedure FormShow(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure btn2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FPDO: TFPDO;

implementation

uses Main;

{$R *.dfm}

procedure TFPDO.FormShow(Sender: TObject);
begin
        {SG1.ColCount:=9;
        SG1.MergeCells(0,0,1,2);
        SG1.MergeCells(1,0,1,2);
        SG1.MergeCells(2,0,1,2);
        SG1.MergeCells(3,0,1,2);
        SG1.MergeCells(4,0,1,2);
        SG1.MergeCells(5,0,1,2);
        SG1.MergeCells(6,0,1,2);
        SG1.MergeCells(7,0,1,2);
        SG1.MergeCells(8,0,1,2);

        SG1.Cells[0,0]:='Группа';
        SG1.Cells[1,0]:='Наименование';
        SG1.Cells[2,0]:='Не готовые';
        SG1.Cells[3,0]:='Предыдущий день';
        SG1.Cells[4,0]:='';
        SG1.Cells[5,0]:='Текущий день';
        SG1.Cells[6,0]:='Не завершенка';
        SG1.Cells[7,0]:='';
        SG1.Cells[8,0]:='Со срывом срока'; }
        btn1.Click;
end;

procedure TFPDO.btn1Click(Sender: TObject);
var Kol_Zap,Kol_Zap_o,I,Res,Res1,Res2:Integer;
Kol_Ger,Kol_Reg,Kol_Tulp,Kol_Stam:Integer;
Kol_Ger_O,Kol_Reg_O,Kol_Tulp_O,Kol_Stam_O,y,Kol_Prin:Integer;

Dat,Dat1:TDateTime;
Dat_S,Dat_S1,Nam,OTK_Kol_Raz,OTK_DATA,Str,Str1,Nom,Zakaz,Dat_Zad:String;
begin
        Memo1.Lines.Clear;
        SG1.ColCount:=9;
        SG1.MergeCells(0,0,1,2);
        SG1.MergeCells(1,0,1,2);
        SG1.MergeCells(2,0,1,2);
        SG1.MergeCells(3,0,2,1);
        SG1.MergeCells(5,0,2,1);
        SG1.MergeCells(7,0,1,2);
        SG1.MergeCells(8,0,1,2);

        SG1.Cells[0,0]:='Группа';
        SG1.Cells[1,0]:='Наименование';
        SG1.Cells[2,0]:='Не готовые';
        SG1.Cells[3,0]:='Предыдущий день';
        SG1.Cells[3,1]:='Запланированно';
        SG1.Cells[4,1]:='Принято ОТК';
        SG1.Cells[5,0]:='Текущий день';
        SG1.Cells[5,1]:='Запланированно';
        SG1.Cells[6,1]:='Принято ОТК';
        SG1.Cells[7,0]:='Не завершенка';
        SG1.Cells[8,0]:='Со срывом срока';

        SG1.Cells[0,2]:='310';
        SG1.Cells[1,2]:='Пожарные клапана';
        SG1.Cells[0,3]:='530';
        SG1.Cells[1,3]:='ГЕРМИК';
        SG1.Cells[0,4]:='520,525';
        SG1.Cells[1,4]:='РЕГУЛЯР,УВК';
        SG1.Cells[0,5]:='400';
        SG1.Cells[1,5]:='ТЮЛЬПАН,КЛАРА';
        SG1.Cells[0,6]:='600';
        SG1.Cells[1,6]:='СТАМ';
        Dat:=Now;
        Dat_S:=FormatDateTime('mm.dd.YYYY',Dat);
        Dat1:=Dat-1;//Вчера
        Dat_S1:=FormatDateTime('mm.dd.YYYY',Dat1);
        Kol_Zap_O :=0;
        //++++++++++++++++++++++++++++++++++=================Не готовые   AND (ОТК IS NULL)
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select * from %s Where (Х IS NULL) '
                +' AND (Отмена IS NULL) AND (['+FN_KOL_ZAP+']='+#39+'0'+#39+') ' , ['Klapana']) then
        exit;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
                Kol_Zap := StrToInt(Form1.ADOQuery1.FieldByName('Кол во').AsString);
                Kol_Zap_O := Kol_Zap_O + (Kol_Zap);
                Form1.ADOQuery1.Next;
        end;
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select * from %s Where (Х IS NULL) '
                +' AND (Отмена IS NULL) AND (ОТК IS NULL)' , ['Запуск']) then
        exit;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
                Kol_Zap := StrToInt(Form1.ADOQuery1.FieldByName('Кол во запущенных').AsString);
                Kol_Zap_O := Kol_Zap_O + (Kol_Zap);
                Form1.ADOQuery1.Next;
        end;
        SG1.Cells[2,2]:=IntToStr(Kol_Zap_o);
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Kol_Zap_O :=0;                                    //Вчера запланированно   AND (ОТК IS NULL)
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select * from %s Where (Х IS NULL) AND ([Планирование]='+#39+Dat_S1+#39+')'
                +' AND (Отмена IS NULL) ' , ['Запуск']) then
        exit;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
                Kol_Zap := StrToInt(Form1.ADOQuery1.FieldByName('Кол во запущенных').AsString);
                Kol_Zap_O := Kol_Zap_O + Kol_Zap;
                Form1.ADOQuery1.Next;
        end;
        SG1.Cells[3,2]:=IntToStr(Kol_Zap_o);
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Kol_Zap_O :=0;                              //Вчера принято ОТК
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select * from %s Where (Х IS NULL) AND ([ОТК]='+#39+Dat_S1+#39+')'
                +' AND (Отмена IS NULL) ' , ['Запуск']) then
        exit;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
                OTK_Kol_Raz:= Form1.ADOQuery1.FieldByName('ОТККолРаз').AsString;
                OTK_Data:= Form1.ADOQuery1.FieldByName('ОТКДата').AsString;
                Nom:= Form1.ADOQuery1.FieldByName('Номер').AsString;
                Zakaz:= Form1.ADOQuery1.FieldByName('Заказ').AsString;
                Res:=Pos('0,',OTK_Kol_Raz);
                If Res<>0 then
                        Delete(OTK_Kol_Raz,1,Res+1);

                Res1:=Pos('0,',OTK_Data);
                If Res1<>0 then
                        Delete(OTK_Data,1,Res1+1);

                for Y:=0 to 20 do
                begin
                  Res:=Pos(',',OTK_Kol_Raz);
                  Res1:=Pos(',',OTK_Data);
                  if Res<>0 then
                  begin
                        Str:=Copy(OTK_Kol_Raz,1,Res-1);
                        Str1:=Copy(OTK_Data,1,Res1-1);
                        Memo1.Lines.Add(Str+' : '+Str1+' : '+Nom+' : '+Zakaz);
                        Delete(OTK_Kol_Raz,1,Res);
                        Delete(OTK_Data,1,Res1);
                  end
                  else
                  begin
                        Memo1.Lines.Add(OTK_Kol_Raz+' : '+OTK_Data+' : '+Nom+' : '+Zakaz);
                        Break;
                  end;

                end;
                //+++++++++++++++++++++++++

                Kol_Zap := StrToInt(Form1.ADOQuery1.FieldByName('Кол принятых').AsString);
                Kol_Zap_O := Kol_Zap_O + Kol_Zap;
                Form1.ADOQuery1.Next;
        end;
        SG1.Cells[4,2]:=IntToStr(Kol_Zap_o);
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Kol_Zap_O :=0;                                   //Сегодня запланированно    AND (ОТК IS NULL)
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select * from %s Where (Х IS NULL) AND ([Планирование]='+#39+Dat_S+#39+')'
                +' AND (Отмена IS NULL) ' , ['Запуск']) then
        exit;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin

                Kol_Zap := StrToInt(Form1.ADOQuery1.FieldByName('Кол во запущенных').AsString);
                Kol_Zap_O := Kol_Zap_O + Kol_Zap;
                Form1.ADOQuery1.Next;
        end;
        SG1.Cells[5,2]:=IntToStr(Kol_Zap_o);
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Kol_Zap_O :=0;                                 //Сегодня принято ОТК
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select * from %s Where (Х IS NULL) AND ([ОТК]='+#39+Dat_S+#39+')'
                +' AND (Отмена IS NULL) ' , ['Запуск']) then
        exit;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
                OTK_Kol_Raz:= Form1.ADOQuery1.FieldByName('ОТККолРаз').AsString;
                OTK_Data:= Form1.ADOQuery1.FieldByName('ОТКДата').AsString;
                Nom:= Form1.ADOQuery1.FieldByName('Номер').AsString;
                Zakaz:= Form1.ADOQuery1.FieldByName('Заказ').AsString;
                Res:=Pos('0,',OTK_Kol_Raz);
                If Res<>0 then
                        Delete(OTK_Kol_Raz,1,Res+1);

                Res1:=Pos('0,',OTK_Data);
                If Res1<>0 then
                        Delete(OTK_Data,1,Res1+1);

                for Y:=0 to 20 do
                begin
                  Res:=Pos(',',OTK_Kol_Raz);
                  Res1:=Pos(',',OTK_Data);
                  if Res<>0 then
                  begin
                        Str:=Copy(OTK_Kol_Raz,1,Res-1);
                        Str1:=Copy(OTK_Data,1,Res1-1);
                        Memo1.Lines.Add(Str+' : '+Str1+' : '+Nom+' : '+Zakaz);
                        Delete(OTK_Kol_Raz,1,Res);
                        Delete(OTK_Data,1,Res1);
                  end
                  else
                  begin
                        Memo1.Lines.Add(OTK_Kol_Raz+' : '+OTK_Data+' : '+Nom+' : '+Zakaz);
                        Break;
                  end;

                end;
                //+++++++++++++++++++++++++
                Kol_Zap := StrToInt(Form1.ADOQuery1.FieldByName('Кол принятых').AsString);
                Kol_Zap_O := Kol_Zap_O + Kol_Zap;
                Form1.ADOQuery1.Next;
        end;
        SG1.Cells[6,2]:=IntToStr(Kol_Zap_o);
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Kol_Zap_O :=0;                               //Не завершенка
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select * from %s Where (Х IS NULL) AND (NOT [Сборка Запуск] IS NULL)'
                +' AND (ОТК IS NULL) AND (Отмена IS NULL) AND ([Заготовка Запуск]<>'+#39+Dat_S1+#39+') ' , ['Запуск']) then
        exit;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
                Kol_Zap := StrToInt(Form1.ADOQuery1.FieldByName('Кол во запущенных').AsString);
                Kol_Zap_O := Kol_Zap_O + Kol_Zap;
                Form1.ADOQuery1.Next;
        end;
        SG1.Cells[7,2]:=IntToStr(Kol_Zap_o);
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                Kol_Zap_O :=0;                               //Срыв
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select * from %s Where (Х IS NULL) AND ([План дата]<'+#39+Dat_S+#39+')'
                +' AND (ОТК IS NULL) AND (Отмена IS NULL) ' , ['Запуск']) then
        exit;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
                Kol_Zap := StrToInt(Form1.ADOQuery1.FieldByName('Кол во запущенных').AsString);
                Kol_Zap_O := Kol_Zap_O + Kol_Zap;
                Form1.ADOQuery1.Next;
        end;
        SG1.Cells[8,2]:=IntToStr(Kol_Zap_o);
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        //530 Гермик
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Kol_Zap_O :=0;
        Kol_Ger_O:=0;
        Kol_Reg_O:=0;
        Kol_Tulp_O:=0;
        Kol_Stam_O:=0;
        //++++++++++++++++++++++++++++++++++=================Не готовые       +' AND (ОТК IS NULL) '
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select * from %s Where (Х IS NULL) AND (Отмена IS NULL) AND (['+FN_KOL_ZAP+']='+#39+'0'+#39+')  '
                , ['KlapanaZap']) then
        exit;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
                Kol_Zap := StrToInt(Form1.ADOQuery1.FieldByName('Кол во').AsString);
               // Kol_Prin := StrToInt(Form1.ADOQuery1.FieldByName('Кол принятых').AsString);
                Nam:=Form1.ADOQuery1.FieldByName('Изделие').AsString;
                Res:=Pos('ГЕРМИК',Nam);
                Res1:=Pos('РЕГЛАН',Nam);
                if (Res<>0) or (Res1<>0) then
                        Kol_Ger_O := Kol_Ger_O + (Kol_Zap);
                //--------------------------------------------------------------
                Res:=Pos('РЕГУЛЯР',Nam);
                Res1:=Pos('УВК',Nam);
                Res2:=Pos('КВР',Nam);
                if (Res<>0) or (Res1<>0)  or (Res2<>0) then
                        Kol_Reg_O := Kol_Reg_O + (Kol_Zap);
                //--------------------------------------------------------------
                Res:=Pos('ТЮЛЬПАН',Nam);
                Res1:=Pos('КЛАРА',Nam);
                if (Res<>0) or (Res1<>0) then
                        Kol_TULP_O := Kol_Tulp_O + (Kol_Zap);
                Form1.ADOQuery1.Next;
        end;
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select * from %s Where (Х IS NULL) '
                +' AND (Отмена IS NULL) AND (ОТК IS NULL)' , ['ЗапускВозд']) then
        exit;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
                Kol_Zap := StrToInt(Form1.ADOQuery1.FieldByName('Кол во запущенных').AsString);
                Nam:=Form1.ADOQuery1.FieldByName('Изделие').AsString;
                Res:=Pos('ГЕРМИК',Nam);
                Res1:=Pos('РЕГЛАН',Nam);
                if (Res<>0) or (Res1<>0) then
                        Kol_Ger_O := Kol_Ger_O + (Kol_Zap);
                //--------------------------------------------------------------
                Res:=Pos('РЕГУЛЯР',Nam);
                Res1:=Pos('УВК',Nam);
                Res2:=Pos('КВР',Nam);
                if (Res<>0) or (Res1<>0)  or (Res2<>0) then
                        Kol_Reg_O := Kol_Reg_O + (Kol_Zap);
                //--------------------------------------------------------------
                Res:=Pos('ТЮЛЬПАН',Nam);
                Res1:=Pos('КЛАРА',Nam);
                if (Res<>0) or (Res1<>0) then
                        Kol_TULP_O := Kol_Tulp_O + (Kol_Zap);
                Form1.ADOQuery1.Next;
        end;
        SG1.Cells[2,3]:=IntToStr(Kol_Ger_o);
        SG1.Cells[2,4]:=IntToStr(Kol_Reg_o);
        SG1.Cells[2,5]:=IntToStr(Kol_Tulp_o);

        //Kol_Ger,Kol_Reg,Kol_Tulp,Kol_Stam:Integer;
        //Kol_G_O,Kol_Reg_O,Kol_Tulp_O,Kol_Stam_O:Integer;
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Kol_Zap_O :=0;
        Kol_Ger_O:=0;
        Kol_Reg_O:=0;
        Kol_Tulp_O:=0;
        Kol_Stam_O:=0;                                   //Вчера запланированно  AND (ОТК IS NULL)
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select * from %s Where (Х IS NULL) AND ([Планирование]='+#39+Dat_S1+#39+')'
                +' AND (Отмена IS NULL)  ' , ['ЗапускВозд']) then
        exit;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
                Kol_Zap := StrToInt(Form1.ADOQuery1.FieldByName('Кол во запущенных').AsString);
                Nam:=Form1.ADOQuery1.FieldByName('Изделие').AsString;
                Res:=Pos('ГЕРМИК',Nam);
                if Res<>0 then
                        Kol_Ger_O := Kol_Ger_O + Kol_Zap;
                //--------------------------------------------------------------
                Res:=Pos('РЕГУЛЯР',Nam);
                Res1:=Pos('УВК',Nam);
                if (Res<>0) or (Res1<>0) then
                        Kol_Reg_O := Kol_Reg_O + Kol_Zap;
                //--------------------------------------------------------------
                Res:=Pos('ТЮЛЬПАН',Nam);
                Res1:=Pos('КЛАРА',Nam);
                if (Res<>0) or (Res1<>0) then
                        Kol_TULP_O := Kol_Tulp_O + Kol_Zap;
                Form1.ADOQuery1.Next;
        end;
        SG1.Cells[3,3]:=IntToStr(Kol_Ger_o);
        SG1.Cells[3,4]:=IntToStr(Kol_Reg_o);
        SG1.Cells[3,5]:=IntToStr(Kol_Tulp_o);
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Kol_Zap_O :=0;
        Kol_Ger_O:=0;
        Kol_Reg_O:=0;
        Kol_Tulp_O:=0;
        Kol_Stam_O:=0;                             //Вчера принято ОТК
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select * from %s Where (Х IS NULL) AND ([ОТК]='+#39+Dat_S1+#39+')'
                +' AND (Отмена IS NULL) ' , ['ЗапускВозд']) then
        exit;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
                Kol_Zap := StrToInt(Form1.ADOQuery1.FieldByName('Кол принятых').AsString);
                Nam:=Form1.ADOQuery1.FieldByName('Изделие').AsString;
                Res:=Pos('ГЕРМИК',Nam);
                if Res<>0 then
                        Kol_Ger_O := Kol_Ger_O + Kol_Zap;
                //--------------------------------------------------------------
                Res:=Pos('РЕГУЛЯР',Nam);
                Res1:=Pos('УВК',Nam);
                if (Res<>0) or (Res1<>0) then
                        Kol_Reg_O := Kol_Reg_O + Kol_Zap;
                //--------------------------------------------------------------
                Res:=Pos('ТЮЛЬПАН',Nam);
                Res1:=Pos('КЛАРА',Nam);
                if (Res<>0) or (Res1<>0) then
                        Kol_TULP_O := Kol_Tulp_O + Kol_Zap;
                Form1.ADOQuery1.Next;
        end;
        SG1.Cells[4,3]:=IntToStr(Kol_Ger_o);
        SG1.Cells[4,4]:=IntToStr(Kol_Reg_o);
        SG1.Cells[4,5]:=IntToStr(Kol_Tulp_o);
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Kol_Zap_O :=0;
        Kol_Ger_O:=0;
        Kol_Reg_O:=0;
        Kol_Tulp_O:=0;
        Kol_Stam_O:=0;                                 //Сегодня запланированно    AND (ОТК IS NULL)
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select * from %s Where (Х IS NULL) AND ([Планирование]='+#39+Dat_S+#39+')'
                +' AND (Отмена IS NULL) ' , ['ЗапускВозд']) then
        exit;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
                Kol_Zap := StrToInt(Form1.ADOQuery1.FieldByName('Кол во запущенных').AsString);
                Nam:=Form1.ADOQuery1.FieldByName('Изделие').AsString;
                Res:=Pos('ГЕРМИК',Nam);
                if Res<>0 then
                        Kol_Ger_O := Kol_Ger_O + Kol_Zap;
                //--------------------------------------------------------------
                Res:=Pos('РЕГУЛЯР',Nam);
                Res1:=Pos('УВК',Nam);
                if (Res<>0) or (Res1<>0) then
                        Kol_Reg_O := Kol_Reg_O + Kol_Zap;
                //--------------------------------------------------------------
                Res:=Pos('ТЮЛЬПАН',Nam);
                Res1:=Pos('КЛАРА',Nam);
                if (Res<>0) or (Res1<>0) then
                        Kol_TULP_O := Kol_Tulp_O + Kol_Zap;
                Form1.ADOQuery1.Next;
        end;
        SG1.Cells[5,3]:=IntToStr(Kol_Ger_o);
        SG1.Cells[5,4]:=IntToStr(Kol_Reg_o);
        SG1.Cells[5,5]:=IntToStr(Kol_Tulp_o);
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Kol_Zap_O :=0;
        Kol_Ger_O:=0;
        Kol_Reg_O:=0;
        Kol_Tulp_O:=0;
        Kol_Stam_O:=0;                                //Сегодня принято ОТК
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select * from %s Where (Х IS NULL) AND ([ОТК]='+#39+Dat_S+#39+')'
                +' AND (Отмена IS NULL) ' , ['ЗапускВозд']) then
        exit;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
                Kol_Zap := StrToInt(Form1.ADOQuery1.FieldByName('Кол принятых').AsString);
                Nam:=Form1.ADOQuery1.FieldByName('Изделие').AsString;
                Res:=Pos('ГЕРМИК',Nam);
                if Res<>0 then
                        Kol_Ger_O := Kol_Ger_O + Kol_Zap;
                //--------------------------------------------------------------
                Res:=Pos('РЕГУЛЯР',Nam);
                Res1:=Pos('УВК',Nam);
                if (Res<>0) or (Res1<>0) then
                        Kol_Reg_O := Kol_Reg_O + Kol_Zap;
                //--------------------------------------------------------------
                Res:=Pos('ТЮЛЬПАН',Nam);
                Res1:=Pos('КЛАРА',Nam);
                if (Res<>0) or (Res1<>0) then
                        Kol_TULP_O := Kol_Tulp_O + Kol_Zap;
                Form1.ADOQuery1.Next;
        end;
        SG1.Cells[6,3]:=IntToStr(Kol_Ger_o);
        SG1.Cells[6,4]:=IntToStr(Kol_Reg_o);
        SG1.Cells[6,5]:=IntToStr(Kol_Tulp_o);
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Kol_Zap_O :=0;
        Kol_Ger_O:=0;
        Kol_Reg_O:=0;
        Kol_Tulp_O:=0;
        Kol_Stam_O:=0;                               //Не завершенка
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select * from %s Where (Х IS NULL) AND (NOT [Сборка Запуск] IS NULL)'
                +' AND (ОТК IS NULL) AND (Отмена IS NULL) AND ([Заготовка Запуск]<>'+#39+Dat_S1+#39+') ' , ['ЗапускВозд']) then
        exit;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
                Kol_Zap := StrToInt(Form1.ADOQuery1.FieldByName('Кол во запущенных').AsString);
                Nam:=Form1.ADOQuery1.FieldByName('Изделие').AsString;
                Res:=Pos('ГЕРМИК',Nam);
                if Res<>0 then
                        Kol_Ger_O := Kol_Ger_O + Kol_Zap;
                //--------------------------------------------------------------
                Res:=Pos('РЕГУЛЯР',Nam);
                Res1:=Pos('УВК',Nam);
                if (Res<>0) or (Res1<>0) then
                        Kol_Reg_O := Kol_Reg_O + Kol_Zap;
                //--------------------------------------------------------------
                Res:=Pos('ТЮЛЬПАН',Nam);
                Res1:=Pos('КЛАРА',Nam);
                if (Res<>0) or (Res1<>0) then
                        Kol_TULP_O := Kol_Tulp_O + Kol_Zap;
                Form1.ADOQuery1.Next;
        end;
        SG1.Cells[7,3]:=IntToStr(Kol_Ger_o);
        SG1.Cells[7,4]:=IntToStr(Kol_Reg_o);
        SG1.Cells[7,5]:=IntToStr(Kol_Tulp_o);
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Kol_Zap_O :=0;
        Kol_Ger_O:=0;
        Kol_Reg_O:=0;
        Kol_Tulp_O:=0;
        Kol_Stam_O:=0;                              //Срыв
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select * from %s Where (Х IS NULL) AND ([Расчетная дата готовности]<'+#39+Dat_S+#39+')'
                +' AND (ОТК IS NULL) AND (Отмена IS NULL)  ' , ['ЗапускВозд']) then
        exit;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
                Kol_Zap := StrToInt(Form1.ADOQuery1.FieldByName('Кол во запущенных').AsString);
                Nam:=Form1.ADOQuery1.FieldByName('Изделие').AsString;
                Res:=Pos('ГЕРМИК',Nam);
                if Res<>0 then
                        Kol_Ger_O := Kol_Ger_O + Kol_Zap;
                //--------------------------------------------------------------
                Res:=Pos('РЕГУЛЯР',Nam);
                Res1:=Pos('УВК',Nam);
                if (Res<>0) or (Res1<>0) then
                        Kol_Reg_O := Kol_Reg_O + Kol_Zap;
                //--------------------------------------------------------------
                Res:=Pos('ТЮЛЬПАН',Nam);
                Res1:=Pos('КЛАРА',Nam);
                if (Res<>0) or (Res1<>0) then
                        Kol_TULP_O := Kol_Tulp_O + Kol_Zap;
                Form1.ADOQuery1.Next;
        end;
        SG1.Cells[8,3]:=IntToStr(Kol_Ger_o);
        SG1.Cells[8,4]:=IntToStr(Kol_Reg_o);
        SG1.Cells[8,5]:=IntToStr(Kol_Tulp_o);
end;

procedure TFPDO.btn2Click(Sender: TObject);
begin
        Form1.ExportGridtoExcel2(SG1);
end;

end.
