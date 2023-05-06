unit Lopatka;

interface

uses
        Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls,
        Forms,
        Dialogs, Grids, StdCtrls, ExtCtrls, ComObj;

type
        TFLopatka = class(TForm)
                StringGrid1: TStringGrid;
                Panel1: TPanel;
                Button1: TButton;
                StringGrid2: TStringGrid;
                Label1: TLabel;
                Splitter1: TSplitter;
                procedure FormShow(Sender: TObject);
                procedure Button1Click(Sender: TObject);
        private
                { Private declarations }
        public
                { Public declarations }
                Nom_Nom: Integer;
        end;

var
        FLopatka: TFLopatka;

implementation

uses Main;

{$R *.dfm}

procedure TFLopatka.FormShow(Sender: TObject);
var
        I, Kol, kol_klap, KolG, Kol_Ob, KolG_Ob, A, N, Poz, Res, Res1: Integer;
        C: Double;
        Nom_Poz, Obozn: string;
begin
        for I := 1 to FLopatka.StringGrid1.RowCount - 1 do
        begin
                Nom_Poz := FLopatka.StringGrid1.Cells[8, i];
                if Nom_Poz <> '' then
                begin
                        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                                'Select * from %s Where ([№Поз] =' + Nom_Poz +
                                ') AND ([№№]=' +
                                IntToStr(Nom_Nom) + ')', ['СпецифЗаказов']) then
                                exit;
                        if Form1.ADOQuery1.RecordCount = 0 then
                                Continue;
                        Kol_Klap := StrToInt(StringGrid1.Cells[15, i]);
                        Kol :=
                                StrToInt(Form1.ADOQuery1.FieldByName('Кол-во').AsString)
                                *
                                Kol_Klap;
                        FLopatka.StringGrid1.Cells[9, i] :=
                                Form1.ADOQuery1.FieldByName('№ПозКомпл').AsString;
                        FLopatka.StringGrid2.Cells[9, i] :=
                                Form1.ADOQuery1.FieldByName('№ПозКомпл').AsString;
                        FLopatka.StringGrid1.Cells[2, i] :=
                                Form1.ADOQuery1.FieldByName('Обозначение').AsString;
                        FLopatka.StringGrid1.Cells[3, i] := IntToStr(Kol);
                        //Form1.ADOQuery1.FieldByName('Кол-во').AsString;
                        FLopatka.StringGrid2.Cells[2, i] :=
                                Form1.ADOQuery1.FieldByName('Обозначение').AsString;
                        FLopatka.StringGrid2.Cells[3, i] := IntToStr(Kol);
                        //Form1.ADOQuery1.FieldByName('Кол-во').AsString;

                        Kol_Ob := Kol_Ob + Kol;
                end;
        end;

        for I := 1 to FLopatka.StringGrid1.RowCount - 1 do
        begin
                Nom_Poz := FLopatka.StringGrid1.Cells[8, i];
                if Nom_Poz <> '' then
                begin
                        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                                'Select * from %s Where ([№Поз] =' + Nom_Poz +
                                ') AND ([№№]=21)',
                                ['СпецифЗаказов']) then
                                exit;
                        if Form1.ADOQuery1.RecordCount = 0 then
                                Continue;
                        //FLopatka.StringGrid1.Cells[9,i]:=Form1.ADOQuery1.FieldByName('№ПозКомпл').AsString;
                        Kol_Klap := StrToInt(StringGrid1.Cells[15, i]);
                        Kol :=
                                StrToInt(Form1.ADOQuery1.FieldByName('Кол-во').AsString)
                                *
                                Kol_Klap;
                        FLopatka.StringGrid1.Cells[14, i] := IntToStr(Kol);
                        //Form1.ADOQuery1.FieldByName('Кол-во').AsString;
                        FLopatka.StringGrid2.Cells[14, i] := IntToStr(Kol);
                        //Form1.ADOQuery1.FieldByName('Кол-во').AsString;
                        FLopatka.StringGrid2.Cells[18, i] :=
                                Form1.ADOQuery1.FieldByName('Н\ч Гибка').AsString;
                        FLopatka.StringGrid2.Cells[19, i] :=
                                Form1.ADOQuery1.FieldByName('Н\ч Резка').AsString;
                        FLopatka.StringGrid1.Cells[18, i] :=
                                Form1.ADOQuery1.FieldByName('Н\ч Гибка').AsString;
                        FLopatka.StringGrid1.Cells[19, i] :=
                                Form1.ADOQuery1.FieldByName('Н\ч Резка').AsString;

                end;
        end;

        for I := 1 to FLopatka.StringGrid1.RowCount - 1 do
        begin
                Nom_Poz := FLopatka.StringGrid1.Cells[9, i];
                if Nom_Poz <> '' then
                begin
                        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                                'Select * from %s Where ([№ПозКомпл] =' + Nom_Poz
                                + ') ',
                                ['СпецифПараметрыДеталей']) then
                                exit;
                        Kol := StrToInt(StringGrid1.Cells[3, i]);
                        KolG :=
                                StrToInt(Form1.ADOQuery1.FieldByName('КолГибов').AsString);
                        FLopatka.StringGrid1.Cells[5, i] :=
                                Form1.ADOQuery1.FieldByName('КолГибов').AsString;
                        FLopatka.StringGrid1.Cells[6, i] := IntToStr(Kol *
                                KolG);
                        KolG_Ob := KolG_Ob + (Kol * KolG);
                        FLopatka.StringGrid2.Cells[4, i] :=
                                Form1.ADOQuery1.FieldByName('Материал').AsString;
                        FLopatka.StringGrid2.Cells[11, i] :=
                                Form1.ADOQuery1.FieldByName('ДлинаРазв').AsString;
                        FLopatka.StringGrid2.Cells[12, i] :=
                                Form1.ADOQuery1.FieldByName('ШиринаРазв').AsString;
                        FLopatka.StringGrid1.Cells[11, i] :=
                                Form1.ADOQuery1.FieldByName('ДлинаРазв').AsString;
                        FLopatka.StringGrid1.Cells[12, i] :=
                                Form1.ADOQuery1.FieldByName('ШиринаРазв').AsString;
                        {if not Form1.mkQuerySelect( Form1.ADOQuery2, 'Select * from %s Where ([№Поз] ='+FLopatka.StringGrid1.Cells[8,i]+') ' ,['СпецифЗаказов'])
                        then exit;

                        FLopatka.StringGrid2.Cells[18,i]:=Form1.ADOQuery2.FieldByName('Н\ч Гибка').AsString;
                        FLopatka.StringGrid2.Cells[19,i]:=Form1.ADOQuery2.FieldByName('Н\ч Резка').AsString;
                        FLopatka.StringGrid1.Cells[18,i]:=Form1.ADOQuery2.FieldByName('Н\ч Гибка').AsString;
                        FLopatka.StringGrid1.Cells[19,i]:=Form1.ADOQuery2.FieldByName('Н\ч Резка').AsString; }
                        if FLopatka.StringGrid1.Cells[14, i] = '' then
                                FLopatka.StringGrid1.Cells[14, i] := '0';
                        //N:=StrToInt(FLopatka.StringGrid1.Cells[14,i]);
                        //A:=StrToInt(FLopatka.StringGrid1.Cells[16,i]);
                        //C:=(((A-24)/(N/2)-8.7)+55.48)-(((A-24)/(N/2)-28)/2)-34.14;
                        //C:=(((A-24)/N-28)/2)-34.14;
                        FLopatka.StringGrid2.Cells[13, i] :=
                                Form1.ADOQuery1.FieldByName('Объем').AsString;
                        FLopatka.StringGrid1.Cells[13, i] :=
                                Form1.ADOQuery1.FieldByName('Объем').AsString;
                end;
        end;
        FLopatka.StringGrid1.RowCount := FLopatka.StringGrid1.RowCount + 1;
        FLopatka.StringGrid1.Cells[2, StringGrid1.RowCount - 1] :=
                'Всего в заданиях:';
        FLopatka.StringGrid1.Cells[4, StringGrid1.RowCount - 1] :=
                'Всего гибов:';
        FLopatka.StringGrid1.Cells[3, StringGrid1.RowCount - 1] :=
                IntToStr(Kol_Ob);
        FLopatka.StringGrid1.Cells[6, StringGrid1.RowCount - 1] :=
                IntToStr(KolG_Ob);
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        FLopatka.StringGrid2.RowCount := FLopatka.StringGrid2.RowCount + 1;
        FLopatka.StringGrid2.Cells[2, StringGrid2.RowCount - 1] :=
                'Всего в заданиях:';
        FLopatka.StringGrid2.Cells[3, StringGrid2.RowCount - 1] :=
                IntToStr(Kol_Ob);
        //Button1.Click;

end;

procedure TFLopatka.Button1Click(Sender: TObject);
var
        XL: Variant;
        L, K, i, j, y, PosTrud, CopyA: integer;
        Color1: TColor;
        TrudStr, S: string;
begin
        S := ExtractFileDir(ParamStr(0)) + '\Лопатки.xls';
        XL := CreateOleObject('Excel.Application');
        XL.Application.EnableEvents := false;
        XL.WorkBooks.Open(S);
        {XL.ActiveSheet.PageSetup.LeftMargin := XL.Application.InchesToPoints(0);
        XL.ActiveSheet.PageSetup.RightMargin :=
                XL.Application.InchesToPoints(0);
        XL.ActiveSheet.PageSetup.TopMargin := XL.Application.InchesToPoints(0);
        XL.ActiveSheet.PageSetup.BottomMargin :=
                XL.Application.InchesToPoints(0); }
        XL.ActiveWorkBook.WorkSheets[3].Range['A1', 'K1'].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[3].Range['A1', 'K1'].Font.Size := 16;
        Y := 4;
        for j := 3 to StringGrid1.RowCount + 3 do
        begin
                for I := 0 to (6) do
                begin

                        XL.ActiveWorkBook.WorkSheets[3].Cells[j + 1, i + 1] :=
                                StringGrid1.Cells[i, j - 3];
                        XL.ActiveWorkBook.WorkSheets[1].Cells[j + 1, i +
                                1].HorizontalAlignment :=
                                3;
                end;
                Inc(Y);
        end;
        XL.ActiveWorkBook.WorkSheets[3].Cells[1, 5] := StringGrid2.Cells[7, 1];
        XL.ActiveWorkBook.WorkSheets[3].Range['A' + IntToStr(Y), 'C' +
                IntToStr(Y)].Merge;
        XL.ActiveWorkBook.WorkSheets[3].Cells[Y, 1] :=
                'Задание сформировал (диспетчер)_______________';
        XL.ActiveWorkBook.WorkSheets[3].Range['D' + IntToStr(Y), 'H' +
                IntToStr(Y)].Merge;
        XL.ActiveWorkBook.WorkSheets[3].Cells[Y, 4] :=
                'Задание выдал (мастер)_______________';

        XL.ActiveWorkBook.WorkSheets[3].Range['A1', 'AB' +
                IntToStr(1)].Borders.LineStyle := 2;
        XL.ActiveWorkBook.WorkSheets[3].Range['A1', 'H' +
                IntToStr(Y)].Borders.LineStyle := 1;
        XL.ActiveWorkBook.WorkSheets[3].Range['C1', 'C7'].Columns.AutoFit;
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        XL.ActiveWorkBook.WorkSheets[2].Range['A1', 'K1'].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[2].Range['A1', 'K1'].Font.Size := 16;
        Y := 4;
        for j := 3 to StringGrid2.RowCount + 3 do
        begin
                for I := 0 to (6) do
                begin

                        XL.ActiveWorkBook.WorkSheets[2].Cells[j + 1, i + 1] :=
                                StringGrid2.Cells[i, j - 3];
                        XL.ActiveWorkBook.WorkSheets[2].Cells[j + 1, i +
                                1].HorizontalAlignment :=
                                3;
                end;
                Inc(Y);
        end;
        XL.ActiveWorkBook.WorkSheets[2].Cells[1, 6] := StringGrid2.Cells[7, 1];
        XL.ActiveWorkBook.WorkSheets[2].Range['A' + IntToStr(Y), 'C' +
                IntToStr(Y)].Merge;
        XL.ActiveWorkBook.WorkSheets[2].Cells[Y, 1] :=
                'Задание сформировал (диспетчер)_______________';
        XL.ActiveWorkBook.WorkSheets[2].Range['D' + IntToStr(Y), 'H' +
                IntToStr(Y)].Merge;
        XL.ActiveWorkBook.WorkSheets[2].Cells[Y, 4] :=
                'Задание выдал (мастер)_______________';

        XL.ActiveWorkBook.WorkSheets[2].Range['A1', 'AB' +
                IntToStr(1)].Borders.LineStyle := 2;
        XL.ActiveWorkBook.WorkSheets[2].Range['A1', 'H' +
                IntToStr(Y)].Borders.LineStyle := 1;
        XL.ActiveWorkBook.WorkSheets[2].Range['C1', 'C7'].Columns.AutoFit;
        XL.ActiveWorkBook.WorkSheets[2].Cells[50, 2] := '';
        //XL.ActiveWorkBook.Worksheets[ 2 ].Select;
        //XL.ActiveWorkBook.WorkSheets[ 2 ].Range['C10:E29'].Select;
        //XL.Selection.Pictures.Insert('C:\Безымянный.bmp');
        //XL.ActiveSheet.Pictures.Insert('C:\Безымянный.bmp');

      //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        XL.ActiveWorkBook.Worksheets[1].Select;
        XL.ActiveWorkBook.WorkSheets[1].Range['A1', 'K1'].Font.Bold := True;
        XL.ActiveWorkBook.WorkSheets[1].Range['A1', 'K1'].Font.Size := 11;
        XL.ActiveWorkBook.WorkSheets[1].Cells[1, 7] := StringGrid2.Cells[7, 1];
        XL.ActiveWorkBook.WorkSheets[1].Range['A1:J8'].Select;
        XL.Selection.Copy;
        CopyA := 8;

        for j := 1 to StrToInt(Label1.Caption) - 1 do
        begin

                XL.ActiveWorkBook.WorkSheets[1].Range['A' + IntToStr(CopyA +
                        1)].Select;
                XL.ActiveWorkBook.WorkSheets[1].Paste;
                CopyA := CopyA + 8;
        end;
        Y := 3;
        L := 5;
        K := 6;

        for j := 1 to StrToInt(Label1.Caption) do
        begin

                XL.ActiveWorkBook.WorkSheets[1].Cells[Y, 1] :=
                        StringGrid2.Cells[10, j];
                XL.ActiveWorkBook.WorkSheets[1].Cells[L, 3] :=
                        StringGrid2.Cells[2, j];
                XL.ActiveWorkBook.WorkSheets[1].Cells[K, 3] :=
                        StringGrid2.Cells[3, j];
                XL.ActiveWorkBook.WorkSheets[1].Cells[L, 6] := 'Лист: ' +
                        StringGrid2.Cells[4, j];
                XL.ActiveWorkBook.WorkSheets[1].Cells[K, 6] := 'Разв: ' +
                        StringGrid2.Cells[11, j] + '*' + StringGrid2.Cells[12, j]
                        + '*' +
                        StringGrid2.Cells[13, j];
                Y := Y + 8;
                L := L + 8;
                K := K + 8;
        end;
        XL.Visible := True;
        XL := UnAssigned;
        //Form1.Button12.Click;

end;

end.
