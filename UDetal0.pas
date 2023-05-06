unit UDetal0;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, StdCtrls;

type
  TFDetal0 = class(TForm)
    SpecGrid: TStringGrid;
    Label1: TLabel;
    Label2: TLabel;
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FDetal0: TFDetal0;

implementation

uses Main, UDetal;

{$R *.dfm}

procedure TFDetal0.FormShow(Sender: TObject);
var
        I: Integer;
begin
        SpecGrid.ColCount := I_FN_NOM1 + 12;
        SpecGrid.Cells[0, 0] := '№';
        if Form1.Spec=1 Then
        SpecGrid.Cells[I_FN_ZAK1, 0] := 'Технолог'
        Else
        SpecGrid.Cells[I_FN_ZAK1, 0] := FN_ZAK1;
        SpecGrid.Cells[I_FN_POS_KOMP, 0] := 'ВидЭлемента';
        SpecGrid.Cells[I_FN_POSS, 0] := 'Элемент';
        SpecGrid.Cells[I_FN_OBOZN, 0] := 'КолНаЕд';

        SpecGrid.Cells[I_FN_TYP_S, 0] := 'Количество';
        SpecGrid.Cells[I_FN_NOM_NOM, 0] := 'ЕИ';

        SpecGrid.Cells[I_FN_ZBOR_ED, 0] := 'Обозначение';
        SpecGrid.Cells[I_FN_MODEL, 0] := 'Диаметр';
        SpecGrid.Cells[I_FN_KATEG, 0] := 'Масса';
        SpecGrid.Cells[I_FN_KOL_S, 0] := 'Длина';
        SpecGrid.Cells[I_FN_SPRAM, 0] := 'Ширина';
        SpecGrid.Cells[I_FN_UKRUG, 0] := 'ДлинаРазв';
        SpecGrid.Cells[I_FN_KPD, 0] := 'ШиринаРазв';
        SpecGrid.Cells[I_FN_KOMPL_W, 0] := 'Объем';
        SpecGrid.Cells[I_FN_NOM_GUR, 0] := 'КолГибов';
        SpecGrid.Cells[I_FN_NOM1, 0] := 'Канбан';
        SpecGrid.Cells[I_FN_NOM1+1, 0] := 'Материал';
        SpecGrid.Cells[I_FN_NOM1+2, 0] := 'Trumph';
        SpecGrid.Cells[I_FN_NOM1+3, 0] := 'Ножницы';
        SpecGrid.Cells[I_FN_NOM1+4, 0] := 'Углоруб';
        SpecGrid.Cells[I_FN_NOM1+5, 0] := 'Гибка';
        SpecGrid.Cells[I_FN_NOM1+6, 0] := 'Прокатка';
        SpecGrid.Cells[I_FN_NOM1+7, 0] := 'Пила';
        SpecGrid.Cells[I_FN_NOM1+8, 0] := 'IdГП';
        SpecGrid.Cells[I_FN_NOM1 + 9, 0] := 'Н\ч Гибка';
        SpecGrid.Cells[I_FN_NOM1 + 10, 0] := 'Н\ч Резка';
        SpecGrid.Cells[I_FN_NOM1 + 11, 0] := 'Дата';
{        if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s Where ['+FN_POSS+']='+#39+Label1.Caption+#39, ['СпецифЗаказов'])
                then
                exit;
        StringGrid2.ColCount := I_FN_NCH_GIB + 1;
        StringGrid2.Cells[0, 0] := '№';
        StringGrid2.Cells[I_FN_POS_KOMP, 0] := FN_POS_KOMP;
        StringGrid2.Cells[I_FN_POSS, 0] := FN_POSS;
        StringGrid2.Cells[I_FN_OBOZN, 0] := FN_OBOZN;
        StringGrid2.Cells[I_FN_ZAK1, 0] := FN_ZAK1;
        StringGrid2.Cells[I_FN_TYP_S, 0] := FN_TYP_S;
        StringGrid2.Cells[I_FN_NOM_NOM, 0] := FN_NOM_NOM;
        StringGrid2.Cells[I_FN_ZBOR_ED, 0] := FN_ZBOR_ED;
        StringGrid2.Cells[I_FN_MODEL, 0] := FN_MODEL;
        StringGrid2.Cells[I_FN_KATEG, 0] := FN_KATEG;
        StringGrid2.Cells[I_FN_KOL_S, 0] := FN_KOL_S;
        StringGrid2.Cells[I_FN_SPRAM, 0] := FN_SPRAM;
        StringGrid2.Cells[I_FN_UKRUG, 0] := FN_UKRUG;
        StringGrid2.Cells[I_FN_KPD, 0] := FN_KPD;
        StringGrid2.Cells[I_FN_KOMPL_W, 0] := FN_KOMPL_W;
        StringGrid2.Cells[I_FN_NOM_GUR, 0] := FN_NOM_GUR;
        StringGrid2.Cells[I_FN_NOM1, 0] := FN_NOM1;
        //StringGrid2.ColCount:=I_FN_NOM1+1;
        if Form1.ADOQuery1.RecordCount = 0 then
        begin
                StringGrid2.RowCount := 1;
                Form1.Clear_StringGrid(StringGrid2);
                exit;
        end;
        Form1.StatusBar1.Panels[2].Text := IntToStr(Form1.ADOQuery1.RecordCount);
        StringGrid2.RowCount := Form1.ADOQuery1.RecordCount + 1;
        Form1.ADOQuery1.First;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
                StringGrid2.Cells[0, I + 1] := IntToStr(I + 1);
                StringGrid2.Cells[I_FN_POS_KOMP, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_POS_KOMP).AsString;
                StringGrid2.Cells[I_FN_POSS, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_POSS).AsString;
                StringGrid2.Cells[I_FN_OBOZN, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_OBOZN).AsString;
                StringGrid2.Cells[I_FN_ZAK1, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_ZAK1).AsString;
                StringGrid2.Cells[I_FN_TYP_S, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_TYP_S).AsString;
                StringGrid2.Cells[I_FN_NOM_NOM, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_NOM_NOM).AsString;
                StringGrid2.Cells[I_FN_ZBOR_ED, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_ZBOR_ED).AsString;
                StringGrid2.Cells[I_FN_MODEL, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_MODEL).AsString;
                StringGrid2.Cells[I_FN_KATEG, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_KATEG).AsString;
                StringGrid2.Cells[I_FN_KOL_S, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_KOL_S).AsString;
                StringGrid2.Cells[I_FN_SPRAM, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_SPRAM).AsString;
                StringGrid2.Cells[I_FN_UKRUG, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_UKRUG).AsString;
                StringGrid2.Cells[I_FN_KPD, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_KPD).AsString;
                StringGrid2.Cells[I_FN_KOMPL_W, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_KOMPL_W).AsString;
                StringGrid2.Cells[I_FN_NOM_GUR, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_NOM_GUR).AsString;
                StringGrid2.Cells[I_FN_NOM1, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_NOM1).AsString;
                StringGrid2.Cells[I_FN_NCH_GIB, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_NCH_GIB).AsString;
                StringGrid2.Cells[I_FN_NCH_REZ, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_NCH_REZ).AsString;
                Form1.ADOQuery1.Next;
        end;   }
end;

end.
