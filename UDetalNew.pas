unit UDetalNew;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, StdCtrls;

type
  TFDetalNew = class(TForm)
    SpecGrid: TStringGrid;
    Label1: TLabel;
    Label2: TLabel;
    ComboBox1: TComboBox;
    Button1: TButton;
    Button2: TButton;
    Label3: TLabel;
    procedure FormShow(Sender: TObject);
    procedure ComboBox1Click(Sender: TObject);
    procedure SpecGridSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure Button2Click(Sender: TObject);
  private
                { Private declarations }
                TmpCol,TmpRow:Integer;
        public
                { Public declarations }
                AColG,ARowG:Integer;
  end;

var
  FDetalNew: TFDetalNew;

implementation

uses Main, UDetal, Reg;

{$R *.dfm}

procedure TFDetalNew.FormShow(Sender: TObject);
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

        SpecGrid.Cells[I_FN_NOM1, 1] := 'False';
        SpecGrid.Cells[I_FN_NOM1+2, 1] := 'False';
        SpecGrid.Cells[I_FN_NOM1+3, 1] := 'False';
        SpecGrid.Cells[I_FN_NOM1+4, 1] := 'False';
        SpecGrid.Cells[I_FN_NOM1+5, 1] := 'False';
        SpecGrid.Cells[I_FN_NOM1+6, 1] := 'False';
        SpecGrid.Cells[I_FN_NOM1+7, 1] := 'False';

end;

procedure TFDetalNew.ComboBox1Click(Sender: TObject);
begin
SpecGrid.Cells[AColG,SpecGrid.Row]:=ComboBox1.Items[ComboBox1.ItemIndex];
end;

procedure TFDetalNew.SpecGridSelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
  Var R,Rect1:TRect;
begin
        if (Form1.FlagDolg=1) OR (Form1.FlagDolg=2) Then
        Begin
        AColG:=ACol;
        ARowG:=ARow;
        //StringGrid2.LeftCol:=0 ;
        SpecGrid.Selection:=TGridRect(Rect(AColG, ARow, ARowG, ARow));
        Rect1:=SpecGrid.CellRect(ACol,ARow);
        TmpCol:=ACol;
        TmpRow:=ARow;
        
ComboBox1.Visible:=False;
  if ((ACol = I_FN_NOM1) OR (ACol = I_FN_NOM1+2) OR (ACol = I_FN_NOM1+3) OR (ACol = I_FN_NOM1+4) OR (ACol = I_FN_NOM1+5) OR (ACol = I_FN_NOM1+6) OR (ACol = I_FN_NOM1+7) OR ((ACol=I_FN_ZAK1) AND (Form1.Spec=1))  ) and (ARow <> 0) then   //КЦКП
  begin
    {Ширина и положение ComboBox должно соответствовать ячейке StringGrid}
    R := SpecGrid.CellRect(ACol, ARow);
    R.Left := R.Left + SpecGrid.Left;
    R.Right := R.Right + SpecGrid.Left;
    R.Top := R.Top + SpecGrid.Top;
    R.Bottom := R.Bottom + SpecGrid.Top;
    ComboBox1.Left := R.Left + 1;
    ComboBox1.Top := R.Top + 1; ComboBox1.Width := (R.Right + 1) - R.Left;
    ComboBox1.Height := (R.Bottom + 1) - R.Top; {Покажем combobox}
    ComboBox1.Visible := True;
    Label3.caption:=SpecGrid.Cells[ACol,0];
    //ComboBox10.SetFocus;
  end;

  end;
end;

procedure TFDetalNew.Button2Click(Sender: TObject);
begin
 if not Form1.mkQueryInsert(Form1.ADOQuery1, 'Insert Into %s '
                                        +
                                        '(ID, Технолог, [ВидЭлемента], [Элемент], [КолНаЕд], [Количество], [ЕИ], [Обозначение],' +
                                        '[Тип], [Категория], [Примечание], [Диаметр], [Масса], [Длина],   [Ширина], [ДлинаРазв],' +
                                        '[ШиринаРазв], [Объем], [КолГибов], [Канбан], [Материал], [КомплВед], [Trumph], [Ножницы],' +
                                        '[Углоруб], [Гибка], [Прокатка], [Пила],IdГП,Изделие,Заказ,Дата) ' +
                                        'Values (%s,%s,%s,%s,%s,%s,%s,%s,' +
                                        '%s,%s,%s,%s,%s,%s,%s,%s,' +
                                        '%s,%s,%s,%s,%s,%s,%s,%s,' +
                                        '%s,%s,%s,%s,%s,%s,%s,%s)',
                                        ['Специф',
                                        #39 + IntToStr(1) + #39, #39 + '1' +
                                                #39, #39 + SpecGrid.Cells[I_FN_POS_KOMP, 1] + #39, #39 + SpecGrid.Cells[I_FN_POSS, 1] + #39,
                                                #39 + SpecGrid.Cells[I_FN_OBOZN, 1] + #39, #39 + SpecGrid.Cells[I_FN_TYP_S, 1] + #39, #39
                                                + SpecGrid.Cells[I_FN_NOM_NOM, 1] + #39, #39 + SpecGrid.Cells[I_FN_ZBOR_ED, 1] + #39,
                                                #39 + 'tip' + #39, #39 + 'Kateg' +
                                                        #39, #39 + 'Prim' + #39, #39 + SpecGrid.Cells[I_FN_MODEL, 1]
                                                        + #39, #39 + 'Massa' + #39, #39 +
                                                        SpecGrid.Cells[I_FN_KOL_S, 1] + #39, #39 + SpecGrid.Cells[I_FN_SPRAM, 1] + #39,
                                                        #39 + SpecGrid.Cells[I_FN_UKRUG, 1] + #39,
                                                #39 + SpecGrid.Cells[I_FN_KPD, 1] + #39, #39 + SpecGrid.Cells[I_FN_KOMPL_W, 1]
                                                        + #39, #39 + SpecGrid.Cells[I_FN_NOM_GUR, 1] + #39, #39
                                                        + SpecGrid.Cells[I_FN_NOM1, 1] + #39, #39 + SpecGrid.Cells[I_FN_NOM1+1, 1] +
                                                        #39, #39 + 'Kompl_Ved' + #39, #39
                                                        + SpecGrid.Cells[I_FN_NOM1+2, 1] + #39, #39 + SpecGrid.Cells[I_FN_NOM1+3, 1] + #39,
                                                #39 + SpecGrid.Cells[I_FN_NOM1+4, 1] + #39, #39 + SpecGrid.Cells[I_FN_NOM1+5, 1]
                                                        + #39, #39 + SpecGrid.Cells[I_FN_NOM1+6, 1] + #39, #39 +
                                                        SpecGrid.Cells[I_FN_NOM1+7, 1] + #39, #39 + SpecGrid.Cells[I_FN_NOM1+8, 1] + #39, #39
                                                        + 'Naim' + #39, #39 + SpecGrid.Cells[I_FN_ZAK1, 1] + #39,
                                                        #39 + SpecGrid.Cells[I_FN_NOM1+11, 1] + #39])
                                                        then
                                        exit;
end;

end.
