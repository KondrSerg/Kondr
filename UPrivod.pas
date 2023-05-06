unit UPrivod;

interface

uses
        Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls,
                Forms,
        Dialogs, ExtCtrls, Grids, StdCtrls, Buttons, ComObj;

type
        TFPrivod = class(TForm)
                SG2: TStringGrid;
                Panel1: TPanel;
                Panel2: TPanel;
                Label1: TLabel;
                Label2: TLabel;
                Label3: TLabel;
                Label4: TLabel;
                Label5: TLabel;
                Label6: TLabel;
                ComboBox1: TComboBox;
                Label7: TLabel;
                Button1: TButton;
    Splitter1: TSplitter;
    SG1: TStringGrid;
    Button2: TButton;
    Button3: TButton;
    Edit1: TEdit;
    Button4: TButton;
    SpeedButton1: TSpeedButton;
    Label8: TLabel;
                procedure FormShow(Sender: TObject);
                procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure SG1SetEditText(Sender: TObject; ACol, ARow: Integer;
      const Value: String);
    procedure Button4Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure ExportGridtoExcel1(Sender: TStringGrid);
        private
                { Private declarations }
                RR:Integer;
        public
                Fl:Integer;
                { Public declarations }
        end;

var
        FPrivod: TFPrivod;

implementation

uses Main;

{$R *.dfm}

procedure TFPrivod.FormShow(Sender: TObject);
var
        I, y, Kol_Ost, Kol_Potreb, Kol_Svobod,Kol,Kol1: Integer;
begin
        if Fl<>1 Then
        Begin
        Button3.Visible:=True;
        SG1.ColCount:=9;
        Form1.Clear_StringGrid(SG2);
        Form1.Clear_StringGrid(SG1);
        //SG1.Cells[0, 0] := 'Заказ';
        SG2.Cells[1, 0] := 'МодПривода';
        SG2.Cells[3, 0] := 'Выдано';
        SG2.Cells[2, 0] := 'Кол-во приводов';

        SG1.Cells[0, 0] := 'Заказ';
        SG1.Cells[1, 0] := 'МодПривода';
        SG1.Cells[2, 0] := 'Кол-во клапанов';
        SG1.Cells[3, 0] := 'Кол-во приводов';
        SG1.Cells[4, 0] := 'Привязано';
        SG1.Cells[5, 0] := 'Привязали';
        SG1.Cells[6, 0] := 'Свободный остаток';
        SG1.Cells[7, 0] := 'Наименование';
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
        'Select distinct (МодПривода) from %s Where (Х IS NULL) AND (Отмена IS NULL) AND ([Кол принятых]<[Кол во запущенных]) AND (['+'Номер'+']='+#39+Form1.ZSG.Cells[0,Form1.ZSG.Row]+#39+') '
        , ['Запуск']) then
        exit;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
                SG2.Cells[1, I + 1] :=
                        Form1.ADOQuery1.FieldByName('МодПривода').AsString;
                Form1.ADOQuery1.Next;
        end;
        For Y:=1 To SG2.RowCount Do
        Begin
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
        'Select * from %s Where ([МодПривода]='+#39+SG2.Cells[1,y]+#39+') AND (['+'Номер'+']='+#39+Form1.ZSG.Cells[0,Form1.ZSG.Row]+#39+') '
        , ['Запуск']) then
        exit;
        SG1.RowCount := Form1.ADOQuery1.RecordCount + 2;
        Kol_Ost:=0;
        Kol_Potreb:=0;
        //Y:=0;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin

                Kol:= Form1.ADOQuery1.FieldByName('Привод').AsInteger;
                Kol1:= Form1.ADOQuery1.FieldByName('Кол во запущенных').AsInteger;
                Kol_Ost:=Kol_Ost+(Kol*Kol1);
                Form1.ADOQuery1.Next;
        End;
        SG2.Cells[2,y]:=IntTostr(Kol_ost);
        End;

        if not Form1.mkQuerySelect(Form1.ADOQuery1,
        'Select * from %s Where (Х IS NULL) AND (Отмена IS NULL) AND ([Кол принятых]<[Кол во запущенных]) AND (['+'Номер'+']='+#39+Form1.ZSG.Cells[0,Form1.ZSG.Row]+#39+') '
        , ['Запуск']) then
        exit;
        SG1.RowCount := Form1.ADOQuery1.RecordCount + 2;
        Kol_Ost:=0;
        Kol_Potreb:=0;
        Y:=0;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
                SG1.Cells[0, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Заказ').AsString;
                SG1.Cells[1, I + 1] :=
                        Form1.ADOQuery1.FieldByName('МодПривода').AsString;
                SG1.Cells[2, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Кол во запущенных').AsString;
                SG1.Cells[7, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Изделие').AsString;
                If SG1.Cells[2, I + 1]='' Then
                Begin
                Kol_Ost:=Kol_Ost+0;
                Kol:=0;
                end
                Else
                Begin
                Kol_Ost:=Kol_Ost+StrToInt(SG1.Cells[2, I + 1]);
                Kol:=StrToInt(SG1.Cells[2, I + 1]);
                end;

                Kol1:=Form1.ADOQuery1.FieldByName('Привод').AsInteger;
                SG1.Cells[3, I + 1] :=IntToStr(Kol*Kol1);
                SG1.Cells[4, I + 1] :=IntToStr(Kol*Kol1);
                If SG1.Cells[3, I + 1]='' Then
                Kol_Potreb:=Kol_Potreb+0 Else
                Kol_Potreb:=Kol_Potreb+StrToInt(SG1.Cells[3, I + 1]);
                Y:=I+1;
                Form1.ADOQuery1.Next;
        end;
        SG1.Cells[0, y + 1] :='Итого';
        SG1.Cells[2, y + 1] :=IntToStr(Kol_Ost);
        SG1.Cells[3, y + 1] :=IntToStr(Kol_Potreb);
        For Y:=1 To SG1.RowCount Do
        Begin
        Form1.ADOQ1S.Close;
        Form1.ADOQ1S.SQL.Clear;
        Form1.ADOQ1S.SQL.Text :='Select * From ПривязкаПриводов Where (Привод='+#39+SG1.Cells[1,Y]+#39+') AND (Заказ='+#39+SG1.Cells[0,Y]+#39+') Order By Привод';
        Form1.ADOQ1S.Open;
        SG1.Cells[6,Y]:=Form1.ADOQ1S.FieldByName('Остаток').AsString;
        if SG1.Cells[6,Y]='' Then
        SG1.Cells[6,Y]:='0';
        if SG1.Cells[6,y]='0' then
                SG1.Cells[4,y]:='0';
        End;
        For Y:=1 To SG1.RowCount Do
        Begin
        Form1.ADOQ1S.Close;
        Form1.ADOQ1S.SQL.Clear;
        Form1.ADOQ1S.SQL.Text :='Select * From Привод1 Where (Привод='+#39+SG1.Cells[1,Y]+#39+')  Order By Привод';
        Form1.ADOQ1S.Open;
        SG1.Cells[6,Y]:=Form1.ADOQ1S.FieldByName('Остаток').AsString;
        if SG1.Cells[6,Y]='' Then
        SG1.Cells[6,Y]:='0';
        //if SG1.Cells[7,y]='0' then
        //        SG1.Cells[4,y]:='0';
        End;
        end
        Else
        Begin
        Button3.Visible:=False;
        Button4.Click;

        end;
end;

procedure TFPrivod.Button1Click(Sender: TObject);
var
        I, y, Kol_Ost, Kol_Potreb, Kol_Svobod: Integer;
begin

        Form1.Clear_StringGrid(SG2);
        Form1.Clear_StringGrid(SG1);
        SG2.Cells[0, 0] := 'Заказ';
        SG2.Cells[1, 0] := 'Кол-во';
        SG2.Cells[2, 0] := 'Потребность';
        SG2.Cells[3, 0] := 'Свободный остаток';
        SG1.Cells[0, 0] := 'Заказ';
        SG1.Cells[1, 0] := 'МодПривода';
        SG1.Cells[2, 0] := 'Кол-во клапанов';
        SG1.Cells[3, 0] := 'Кол-во приводов';

        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Form1.ADOQ1S.Close;
        Form1.ADOQ1S.SQL.Clear;
        Form1.ADOQ1S.SQL.Text := 'Select * From ПривязкаПриводов WHERE Привод='
                + #39 + 'Электропривод ' + ComboBox1.Text + #39 +
                ' Order By Привод';
        Form1.ADOQ1S.Open;
        SG2.RowCount := Form1.ADOQ1S.RecordCount + 2;
        Kol_Ost:=0;
        for I := 0 to Form1.ADOQ1S.RecordCount - 1 do
        begin
                SG2.Cells[0, I + 1] :=
                        Form1.ADOQ1S.FieldByName('Заказ').AsString;
                SG2.Cells[1, I + 1] :=
                        Form1.ADOQ1S.FieldByName('Остаток').AsString;
                Kol_Ost:=Kol_Ost+Form1.ADOQ1S.FieldByName('Остаток').AsInteger;
                SG2.Cells[2, I + 1] :=
                        Form1.ADOQ1S.FieldByName('Потребность').AsString;
                Kol_Potreb:=Kol_Potreb+Form1.ADOQ1S.FieldByName('Потребность').AsInteger;
                SG2.Cells[3, I + 1] :=
                        Form1.ADOQ1S.FieldByName('Свободно').AsString;
                Kol_Svobod:=Kol_Svobod+Form1.ADOQ1S.FieldByName('Свободно').AsInteger;

                Y:=I+1;
                Form1.ADOQ1S.Next;
        end;
        SG2.Cells[0, y + 1] :='Итого';
        SG2.Cells[1, y + 1] :=IntToStr(Kol_Ost);
        SG2.Cells[2, y + 1] :=IntToStr(Kol_Potreb);
        SG2.Cells[3, y + 1] :=IntToStr(Kol_Svobod);
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Form1.ADOQ1S.Close;
        Form1.ADOQ1S.SQL.Clear;
        Form1.ADOQ1S.SQL.Text := 'Select * From Привод WHERE Привод=' + #39 +
                'Электропривод ' + ComboBox1.Text + #39 + ' Order By Привод';
        Form1.ADOQ1S.Open;
        Label2.Caption := Form1.ADOQ1S.FieldByName('Остаток').AsString + ' шт.';
        Label3.Caption := Form1.ADOQ1S.FieldByName('Потребность').AsString +
                ' шт.';
        Label4.Caption := Form1.ADOQ1S.FieldByName('Свободно').AsString +
                ' шт.';
        //---------------------------------------------------------------------
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
        'Select * from %s Where (Х IS NULL) AND (Отмена IS NULL) AND ([Кол принятых]<[Кол во запущенных]) AND (['+'Номер'+']='+#39+Form1.ZSG.Cells[0,Form1.ZSG.Row]+#39+') '
        , ['Запуск']) then
        exit;
        SG1.RowCount := Form1.ADOQuery1.RecordCount + 2;
        Kol_Ost:=0;
        Kol_Potreb:=0;
        Y:=0;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin



                SG1.Cells[0, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Заказ').AsString;
                SG1.Cells[1, I + 1] :=
                        Form1.ADOQuery1.FieldByName('МодПривода').AsString;
                SG1.Cells[2, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Кол во запущенных').AsString;
                If SG1.Cells[2, I + 1]='' Then
                Kol_Ost:=Kol_Ost+0 Else
                Kol_Ost:=Kol_Ost+StrToInt(SG1.Cells[2, I + 1]);
                SG1.Cells[3, I + 1] :=
                        Form1.ADOQuery1.FieldByName('Привод').AsString;
                If SG1.Cells[3, I + 1]='' Then
                Kol_Potreb:=Kol_Potreb+0 Else
                Kol_Potreb:=Kol_Potreb+StrToInt(SG1.Cells[3, I + 1]);
                Y:=I+1;
                Form1.ADOQuery1.Next;
        end;
        SG1.Cells[0, y + 1] :='Итого';
        SG1.Cells[2, y + 1] :=IntToStr(Kol_Ost);
        SG1.Cells[3, y + 1] :=IntToStr(Kol_Potreb);
end;

procedure TFPrivod.Button2Click(Sender: TObject);
begin
        Close;
end;

procedure TFPrivod.Button3Click(Sender: TObject);
Var Svob,Vidan,Kol,I:Integer;
begin
        For I:=1 To SG1.RowCount Do
        Begin
        SG1.Cells[5, i]:=SG1.Cells[4, i];
        Svob:=StrToInt(SG1.Cells[6, i]);
        Kol:=StrToInt(SG1.Cells[4, i]);
        Vidan:=StrToInt(SG1.Cells[5, i]);
        Form1.ADOQ1S.Close;
        Form1.ADOQ1S.SQL.Clear;
        Form1.ADOQ1S.SQL.Text := 'Update  ПривязкаПриводов Set Остаток='+#39+IntToStr(Svob-Kol)+#39+' WHERE (Привод='
                + #39 + SG1.Cells[1, i] + #39 +') AND (Заказ='+#39+SG1.Cells[0, i]+#39+')';
        Form1.ADOQ1S.ExecSQL;
        //++++++++++++++++++++++++++++++++++++++++++++++++++
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [Кол выданных]=' +
                #39 +SG1.Cells[5, i]
                 + #39 +
                ' WHERE ([МодПривода]=' + #39 + SG1.Cells[1, i] + #39+') AND (Заказ='+#39+SG1.Cells[0, i]+#39+') AND (Номер='+#39+FPrivod.Caption+#39+') AND (Изделие='+#39+SG1.Cells[7, i]+#39+')',
                ['Запуск']) then
                Exit;
        SG1.Cells[5, i]:=SG1.Cells[4, i];
        SG1.Cells[4, i]:='0';

        end;
end;

procedure TFPrivod.SG1SetEditText(Sender: TObject; ACol, ARow: Integer;
  const Value: String);
begin
      Edit1.Text:=SG1.Cells[ACol,ARow];
      RR:=ARow;
end;

procedure TFPrivod.Button4Click(Sender: TObject);
Var I:Integer;
begin
        //ADOConnection1.ConnectionString:=MSSQL_CONN_STR;
        //ADOConnection1.Connected:=True;
        SG1.ColCount:=6;
        SG1.Cells[0,0]:='№';
        SG1.Cells[1,0]:='Привод';
        SG1.Cells[2,0]:='Остаток';
        SG1.Cells[3,0]:='Потребность';
        SG1.Cells[4,0]:='Заказано';
        SG1.Cells[5,0]:='Свободно для привязки';
        Form1.ADOQ1S.Close;
        Form1.ADOQ1S.SQL.Clear;
        Form1.ADOQ1S.SQL.Text :='Select * From Привод1 Order By Привод';
        Form1.ADOQ1S.Open;
        SG1.RowCount:= Form1.ADOQ1S.RecordCount+1;

        For I:=0 To  Form1.ADOQ1S.RecordCount-1 Do
        Begin
        SG1.Cells[0,i+1]:=IntToStr(I+1);
        SG1.Cells[1,i+1]:=Form1.ADOQ1S.FieldByName('Привод').AsString;
        SG1.Cells[2,i+1]:=Form1.ADOQ1S.FieldByName('Остаток').AsString;
        SG1.Cells[3,i+1]:=Form1.ADOQ1S.FieldByName('Потребность').AsString;
        SG1.Cells[4,i+1]:=Form1.ADOQ1S.FieldByName('Заказано').AsString;
        SG1.Cells[5,i+1]:=Form1.ADOQ1S.FieldByName('Свободно').AsString;
        Form1.ADOQ1S.Next;
        end;
end;
procedure TFPrivod.ExportGridtoExcel1(Sender: TStringGrid);
var
        XL: Variant;
        i, j: integer;
begin
        XL := CreateOleObject('Excel.Application');
        XL.Application.EnableEvents := false;
        XL.WorkBooks.Add;
        {XL.ActiveSheet.PageSetup.LeftMargin := XL.Application.InchesToPoints(0);
        XL.ActiveSheet.PageSetup.RightMargin :=
                XL.Application.InchesToPoints(0);
        XL.ActiveSheet.PageSetup.TopMargin := XL.Application.InchesToPoints(0);
        XL.ActiveSheet.PageSetup.BottomMargin :=
                XL.Application.InchesToPoints(0);}
        for j := 0 to Sender.RowCount - 1 do
        begin
                for I := 0 to (Sender.ColCount) - 1 do
                begin

                        if Sender.ColWidths[I] <> 0 then
                        begin

                                XL.ActiveWorkBook.WorkSheets[1].Cells[j + 1, i +
                                        1] := Sender.Cells[i,
                                        j];
                                XL.ActiveWorkBook.WorkSheets[1].Cells[j + 1, i +
                                        1].HorizontalAlignment
                                        := 3;
                        end;
                end;
        end;
        XL.ActiveWorkBook.WorkSheets[1].Range['A1', 'H'+IntToStr(Sender.RowCount) ].Borders.LineStyle := 2;
        XL.ActiveWorkBook.WorkSheets[1].Range['A1', 'H'+IntToStr(Sender.RowCount) ].Borders.LineStyle := 1;
        XL.ActiveWorkBook.WorkSheets[1].Range['A1', 'H1'].Columns.AutoFit;
        XL.ActiveWorkBook.WorkSheets[1].Range['A2', 'A2'].Columns.AutoFit;
        XL.ActiveWorkBook.WorkSheets[1].Range['H2', 'H2'].Columns.AutoFit;
        XL.ActiveWorkBook.WorkSheets[1].Range['A1', 'H'+IntToStr(Sender.RowCount)].HorizontalAlignment := 3;
        XL.Visible := True;
        XL := UnAssigned;
end;
procedure TFPrivod.SpeedButton1Click(Sender: TObject);
begin
ExportGridtoExcel1(SG1);
end;

end.


