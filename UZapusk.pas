unit UZapusk;

interface

uses
        Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls,
        Forms,
        Dialogs, StdCtrls, Grids, ExtCtrls, ComCtrls, ComObj;

type
        TFZapusk = class(TForm)
                Panel1: TPanel;
                SG1: TStringGrid;
                Button1: TButton;
                Button2: TButton;
                Panel2: TPanel;
                Label14: TLabel;
                Label15: TLabel;
                Panel3: TPanel;
                SG2: TStringGrid;
                Button3: TButton;
                Edit1: TEdit;
                Edit2: TEdit;
                Label13: TLabel;
                Label12: TLabel;
                Label1: TLabel;
                Label9: TLabel;
                Label11: TLabel;
                Memo1: TMemo;
                DateTimePicker1: TDateTimePicker;
                Label2: TLabel;
                ProgressBar1: TProgressBar;
                Label3: TLabel;
                Label4: TLabel;
                SD1: TStringGrid;
                ComboBox1: TComboBox;
                Edit3: TEdit;
                Edit4: TEdit;
                Edit5: TEdit;
                Edit6: TEdit;
                procedure FormShow(Sender: TObject);
                procedure FormClose(Sender: TObject; var Action: TCloseAction);
                procedure Button3Click(Sender: TObject);
                procedure Button1Click(Sender: TObject);
                procedure Button2Click(Sender: TObject);
                procedure Clear_Briket_IZ_MaskEdit(Str: string; var Br, Br1,
                        Br2: string);
                procedure SG2DblClick(Sender: TObject);
                procedure DateTimePicker1Enter(Sender: TObject);
                procedure SG2SelectCell(Sender: TObject; ACol, ARow: Integer;
                        var CanSelect: Boolean);
                function Xls_To_StringGrid(AGrid: TStringGrid; AXLSFile: string;
                        List:
                        string; kpd: Integer; L: Integer; Coll: Integer; Roww:
                        Integer): Boolean;
    procedure FormKeyPress(Sender: TObject; var Key: Char);
        private
                { Private declarations }
        public
                { Public declarations }
                Summa: Integer;
                F_Rasf: Integer;
        end;

var
        FZapusk: TFZapusk;

implementation

uses Main;

{$R *.dfm}
//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

function TFZapusk.Xls_To_StringGrid(AGrid: TStringGrid; AXLSFile: string; List:
        string; kpd: Integer; L: Integer; Coll: Integer; Roww: Integer):
        Boolean;
const
        xlCellTypeLastCell = $0000000B;
var
        XLApp, Sheet: OLEVariant;
        RangeMatrix: Variant;
        x, y, k, r, Res, I, J, Col, Posic, Kod, PosZ, F: Integer;
        Nam, Str, Str1, StrDat, Dat: string;
begin
        Result := False;
        Form1.Clear_StringGrid(SD1);
        XLApp := CreateOleObject('Excel.Application');
        try
                XLApp.Visible := False;
                XLApp.Workbooks.Open(AXLSFile);
                for Col := 1 to
                        XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets.Count do
                begin
                        Nam :=
                                XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets[Col].Name;
                        Res := AnsiCompareStr(Nam, List);
                        if Res = 0 then
                        begin
                                XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets.Item[Col].Activate;
                                Sheet :=
                                        XLApp.Workbooks[ExtractFileName(AXLSFile)].WorkSheets[Col];
                                Break;
                        end;
                end;
                Nam := Sheet.Name;
                Sheet.Cells.SpecialCells(xlCellTypeLastCell,
                        EmptyParam).Activate;
                if (kpd = 1) and (L = 1) then
                begin
                        x := Roww;
                        y := Coll;
                        AGrid.RowCount := x;
                        AGrid.ColCount := y + 2;
                        RangeMatrix := XLApp.Range['A1', XLApp.Cells.Item[X,
                                Y]].Value;
                        k := 1;
                        repeat
                                for r := 1 to y do
                                begin
                                        if r = 19 then
                                                Continue;
                                        Str := string(RangeMatrix[K, R]);
                                        AGrid.Cells[(r), (k - 1)] :=
                                                string(RangeMatrix[K, R]);
                                end;
                                Inc(k, 1);
                                AGrid.RowCount := k + 1;
                        until k > x;
                        RangeMatrix := Unassigned;
                end;

                if (kpd = 2) and (L = 2) then //Ножницы
                begin
                        x := Roww; //XLApp.ActiveCell.Row;
                        y := Coll;
                        AGrid.RowCount := x;
                        AGrid.ColCount := y + 2;
                        RangeMatrix := XLApp.Range['A1', XLApp.Cells.Item[X,
                                Y]].Value;
                        k := 1;
                        repeat
                                for r := 1 to y do
                                begin
                                        if (r = 8) or (R = 16) then
                                                Continue;
                                        if r = 9 then
                                                Continue;
                                        Str := string(RangeMatrix[K, R]);
                                        AGrid.Cells[(r), (k - 1)] :=
                                                string(RangeMatrix[K, R]);
                                end;
                                Inc(k, 1);
                                AGrid.RowCount := k + 1;
                        until k > x;
                        RangeMatrix := Unassigned;
                end;
                if (Kpd = 0) and (L = 0) then

                begin
                        x := Roww; //10;//XLApp.ActiveCell.Row;
                        y := Coll; //XLApp.ActiveCell.Column; //30;
                        AGrid.RowCount := x;
                        AGrid.ColCount := y + 2;
                        RangeMatrix := XLApp.Range['A1', XLApp.Cells.Item[X,
                                Y]].Value;
                        k := 1;
                        repeat
                                for r := 1 to y do
                                begin

                                        Str := string(RangeMatrix[K, R]);
                                        AGrid.Cells[(r), (k - 1)] :=
                                                string(RangeMatrix[K, R]);
                                end;
                                Inc(k, 1);
                                AGrid.RowCount := k + 1;
                        until k > x;
                        RangeMatrix := Unassigned;

                end;
                //=\===================================================Гибка
                if (Kpd = 0) and (L = 1) then

                begin
                        x := Roww; //10;//XLApp.ActiveCell.Row;
                        y := Coll; //XLApp.ActiveCell.Column; //30;
                        AGrid.RowCount := x;
                        AGrid.ColCount := y + 2;
                        RangeMatrix := XLApp.Range['A1', XLApp.Cells.Item[X,
                                Y]].Value;
                        k := 1;
                        repeat
                                for r := 1 to y do
                                begin
                                        if (R = 31) or (R = 19) or (R = 20) then
                                                Continue;
                                        Str := string(RangeMatrix[K, R]);
                                        AGrid.Cells[(r), (k - 1)] :=
                                                string(RangeMatrix[K, R]);
                                end;
                                Inc(k, 1);
                                AGrid.RowCount := k + 1;
                        until k > x;
                        RangeMatrix := Unassigned;

                end;
                //=\===================================================Программа
                if (Kpd = 0) and (L = 2) then

                begin
                        x := 16; //10;//XLApp.ActiveCell.Row;
                        y := 30; //XLApp.ActiveCell.Column; //30;
                        AGrid.RowCount := x;
                        AGrid.ColCount := y + 2;
                        for R := 0 to 16 do
                        begin
                                for K := 0 to 30 do
                                begin
                                        AGrid.Cells[(r), (k)] := Sheet.Cells[R +
                                                30, k];

                                end;
                        end;
                        {RangeMatrix := XLApp.Range['E29', 'AD44'].Value;
                        k := 29;
                        repeat
                                for r := 5 to y do
                                begin
                                        //if (R=31) OR (R=19) OR (R=20) Then
                                        //Continue;
                                        AGrid.Cells[(r), (k - 1)] := String(RangeMatrix[K, R]);//Str:= String(RangeMatrix[K, R]);

                                end;
                                Inc(k, 1);
                                //AGrid.RowCount := k + 1;
                        until k > 46;
                        RangeMatrix := Unassigned;}

                end;
        finally
                // Quit Excel
                if not VarIsEmpty(XLApp) then
                begin
                        // XLApp.DisplayAlerts := False;
                        XLApp.Quit;
                        XLAPP := Unassigned;
                        Sheet := Unassigned;
                        Result := True;
                end;
        end;
end;

procedure TFZapusk.Clear_Briket_IZ_MaskEdit(Str: string; var Br, Br1, Br2:
        string);
var
        Res: Integer;
        StrNed: string;
begin

        Str := StringReplace(Str, ' ', '', [RFReplaceall]);
        StrNed := Str;
        Res := Pos('(', StrNed);
        if Res <> 0 then
                Delete(StrNed, Res, 10);
        Br1 := StrNed;

        StrNed := Str;
        Res := Pos('-', StrNed);
        if Res <> 0 then
                Delete(StrNed, 1, Res);

        Res := Pos('(', StrNed);
        if Res <> 0 then
                Delete(StrNed, Res, 10);
        Br2 := StrNed;

        Br := Str;

end;
//------------------------------------------------------------------------------

procedure TFZapusk.FormShow(Sender: TObject);
var
        I, Kol, J, y, Res, Nom1, Nom2: Integer;
        Nom_Poz, Zak, IDGP, Kol_Zap, Kol_Raz: string;
begin

        Button1.Enabled := False;
        Summa := 0;
        SG1.Enabled := False;
        SG1.RowCount := 2;
        Form1.Clear_StringGrid(SG1);
        Form1.Clear_StringGrid(SG2);
        Form1.Clear_StringGrid(SD1);
        SG2.Cells[0, 0] := 'Заказ';
        SG2.Cells[1, 0] := 'Расчетная дата';
        SG2.Cells[2, 0] := 'Кол во запущенных';
        SG2.Cells[3, 0] := 'Номер';
        SG2.Cells[4, 0] := 'ID';
        SG2.ColWidths[0] := 0;
        SG2.ColWidths[2] := 0;
        SG2.ColWidths[4] := 0;
        Edit1.Text := '0';
        Memo1.Lines.Clear;
        Nom2 := 0;
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                'Select * from %s Where  ( [Номер]<>0) AND ([Заготовка Запуск] IS NULL) AND ([Отмена] IS NULL) AND (NOT [Технолог Готов] IS NULL) Order By Номер',
                ['Запуск']) then
                exit;
        Y := 1;
        j := 1;
        SG2.RowCount := Form1.ADOQuery2.RecordCount + 1;
        Memo1.Lines.Clear;
        if Form1.ADOQuery2.RecordCount <> 0 then
        begin
                for I := 0 to Form1.ADOQuery2.RecordCount - 1 do
                begin
                        Nom1 := Form1.ADOQuery2.FieldByName('Номер').AsInteger;
                        if j > 1 then
                                Nom2 := StrToInt(SG2.Cells[3, j - 1]);
                        if Nom1 = Nom2 then
                        begin
                                Form1.ADOQuery2.Next;
                                Continue;
                        end;
                        SG2.Cells[0, j] :=
                                Form1.ADOQuery2.FieldByName('Заказ').AsString;
                        SG2.Cells[1, j] :=
                                Form1.ADOQuery2.FieldByName('Расчетная дата готовности').AsString;
                        SG2.Cells[2, j] :=
                                Form1.ADOQuery2.FieldByName('Кол во запущенных').AsString;
                        SG2.Cells[3, j] :=
                                Form1.ADOQuery2.FieldByName('Номер').AsString;
                        SG2.Cells[4, j] :=
                                Form1.ADOQuery2.FieldByName('ID').AsString;
                        Nom_Poz := Form1.ADOQuery2.FieldByName('№Поз').AsString;
                        Inc(y);
                        Inc(j);
                        Form1.ADOQuery2.Next;
                end;
        end;
        SG2.RowCount := Y + 2;
        {if not Form1.mkQuerySelect( Form1.ADOQuery2, 'Select * from %s Where  ( [Номер]<>0) AND ([Заготовка Запуск] IS NULL) AND ([Кол во запущенных]=[Кол во] )',['Klapana'] )
        then exit;

      if Form1.ADOQuery2.RecordCount<>0 Then
      Begin
      SG2.RowCount:=J+ Form1.ADOQuery2.RecordCount+1;
      For i:=j To Form1.ADOQuery2.RecordCount+J Do
      Begin
        Kol_Raz:= Form1.ADOQuery2.FieldByName('Кол раз').AsString;
        Res:=Pos(',',Kol_Raz);
        if Res=0 Then
        Begin
                SG2.Cells[0,i+1]:=Form1.ADOQuery2.FieldByName('Заказ').AsString;
                SG2.Cells[1,i+1]:=Form1.ADOQuery2.FieldByName('Расчетная дата готовности').AsString;
                SG2.Cells[2,i+1]:=Form1.ADOQuery2.FieldByName('Кол во запущенных').AsString;
                SG2.Cells[3,i+1]:=Form1.ADOQuery2.FieldByName('Номер').AsString;
                SG2.Cells[4,i+1]:=Form1.ADOQuery2.FieldByName('ID').AsString;
        End;
        Form1.ADOQuery2.Next;
      End;
      End; }

end;

procedure TFZapusk.FormClose(Sender: TObject; var Action: TCloseAction);
begin
        Summa := 0;
        Label11.Caption := '0';
        Label9.Caption := '0';
        Label1.Caption := '0';
end;

procedure TFZapusk.Button3Click(Sender: TObject);
var
        I, Kol, Y, j, Pos, Pos1, x: Integer;
        Nom_Poz, Zak, IDGP, Briket, Briket1, Briket2, Br, Br1, Br2, Kol_Zap:
        string;
begin
        SG1.Enabled:=True;
        Form1.Clear_StringGrid(SG1);
        SG1.ColCount := 8;
        SG1.Cells[0, 0] := 'Заказ';
        SG1.Cells[1, 0] := 'Кол во запущенных';
        SG1.Cells[2, 0] := 'Номер';
        SG1.Cells[3, 0] := 'ID';
        SG1.Cells[4, 0] := 'Клапан';
        SG1.Cells[5, 0] := 'Сборка';
        SG1.Cells[6, 0] := 'Сварка';
        SG1.Cells[7, 0] := 'Привод';
        //  ([Заказ]= '+#39+Zak+#39+')AND
        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                'Select * from %s Where   (Номер='
                + #39 + Label15.Caption + #39 + ') ', ['Запуск']) then
                exit;
        SG1.RowCount:=1;;
        if Form1.ADOQuery2.RecordCount <> 0 then
        begin
                SG1.RowCount := Form1.ADOQuery2.RecordCount + 1;

                for i := 0 to Form1.ADOQuery2.RecordCount - 1 do
                begin
                        SG1.Cells[0, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Заказ').AsString;
                        SG1.Cells[1, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Кол во запущенных').AsString;
                        SG1.Cells[2, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Номер').AsString;
                        SG1.Cells[3, i + 1] :=
                                Form1.ADOQuery2.FieldByName('ID').AsString;
                        SG1.Cells[4, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Изделие').AsString;
                        SG1.Cells[5, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Н\ч Сборка Клапана').AsString;
                        SG1.Cells[6, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Н\ч Сварка').AsString;
                        SG1.Cells[7, i + 1] :=
                                Form1.ADOQuery2.FieldByName('МодПривода').AsString;
                        Form1.ADOQuery2.Next;
                end;
        end;
end;

procedure TFZapusk.Button1Click(Sender: TObject);
var
        Kol_ZAP, Kol, Kol_Zap_Ranee, Res, Nomer, Zak_Int, i, A, B, D, y, F_KPU,
                F_KPD,
                xy, Res1, e, Zag, Zap, j, g, k, List, Err, X_kol, Kol1, Kol2,NI24:
        Integer;
        Zak, Dat, Plan_Dat, Vn_Dat, Nom, Pos_Vst, Pos_Ml, R, Pereh, Privod,
                Zag_S,
                Zap_S, Dir_main, God, Mes, Fil, A_S, B_S,N22,n24,n26,n29,n30: string;
        Svar, Sbor, Izdel, Pos1, Pos2, Pos3, Pos4, Pos5, Pos_Isp, Pos_Flan,
                Pos_Flan1,
                Pos_Privod, Pos_Sn, Pos_Dop, Pos_Ram, Dir,
                R1, R2, R3, R4, R5, r6, r7, r8, r9, r10, r11, r12, r13, r14,
                r15, r16, r17,
                r18: string;
        NC,ND22: Double;
        Dat1, Dat2, Dat3: TDate;
        XL, XL1, XL2, XL_Temp, V1: Variant;
begin

        F_KPU := 0;
        F_KPD := 0;
        y := 1;
        Xy := 1;
        e := 30;
        K := 30;
        Err := 0;
        Vn_Dat := FormatDateTime('mm.dd.yyyy', DateTimePicker1.Date);
        Dir_main := ExtractFileDir(ParamStr(0));
        God := FormatDateTime('yyyy', Now);
        Mes := FormatDateTime('mmmm', Now);

        if (MessageDlg('Сохранить именно эту дату?',mtInformation,[mbYes,MbNo],0)=mrYes) Then
        Begin
        for i := 1 to SG1.RowCount do
        begin

                //ProgressBar1.Max := SG1.RowCount;
                if SG1.Cells[2, i] <> '' then
                begin
                        Nom := SG1.Cells[2, i];
                        Izdel := SG1.Cells[4, i];
                        Kol_Zap := StrToInt(SG1.Cells[1, i]);
                        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                                'Select Запуск,Заготовка from %s Where   (Заказ=' + #39 +
                                SG1.Cells[0, i] + #39 +
                                ') AND (Изделие=' + #39 + SG1.Cells[4, i] + #39
                                +
                                ')', ['Klapana']) then
                                exit;
                       // Fil := Form1.ADOQuery2.FieldByName('Файл').AsString;

                        //-----------------------------------------------
                        Zap_S := Form1.ADOQuery2.FieldByName('Запуск').AsString;
                        if Zap_S = '' then
                                Zap := 0
                        else
                                Zap := StrToInt(Zap_s);
                        Zap := Zap + StrToInt(SG1.Cells[1, i]);
                        //-----------------------------------------------
                        Zag_S :=
                                Form1.ADOQuery2.FieldByName('Заготовка').AsString;
                        if Zag_S = '' then
                                Zag := 0
                        else
                                Zag := StrToInt(Zag_s);
                        Zag := Zag + StrToInt(SG1.Cells[1, i]);

                        //++++++++++++++++++++++++++++++++++++++++++++++
                        if SG1.Cells[6, i] = '' then
                                Svar := '0'
                        else
                                Svar := (SG1.Cells[6, i]);
                        if SG1.Cells[5, i] = '' then
                                Sbor := '0'
                        else
                                Sbor := (SG1.Cells[5, i]);
                        Res := Pos(',', Svar);
                        Delete(Svar, Res, 1);
                        if Res <> 0 then
                                Insert('.', Svar, Res);

                        Res := Pos(',', Sbor);
                        Delete(Sbor, Res, 1);
                        if Res <> 0 then
                                Insert('.', Sbor, Res);
                        Privod := SG1.Cells[7, i];
                        //-----------------------------------------------
                        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                                'UPDATE %s SET [Заготовка Запуск]=' + #39 +
                                Vn_Dat + #39 +
                                ',Заготовка=' + #39 + IntToStr(Zag) + #39 +
                                ' ,Запуск=' + #39 +
                                IntToStr(Zap) + #39 + ' WHERE ([Заказ]=' + #39 +
                                SG1.Cells[0, i] + #39 +
                                ') AND (Изделие=' + #39 + SG1.Cells[4, i] + #39
                                        +
                                ')', ['Klapana']) then
                                Exit;
                        //++++++++++++++++++++++++++++++++++++++++++++++
                        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                                'UPDATE %s SET [Заготовка Запуск]=' + #39 +
                                Vn_Dat + #39 +
                                ' WHERE  ([Номер]=' + #39 + SG1.Cells[2, i] + #39
                                + ') ', ['Запуск']) then
                                Exit;

                end;

        end;

        Form1.Button16.Click;
        FZapusk.Close;
        end;


end;

procedure TFZapusk.Button2Click(Sender: TObject);
begin
        Close;
end;

procedure TFZapusk.SG2DblClick(Sender: TObject);
begin
        //If (SG2.Col=0) Then
        Label15.Caption := SG2.Cells[3, SG2.Row];
        Button3.Click;
end;

procedure TFZapusk.DateTimePicker1Enter(Sender: TObject);
begin
        Button1.Enabled := True;
end;

procedure TFZapusk.SG2SelectCell(Sender: TObject; ACol, ARow: Integer;
        var CanSelect: Boolean);
begin
        Label15.Caption := SG2.Cells[3, ARow];
        Button3.Click;
end;

procedure TFZapusk.FormKeyPress(Sender: TObject; var Key: Char);
begin
     If Key=#27 Then
     Close;
end;

end.

