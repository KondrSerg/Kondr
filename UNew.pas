unit UNew;

interface

uses
        Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls,
        Forms,
        Dialogs, Grids, StdCtrls, ComObj;

type
        TFVozKl = class(TForm)
                Edit1: TEdit;
                Edit2: TEdit;
                ComName: TComboBox;
                ComIspol: TComboBox;
                ComEPV: TComboBox;
                Label1: TLabel;
                Label2: TLabel;
                ComA: TComboBox;
                ComB: TComboBox;
                Label3: TLabel;
                ComL: TComboBox;
                Label4: TLabel;
                Label5: TLabel;
                ComboBox1: TComboBox;
                Label6: TLabel;
                Label7: TLabel;
                        Name: TLabel;
                Al: TLabel;
                Bl: TLabel;
                lPos2: TLabel;
                lPos1: TLabel;
                lPos3: TLabel;
                lPos4: TLabel;
                lPos5: TLabel;
                lPos6: TLabel;
                lPos7: TLabel;
                SGL: TStringGrid;
                Button1: TButton;
                OpenDialog1: TOpenDialog;
                SGR: TStringGrid;
                Button2: TButton;
                procedure Button1Click(Sender: TObject);
                function Xls_To_StringGrid(AGrid: TStringGrid; AXLSFile: string;
                        List:
                        string): Boolean;
                procedure Button2Click(Sender: TObject);
                procedure SGRClick(Sender: TObject);
                procedure FormShow(Sender: TObject);
        private
                { Private declarations }
        public
                { Public declarations }
        end;

var
        FVozKl: TFVozKl;

implementation

uses Main;

{$R *.dfm}

function TFVozKl.Xls_To_StringGrid(AGrid: TStringGrid; AXLSFile: string; List:
        string): Boolean;
const
        xlCellTypeLastCell = $0000000B;
var
        XLApp, Sheet: OLEVariant;
        RangeMatrix: Variant;
        x, y, k, r, Col, Res, I, J, Posic, Kod, PosZ, F: Integer;
        Nam, Str, Str1, StrDat, Dat: string;
begin
        Result := False;
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
                x := XLApp.ActiveCell.Row;
                y := 30; //XLApp.ActiveCell.Column;
                AGrid.RowCount := x;
                AGrid.ColCount := y + 2;
                RangeMatrix := XLApp.Range['A1', XLApp.Cells.Item[X, Y]].Value;
                k := 1;
                repeat
                        for r := 1 to y do
                        begin
                                AGrid.Cells[(r), (k - 1)] := RangeMatrix[K, R];
                        end;
                        Inc(k, 1);
                        AGrid.RowCount := k + 1;
                until k > x;
                RangeMatrix := Unassigned;

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

procedure TFVozKl.Button1Click(Sender: TObject);
var
        Str: string;
begin
        if OpenDialog1.Execute then
        begin
                Form1.Clear_StringGrid(SGL);
                Form1.Clear_StringGrid(SGR);
                Str := OpenDialog1.FileName;
                Xls_To_StringGrid(SGL, Str, 'ÃË‡ÒÒ');

        end;
end;

procedure TFVozKl.Button2Click(Sender: TObject);
var
        I, j, x, Kod: Integer;
begin
        J := 1;
        SGR.ColCount := SGL.ColCount;
        for I := 0 to SGL.RowCount - 1 do
        begin
                if SGL.Cells[1, i] <> '' then
                begin
                        try
                                Kod := StrToInt(SGL.Cells[1, i]);
                        except
                                Kod := 0;
                        end;
                        if (Kod = 520) or (Kod = 525) or (Kod = 530) or (Kod =
                                400) then
                        begin
                                for X := 0 to SGL.ColCount - 1 do
                                begin
                                        SGR.Cells[X, J] := SGL.Cells[X, I];
                                end;
                                Inc(J);
                                SGR.RowCount := SGR.RowCount + j;
                        end;

                end;
        end;

end;

procedure TFVozKl.SGRClick(Sender: TObject);
var
        Izdel, Nam, A, B, POS_Privod, Pos_Sn, Pos_Dop, Pos_Ram, Pos1, Pos2,
                Pos3,
                Pos4, Pos5, Pos6, Pos7: string;
        Res, Kod, Pos_ZV: Integer;
begin
        Izdel := SGr.Cells[7, SGR.Row];
        Kod := StrToInt(SGr.Cells[1, SGR.Row]);
        Name.Caption := '';
        Al.Caption := '';
        Bl.Caption := '';
        lPos1.Caption := '';
        lPos2.Caption := '';
        lPos3.Caption := '';
        lPos4.Caption := '';
        lPos5.Caption := '';
        lPos6.Caption := '';
        lPos7.Caption := '';
        if Kod = 520 then
        begin
                // Î‡Ô‡Ì –≈√”Àﬂ–-0525-0575-Õ-œ-12-00-00-”2 ‰Îˇ  ÓÌ‰ËˆËÓÌÂ  ÷ œ-3,15
               // Î‡Ô‡Ì –≈√”Àﬂ–-1075-1135-Õ-œ-04-00-00-”2
                Pos_ZV := Pos('*', Izdel);
                if Pos_Zv = 0 then
                begin
                        Res := Pos(' ', Izdel);
                        Delete(Izdel, 1, Res);
                        //======================================== –≈√”Àﬂ–
                        Res := Pos('-', Izdel);
                        Nam := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);

                        //========================================
                        Res := Pos('-', Izdel);
                        A := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        B := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        Pos1 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        Pos2 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        Pos3 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        Pos4 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        Pos5 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos(' ', Izdel);
                        Pos6 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Name.Caption := Nam;
                        Al.Caption := A;
                        Bl.Caption := B;
                        lPos1.Caption := Pos1;
                        lPos2.Caption := Pos2;
                        lPos3.Caption := Pos3;
                        lPos4.Caption := Pos4;
                        lPos5.Caption := Pos5;
                        lPos6.Caption := Pos6;
                end;
                // Î‡Ô‡Ì –≈√”Àﬂ–-1000*1000-Õ-M230-S_NM230A-S-œ
                Pos_ZV := Pos('_', Izdel);
                if Pos_Zv <> 0 then
                begin
                        Res := Pos(' ', Izdel);
                        Delete(Izdel, 1, Res);
                        //======================================== –≈√”Àﬂ–
                        Res := Pos('-', Izdel);
                        Nam := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);

                        //========================================
                        Res := Pos('*', Izdel);
                        A := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        B := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        Pos1 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        Pos2 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        Pos3 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        Pos4 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        Pos5 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Pos6 := Izdel;
                        //========================================
                        Name.Caption := Nam;
                        Al.Caption := A;
                        Bl.Caption := B;
                        lPos1.Caption := Pos1;
                        lPos2.Caption := Pos2;
                        lPos3.Caption := Pos3;
                        lPos4.Caption := Pos4;
                        lPos5.Caption := Pos5;
                        lPos6.Caption := Pos6;
                end;
                // Î‡Ô‡Ì –≈√”Àﬂ–-1000*820-Õ-1*Û˜ÌÓÈ-œ
                Pos_ZV := Pos('*', Izdel);
                if Pos_Zv <> 0 then
                begin
                        Res := Pos(' ', Izdel);
                        Delete(Izdel, 1, Res);
                        //======================================== –≈√”Àﬂ–
                        Res := Pos('-', Izdel);
                        Nam := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);

                        //========================================
                        Res := Pos('*', Izdel);
                        A := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        B := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        Pos1 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        Pos2 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('*', Izdel);
                        Pos3 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        Pos4 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        Pos5 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Pos6 := Izdel;
                        //========================================
                        Name.Caption := Nam;
                        Al.Caption := A;
                        Bl.Caption := B;
                        lPos1.Caption := Pos1;
                        lPos2.Caption := Pos2;
                        lPos3.Caption := Pos3;
                        lPos4.Caption := Pos4;
                        lPos5.Caption := Pos5;
                        lPos6.Caption := Pos6;
                end;
        end;
        // Î‡Ô‡Ì ÎÂÔÂÒÚÍÓ‚˚È “ﬁÀ‹œ¿Õ-2-1090*1090-Õ-0
        if Kod = 400 then
        begin
                Res := Pos(' ', Izdel);
                Delete(Izdel, 1, Res);
                //========================================
                Res := Pos(' ', Izdel);
                Delete(Izdel, 1, Res);
                //========================================“ﬁÀ‹œ¿Õ
                Res := Pos('-', Izdel);
                Nam := Copy(Izdel, 1, Res - 1);
                Delete(Izdel, 1, Res);
                //======================================== 2
                Res := Pos('-', Izdel);
                Pos1 := Copy(Izdel, 1, Res - 1);
                Delete(Izdel, 1, Res);
                //========================================
                Res := Pos('*', Izdel);
                A := Copy(Izdel, 1, Res - 1);
                Delete(Izdel, 1, Res);
                //========================================
                Res := Pos('-', Izdel);
                B := Copy(Izdel, 1, Res - 1);
                Delete(Izdel, 1, Res);
                //========================================
                Res := Pos('-', Izdel);
                Pos2 := Copy(Izdel, 1, Res - 1);
                Delete(Izdel, 1, Res);
                //========================================
                Res := Pos('-', Izdel);
                Pos3 := Copy(Izdel, 1, Res - 1);
                Delete(Izdel, 1, Res);
                //========================================
                Res := Pos('-', Izdel);
                Pos4 := Copy(Izdel, 1, Res - 1);
                Delete(Izdel, 1, Res);
                //========================================
                Res := Pos('-', Izdel);
                Pos5 := Copy(Izdel, 1, Res - 1);
                Delete(Izdel, 1, Res);
                //========================================
                Pos6 := Izdel;
                //========================================
                Name.Caption := Nam;
                Al.Caption := A;
                Bl.Caption := B;
                lPos1.Caption := Pos1;
                lPos2.Caption := Pos2;
                lPos3.Caption := Pos3;
                lPos4.Caption := Pos4;
                lPos5.Caption := Pos5;
                lPos6.Caption := Pos6;
        end;

        if Kod = 530 then
        begin
                Pos_Zv := Pos('*', Izdel);
                if Pos_Zv <> 0 then
                begin
                        // Î‡Ô‡Ì √≈–Ã» -œ-475*1707-Õ-F24-S_NF24A-S2-1
                        Res := Pos(' ', Izdel);
                        Delete(Izdel, 1, Res);
                        //========================================√≈–Ã» 
                        Res := Pos('-', Izdel);
                        Nam := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //======================================== œ
                        Res := Pos('-', Izdel);
                        Pos1 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('*', Izdel);
                        A := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        B := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //======================================== H
                        Res := Pos('-', Izdel);
                        Pos2 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //======================================== F24
                        Res := Pos('-', Izdel);
                        Pos3 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //======================================== F24
                        Res := Pos('-', Izdel);
                        Pos4 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //======================================== F24
                        Res := Pos('-', Izdel);
                        Pos5 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Pos6 := Izdel;
                        //========================================
                        Name.Caption := Nam;
                        Al.Caption := A;
                        Bl.Caption := B;
                        lPos1.Caption := Pos1;
                        lPos2.Caption := Pos2;
                        lPos3.Caption := Pos3;
                        lPos4.Caption := Pos4;
                        lPos5.Caption := Pos5;
                        lPos6.Caption := Pos6;
                end
                else
                begin
                        // Î‡Ô‡Ì √≈–Ã» -—-0625-0875-Õ-œ-12-00-00-”2 ‰Îˇ  ÓÌ‰ËˆËÓÌÂ  ÷ œ-—1-5
                        Res := Pos(' ', Izdel);
                        Delete(Izdel, 1, Res);
                        //========================================√≈–Ã» 
                        Res := Pos('-', Izdel);
                        Nam := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //======================================== C
                        Res := Pos('-', Izdel);
                        Pos1 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        A := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        B := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================H
                        Res := Pos('-', Izdel);
                        Pos2 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        Pos3 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        Pos4 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        Pos5 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos('-', Izdel);
                        Pos6 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Res := Pos(' ', Izdel);
                        Pos7 := Copy(Izdel, 1, Res - 1);
                        Delete(Izdel, 1, Res);
                        //========================================
                        Name.Caption := Nam;
                        Al.Caption := A;
                        Bl.Caption := B;
                        lPos1.Caption := Pos1;
                        lPos2.Caption := Pos2;
                        lPos3.Caption := Pos3;
                        lPos4.Caption := Pos4;
                        lPos5.Caption := Pos5;
                        lPos6.Caption := Pos6;
                        lPos7.Caption := Pos7;
                end;
        end;

        // Î‡Ô‡Ì ”¬ -0510-1050-12-00-”2 ‰Îˇ  ÓÌ‰ËˆËÓÌÂ  ÷ œ-√2-6,3
        if Kod = 525 then
        begin
                Res := Pos(' ', Izdel);
                Delete(Izdel, 1, Res);
                //======================================== ”¬ 
                Res := Pos('-', Izdel);
                Nam := Copy(Izdel, 1, Res - 1);
                Delete(Izdel, 1, Res);
                //========================================
                Res := Pos('-', Izdel);
                A := Copy(Izdel, 1, Res - 1);
                Delete(Izdel, 1, Res);
                ;
                //========================================
                Res := Pos('-', Izdel);
                B := Copy(Izdel, 1, Res - 1);
                Delete(Izdel, 1, Res);
                //========================================
                Res := Pos('-', Izdel);
                Pos1 := Copy(Izdel, 1, Res - 1);
                Delete(Izdel, 1, Res);
                //========================================
                Res := Pos('-', Izdel);
                Pos2 := Copy(Izdel, 1, Res - 1);
                Delete(Izdel, 1, Res);
                //========================================
                Res := Pos('-', Izdel);
                Pos3 := Copy(Izdel, 1, Res - 1);
                Delete(Izdel, 1, Res);
                //========================================
                Pos4 := Izdel;
                //========================================
                Name.Caption := Nam;
                Al.Caption := A;
                Bl.Caption := B;
                lPos1.Caption := Pos1;
                lPos2.Caption := Pos2;
                lPos3.Caption := Pos3;
                lPos4.Caption := Pos4;
                lPos5.Caption := Pos5;
                lPos6.Caption := Pos6;
        end;
end;

procedure TFVozKl.FormShow(Sender: TObject);
begin
        Form1.Clear_StringGrid(SGL);
        Form1.Clear_StringGrid(SGR);
end;

end.
