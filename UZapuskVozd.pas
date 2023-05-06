unit UZapuskVozd;

interface

uses
        Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls,
        Forms,
        Dialogs, StdCtrls, Grids, ComCtrls, ExtCtrls;

type
        TFZapuskVozd = class(TForm)
                Panel2: TPanel;
                Label3: TLabel;
                Label4: TLabel;
                Label1: TLabel;
                Label5: TLabel;
                Label6: TLabel;
                Label2: TLabel;
                DateTimePicker1: TDateTimePicker;
                Panel3: TPanel;
                Label7: TLabel;
                Label8: TLabel;
                SG2: TStringGrid;
                Button3: TButton;
                Edit1: TEdit;
                Edit2: TEdit;
                Panel1: TPanel;
                SG1: TStringGrid;
                Button1: TButton;
                Button2: TButton;
                Memo1: TMemo;
                Label11: TLabel;
                Label9: TLabel;
                Label15: TLabel;
                procedure FormShow(Sender: TObject);
                procedure FormClose(Sender: TObject; var Action: TCloseAction);
                procedure Button3Click(Sender: TObject);
                procedure Button1Click(Sender: TObject);
                procedure Button2Click(Sender: TObject);
                procedure SG2DblClick(Sender: TObject);
                procedure DateTimePicker1Enter(Sender: TObject);
                procedure SG2SelectCell(Sender: TObject; ACol, ARow: Integer;
                        var CanSelect: Boolean);
        private
                { Private declarations }
        public
                { Public declarations }
        end;

var
        FZapuskVozd: TFZapuskVozd;

implementation

uses Main;

{$R *.dfm}

procedure TFZapuskVozd.FormShow(Sender: TObject);
var
        I, Kol, J, y, Res, Nom1, Nom2: Integer;
        Nom_Poz, Zak, IDGP, Kol_Zap, Kol_Raz: string;
begin

        Button1.Enabled := False;
        SG1.RowCount := 2;
        Form1.Clear_StringGrid(SG1);
        DateTimePicker1.DateTime:=Now;
        SG2.Cells[0, 0] := 'Заказ';
        SG2.Cells[1, 0] := 'Расчетная дата';
        SG2.Cells[2, 0] := 'Кол во запущенных';
        SG2.Cells[3, 0] := 'Номер';
        SG2.Cells[4, 0] := 'ID';
        SG2.Cells[5, 0] := 'IDГП';
        SG2.Cells[6, 0] := 'IDКО';
        SG2.Cells[7, 0] := 'БЗ';
        SG2.ColWidths[0] := 0;
        SG2.ColWidths[2] := 0;
        SG2.ColWidths[4] := 0;
        Edit1.Text := '0';
        Memo1.Lines.Clear;
        Nom2 := 0;
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                'Select * from %s Where  ( [Номер]<>0) AND ([Заготовка Запуск] IS NULL)  Order By Номер',
                ['ЗапускВозд']) then
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
                        SG2.Cells[4, j] :=Form1.ADOQuery2.FieldByName('ID').AsString;
                        SG2.Cells[5, j] :=Form1.ADOQuery2.FieldByName('IDГП').AsString;
                        SG2.Cells[6, j] :=Form1.ADOQuery2.FieldByName('IDКО').AsString;
                        SG2.Cells[7, j] :=Form1.ADOQuery2.FieldByName('БЗ').AsString;
                        Nom_Poz := Form1.ADOQuery2.FieldByName('№Поз').AsString;
                        Inc(y);
                        Inc(j);
                        Form1.ADOQuery2.Next;
                end;
        end;
        SG2.RowCount := Y + 2;
end;

procedure TFZapuskVozd.FormClose(Sender: TObject;
        var Action: TCloseAction);
begin

        Label11.Caption := '0';
        Label9.Caption := '0';
        Label1.Caption := '0';
end;

procedure TFZapuskVozd.Button3Click(Sender: TObject);
var
        I, Kol, Y, j, Pos, Pos1, x: Integer;
        Nom_Poz, Zak, IDGP, Briket, Briket1, Briket2, Br, Br1, Br2, Kol_Zap:
        string;
begin
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
        {if not Form1.mkQuerySelect( Form1.ADOQuery2, 'Select * from %s Where  (Номер='+#39+Label15.Caption+#39+') AND ([Кол во]=[Кол во запущенных] )',['Klapana'] )
        then exit;
        SG1.RowCount:=Form1.ADOQuery2.RecordCount+1;
        if Form1.ADOQuery2.RecordCount<>0 Then
        Begin
                For I:=0 To Form1.ADOQuery2.RecordCount-1 Do
                Begin

                        Kol_Zap:= Form1.ADOQuery2.FieldByName('Кол во запущенных').AsString;
                        if Kol_Zap='' Then
                        Kol_Zap:='0';

                        Summa:= Summa+ StrToInt(Kol_Zap);
                        SG1.Cells[0,i+1]:=Form1.ADOQuery2.FieldByName('Заказ').AsString;
                        SG1.Cells[1,i+1]:=Form1.ADOQuery2.FieldByName('Кол во запущенных').AsString;
                        SG1.Cells[2,i+1]:=Form1.ADOQuery2.FieldByName('Номер').AsString;
                        SG1.Cells[3,i+1]:=Form1.ADOQuery2.FieldByName('ID').AsString;
                        SG1.Cells[4,i+1]:=Form1.ADOQuery2.FieldByName('Изделие').AsString;
                        InC(J);
                        Form1.ADOQuery2.Next;
                End;
        Label9.Caption:=IntToStr(Summa);
        Label11.Caption:=IntToStr(Kol-Summa);
        End
        Else
        Begin
        Label9.Caption:=IntToStr(0);
        Label11.Caption:=IntToStr(Kol);
        End; }//  ([Заказ]= '+#39+Zak+#39+')AND
        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                'Select * from %s Where   (Номер='
                + #39 + Label15.Caption + #39 + ') ', ['ЗапускВозд']) then
                exit;
        X := SG1.RowCount;
        if Form1.ADOQuery2.RecordCount <> 0 then
        begin // J+
                SG1.RowCount := Form1.ADOQuery2.RecordCount + 1;
                // j                              //+J
                for i := 0 to Form1.ADOQuery2.RecordCount - 1 do
                begin
                        {Pos:=Form1.ADOQuery2.FieldByName(FN_POS).AsInteger;
                        for y:=1 To X Do
                        Begin
                              if SG1.Cells[2,y]='' Then
                                Pos1:=0
                              Else
                                Pos1:=StrToInt(SG1.Cells[2,y]);
                              if Pos<>Pos1 Then
                              Begin }
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
                        // End;
                  // End;
                        Form1.ADOQuery2.Next;
                end;
        end;
end;

procedure TFZapuskVozd.Button1Click(Sender: TObject);
var
        Kol_ZAP, Kol, Kol_Zap_Ranee, Res, Nomer, Zak_Int, i, A, B, D, y, F_KPU,
                F_KPD,
                x, Res1, e: Integer;
        Zak, Dat, Plan_Dat, Vn_Dat, Nom, Pos_Vst, Pos_Ml, R, Pereh, Privod:
        string;
        Svar, Sbor, Izdel, Pos1, Pos2, Pos3, Pos4, Pos5, Pos_Isp, Pos_Flan,
                Pos_Flan1,
                Pos_Privod, Pos_Sn, Pos_Dop, Pos_Ram, Dir,S,God: string;
        Dat1, Dat2, Dat3: TDate;
        XL, XL1, XL2: Variant;
begin

        F_KPU := 0;
        F_KPD := 0;
        y := 1;
        X := 1;
        e := 1;
        Vn_Dat := FormatDateTime('mm.dd.yyyy', DateTimePicker1.Date);
        S := FormatDateTime('dd.mm.yyyy', DateTimePicker1.Date);
        God := FormatDateTime('yyyy', Now);
        Dir := ExtractFileDir(ParamStr(0));
        Izdel := SG1.Cells[4, i];
        for i := 1 to SG1.RowCount do
        begin

                if SG1.Cells[2, i] <> '' then
                begin
                        Nom := SG1.Cells[2, i];
                        Izdel := SG1.Cells[4, i];
                        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                                'UPDATE %s SET [Заготовка Запуск]=' + #39 +
                                Vn_Dat + #39 +
                                ' WHERE ([Номер]=' + #39 + SG1.Cells[2, i] + #39
                                + ') ', ['KlapanaZap']) then
                                Exit;
                        //++++++++++++++++++++++++++++++++++++++++++++++
                        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                                'UPDATE %s SET [Заготовка Запуск]=' + #39 +
                                Vn_Dat + #39 +
                                ' WHERE ([ID]=' + #39 + SG1.Cells[3, i] + #39 +
                                ') AND ([Номер]=' + #39
                                + SG1.Cells[2, i] + #39 + ') ', ['ЗапускВозд'])
                                        then
                                Exit;
                        //++++++++++++++++++++++++++++++++++++++++++++++
                        if SG1.Cells[6, i] = '' then
                                Svar := '0'
                        else
                                Svar := (SG1.Cells[5, i]);
                        if SG1.Cells[5, i] = '' then
                                Sbor := '0'
                        else
                                Sbor := (SG1.Cells[6, i]);
                        Res := Pos(',', Svar);
                        Delete(Svar, Res, 1);
                        if Res <> 0 then
                                Insert('.', Svar, Res);

                        Res := Pos(',', Sbor);
                        Delete(Sbor, Res, 1);
                        if Res <> 0 then
                                Insert('.', Sbor, Res);
                        Privod := SG1.Cells[7, i];
                        //++++++++++++++++++++++++++++++++++++++++++++++

                end;
        end;

end;

procedure TFZapuskVozd.Button2Click(Sender: TObject);
begin
        Close;
end;

procedure TFZapuskVozd.SG2DblClick(Sender: TObject);
begin
        Label15.Caption := SG2.Cells[3, SG2.Row];
        Button3.Click;
end;

procedure TFZapuskVozd.DateTimePicker1Enter(Sender: TObject);
begin
        Button1.Enabled := True;
end;

procedure TFZapuskVozd.SG2SelectCell(Sender: TObject; ACol, ARow: Integer;
        var CanSelect: Boolean);
begin
        Label15.Caption := SG2.Cells[3, ARow];
        Button3.Click;
end;

end.
