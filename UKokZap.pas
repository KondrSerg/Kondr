unit UKokZap;

interface

uses
        Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls,
        Forms,
        Dialogs, StdCtrls, Grids, ExtCtrls, ComCtrls;

type
        TFKolZap = class(TForm)
                Panel1: TPanel;
                Panel2: TPanel;
                Button1: TButton;
                Button2: TButton;
                Label2: TLabel;
                Edit1: TEdit;
                Label1: TLabel;
                Label3: TLabel;
                Label4: TLabel;
                Label5: TLabel;
                Label6: TLabel;
                Panel3: TPanel;
                SGS: TStringGrid;
                Button3: TButton;
                DateTimePicker1: TDateTimePicker;
                Label7: TLabel;
                procedure Button1Click(Sender: TObject);
                procedure FormShow(Sender: TObject);
                procedure Button3Click(Sender: TObject);
                procedure Button2Click(Sender: TObject);
                procedure SGSSelectCell(Sender: TObject; ACol, ARow: Integer;
                        var CanSelect: Boolean);
                procedure SGSSetEditText(Sender: TObject; ACol, ARow: Integer;
                        const Value: string);
        private
                { Private declarations }
        public
                { Public declarations }
        end;

var
        FKolZap: TFKolZap;

implementation

uses Main;

{$R *.dfm}

procedure TFKolZap.Button1Click(Sender: TObject);
var
        Kol, Kol_Zap, Kol_Zap_R, Svob, I: Integer;
        Nam, Str, Nom, Zak: string;
begin
        if (DateTimePicker1.Visible = True) and (Label1.Caption <> '') then
        begin
                Str := FormatDateTime('mm.dd.yyyy', DateTimePicker1.Date);
                Nom := Label1.Caption;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                        'UPDATE %s SET [ОТК]=' + #39 +
                        Str + #39 +
                        ' WHERE ([' + FN_NoM + ']=' + #39 + Nom + #39 + ')',
                        ['Запуск']) then
                        Exit;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                        'UPDATE %s SET [ОТК]=' + #39 +
                        Str + #39 +
                        ' WHERE ([' + FN_NoM + ']=' + #39 + Nom + #39 + ')',
                        ['ЗапускВозд']) then
                        Exit;
                Kol := 0;
                Kol_Zap := 0;
                for i := 0 to SGS.RowCount - 1 do
                begin
                        Nam := SGS.Cells[3, i + 1];
                        Zak := SGS.Cells[0, I + 1];
                        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                                'Select * from %s Where   (Изделие=' + #39 + Nam
                                + #39 + ') And (Заказ=' +
                                #39 + Zak + #39 + ') Order By Заказ ',
                                ['Запуск']) then
                                exit;
                        if Form1.ADOQuery2.RecordCount <> 0 then
                        begin
                                Kol :=
                                        StrToInt(Form1.ADOQuery2.FieldByName('Кол во').AsString);
                                Kol_Zap := Kol_Zap +
                                        StrToInt(Form1.ADOQuery2.FieldByName('Кол во запущенных').AsString);

                        end;
                end;
                if (Kol_Zap = Kol) and (Kol <> 0) then
                begin
                        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                                'UPDATE %s SET [ОТК]=' + #39 +
                                Str + #39 +
                                ' WHERE ([' + FN_NoM + ']=' + #39 + Nom + #39 +
                                ')', ['Klapana']) then
                                Exit;
                end;
                Kol := 0;
                Kol_Zap := 0;
                for i := 0 to SGS.RowCount - 1 do
                begin
                        Nam := SGS.Cells[3, i + 1];
                        Zak := SGS.Cells[0, I + 1];
                        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                                'Select * from %s Where   (Изделие=' + #39 + Nam
                                + #39 + ') And (Заказ=' +
                                #39 + Zak + #39 + ') Order By Заказ ',
                                ['ЗапускВозд']) then
                                exit;
                        if Form1.ADOQuery2.RecordCount <> 0 then
                        begin
                                Kol :=
                                        StrToInt(Form1.ADOQuery2.FieldByName('Кол во').AsString);
                                Kol_Zap := Kol_Zap +
                                        StrToInt(Form1.ADOQuery2.FieldByName('Кол во запущенных').AsString);

                        end;
                end;
                if (Kol_Zap = Kol) and (Kol <> 0) then
                begin
                        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                                'UPDATE %s SET [ОТК]=' + #39 +
                                Str + #39 +
                                ' WHERE ([' + FN_NoM + ']=' + #39 + Nom + #39 +
                                ')', ['KlapanaZap']) then
                                Exit;
                end;
        end;
end;

procedure TFKolZap.FormShow(Sender: TObject);
var
        I: Integer;
begin
        Form1.Clear_StringGrid(SGS);
        SGS.ColCount := 5;
        SGS.Cells[0, 0] := 'Заказ';
        SGS.Cells[1, 0] := 'Кол во запущенных';
        SGS.Cells[2, 0] := 'Клапан';
        SGS.Cells[3, 0] := 'Привод';
        SGS.Cells[4, 0] := 'Кол-во готовых';
        if Label1.Caption <> '' then
        begin
                if not Form1.mkQuerySelect(Form1.ADOQuery2,
                        'Select * from %s Where   (Номер=' + #39 + Label1.Caption
                        + #39 +
                        ') Order By Заказ ', ['Запуск']) then
                        exit;
                if Form1.ADOQuery2.RecordCount <> 0 then
                begin
                        SGS.RowCount := Form1.ADOQuery2.RecordCount + 1;
                        for i := 0 to Form1.ADOQuery2.RecordCount - 1 do
                        begin
                                SGS.Cells[0, i + 1] :=
                                        Form1.ADOQuery2.FieldByName('Заказ').AsString;
                                SGS.Cells[1, i + 1] :=
                                        Form1.ADOQuery2.FieldByName('Кол во запущенных').AsString;
                                SGS.Cells[2, i + 1] :=
                                        Form1.ADOQuery2.FieldByName('Изделие').AsString;
                                SGS.Cells[3, i + 1] :=
                                        Form1.ADOQuery2.FieldByName('МодПривода').AsString;
                                //SGS.Cells[4,i+1]:=Form1.ADOQuery2.FieldByName('Кол-во готовых').AsString;
                                Form1.ADOQuery2.Next;
                        end;
                end;
        end;
end;

procedure TFKolZap.Button3Click(Sender: TObject);
var
        I: Integer;
begin
        Form1.Clear_StringGrid(SGS);
        SGS.ColCount := 5;
        SGS.Cells[0, 0] := 'Заказ';
        SGS.Cells[1, 0] := 'Кол во запущенных';
        SGS.Cells[2, 0] := 'Клапан';
        SGS.Cells[3, 0] := 'Привод';
        SGS.Cells[4, 0] := 'Кол-во готовых';
        Label1.Caption := Edit1.Text;
        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                'Select * from %s Where   (Номер='
                + #39 + Edit1.Text + #39 + ') Order By Заказ ', ['Запуск']) then
                exit;
        if Form1.ADOQuery2.RecordCount <> 0 then
        begin
                SGS.RowCount := Form1.ADOQuery2.RecordCount + 1;
                for i := 0 to Form1.ADOQuery2.RecordCount - 1 do
                begin

                        SGS.Cells[0, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Заказ').AsString;
                        SGS.Cells[1, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Кол во запущенных').AsString;
                        SGS.Cells[2, i + 1] :=
                                Form1.ADOQuery2.FieldByName('Изделие').AsString;
                        SGS.Cells[3, i + 1] :=
                                Form1.ADOQuery2.FieldByName('МодПривода').AsString;
                        //SGS.Cells[4,i+1]:=Form1.ADOQuery2.FieldByName('Кол-во готовых').AsString;
                        Form1.ADOQuery2.Next;
                end;
        end;
end;

procedure TFKolZap.Button2Click(Sender: TObject);
begin
        Close;
end;

procedure TFKolZap.SGSSelectCell(Sender: TObject; ACol, ARow: Integer;
        var CanSelect: Boolean);
begin
        if ACol = 4 then
        begin
                SGS.Options := SGS.Options + [goEditing];

        end
        else
                SGS.Options := SGS.Options - [goEditing];
end;

procedure TFKolZap.SGSSetEditText(Sender: TObject; ACol, ARow: Integer;
        const Value: string);
var
        Kol, Kol_Got, I: Integer;

begin
        if SGS.Cells[4, ARow] = '' then
                Kol_Got := 0
        else
                Kol_Got := StrToInt(SGS.Cells[4, ARow]);
        if SGS.Cells[1, ARow] = '' then
                Kol := 0
        else
                Kol := StrToInt(SGS.Cells[1, ARow]);
        i := Kol - Kol_Got;
        if I < 0 then
        begin
                MessageDlg('Их всего ' + SGS.Cells[4, ARow] + ' !', mtError,
                        [mbOk], 0);
                SGS.Cells[4, ARow] := '';
                Exit;
        end;
end;

end.
