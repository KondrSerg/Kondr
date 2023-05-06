unit UNewSbor;

interface

uses
        Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls,
        Forms,
        Dialogs, ExtCtrls, Grids, StdCtrls, Menus;

type
        TFNewSbor = class(TForm)
                SGSB: TStringGrid;
                Panel1: TPanel;
                Panel2: TPanel;
                Button1: TButton;
                Button2: TButton;
                PopupMenu1: TPopupMenu;
                N1: TMenuItem;
                N2: TMenuItem;
                procedure FormShow(Sender: TObject);
                procedure N1Click(Sender: TObject);
                procedure N2Click(Sender: TObject);
                procedure Button1Click(Sender: TObject);
                procedure Button2Click(Sender: TObject);
    procedure SGSBKeyPress(Sender: TObject; var Key: Char);
        private
                { Private declarations }
        public
                { Public declarations }
        end;

var
        FNewSbor: TFNewSbor;

implementation

uses Main;

{$R *.dfm}

procedure TFNewSbor.FormShow(Sender: TObject);
var
        I: Integer;
begin
        Form1.Clear_StringGrid(SGSB);
        SGSB.Cells[0, 0] := '�';
        SGSB.Cells[1, 0] := '�������';
        SGSB.Cells[2, 0] := '�������';
        SGSB.Cells[3, 0] := '���� ��������';
        SGSB.Cells[4, 0] := '���������';
        SGSB.Cells[5, 0] := '���������';
        if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s  ',
                ['�������']) then
                exit;
        SGSB.RowCount := Form1.ADOQuery1.RecordCount + 1;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
                SGSB.Cells[0, i + 1] := IntToStr(I + 1);
                SGSB.Cells[1, i + 1] :=
                        Form1.ADOQuery1.FieldByName('�������').AsString;
                SGSB.Cells[2, i + 1] :=
                        Form1.ADOQuery1.FieldByName('�������').AsString;
                SGSB.Cells[3, i + 1] :=
                        Form1.ADOQuery1.FieldByName('���� ��������').AsString;
                SGSB.Cells[4, i + 1] :=
                        Form1.ADOQuery1.FieldByName('���������').AsString;
                SGSB.Cells[5, i + 1] :=
                        Form1.ADOQuery1.FieldByName('���������').AsString;
                Form1.ADOQuery1.Next;
        end;
end;

procedure TFNewSbor.N1Click(Sender: TObject);
begin
        SGSB.RowCount := SGSB.RowCount + 1;
        Form1.FSelectRow(SGSB, SGSB.RowCount - 1);
end;

procedure TFNewSbor.N2Click(Sender: TObject);
begin
        Form1.DeleteARow(SGSB, SGSB.Row);
end;

procedure TFNewSbor.Button1Click(Sender: TObject);
begin
        Close;
end;

procedure TFNewSbor.Button2Click(Sender: TObject);
var
        I: Integer;
begin
        if not Form1.mkQueryDelete(Form1.ADOQuery1, 'DELETE FROM %s  ',
                ['�������']) then
                Exit;
        for I := 1 to SGSB.RowCount - 1 do
        begin
                if not Form1.mkQueryInsert(Form1.ADOQuery1, 'Insert Into %s ' +
                        '(�������,�������,[���� ��������],���������,��������� ) ' +
                        'Values (%s,%s,%s,%s,%s)', ['�������', #39 + SGSB.Cells[1, i] +
                        #39, #39 + SGSB.Cells[2, i] + #39,#39 + SGSB.Cells[3, i] + #39,#39 + SGSB.Cells[4, i] + #39
                        ,#39 + SGSB.Cells[5, i] + #39] ) then
                        exit;

        end;

        Form1.ComboBox6.Items.Clear;
        Form1.ComboBox6.Items.Add('');
        Form1.ComboBox22.Items.Clear;
        Form1.ComboBox22.Items.Add('');
        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                'Select * from %s  Order by �������',
                ['�������']) then
                exit;
        for I := 0 to Form1.ADOQuery2.RecordCount - 1 do
        begin
                Form1.ComboBox6.Items.Add(Form1.ADOQuery2.FieldByName('�������').AsString);
                Form1.ComboBox22.Items.Add(Form1.ADOQuery2.FieldByName('�������').AsString);
                Form1.ADOQuery2.Next;
        end;
        Close;
end;

procedure TFNewSbor.SGSBKeyPress(Sender: TObject; var Key: Char);
begin
     If Key=#27 Then
     Close;
        if (Key = #13) then
                Button2Click(nil);
end;

end.
