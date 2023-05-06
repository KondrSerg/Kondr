unit UDetal;

interface

uses
        Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls,
        Forms,
        Dialogs, Grids, Menus, StdCtrls;

type
        TFDetal = class(TForm)
                StringGrid3: TStringGrid;
                PopupMenu1: TPopupMenu;
                True1: TMenuItem;
                False1: TMenuItem;
    Label1: TLabel;
    Edit1: TEdit;
    Button1: TButton;
    Button2: TButton;
    Label2: TLabel;
                procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
        private
                { Private declarations }
        public
                { Public declarations }
                KOz,Fl:Integer;
        end;

var
        FDetal: TFDetal;

implementation

uses Main, USpec;

{$R *.dfm}

procedure TFDetal.FormShow(Sender: TObject);
var
        I: Integer;
begin
        //If ComboBox2.Text<>'' Then
        //     SQLText1:='';
{        if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s ' +
                Form1.SQLText1, ['Специф']) then
                exit;
        StringGrid3.ColCount := I_FN_NOM2 + 1;
        StringGrid3.Cells[0, 0] := '№';
        StringGrid3.Cells[I_FN_POS_KOMP1, 0] := FN_POS_KOMP1;
        StringGrid3.Cells[I_FN_MATERIAL, 0] := FN_MATERIAL;
        StringGrid3.Cells[I_FN_MASSA, 0] := FN_MASSA;
        StringGrid3.Cells[I_FN_OBYM, 0] := FN_OBYM;
        StringGrid3.Cells[I_FN_DLINA, 0] := FN_DLINA;
        StringGrid3.Cells[I_FN_SHIR, 0] := FN_SHIR;
        StringGrid3.Cells[I_FN_DIAM, 0] := FN_DIAM;
        StringGrid3.Cells[I_FN_DLIN_RAZV, 0] := FN_DLIN_RAZV;
        StringGrid3.Cells[I_FN_SHIR_RAZV, 0] := FN_SHIR_RAZV;
        StringGrid3.Cells[I_FN_KOL_GIB, 0] := FN_KOL_GIB;
        StringGrid3.Cells[I_FN_KOL_RES, 0] := FN_KOL_RES;
        StringGrid3.Cells[I_FN_STAND, 0] := FN_STAND;
        StringGrid3.Cells[I_FN_Trumpf, 0] := FN_Trumpf;
        StringGrid3.Cells[I_FN_NOG, 0] := FN_NOG;
        StringGrid3.Cells[I_FN_GIBKA, 0] := FN_GIBKA;
        StringGrid3.Cells[I_FN_PROKAT, 0] := FN_PROKAT;
        StringGrid3.Cells[I_FN_ZAK2, 0] := FN_ZAK2;
        StringGrid3.Cells[I_FN_NOM2, 0] := FN_NOM2;

        if Form1.ADOQuery1.RecordCount = 0 then
        begin
                StringGrid3.RowCount := 1;
                Form1.Clear_StringGrid(StringGrid3);
                exit;
        end;
        StringGrid3.RowCount := Form1.ADOQuery1.RecordCount + 1;
        Form1.StatusBar1.Panels[2].Text :=
                IntToStr(Form1.ADOQuery1.RecordCount);
        Form1.ADOQuery1.First;
        for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
        begin
                StringGrid3.Cells[0, I + 1] := IntToStr(I + 1);
                StringGrid3.Cells[I_FN_POS_KOMP1, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_POS_KOMP1).AsString;
                StringGrid3.Cells[I_FN_MATERIAL, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_MATERIAL).AsString;
                StringGrid3.Cells[I_FN_MASSA, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_MASSA).AsString;
                StringGrid3.Cells[I_FN_OBYM, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_OBYM).AsString;
                StringGrid3.Cells[I_FN_DLINA, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_DLINA).AsString;
                StringGrid3.Cells[I_FN_SHIR, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_SHIR).AsString;
                StringGrid3.Cells[I_FN_DIAM, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_DIAM).AsString;
                StringGrid3.Cells[I_FN_DLIN_RAZV, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_DLIN_RAZV).AsString;
                StringGrid3.Cells[I_FN_SHIR_RAZV, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_SHIR_RAZV).AsString;
                StringGrid3.Cells[I_FN_KOL_GIB, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_KOL_GIB).AsString;
                StringGrid3.Cells[I_FN_KOL_RES, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_KOL_RES).AsString;
                StringGrid3.Cells[I_FN_STAND, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_STAND).AsString;
                StringGrid3.Cells[I_FN_Trumpf, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_Trumpf).AsString;
                StringGrid3.Cells[I_FN_NOG, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_NOG).AsString;
                StringGrid3.Cells[I_FN_GIBKA, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_GIBKA).AsString;
                StringGrid3.Cells[I_FN_PROKAT, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_PROKAT).AsString;
                StringGrid3.Cells[I_FN_ZAK2, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_ZAK2).AsString;
                StringGrid3.Cells[I_FN_NOM2, I + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_NOM2).AsString;
                //+++++++++++++++++++++++++++++++++++++++++++
                Form1.ADOQuery1.Next;
        end; }
end;

procedure TFDetal.Button1Click(Sender: TObject);
begin
  Close;
end;

procedure TFDetal.Button2Click(Sender: TObject);
begin
        if Form1.Spec=1 Then
        Begin
        if FSpec.Fl=0 Then
        Begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' + FDetal.Caption +
                ']=' +
                #39 + Edit1.Text + #39 +
                ' WHERE ([Элемент]=' + #39 + Label1.Caption + #39+') AND ([ОбозначениеШ]=' + #39 + Label2.Caption + #39+')',
                ['Шаблон']) then
                Exit;
                Close;
        end;

        if FSpec.Fl=1 Then
        Begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' + FDetal.Caption +
                ']=' +
                #39 + Edit1.Text + #39 +
                ' WHERE ([Элемент]=' + #39 + Label1.Caption + #39+') AND ([ОбозначениеШ]=' + #39 + Label2.Caption + #39+')',
                ['ШаблонВоз']) then
                Exit;
                Close;
        end;
        end;
    FSpec.SpecGrid.Cells[I_FN_NOM1+5, FSpec.SpecGrid.Row]:=Edit1.Text;

end;

end.
