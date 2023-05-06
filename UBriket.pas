unit UBriket;

interface

uses
        Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls,
        Forms,
        Dialogs, StdCtrls, ExtCtrls, GradientLabel;

type
        TFBriket = class(TForm)
                Label3: TLabel;
                Panel2: TPanel;
                Button1: TButton;
                Edit3: TEdit;
                Label4: TLabel;
                Label1: TLabel;
    Button2: TButton;
    GradientLabel1: TGradientLabel;
    GradientLabel2: TGradientLabel;
                procedure Button1Click(Sender: TObject);
                procedure Clear_Briket_IZ_MaskEdit(Str: string; var Br, Br1,
                        Br2: string);
    procedure FormShow(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Edit3KeyPress(Sender: TObject; var Key: Char);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure Button2KeyPress(Sender: TObject; var Key: Char);
        private
                { Private declarations }
        public
                { Public declarations }
        end;

var
        FBriket: TFBriket;

implementation

uses Main;

{$R *.dfm}

procedure TFBriket.Clear_Briket_IZ_MaskEdit(Str: string; var Br, Br1, Br2:
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

procedure TFBriket.Button1Click(Sender: TObject);
var
        Briket1, Briket2, Briket, Br, Br1, Br2, IDGP, Nom, Zak,Nam: string;
        I, Kol, KOl_ZAP: Integer;
begin

                Nam := Form1.ZSG.Cells[I_FN_KOL_ZAP + 7, Form1.ZSG.row];
                Zak := Form1.ZSG.Cells[I_FN_ZAK, Form1.ZSG.row];
                Nom := Form1.ZSG.Cells[0, Form1.ZSG.row];
                if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                        'UPDATE %s SET ['+Label4.Caption+']=' + #39 +
                        Edit3.Text + #39 +
                        ' WHERE ([' + FN_NAM + ']=' + #39 + Nam + #39 + ') AND ([' + FN_ZAK + ']=' + #39 + Zak + #39 + ')AND ([' + FN_NOM + ']=' + #39 + Nom + #39 + ')',
                        ['Запуск']) then
                        Exit;
                Form1.ZSG.Cells[StrToInt(Label1.Caption), Form1.ZSG.row]:=Edit3.Text;
        Close;
end;

procedure TFBriket.FormShow(Sender: TObject);
begin
    Edit3.Text:='';
    Edit3.SetFocus;
end;

procedure TFBriket.Button2Click(Sender: TObject);
begin
     Close;

end;

procedure TFBriket.Edit3KeyPress(Sender: TObject; var Key: Char);
begin
     If Key=#27 Then
     Close;
        if (Key = #13) then
                Button1Click(nil);
end;

procedure TFBriket.FormKeyPress(Sender: TObject; var Key: Char);
begin
     If Key=#27 Then
     Close;
        if (Key = #13) then
                Button1Click(nil);
end;

procedure TFBriket.Button2KeyPress(Sender: TObject; var Key: Char);
begin
     If Key=#27 Then
     Close;
end;

end.

