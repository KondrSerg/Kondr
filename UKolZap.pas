unit UKolZap;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Grids;

type
  TForm5 = class(TForm)
    pnl1: TPanel;
    mmo1: TMemo;
    mmo2: TMemo;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    btn1: TButton;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    SpecGrid1: TStringGrid;
    mmo3: TMemo;
    mmo4: TMemo;
    LabKO1: TLabel;
    LabKO2: TLabel;
    procedure FormShow(Sender: TObject);
    procedure btn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form5: TForm5;

implementation

uses Main;

{$R *.dfm}

procedure TForm5.FormShow(Sender: TObject);
var Kol,Br:string;
Res1,I,Kol_I:Integer;
begin
        mmo3.Lines.Clear;
        mmo4.Lines.Clear;
        if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s WHERE (IDГП='+#39+Label1.Caption+#39+') AND (L=0) ', ['СТАМ']) then
                exit;
        Br:=Form1.ADOQuery1.FieldByName('Брикет').AsString;
        Delete(Br,1,2);
        for I:=0 to 30  do
        begin
                Res1:=Pos(',',Br);
                If Res1<>0 then
                begin
                  mmo3.Lines.Add(Copy(Br,1,Res1-1));
                  Delete(Br,1,Res1);
                end
                else
                begin
                        mmo3.Lines.Add(Br);
                        Break;
                end;
        end;

        Kol:=Form1.ADOQuery1.FieldByName('Кол Раз').AsString;
        Delete(Kol,1,2);
        for I:=0 to 30  do
        begin
                Res1:=Pos(',',Kol);
                If Res1<>0 then
                begin
                  mmo4.Lines.Add(Copy(Kol,1,Res1-1));
                  Delete(Kol,1,Res1);
                end
                else
                begin
                        mmo4.Lines.Add(Kol);
                        Break;
                end;
        end;
        SpecGrid1.Cells[0,0]:='Расчетная дата';
        SpecGrid1.Cells[1,0]:='Кол-во в брикете';
        //SpecGrid1.Cells[2,0]:='Всего ';
        if mmo3.Lines.Count<>0 then
        SpecGrid1.RowCount:=mmo3.Lines.Count+1;
        for I:=0 to mmo3.Lines.Count do
        begin
                if mmo4.Lines.Strings[i]<>'' then
                        Kol_I:=Kol_i+StrToInt(mmo4.Lines.Strings[i]);
                SpecGrid1.Cells[0,I+1]:=mmo3.Lines.Strings[i];
                SpecGrid1.Cells[1,I+1]:=mmo4.Lines.Strings[i];
        end;
        LabKO2.Caption:=IntToStr(Kol_I);
        mmo1.Text:=Br;
        mmo2.Text:=Kol;
end;

procedure TForm5.btn1Click(Sender: TObject);
begin
        Close;
end;

end.
