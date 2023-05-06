unit USpecOb;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Grids, AdvObj, BaseGrid, AdvGrid, Menus,Clipbrd,
  ActnList, System.Actions;

type
  TFSpecOb = class(TForm)
    SpecGrid: TAdvStringGrid;
    Memo3: TMemo;
    pnl1: TPanel;
    Label1: TLabel;
    pm1: TPopupMenu;
    NNCopy1: TMenuItem;
    Memo1: TMemo;
    Memo2: TMemo;
    Memo4: TMemo;
    actlst1: TActionList;
    act1: TAction;
    procedure NNCopy1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure SpecGridDblClick(Sender: TObject);
    procedure SpecGridDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure act1Execute(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FSpecOb: TFSpecOb;

implementation

{$R *.dfm}

procedure TFSpecOb.NNCopy1Click(Sender: TObject);
var
        Str1: string;
begin
        Str1 := SpecGrid.Cells[SpecGrid.Col, SpecGrid.Row];
        Memo1.Lines.Clear;
        Memo1.Lines.Add(Str1);
        Clipboard.AsText := Memo1.Lines.Strings[0];
end;

procedure TFSpecOb.FormShow(Sender: TObject);
Var
S:String;
I:Integer;
begin
      S := ExtractFileDir(ParamStr(0));
      Memo2.Lines.Clear;
      Memo2.Lines.LoadFromFile(S + '\SpecGrid1.ini');
     // SpecGrid.ColCount := Memo2.Lines.Count;
      for I := 0 to Memo2.Lines.Count - 1 do
        SpecGrid.ColWidths[i] := StrToInt(Memo2.Lines.Strings[i]); 
end;

procedure TFSpecOb.FormClose(Sender: TObject; var Action: TCloseAction);
Var
S:String;
I:Integer;
begin
      S := ExtractFileDir(ParamStr(0));
      Memo2.Lines.Clear;
      for I := 0 to SpecGrid.ColCount - 1 do
                Memo2.Lines.Add(IntToStr(SpecGrid.ColWidths[i]));
      Memo2.Lines.SaveToFile(S + '\SpecGrid1.ini');
      //+++++++++++++++++++++++++++++++++++++
end;

procedure TFSpecOb.SpecGridDblClick(Sender: TObject);
begin
       SpecGrid.Cells[0,SpecGrid.Row]:='0';
end;

procedure TFSpecOb.SpecGridDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
  Var Res,Kol:Integer;
begin
        Kol:=  SpecGrid.ColCount-1;
        Res:=AnsiCompareStr('0',SpecGrid.Cells[0,ARow]);
        if Res=0 Then
         SpecGrid.ColorRect(0,ARow,Kol,ARow,clYellow);
end;

procedure TFSpecOb.act1Execute(Sender: TObject);
begin
        if Memo4.Visible = False then
        begin
                Memo4.Visible := True;
                Memo2.Visible := True;
        end
        else
        begin
                Memo4.Visible := False;
                Memo2.Visible := False;
        end;
end;

end.
