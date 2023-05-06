unit UPDM;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, ExtCtrls;

type
  TFPDM = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    SG1: TStringGrid;
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FPDM: TFPDM;

implementation

uses Main;

{$R *.dfm}

procedure TFPDM.FormShow(Sender: TObject);
Var I:Integer;
begin                             // 
                        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                        'Select DISTINCT(Дата) from %s Where ((Статус=' + #39 +
                        'ВНЕС' + #39 + ') OR (Статус IS NULL)) AND (Х IS NULL)  AND (Отмена IS NULL)', ['Klapana']) then
                        exit;
                        SG1.RowCount:=Form1.ADOQuery1.RecordCount+1;
                        SG1.Cells[0,0]:='№';
                        SG1.Cells[1,0]:='Дата Задания';
                        For I:=0 To Form1.ADOQuery1.RecordCount-1 Do
                        Begin
                              SG1.Cells[1,I+1]:=IntToStr(I+1);
                              SG1.Cells[1,I+1]:=Form1.ADOQuery1.FieldByName(FN_DAT).AsString;
                              Form1.ADOQuery1.Next;
                        end;
end;

end.
