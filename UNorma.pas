unit UNorma;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, UNewNakl, ComCtrls, StdCtrls, Grids, Buttons, ExtCtrls;

type
  TFNewNakl5 = class(TFNewNakl)
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FNewNakl5: TFNewNakl5;

implementation

{$R *.dfm}

procedure TFNewNakl5.FormShow(Sender: TObject);
begin
        SG2.Cells[0, 0] := 'Заказ';
        SG2.Cells[1, 0] := 'Расчетная дата';
        SG2.Cells[2, 0] := 'Кол во запущенных';
        SG2.Cells[3, 0] := 'Номер';
        SG2.Cells[4, 0] := 'ID';
        SG2.Cells[5, 0] := 'IDgp';

end;

end.
