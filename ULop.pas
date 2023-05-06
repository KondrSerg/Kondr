unit ULop;

interface

uses
        Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls,
        Forms,
        Dialogs, Grids, ExtCtrls, StdCtrls;

type
        TFLop = class(TForm)
                Panel1: TPanel;
                StringGrid1: TStringGrid;
                Label1: TLabel;
                Panel2: TPanel;
                Label2: TLabel;
                StringGrid2: TStringGrid;
                Panel3: TPanel;
                Label3: TLabel;
                StringGrid3: TStringGrid;
        private
                { Private declarations }
        public
                { Public declarations }
        end;

var
        FLop: TFLop;

implementation

{$R *.dfm}

end.
