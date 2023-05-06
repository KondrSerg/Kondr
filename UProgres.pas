unit UProgres;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, AdvWiiProgressBar, AdvCircularProgress, AdvProgressBar;

type
  TFProgres = class(TForm)
    APB1: TAdvProgressBar;
    ACP1: TAdvCircularProgress;
    AWP1: TAdvWiiProgressBar;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FProgres: TFProgres;

implementation



{$R *.dfm}

end.
