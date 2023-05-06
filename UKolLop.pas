unit UKolLop;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TFKolLop = class(TForm)
    btn1: TButton;
    btn2: TButton;
    Label7: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Edit1: TEdit;
    procedure btn2Click(Sender: TObject);
    procedure btn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FKolLop: TFKolLop;

implementation

uses Main, UNorm;

{$R *.dfm}

procedure TFKolLop.btn2Click(Sender: TObject);
var
         SB,BZ: string;
        Res: Integer;
begin
        SB := Edit1.Text;
        BZ:=Label7.Caption;
        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                'UPDATE %s SET [РаскрЛопаток]=' + #39 + SB + #39 +
                ' WHERE ([Заказ]=' + #39 + Label2.Caption + #39 + ') AND ([' +
                FN_NAM + ']='
                + #39 + FKolLop.Caption + #39 + ') AND (БЗ='+#39+BZ+#39+')', ['KlapanaZap']) then
                Exit;
        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                'UPDATE %s SET [КолЛопаток]=' + #39 + SB + #39 +
                ' WHERE ([Заказ]=' + #39 + Label2.Caption + #39 + ') AND ([' +
                FN_NAM + ']='
                + #39 + FKolLop.Caption + #39 + ') AND (БЗ='+#39+BZ+#39+')', ['ЗапускВозд']) then
                Exit;
end;

procedure TFKolLop.btn1Click(Sender: TObject);
begin
        Close;
end;

end.
