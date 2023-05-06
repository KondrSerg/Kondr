unit UPrim1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls;

type
  TFPrim = class(TForm)
    Panel1: TPanel;
    Memo1: TMemo;
    Button1: TButton;
    Button2: TButton;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FPrim: TFPrim;

implementation

uses Main;

{$R *.dfm}

procedure TFPrim.Button1Click(Sender: TObject);
begin
Close;
end;

procedure TFPrim.Button2Click(Sender: TObject);
Var
I,Res:Integer;
Str,Nom_Poz,Zak:String;
begin

       Nom_Poz:=Label1.Caption;
       Zak:=Label2.Caption;
       if not Form1.mkQueryUpdate( Form1.ADOQuery1, 'UPDATE %s SET ['+Label3.Caption+']='+#39+Memo1.Text+#39+
       ' WHERE ['+FN_IDGP+']='+ #39+FPrim.Caption+#39 ,['Klapana'])
       then
               Exit;
       Form1.StringGrid10.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:=Memo1.Text;
       Close;
end;

end.
