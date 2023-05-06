unit U710;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, Vcl.ExtCtrls;

type
  TF710 = class(TForm)
    btn1: TButton;
    btn2: TButton;
    dtp1: TDateTimePicker;
    Label2: TLabel;
    Button1: TButton;
    rg1: TRadioGroup;
    Memo1: TMemo;
    Edit1: TEdit;
    Label1: TLabel;
    procedure btn1Click(Sender: TObject);
    procedure btn2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure rg1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    Fam:string;
  end;

var
  F710: TF710;

implementation

uses Main;

{$R *.dfm}

procedure TF710.btn1Click(Sender: TObject);
begin
Close;
end;

procedure TF710.btn2Click(Sender: TObject);
var Dat,IDGP,S,Kol:string;
R,K:Integer;
begin
        IDGP:=Label2.Caption;
        Dat:=FormatDateTime('mm.dd.YYYY',dtp1.Date);
        S:='';
        R:=Pos('ОТК',F710.Caption);
        if R<>0 then
        Begin
          if Fam='' then
          Begin
                MessageDlg('Выберите фамилию!', mtError,
                        [mbOk], 0);
                Exit;
          end;
          S:=', ОТКФам='+#39+Fam+#39;
          K:=StrToInt(Edit1.Text);
          if k=0 then
          Begin
                MessageDlg('Поставьте кол-во!', mtError,
                        [mbOk], 0);
                Exit;
          end;
          Kol:=' ,[Кол принятых]='+#39+Edit1.Text+#39;
        End;
        //
        R:=Pos('Прим',F710.Caption);
        if R<>0 then
        Begin
            Dat:=Memo1.Text;
        End;
        //
        if Form1.F7=0 then
        begin
              if not Form1.mkQueryUpdate2(Form1.ADOQuery1,
              'UPDATE %s SET ['+F710.Caption+']=' +
              #39+ Dat + #39+S+Kol+
              ' WHERE ([IdГП]=' + #39 + IDGP + #39 +
              ') ', ['[710]']) then  Exit;
              Form1.Button7.Click;
        end;

        if Form1.F7=1 then
        begin
              if not Form1.mkQueryUpdate2(Form1.ADOQuery1,
              'UPDATE %s SET ['+F710.Caption+']=' +
              #39+ Dat + #39 +S+
              ' WHERE ([IdГП]=' + #39 + IDGP + #39 +
              ') ', ['[750]']) then Exit;
              R:=Pos('Прим',F710.Caption);
              if R<>0 then
              Begin


              if not Form1.mkQueryUpdate2(Form1.ADOQuery1,
              'UPDATE %s SET [Сборка Примечание]=' +
              #39+ Dat + #39 +S+
              ' WHERE ([IdГП]=' + #39 + IDGP + #39 +
              ') ', ['[Запуск750]']) then Exit;
              End ;
              Form1.Btn55.Click;
        end;
        //
        rg1.ItemIndex:=-1;
        Fam:='';
        Memo1.Text:='';
        Close;
end;

procedure TF710.Button1Click(Sender: TObject);
var IDGP:string;
begin
        IDGP:=Label2.Caption;
        if Form1.F7=0 then
        begin
        if not Form1.mkQueryUpdate2(Form1.ADOQuery1,
        'UPDATE %s SET ['+F710.Caption+']=NULL' +
        ' WHERE ([IdГП]=' + #39 + IDGP + #39 +
        ') ', ['[710]']) then
        Exit;
        Form1.Button7.Click;
        end;

        if Form1.F7=1 then
        begin
        if not Form1.mkQueryUpdate2(Form1.ADOQuery1,
        'UPDATE %s SET ['+F710.Caption+']=NULL' +
        ' WHERE ([IdГП]=' + #39 + IDGP + #39 +
        ') ', ['[750]']) then
        Exit;
        Form1.Btn55.Click;
        end;
        Close;
end;

procedure TF710.FormShow(Sender: TObject);
var R:Integer;
begin
  rg1.Items.Clear;
  rg1.Items.LoadFromFile( 'V:\ОИТ\Cklapana2\2013\OTK.ini');
        R:=Pos('Прим',F710.Caption);
        if R<>0 then
        Begin
            Memo1.Visible:=True;
            dtp1.Visible:=False;
           // rg1.Visible:=False;
        End
        else
        Begin
            Memo1.Visible:=False;
            dtp1.Visible:=True;
            //rg1.Visible:=True;
        End;
end;

procedure TF710.rg1Click(Sender: TObject);
begin
Fam:=rg1.Items[rg1.ItemIndex];
end;

end.
