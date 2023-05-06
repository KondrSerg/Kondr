unit UPrim;

interface

uses
        Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls,
        Forms,
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
    Label6: TLabel;
    Memo2: TMemo;
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
var
        I, Res: Integer;
        Str, Nam, Zak,BZ: string;
begin
        BZ:=Label6.Caption;
        Nam := FPrim.Caption;
        Zak := Label2.Caption;
        if Form1.Vozduh=1 Then
        Begin
          if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label3.Caption
                + ']=' + #39 + Memo1.Text + #39 +
                ' WHERE ([IDГП]=' + #39 + Label6.Caption + #39 +
                ') AND ([IDКО]=' + #39 + Label2.Caption + #39 + ')', ['Klapana']) then
                Exit;
          Res:=Pos('Технолог',Label3.Caption);
          if Res<>0 then
          Begin
             if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [Сборка Примечание]=concat([Сборка Примечание],' + #39 + Memo1.Text + #39 +
             ') WHERE ([IDГП]=' + #39 + Label6.Caption + #39 +
                ') AND ([IDКО]=' + #39 + Label2.Caption + #39 + ')', ['Запуск']) then
                Exit;
          End
          else
          Begin
            if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label3.Caption
                + ']=' + #39 + Memo1.Text + #39 +
                ' WHERE ([IDГП]=' + #39 + Label6.Caption + #39 +
                ') AND ([IDКО]=' + #39 + Label2.Caption + #39 + ')', ['Запуск']) then
                Exit;
          End;
          Form1.Button16.Click;
        End;
        if Form1.Vozduh=0 Then
        Begin
        if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label3.Caption
                + ']=' + #39 + Memo1.Text + #39 +
                ' WHERE ([IDГП]=' + #39 + Label6.Caption + #39 +
                ') AND ([IDКО]=' + #39 + Label2.Caption + #39 + ')', ['KlapanaZap']) then
                Exit;
                  Res:=Pos('Технолог',Label3.Caption);
          if Res<>0 then
          Begin
             if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [Сборка Примечание]=concat([Сборка Примечание],' + #39 + Memo1.Text + #39 +
             ') WHERE ([IDГП]=' + #39 + Label6.Caption + #39 +
                ') AND ([IDКО]=' + #39 + Label2.Caption + #39 + ')', ['ЗапускВозд']) then
                Exit;
          End
          else
          Begin
              if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label3.Caption
                + ']=' + #39 + Memo1.Text + #39 +
                ' WHERE ([IDГП]=' + #39 + Label6.Caption + #39 +
                ') AND ([IDКО]=' + #39 + Label2.Caption + #39 + ')', ['ЗапускВозд']) then
                Exit;
          End;
          Form1.Button28.Click;
        End;
        //Form1.SG6.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:=Memo1.Text; .
       // FTSP.Send( Label3.Caption +  ' = '+Memo1.Text);

        Close;
end;

end.
