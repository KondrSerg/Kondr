unit Unit3;

interface

uses
        Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls,
        Forms,
        Dialogs, StdCtrls, ExtCtrls;

type
        TForm3 = class(TForm)
                Panel1: TPanel;
                Label1: TLabel;
                Label2: TLabel;
                Label3: TLabel;
                Label4: TLabel;
                Label5: TLabel;
                Button1: TButton;
                Button2: TButton;
                Memo1: TMemo;
                Label6: TLabel;
                Label7: TLabel;
    LabKO: TLabel;
    LabGP: TLabel;
                procedure Button1Click(Sender: TObject);
                procedure Button2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
        private
                { Private declarations }
        public
                { Public declarations }
        end;

var
        Form3: TForm3;

implementation

uses Main;

{$R *.dfm}

procedure TForm3.Button1Click(Sender: TObject);
begin
        Close;
end;

procedure TForm3.Button2Click(Sender: TObject);
var
        I, Res, Res1: Integer;
        Str, Nam, Zak, Nom,IDGP,IDKO: string;
begin

        Nam := Label6.Caption;
        Zak := Label7.Caption;
        Nom := Label3.Caption;
        IDGP:=LabGP.Caption;
        IDKO:=LabKO.Caption;
        Res := AnsiCompareStr(Label2.Caption, '0');
        Str:= StringReplace(Memo1.Text,#13#10, ' ',[rfReplaceAll,rfIgnoreCase]);
        if Form1.Vozduh=0 Then
        Begin
        if Res = 0 then
        begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [ОТК Прим]=' + #39 + Memo1.Text + #39 +
                        ' WHERE ([Изделие]=' + #39 + Nam + #39 +
                        ') AND ([Заказ]=' + #39 + Zak +
                        #39 + ') AND ([Номер]=' + #39 + Nom + #39 + ') AND ([IdГП]=' + #39 + IDGP + #39 + ') AND ([Id]=' + #39 + IDKO + #39 + ')',
                        ['Запуск']) then
                        Exit;

                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [ОТК Прим]=' + #39 + Memo1.Text + #39 +
                        ' WHERE ([Изделие]=' + #39 + Nam + #39 +
                        ') AND ([Заказ]=' + #39 + Zak +
                        #39 + ') AND ([IdГП]=' + #39 + IDGP + #39 + ') AND ([Id]=' + #39 + IDKO + #39 + ')',
                        ['Klapana']) then
                        Exit;
        end
        else
        begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label1.Caption + ' ' + Label2.Caption + ']=' + #39 +
                        Memo1.Text + #39 +
                        ' WHERE ([Изделие]=' + #39 + Nam + #39 +
                        ') AND ([Заказ]=' + #39 + Zak +
                        #39 + ') AND ([Номер]=' + #39 + Nom + #39 + ') AND ([IdГП]=' + #39 + IDGP + #39 + ') AND ([Id]=' + #39 + IDKO + #39 + ')',
                        ['Запуск']) then
                        Exit;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label1.Caption + ' ' + Label2.Caption + ']=' + #39 +
                        Memo1.Text + #39 +
                        ' WHERE ([Изделие]=' + #39 + Nam + #39 +
                        ') AND ([Заказ]=' + #39 + Zak +
                        #39 + ') AND ([IdГП]=' + #39 + IDGP + #39 + ') AND ([Id]=' + #39 + IDKO + #39 + ') ',
                        ['Klapana']) then
                        Exit;

        end;
                {FTSP.Send(Label1.Caption+ ' ' + Label2.Caption + ']=' + #39 + Memo1.Text + #39 +'([Изделие]=' + #39 +
                        Nam + #39 + ') AND ([Заказ]=' + #39 +
                        Zak + #39 +
                        ')  ');}
        Form1.ZSG.Cells[StrToInt(Label5.Caption), StrToInt(Label4.Caption)] :=
                Str;

        end;
        if Form1.Vozduh=1 Then
        Begin
        if Res = 0 then
        begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [ОТК Прим]=' + #39 + Memo1.Text + #39 +
                        ' WHERE ([Изделие]=' + #39 + Nam + #39 +
                        ') AND ([Заказ]=' + #39 + Zak +
                        #39 + ') AND ([Номер]=' + #39 + Nom + #39 + ')',
                        ['ЗапускВозд']) then
                        Exit;

                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [ОТК Прим]=' + #39 + Memo1.Text + #39 +
                        ' WHERE (IdГП=' + #39 +
                LabGP.Caption
                + #39 + ') AND (IdКО=' + #39 +
                LabKO.Caption
                + #39 + ')',
                        ['KlapanaZap']) then
                        Exit;
        end
        else
        begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label1.Caption + ' ' + Label2.Caption + ']=' + #39 +
                        Memo1.Text + #39 +
                        ' WHERE (IdГП=' + #39 +LabGP.Caption
                + #39 + ') AND ([Номер]=' + #39 + Nom + #39 + ') AND (IdКО=' + #39 +LabKO.Caption
                + #39 + ') ',
                        ['ЗапускВозд']) then
                        Exit;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label1.Caption + ' ' + Label2.Caption + ']=' + #39 +
                        Memo1.Text + #39 +' WHERE (IdГП=' + #39 +LabGP.Caption
                + #39 + ') AND (IdКО=' + #39 +LabKO.Caption
                + #39 + ') ',
                        ['KlapanaZap']) then
                        Exit;

        end;
             Form1.ZCV.Cells[StrToInt(Label5.Caption), StrToInt(Label4.Caption)] :=
                Str;
        end;
        //
        if Form1.Vozduh=2 Then
        Begin
        if Res = 0 then
        begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE [%s] SET [ОТК Прим]=' + #39 + Memo1.Text + #39 +
                        ' WHERE ([Изделие]=' + #39 + Nam + #39 +
                        ') AND ([Заказ]=' + #39 + Zak +
                        #39 + ') AND ([Номер]=' + #39 + Nom + #39 + ')',
                        ['Запуск750']) then
                        Exit;

                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE [%s] SET [ОТК Прим]=' + #39 + Memo1.Text + #39 +
                        ' WHERE (IdГП=' + #39 +
                LabGP.Caption
                + #39 + ') AND (IdКО=' + #39 +
                LabKO.Caption
                + #39 + ')',
                        ['750']) then
                        Exit;
        end
        else
        begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE [%s] SET [' +
                        Label1.Caption + ' ' + Label2.Caption + ']=' + #39 +
                        Memo1.Text + #39 +
                        ' WHERE (IdГП=' + #39 +LabGP.Caption
                + #39 + ') AND ([Номер]=' + #39 + Nom + #39 + ') AND (IdКО=' + #39 +LabKO.Caption
                + #39 + ') ',
                        ['Запуск750']) then
                        Exit;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE [%s] SET [' +
                        'Примечания' + ']=' + #39 +
                        Memo1.Text + #39 +' WHERE (IdГП=' + #39 +LabGP.Caption
                + #39 + ') AND (IdКО=' + #39 +LabKO.Caption
                + #39 + ') ',
                        ['750']) then
                        Exit;

        end;
             Form1.ZCV.Cells[StrToInt(Label5.Caption), StrToInt(Label4.Caption)] :=
                Str;
        end;
        if Form1.Vozduh=4 Then
        Begin

          if Form1.Luk=0 then
          Begin
          if Form1.OTK_Prim=0 then
           begin
              if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE [%s] SET [' +
                        'Примечания' + ']=' + #39 +
                        Memo1.Text + #39 +' WHERE (IdГП=' + #39 +Label6.Caption
                + #39 + ') ',
                        ['СТАМ']) then
                        Exit;
              if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE [%s] SET [' +
                        'Примечание' + ']=' + #39 +
                        Memo1.Text + #39 +' WHERE (IdГП=' + #39 +Label6.Caption
                + #39 + ') ',//AND (Номер=' + #39 +Label3.Caption+ #39 + ')',
                        ['ЗапускСТАМ']) then
                        Exit;
           end
           else
              if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE [%s] SET [' +
                        'ОТК Прим' + ']=' + #39 +
                        Memo1.Text + #39 +' WHERE (IdГП=' + #39 +Label6.Caption
                + #39 + ') ',//AND (Номер=' + #39 +Label3.Caption+ #39 + ')',
                        ['ЗапускСТАМ']) then
                        Exit;
          End;

          if Form1.Luk=1 then
          begin
          if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        'Примечания' + ']=' + #39 +
                        Memo1.Text + #39 +' WHERE (IdГП=' + #39 +Label6.Caption
                + #39 + ')  AND (Номер=' + #39 +Label3.Caption+ #39 + ')',
                        ['Люк']) then
                        Exit;

          if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        'Примечание' + ']=' + #39 +
                        Memo1.Text + #39 +' WHERE (IdГП=' + #39 +Label6.Caption
                + #39 + ')  AND (Номер=' + #39 +Label3.Caption+ #39 + ')',
                        ['ЗапускЛюк']) then
                        Exit;
          end;
        end;
        Close;
end;

procedure TForm3.FormShow(Sender: TObject);
begin

        Memo1.SetFocus;
        Memo1.SelectAll;
end;

end.
