unit UBrigada;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls;

type
  TFBrigada = class(TForm)
    Label6: TLabel;
    Label276: TLabel;
    Label269: TLabel;
    Label278: TLabel;
    Label271: TLabel;
    Label280: TLabel;
    Label281: TLabel;
    Label275: TLabel;
    Label279: TLabel;
    Label270: TLabel;
    Label277: TLabel;
    Label68: TLabel;
    CBB1: TComboBox;
    CBB2: TComboBox;
    CBB3: TComboBox;
    CBB4: TComboBox;
    CBB5: TComboBox;
    CBB6: TComboBox;
    Memo1: TMemo;
    Button1: TButton;
    Button2: TButton;
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FBrigada: TFBrigada;

implementation

uses
  Main;

{$R *.dfm}

procedure TFBrigada.Button1Click(Sender: TObject);
begin
      Close;
end;

procedure TFBrigada.Button2Click(Sender: TObject);
var
Nom,Zak,Nam,Kol,S:string;
I,R:Integer;
begin
     for I := 0 to Memo1.Lines.Count-1 do
     begin
        S:=Memo1.Lines.Strings[i];
                                       //Nom                 FN_ZAK         FN_KOL_ZAP                          FN_NAM
     //Form1.Memo32.Lines.Add(Form1.ZSG.Cells[0,R]+';'+Form1.ZSG.Cells[2,R]+';'+Form1.ZSG.Cells[4,R]+';'+Form1.ZSG.Cells[11,R]+';');

        R:=Pos(';',S);
        if R<>0 then
        begin
            Nom:=Copy(S,1,R-1);
            Delete(S,1,R);
        end;
        //
        R:=Pos(';',S);
        if R<>0 then
        begin
            Zak:=Copy(S,1,R-1);
            Delete(S,1,R);
        end;
        //
        R:=Pos(';',S);
        if R<>0 then
        begin
            Kol:=Copy(S,1,R-1);
            Delete(S,1,R);
        end;
        //
        R:=Pos(';',S);
        if R<>0 then
        begin
            Nam:=Copy(S,1,R-1);
            Delete(S,1,R);
        end;
        //
          {  if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                'UPDATE %s SET [%s]=%s WHERE ([Номер]=%s) ',
                ['Klapana', 'Сборщик', #39 + CBB1.Text
                + #39, #39 + Nom + #39]) then
                Exit;
            //
            if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                'UPDATE %s SET [%s]=%s,[%s]=%s,[%s]=%s,[%s]=%s WHERE ([Номер]=%s)'+
                ' AND (Заказ=%s) AND (Изделие=%s)',
                ['Запуск', 'Сборщик', #39 + CBB1.Text
                + #39,
                 'Подсборка1', #39 + CBB2.Text
                + #39,
                 'СборЛопатка', #39 + CBB3.Text
                + #39,
                 'СборТяга', #39 + CBB4.Text
                + #39,
                'СборРеш', #39 + CBB5.Text
                + #39,
                'СборЭл', #39 + CBB6.Text
                + #39,
                #39 + Nom + #39,#39 + Zak + #39,#39 + Nam + #39]) then
                Exit;   }
     end;
end;

procedure TFBrigada.FormShow(Sender: TObject);
Var
I:integer;
begin

        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                'Select * from %s  Order by Фамилия',
                ['Сборщик']) then
                exit;

        CBB1.Items.Add('');
        CBB1.DropDownCount := 30;
        CBB2.Items.Add('');
        CBB2.DropDownCount := 30;
        CBB3.Items.Add('');
        CBB3.DropDownCount := 30;

        CBB4.Items.Add('');
        CBB4.DropDownCount := 30;
        CBB5.Items.Add('');
        CBB5.DropDownCount := 30;
        CBB6.Items.Add('');
        CBB6.DropDownCount := 30;

        for I := 0 to Form1.ADOQuery2.RecordCount - 1 do
        begin
                CBB1.Items.Add(Form1.ADOQuery2.FieldByName('Фамилия').AsString);
                CBB2.Items.Add(Form1.ADOQuery2.FieldByName('Фамилия').AsString);
                CBB3.Items.Add(Form1.ADOQuery2.FieldByName('Фамилия').AsString);

                CBB4.Items.Add(Form1.ADOQuery2.FieldByName('Фамилия').AsString);
                CBB5.Items.Add(Form1.ADOQuery2.FieldByName('Фамилия').AsString);
                CBB6.Items.Add(Form1.ADOQuery2.FieldByName('Фамилия').AsString);
                Form1.ADOQuery2.Next;
        end;
        Cbb1.SetFocus;
end;

end.
