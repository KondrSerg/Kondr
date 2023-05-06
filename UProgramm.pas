unit UProgramm;

interface

uses
        Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls,
        Forms,
        Dialogs, StdCtrls, ComCtrls;

type
        TFProgramm = class(TForm)
                Edit1: TEdit;
                Edit2: TEdit;
                Edit3: TEdit;
                Edit4: TEdit;
                Edit5: TEdit;
                Button1: TButton;
                Button2: TButton;
                ComboBox1: TComboBox;
                Label1: TLabel;
                Label2: TLabel;
                Label3: TLabel;
                Label4: TLabel;
                Label5: TLabel;
                Label6: TLabel;
                Label7: TLabel;
                Label8: TLabel;
                Label9: TLabel;
                Label10: TLabel;
                Label11: TLabel;
                DateTimePicker1: TDateTimePicker;
                Label12: TLabel;
                procedure Button2Click(Sender: TObject);
                procedure Button1Click(Sender: TObject);
        private
                { Private declarations }
        public
                { Public declarations }
        end;

var
        FProgramm: TFProgramm;

implementation

uses Main;

{$R *.dfm}

procedure TFProgramm.Button2Click(Sender: TObject);
begin
        Close;
end;

procedure TFProgramm.Button1Click(Sender: TObject);
var
        Str: string;
begin
        Str := FormatDateTime('mm.dd.yyyy', DateTimePicker1.Date);
        if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [Файл]=' + #39
                +
                Edit1.Text + #39 +
                ',[Машинное время]=' + #39 + Edit2.Text + #39 + ',[Программист]='
                + #39 +
                Edit3.Text + #39 +
                ',[Мат]=' + #39 + ComboBox1.Text + #39 + ',Х=' + #39 + Edit4.Text
                + #39 +
                ',У=' + #39 + Edit5.Text + #39 +
                ',[Технолог Готов]=' + #39 + Str + #39 + ' WHERE ([Заказ]=' + #39
                +
                Label8.Caption + #39 + ') AND ([' + FN_NAM + ']=' + #39 +
                Label10.Caption +
                #39 + ')', ['Klapana']) then
                Exit;
        if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [Файл]=' + #39
                +
                Edit1.Text + #39 +
                ',[Машинное время]=' + #39 + Edit2.Text + #39 + ',[Программист]='
                + #39 +
                Edit3.Text + #39 +
                ',[Мат]=' + #39 + ComboBox1.Text + #39 + ',Х=' + #39 + Edit4.Text
                + #39 +
                ',У=' + #39 + Edit5.Text + #39 +
                ',[Технолог Готов]=' + #39 + Str + #39 + ' WHERE ([Заказ]=' + #39
                +
                Label8.Caption + #39 + ') AND ([' + FN_NAM + ']=' + #39 +
                Label10.Caption +
                #39 + ')', ['Запуск']) then
                Exit;
end;

end.
