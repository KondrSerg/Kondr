unit New;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, DB, ADODB;

type
  TFNew = class(TForm)
    DateTimePicker1: TDateTimePicker;
    Label1: TLabel;
    Button1: TButton;
    Button2: TButton;
    Memo1: TMemo;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    Memo2: TMemo;
    procedure FormShow(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FNew: TFNew;

implementation

uses Main;

{$R *.dfm}

procedure TFNew.FormShow(Sender: TObject);
begin
      DateTimePicker1.DateTime:=Now;
end;

procedure TFNew.Button2Click(Sender: TObject);
begin
        Close;
end;

procedure TFNew.Button1Click(Sender: TObject);
Var
Str2,Model,IDGP,IDKO,DatZak,Zakaz,BZ,Tip,Grup,Izdel,GrupGP,IzdelGP,Kol,Poluch,
DatGot,DatOtmen,Prim,NachNom,KonNom,NachKod,KonKod:String;
i,Kod:Integer;
Otmen:String;
begin
            Memo1.Lines.Clear;
            Memo1.Lines.Add(ADOConnection1.ConnectionString);
            ADOConnection1.Connected:=True;
            Memo1.Lines.Add('ADOConnection1.Connection=True');
            ADOQuery1.Close;
            ADOQuery1.SQL.Clear;
            Str2:='select [IdГП], [IdКО], [ДатаЗаказа], [Заказ], [БЗ], [Тип], [Группа], [Изделие], [ГруппаГП], [ИзделиеГП], [Кол-во], [Получатель], [ДатаГотОфис],'+
            '[Отмена], [ДатаОтмены], [Особенности]'+
            'from dbo.[ИмпортЗаказовКП_МИ]('+#39+FormatDateTime('dd.mm.yyyy',DateTimePicker1.DateTime)+#39+')  AS ИмпортЗаказовКП_МИ_1';
            ADOQuery1.SQL.Add(Str2);
            Memo1.Lines.Add(Str2);
            ADOQuery1.Open;
            Memo1.Lines.Add(IntToStr(ADOQuery1.RecordCount));
            For I:=0 To ADOQuery1.RecordCount-1 Do
            Begin
                IDGP:=ADOQuery1.FieldByName('IdГП').AsString;
                Kod:=ADOQuery1.FieldByName('Группа').AsInteger;
                if Kod=310 Then
                Begin
                        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                                'Select * from %s Where (IdГП=' + #39 +
                                IDGP + #39 + ') ', ['Klapana']) then
                                exit;
                        if Form1.ADOQuery2.RecordCount = 0 then
                        Begin

                                IDKO:=ADOQuery1.FieldByName('IdКО').AsString;
                                DatZak:=ADOQuery1.FieldByName('ДатаЗаказа').AsString;
                                Zakaz:=ADOQuery1.FieldByName('Заказ').AsString;
                                BZ:=ADOQuery1.FieldByName('БЗ').AsString;
                                Tip:=ADOQuery1.FieldByName('Тип').AsString;
                                Grup:=ADOQuery1.FieldByName('Группа').AsString;
                                Izdel:=ADOQuery1.FieldByName('Изделие').AsString;
                                GrupGP:=ADOQuery1.FieldByName('ГруппаГП').AsString;
                                IzdelGP:=ADOQuery1.FieldByName('ИзделиеГП').AsString;
                                Kol:=ADOQuery1.FieldByName('Кол-во').AsString;
                                Poluch:=ADOQuery1.FieldByName('Получатель').AsString;
                                DatGot:=ADOQuery1.FieldByName('ДатаГотОфис').AsString;
                                Otmen:=ADOQuery1.FieldByName('Отмена').AsString;
                                DatOtmen:=ADOQuery1.FieldByName('ДатаОтмены').AsString;
                                Prim:=ADOQuery1.FieldByName('Особенности').AsString;
                                Memo1.Lines.Add(' '+IDGP+' '+IDKO+' '+DatZak+' '+Zakaz+' '+BZ+' '+Tip+' '+Grup+
                                ' '+Izdel+' '+GrupGP+' '+IzdelGP+' '+Kol+' '+Poluch+' '+DatGot+' '+Otmen+' '+DatOtmen+' '+Prim+' ');


                        end;
                End
                Else
                Begin
                        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                                'Select * from %s Where (IdГП=' + #39 +
                                IDGP + #39 + ') ', ['KlapanaZap']) then
                                exit;
                        if Form1.ADOQuery2.RecordCount = 0 then
                        Begin

                                IDKO:=ADOQuery1.FieldByName('IdКО').AsString;
                                DatZak:=ADOQuery1.FieldByName('ДатаЗаказа').AsString;
                                Zakaz:=ADOQuery1.FieldByName('Заказ').AsString;
                                BZ:=ADOQuery1.FieldByName('БЗ').AsString;
                                Tip:=ADOQuery1.FieldByName('Тип').AsString;
                                Grup:=ADOQuery1.FieldByName('Группа').AsString;
                                Izdel:=ADOQuery1.FieldByName('Изделие').AsString;
                                GrupGP:=ADOQuery1.FieldByName('ГруппаГП').AsString;
                                IzdelGP:=ADOQuery1.FieldByName('ИзделиеГП').AsString;
                                Kol:=ADOQuery1.FieldByName('Кол-во').AsString;
                                Poluch:=ADOQuery1.FieldByName('Получатель').AsString;
                                DatGot:=ADOQuery1.FieldByName('ДатаГотОфис').AsString;
                                Otmen:=ADOQuery1.FieldByName('Отмена').AsString;
                                DatOtmen:=ADOQuery1.FieldByName('ДатаОтмены').AsString;
                                Prim:=ADOQuery1.FieldByName('Особенности').AsString;
                                Memo1.Lines.Add(' '+IDGP+' '+IDKO+' '+DatZak+' '+Zakaz+' '+BZ+' '+Tip+' '+Grup+
                                ' '+Izdel+' '+GrupGP+' '+IzdelGP+' '+Kol+' '+Poluch+' '+DatGot+' '+Otmen+' '+DatOtmen+' '+Prim+' ');
                                {if not mkQueryInsert(ADOQuery1, 'Insert Into %s '
                                        +
                                        '(Код,[Дата],Заказ,Изделие,[Кол во],' +
                                        'Заказчик,Подр,Пр,Пол,СГП,' +
                                        '[план Неделя],[план. Дата],[Заводской номер],[МодПривода],[Н\ч Сварка],[Н\ч Сборка Клапана],[A],[B],Штрих,ШтрихПос ) ' +
                                        'Values (%s,%s,%s,%s,%s' +
                                        ',%s,%s,%s,%s,%s,' +
                                        '%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)',
                                        ['KlapanaZap',
                                        #39 + SG1.Cells[1, i] + #39, #39 + Dat1
                                                + #39, #39 + SG1.Cells[3, i] +
                                                #39, #39 + SG1.Cells[7, i] +
                                                #39, #39 + SG1.Cells[8, i] +
                                                #39,
                                                #39 + SG1.Cells[5, i] + #39, #39
                                                + SG1.Cells[6, i] + #39, #39 +
                                                SG1.Cells[11, i] + #39, #39 +
                                                SG1.Cells[12, i] + #39, #39 +
                                                SG1.Cells[13, i] + #39,
                                                #39 + SG1.Cells[10, i] + #39, #39
                                                + Dat2 + #39, #39 +
                                                SG1.Cells[16,
                                                i] + #39, #39 + '' + #39, #39 +
                                                NC_Svar + #39, #39 + NC_Sbor +
                                                #39,
                                                #39 + A + #39, #39 + B + #39, #39 +
                                                SG1.Cells[20, i] + #39, #39 +
                                                SG1.Cells[21, i] + #39])
                                                        then
                                        exit; }
                        end;
                end;
                ADOQuery1.Next;

            end;

end;

end.
