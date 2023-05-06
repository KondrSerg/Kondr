unit UNorm;

interface

uses
        Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls,
        Forms,
        Dialogs, StdCtrls;

type
        TFNorm = class(TForm)
                Edit1: TEdit;
                Label1: TLabel;
                Edit2: TEdit;
                Label2: TLabel;
                Button1: TButton;
                Button2: TButton;
                Label3: TLabel;
                Label4: TLabel;
    Button3: TButton;
    Label5: TLabel;
    lbliDGP: TLabel;
    lbliDKO: TLabel;
                procedure Button2Click(Sender: TObject);
                procedure Button1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
        private
                { Private declarations }
        public
                { Public declarations }
        end;

var
        FNorm: TFNorm;

implementation

uses Main;

{$R *.dfm}

procedure TFNorm.Button2Click(Sender: TObject);
begin
        Close;
end;

procedure TFNorm.Button1Click(Sender: TObject);
var
         SB,BZ,Sb_o,Nom: string;
        Res,Kol_Zap,I: Integer;
        SB_F:Double;
begin
        SyStem.SysUtils.FormatSettings.DecimalSeparator :=('.');
        BZ:=Label5.Caption;
        SB := Edit2.Text;
        Res := Pos(',', Sb);
        Delete(Sb, Res, 1);
        if Res <> 0 then
                Insert('.', Sb, Res);
        SyStem.SysUtils.FormatSettings.DecimalSeparator :=('.');
        SB_F:=StrToFloat(SB);
        //========================
        SB := Edit2.Text;
        Res := Pos(',', Sb);
        Delete(Sb, Res, 1);
        if Res <> 0 then
                Insert('.', Sb, Res);
        //========================
        if Form1.Vozduh=2 Then
        Begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                'UPDATE [%s] SET [Н\ч Сборка Клапана]=' + #39 + SB + #39 +' WHERE ([Заказ]=' + #39 + FNorm.Caption + #39 + ') AND ([' +
                FN_NAM + ']='
                + #39 + Label3.Caption + #39 + ')', ['750']) then
                Exit;
                //========================================================
                if not Form1.mkQuerySelect1(Form1.ADOQuery2,
                'Select * from [%s]  WHERE ([Заказ]=' + #39 + FNorm.Caption + #39 + ') AND ([' +
                FN_NAM + ']='
                + #39 + Label3.Caption + #39 + ')', ['Запуск750']) then
                exit;
                for I:=0 to Form1.ADOQuery2.RecordCount-1 do
                begin
                        Kol_Zap:=Form1.ADOQuery2.FieldByName('Кол во запущенных').AsInteger;
                        Nom:=Form1.ADOQuery2.FieldByName('Номер').AsString;
                        SB_O:=FloatToStr(SB_F*Kol_Zap);
                        Res := Pos(',', Sb_O);
                        Delete(Sb_O, Res, 1);
                        if Res <> 0 then
                                Insert('.', Sb_O, Res);
                        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                        'UPDATE %s SET [Н\ч Сборка Клапана]=' + #39 + Sb_O+ #39 +
                        ' WHERE ([Заказ]=' + #39 + FNorm.Caption + #39 + ')AND ([Номер]=' + #39 + Nom + #39 + ') AND ([' +
                        FN_NAM + ']='
                        + #39 + Label3.Caption + #39 + ')', ['Запуск750']) then
                        Exit;
                        Form1.ADOQuery2.Next
                end;

                Form1.SG10.Cells[15 ,Form1.SG10.Row ] := SB;
        end;
        //========================
        if Form1.Vozduh=3 Then //STAM
        Begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                'UPDATE [%s] SET [Н\ч Сборка]=' + #39 + SB + #39 +' WHERE ([IDГП]=' + #39 + lbliDGP.Caption + #39 + ')', ['СТАМ']) then
                Exit;
                //========================================================
                if not Form1.mkQuerySelect1(Form1.ADOQuery2,
                'Select * from [%s]  WHERE ([IDГП]=' + #39 + lbliDGP.Caption + #39 + ')', ['ЗапускСТАМ']) then
                exit;
                for I:=0 to Form1.ADOQuery2.RecordCount-1 do
                begin
                        Kol_Zap:=Form1.ADOQuery2.FieldByName('Кол во запущенных').AsInteger;
                        Nom:=Form1.ADOQuery2.FieldByName('Номер').AsString;
                        SB_O:=FloatToStr(SB_F*Kol_Zap);
                        Res := Pos(',', Sb_O);
                        Delete(Sb_O, Res, 1);
                        if Res <> 0 then
                                Insert('.', Sb_O, Res);
                        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                        'UPDATE %s SET [Н\ч Сборка Клапана]=' + #39 + Sb_O+ #39 +
                        ' WHERE ([Заказ]=' + #39 + FNorm.Caption + #39 + ')AND ([Номер]=' + #39 + Nom + #39 + ') AND ([' +
                        FN_NAM + ']='
                        + #39 + Label3.Caption + #39 + ')', ['Запуск750']) then
                        Exit;
                        Form1.ADOQuery2.Next
                end;

                Form1.SG10.Cells[15 ,Form1.SG10.Row ] := SB;
        end;
        if Form1.Vozduh=1 Then
        Begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                'UPDATE %s SET [Н\ч Сборка Клапана]=' + #39 + SB + #39 +' WHERE ([Заказ]=' + #39 + FNorm.Caption + #39 + ') AND ([' +
                FN_NAM + ']='
                + #39 + Label3.Caption + #39 + ')', ['Klapana']) then
                Exit;
                //========================================================
                if not Form1.mkQuerySelect1(Form1.ADOQuery2,
                'Select * from %s  WHERE ([Заказ]=' + #39 + FNorm.Caption + #39 + ') AND ([' +
                FN_NAM + ']='
                + #39 + Label3.Caption + #39 + ')', ['Запуск']) then
                exit;
                for I:=0 to Form1.ADOQuery2.RecordCount-1 do
                begin
                        Kol_Zap:=Form1.ADOQuery2.FieldByName('Кол во запущенных').AsInteger;
                        Nom:=Form1.ADOQuery2.FieldByName('Номер').AsString;
                        SB_O:=FloatToStr(SB_F*Kol_Zap);
                        Res := Pos(',', Sb_O);
                        Delete(Sb_O, Res, 1);
                        if Res <> 0 then
                                Insert('.', Sb_O, Res);
                        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                        'UPDATE %s SET [Н\ч Сборка Клапана]=' + #39 + Sb_O+ #39 +
                        ' WHERE ([Заказ]=' + #39 + FNorm.Caption + #39 + ')AND ([Номер]=' + #39 + Nom + #39 + ') AND ([' +
                        FN_NAM + ']='
                        + #39 + Label3.Caption + #39 + ')', ['Запуск']) then
                        Exit;
                        Form1.ADOQuery2.Next
                end;

                Form1.SG6.Cells[I_FN_SBOR_KLAP_NC ,Form1.SG6.Row ] := SB;
        end;
        if Form1.Vozduh=0 Then
        Begin
        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                'UPDATE %s SET [Н\ч Сборка Клапана]=' + #39 + SB + #39 +
                ' WHERE ([Заказ]=' + #39 + FNorm.Caption + #39 + ') AND ([' +
                FN_NAM + ']='
                + #39 + Label3.Caption + #39 + ') AND (БЗ='+#39+BZ+#39+')', ['KlapanaZap']) then
                Exit;
                        //========================================================
                if not Form1.mkQuerySelect1(Form1.ADOQuery2,
                'Select * from %s  WHERE ([Заказ]=' + #39 + FNorm.Caption + #39 + ') AND ([' +
                FN_NAM + ']='
                + #39 + Label3.Caption + #39 + ')', ['ЗапускВозд']) then
                exit;
                for I:=0 to Form1.ADOQuery2.RecordCount-1 do
                begin
                        Kol_Zap:=Form1.ADOQuery2.FieldByName('Кол во запущенных').AsInteger;
                        Nom:=Form1.ADOQuery2.FieldByName('Номер').AsString;
                        SB_O:=FloatToStr(SB_F*Kol_Zap);
                        Res := Pos(',', Sb_O);
                        Delete(Sb_O, Res, 1);
                        if Res <> 0 then
                                Insert('.', Sb_O, Res);
                        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                        'UPDATE %s SET [Н\ч Сборка Клапана]=' + #39 + Sb_O+ #39 +
                        ' WHERE ([Заказ]=' + #39 + FNorm.Caption + #39 + ')AND ([Номер]=' + #39 + Nom + #39 + ') AND ([' +
                        FN_NAM + ']='
                        + #39 + Label3.Caption + #39 + ')', ['ЗапускВозд']) then
                        Exit;
                  Form1.ADOQuery2.Next
                end;
        {if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                'UPDATE %s SET [Н\ч Сборка Клапана]=' + #39 + SB + #39 +
                '*[кол во запущенных] WHERE ([Заказ]=' + #39 + FNorm.Caption + #39 + ') AND ([' +
                FN_NAM + ']='
                + #39 + Label3.Caption + #39 + ') AND (БЗ='+#39+BZ+#39+')', ['ЗапускВозд']) then
                Exit;}
        Form1.Button28.Click;
        end;
        Close;
        {         if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                'UPDATE %s SET [Статус]=' + #39
                + '0' + #39 + ' WHERE ([IdГП]=' + 
                #39+SGL.Cells[I_FN_SGP + 2, StringGrid6.Row]+#39+') AND'+
                ' (IdКО= '+#39+SGL.Cells[I_FN_SGP + 8, StringGrid6.Row]+#39+') ' ,
                ['KlapanaZAP']) then
                Exit;}
end;

procedure TFNorm.Button3Click(Sender: TObject);
var
       SV_O, Sv,BZ,Nom: string;
        Res,Kol_Zap,I: Integer;
        SV_F:Double;
begin
        SV := Edit1.Text;
        BZ:=Label5.Caption;
        Res := Pos(',', Sv);
        Delete(Sv, Res, 1);
        if Res <> 0 then
                Insert('.', Sv, Res);

        SyStem.SysUtils.FormatSettings.DecimalSeparator :=('.');
        SV := Edit1.Text;
        Res := Pos(',', SV);
        Delete(SV, Res, 1);
        if Res <> 0 then
                Insert('.', SV, Res);
        SyStem.SysUtils.FormatSettings.DecimalSeparator :=('.');
        SV_F:=StrToFloat(SV);
        //========================
        SV := Edit1.Text;
        Res := Pos(',', SV);
        Delete(SV, Res, 1);
        if Res <> 0 then
                Insert('.', SV, Res);
        if Form1.Vozduh=1 Then
        Begin
        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                'UPDATE %s SET [Н\ч Сварка]=' + #39 + SV + #39 +
                ' WHERE ([Заказ]=' + #39 + FNorm.Caption + #39 + ') AND ([' +
                FN_NAM + ']='
                + #39 + Label3.Caption + #39 + ')', ['Klapana']) then
                Exit;
       { if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                'UPDATE %s SET [Н\ч Сварка]=' + #39 + SV + #39 +
                ' WHERE ([Заказ]=' + #39 + FNorm.Caption + #39 + ') AND ([' +
                FN_NAM + ']='
                + #39 + Label3.Caption + #39 + ')', ['Запуск']) then
                Exit; }
        //========================================================
                if not Form1.mkQuerySelect1(Form1.ADOQuery2,
                'Select * from %s  WHERE ([Заказ]=' + #39 + FNorm.Caption + #39 + ') AND ([' +
                FN_NAM + ']='
                + #39 + Label3.Caption + #39 + ')', ['Запуск']) then
                exit;
                for I:=0 to Form1.ADOQuery2.RecordCount-1 do
                begin
                        Kol_Zap:=Form1.ADOQuery2.FieldByName('Кол во запущенных').AsInteger;
                        Nom:=Form1.ADOQuery2.FieldByName('Номер').AsString;
                        SV_O:=FloatToStr(SV_F*Kol_Zap);
                        Res := Pos(',', SV_O);
                        Delete(SV_O, Res, 1);
                        if Res <> 0 then
                                Insert('.', SV_O, Res);
                        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                        'UPDATE %s SET [Н\ч Сварка]=' + #39 + SV_O+ #39 +
                        ' WHERE ([Заказ]=' + #39 + FNorm.Caption + #39 + ')AND ([Номер]=' + #39 + Nom + #39 + ') AND ([' +
                        FN_NAM + ']='
                        + #39 + Label3.Caption + #39 + ')', ['Запуск']) then
                        Exit;
                        Form1.ADOQuery2.Next
                end;

            Form1.SG6.Cells[I_FN_SVARKA_NC ,Form1.SG6.Row ] := SV;
            //Form1.Button16.Click;

        end;
        if Form1.Vozduh=0 Then
        Begin
        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                'UPDATE %s SET [Н\ч Сварка]=' + #39 + SV + #39 +
                ' WHERE ([Заказ]=' + #39 + FNorm.Caption + #39 + ') AND ([' +
                FN_NAM + ']='
                + #39 + Label3.Caption + #39 + ') AND (БЗ='+#39+BZ+#39+')', ['KlapanaZap']) then
                Exit;
        {if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                'UPDATE %s SET [Н\ч Сварка]=' + #39 + SV + #39 +
                ' WHERE ([Заказ]=' + #39 + FNorm.Caption + #39 + ') AND ([' +
                FN_NAM + ']='
                + #39 + Label3.Caption + #39 + ') AND (БЗ='+#39+BZ+#39+')', ['ЗапускВозд']) then
                Exit; }

        //========================================================
                if not Form1.mkQuerySelect1(Form1.ADOQuery2,
                'Select * from %s  WHERE ([Заказ]=' + #39 + FNorm.Caption + #39 + ') AND ([' +
                FN_NAM + ']='
                + #39 + Label3.Caption + #39 + ')', ['ЗапускВозд']) then
                exit;
                for I:=0 to Form1.ADOQuery2.RecordCount-1 do
                begin
                        Kol_Zap:=Form1.ADOQuery2.FieldByName('Кол во запущенных').AsInteger;
                        Nom:=Form1.ADOQuery2.FieldByName('Номер').AsString;
                        SV_O:=FloatToStr(SV_F*Kol_Zap);
                        Res := Pos(',', SV_O);
                        Delete(SV_O, Res, 1);
                        if Res <> 0 then
                                Insert('.', SV_O, Res);
                        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                        'UPDATE %s SET [Н\ч Сварка]=' + #39 + SV_O+ #39 +
                        ' WHERE ([Заказ]=' + #39 + FNorm.Caption + #39 + ')AND ([Номер]=' + #39 + Nom + #39 + ') AND ([' +
                        FN_NAM + ']='
                        + #39 + Label3.Caption + #39 + ')', ['ЗапускВозд']) then
                        Exit;
                        Form1.ADOQuery2.Next
                end;
        Form1.Button28.Click;
        end;
        Close;
end;

end.
