unit UNomerProizv;

interface

uses
        Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls,
        Forms,
        Dialogs, StdCtrls, ComCtrls,UConnCeh,UConnKlap,UOsnova;

type
        TForm2 = class(TForm)
                Button1: TButton;
                Button2: TButton;
                DateTimePicker1: TDateTimePicker;
                Label1: TLabel;
                Label2: TLabel;
                Label3: TLabel;
                Label4: TLabel;
                Label5: TLabel;
                Label6: TLabel;
                Label7: TLabel;
                Label8: TLabel;
    Label9: TLabel;
    Button3: TButton;
    Label10: TLabel;
    LabGP: TLabel;
    LabKO: TLabel;
    Label3bz: TLabel;
    LabCVET: TLabel;
    Plan: TLabel;
                procedure Button1Click(Sender: TObject);
                procedure Button2Click(Sender: TObject);
    procedure Button1KeyPress(Sender: TObject; var Key: Char);
    procedure Button3Click(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure DateTimePicker1KeyPress(Sender: TObject; var Key: Char);
    procedure Button2KeyPress(Sender: TObject; var Key: Char);
    procedure FormShow(Sender: TObject);
        private
                { Private declarations }
        public
                { Public declarations }
                Conn_Ceh:Connect_Miass_Ceh;
                Conn_Klap:Connect_Miass_Klap;
                UOsnova_Main:Osnova_Main ;
                Cvet1,Cvet:Integer;
        end;

var
        Form2: TForm2;

implementation

uses Main;

{$R *.dfm}

procedure TForm2.Button1Click(Sender: TObject);
begin
        Close;
end;

procedure TForm2.Button2Click(Sender: TObject);
var
        Str, Zag_S, Sbor_S,Zap,Str1,Zakaz,Izdelie,Nomer,Kol_Zap,Dat,BZ,fileName,Vn_Dat: string;

        N, Res, Res1, Res2, Res3, Zag, Sbor,I,L9,Otgr,Zap1,Res4,Pods,C: Integer;
        S,S1:string;
D,D1:TDateTime;

begin
        C:=StrToInt(LabCVET.Caption);
        Str := FormatDateTime('mm.dd.yyyy', DateTimePicker1.Date);
        Str1 := FormatDateTime('dd.mm.yyyy', DateTimePicker1.Date);
        Res := AnsiCompareStr(Label1.Caption + ' ' + Label2.Caption,
                'Заготовка Запуск');
        Res1 := AnsiCompareStr(Label1.Caption + ' ' + Label2.Caption,
                'Заготовка Готовность');
        Res2 := AnsiCompareStr(Label1.Caption + ' ' + Label2.Caption,
                'Сборка Запуск');
        Res3 := AnsiCompareStr(Label1.Caption + ' ' + Label2.Caption,
                'Сборка Готовность');
        Res4 := AnsiCompareStr(Label2.Caption,
                'Планирование');
        Otgr := AnsiCompareStr(Label2.Caption,
                'Упаковка');
        Pods := AnsiCompareStr(Label2.Caption,
                'Подсборка');
        L9:=StrToInt(Label9.Caption);
        if Form1.Vozduh=0 Then
        Begin
          if (Otgr=0 ) Then
          Begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 Label2.Caption + ']=' + #39 + Str + #39 +
                ' WHERE ([ID]=' + #39 + Form1.ZSG.Cells[I_FN_KOL_ZAP + 14,Form1.ZSG.Row] + #39 + ') AND  ([Заказ]=' + #39 + Label6.Caption + #39 + ') AND (' +
                FN_NAM + '=' +
                #39 + Label3.Caption + #39 + ') AND (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') ', ['Запуск']) then
                Exit;
                Form1.ZSG.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:=Str1;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 Label2.Caption + ']=' + #39 + Str + #39 +
                ' WHERE ([Заказ]=' + #39 + Label6.Caption + #39 + ') AND (' +
                FN_NAM + '=' +
                #39 + Label3.Caption + #39 + ')  ', ['Klapana']) then
                Exit;
          end;
          if (Pods=0 ) Then
          Begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 Label2.Caption + ']=' + #39 + Str + #39 +
                ' WHERE  ([ID]=' + #39 + Form1.ZSG.Cells[I_FN_KOL_ZAP + 14,Form1.ZSG.Row] + #39 + ') AND ([Заказ]=' + #39 + Label6.Caption + #39 + ') AND (' +
                FN_NAM + '=' +
                #39 + Label3.Caption + #39 + ') AND (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') ', ['Запуск']) then
                Exit;
                Form1.ZSG.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:=Str1;

          end;
          if Res=0 Then //Заготовка Запуск      ([ID]=' + #39 + Form1.ZSG.Cells[I_FN_KOL_ZAP + 14,Form1.ZSG.Row] + #39 + ') AND ([Заказ]=' + #39 + Label6.Caption + #39 + ') AND (' +
                //FN_NAM + '=' +
                //#39 + Label3.Caption + #39 + ') AND'}
          Begin
            if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label1.Caption
                + ' ' + Label2.Caption + ']=' + #39 + Str + #39 +
                ' WHERE  (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') ', ['Запуск']) then
                Exit;
            if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label1.Caption
                + ' ' + Label2.Caption + ']=' + #39 + Str + #39 +
                ' WHERE  (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') ', ['Klapana']) then
                Exit;
                Form1.ZSG.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:=Str1;
          end;
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
          if Res1 = 0 then //Заготовка Готовность
          begin
               { }

                if not Form1.mkQuerySelect1(Form1.ADOQuery2,
                        'Select * from %s Where  (Заказ=' + #39 + Label6.Caption + #39 + ') AND (Изделие=' + #39 + Label3.Caption + #39 + ') ',
                        ['Klapana']) then
                        exit;

                Zag_S := Form1.ADOQuery2.FieldByName('Заготовка').AsString;
                Kol_Zap := Label7.Caption;
                Dat:=Label10.Caption;
                if Zag_S = '' then
                        Zag := 0
                else
                        Zag := StrToInt(Zag_S);
                if Dat='' Then
                        Zag := Zag - StrToInt(Kol_Zap)
                Else
                        Zag:=Zag;
                if Zag<0 Then
                Zag:=0;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label1.Caption + ' ' + Label2.Caption + ']=' + #39 + Str
                        + #39 +
                        ',Заготовка=' + #39 + IntToStr(Zag) + #39 +
                        ' WHERE ([Изделие]=' + #39 +
                        Label3.Caption + #39 + ') AND ([Заказ]=' + #39 +
                        Label6.Caption + #39 +
                        ') '  , ['Klapana']) then
                        Exit;
                        Form1.ADOQuery2.Next;

                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label1.Caption
                + ' ' + Label2.Caption + ']=' + #39 + Str + #39 +
                ' WHERE  ([ID]=' + #39 + Form1.ZSG.Cells[I_FN_KOL_ZAP + 14,Form1.ZSG.Row] + #39 + ') AND  (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') AND ([Изделие]=' + #39 +
                        Label3.Caption + #39 + ') AND ([Заказ]=' + #39 +
                        Label6.Caption + #39 +
                        ')  ', ['Запуск']) then
                Exit;

                Form1.ZSG.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:=Str1;
          end;
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
          if Res2 = 0 then //Сборка Запуск
          begin
                if not Form1.mkQuerySelect1(Form1.ADOQuery2,
                        'Select * from %s Where  (Заказ=' + #39 + Label6.Caption + #39 + ') AND (Изделие=' + #39 + Label3.Caption + #39 + ') ',
                ['Klapana']) then
                exit;
                Sbor_S := Form1.ADOQuery2.FieldByName('Сборка').AsString;
                Zap1:=StrToInt(Form1.ADOQuery2.FieldByName('Запуск1').AsString);
                Dat:=Label10.Caption;
                Kol_Zap := Label7.Caption;
                if Sbor_S = '' then
                        Sbor := 0
                else
                        Sbor := StrToInt(Sbor_S);
                if Dat<>'' Then
                        Sbor := Sbor + (0)
                Else
                        Sbor := Sbor + StrToInt(Kol_ZAp);
                if Dat<>'' Then
                        Zap1 := Zap1 + (0)
                Else
                        Zap1 := Zap1 + StrToInt(Kol_ZAp);

                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label1.Caption + ' ' + Label2.Caption + ']=' + #39 + Str
                        + #39 +
                        ',Сборка=' + #39 + IntToStr(Sbor) + #39 +
                        ',Запуск1=' + #39 + IntToStr(Zap1) + #39 +' WHERE  ([Изделие]=' + #39 +
                        Label3.Caption + #39 + ') AND ([Заказ]=' + #39 +
                        Label6.Caption + #39 +
                        ') '  , ['Klapana']) then
                        Exit;

                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label1.Caption
                + ' ' + Label2.Caption + ']=' + #39 + Str + #39 +
                ' WHERE  ([ID]=' + #39 + Form1.ZSG.Cells[I_FN_KOL_ZAP + 14,Form1.ZSG.Row] + #39 + ') AND  (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') AND ([Изделие]=' + #39 +
                        Label3.Caption + #39 + ') AND ([Заказ]=' + #39 +
                        Label6.Caption + #39 +
                        ')  ', ['Запуск']) then
                Exit;

                Form1.ZSG.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:=Str1;
          end;
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
          if Res3 = 0 then //Сборка Готовность
          begin
                if not Form1.mkQuerySelect1(Form1.ADOQuery2,
                        'Select * from %s Where   (Заказ=' + #39 + Label6.Caption
                        + #39 +
                        ') AND (Изделие=' + #39 + Label3.Caption + #39 + ') ',
                        ['Klapana']) then
                        exit;
                Sbor_S := Form1.ADOQuery2.FieldByName('Сборка').AsString;
                Zap1:=StrToInt(Form1.ADOQuery2.FieldByName('Запуск1').AsString);
                Dat:=Label10.Caption;
                Kol_Zap := Label7.Caption;
                if Sbor_S = '' then
                        Sbor := 0
                else
                        Sbor := StrToInt(Sbor_S);
                if Dat<>'' Then
                        Sbor:=Sbor
                Else
                        Sbor := Sbor - StrToInt(Kol_Zap);
                if Sbor<0 Then
                        Sbor:=0;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label1.Caption + ' ' + Label2.Caption + ']=' + #39 + Str
                + #39 +',Сборка=' + #39 + IntToStr(Sbor) + #39 +
                ' WHERE   ([Изделие]=' + #39 +
                Label3.Caption + #39 + ') AND ([Заказ]=' + #39 +
                Label6.Caption + #39 +
                ') ', ['Klapana']) then
                Exit;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label1.Caption
                + ' ' + Label2.Caption + ']=' + #39 + Str + #39 +
                ' WHERE  ([ID]=' + #39 + Form1.ZSG.Cells[I_FN_KOL_ZAP + 14,Form1.ZSG.Row] + #39 + ') AND ([Заказ]=' + #39 + Label6.Caption + #39 + ') AND (' +
                FN_NAM + '=' +
                #39 + Label3.Caption + #39 + ') AND (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') ', ['Запуск']) then
                Exit;
                        Form1.ZSG.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:=Str1;
          end;
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
          if Res4 = 0 then  //([Заказ]=' + #39 + Label6.Caption + #39 + ') AND (' +
                //FN_NAM + '=' +
                //#39 + Label3.Caption + #39 + ') AND
         { begin

         S:=FormatDateTime('mm.dd.YYYY',DateTimePicker1.Date);
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select  max(планирование) as S from %s Where планирование<'+#39+S+#39 ,
                ['ЗапускВозд']) then
                exit;
        D:= Form1.ADOQuery1.FieldByName('S').AsDateTime;
//+++++++++++++++++++++++++++++++
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select  max(планирование) as S from %s Where планирование<'+#39+S+#39 ,
                ['Запуск']) then
                exit;
        D1:= Form1.ADOQuery1.FieldByName('S').AsDateTime;

        if D>D1 then
        Begin
        S1:=FormatDateTime('mm.dd.YYYY',D);
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select  * from %s Where планирование='+#39+S1+#39 ,
                ['ЗапускВозд']) then
                exit;
        End
        Else
        Begin
        S1:=FormatDateTime('mm.dd.YYYY',D1);
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select  * from %s Where планирование='+#39+S1+#39 ,
                ['Запуск']) then
                exit;
        End;
        I:=Form1.ADOQuery1.RecordCount;
        Cvet:= Form1.ADOQuery1.FieldByName('Cvet').AsInteger;
         case Cvet of
            0:
            begin
              ///Label16.Font.Color:=RGB(255,255,255);//Белый
              //Label17.Font.Color:=RGB(234,239,45); //Желтый
              Cvet1:=1;
            End;
            1:
            begin
              //Label16.Font.Color:=RGB(234,239,45); //Желтый
              //Label17.Font.Color:=RGB(41,243,104);//Зеленый
              Cvet1:=2;
            End;
            2:
            begin
             // Label16.Font.Color:=RGB(41,243,104);//Зеленый
              //Label17.Font.Color:=RGB(122,156,250); //Голубой
              Cvet1:=3;
            End;
            3:
            begin
              //Label16.Font.Color:=RGB(122,156,250); //Голубой
              //Label17.Font.Color:=RGB(234,239,45); //Желтый
              Cvet1:=1;
            End;
         end;
                UOsnova_Main:=Osnova_Main.Create() ;
                fileName :=Plan.Caption;// ExtractFileDir(ParamStr(0));//+'\Klapan.EXE';
                Vn_Dat := FormatDateTime('dd.mm.yyyy', DateTimePicker1.Date);
                UOsnova_Main.Flag_Error :=0;
                UOsnova_Main.Osnova( Vn_Dat,'Запуск','Специф',#39+Label8.Caption+#39,fileName, 1,Cvet1);
                //++++++++++++++++++++++++++++++++++++++++++++++++++
                if  UOsnova_Main.Flag_Error =2 Then //ДЕТАЛИ НЕ ВСЕ
                //ЗАНЕСТИ В ЗАПУСК И Поставить флаг ФлагЗаготовки и убрать Дату Планирования
                Begin
                        //++++++++++++++++++++++++++++++++++++++++++++++++++++
                        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                        'UPDATE %s SET [ФлагЗаготовки]=' + #39
                        + '0' + #39+ ',Планирование=NULL WHERE [Номер]=' + #39 +
                        (Label8.Caption) + #39, ['Запуск'])
                        then
                        Exit;
                        Form1.Memo2.Lines.Add('UOsnova_Main.Osnova( ДЕТАЛИ НЕ ВСЕ,Запуск,Специф,'+Label8.Caption+','+fileName+' )UOsnova_Main.Flag_Error =2') ;
                        Exit;
                end;
                //++++++++++++++++++++++++++++++++++++++++++++++++++
                if  UOsnova_Main.Flag_Error =1 Then //ДАТА ЛЕВАЯ УДАЛИТЬ Из ЗАПУСКА
                                                //И UPDATE КЛАПАНА НЕ ДЕЛАТЬ
                Begin

                        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                        'UPDATE %s SET Планирование=NULL WHERE [Номер]=' + #39 +
                        (Label8.Caption) + #39, ['Запуск'])
                        then
                        Exit;
                        Form1.Memo2.Lines.Add('UOsnova_Main.Osnova( ДАТА НЕ ТА - '+Vn_Dat+',Запуск,Специф,'+Label8.Caption+','+fileName+' )UOsnova_Main.Flag_Error =1') ;
                        Exit;
                end;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label2.Caption + ']=' + #39 + Str + #39 +',Cvet=' + #39 + IntToStr(Cvet1) + #39 +
                ' WHERE (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') ', ['Запуск']) then
                Exit;
                        Form1.ZSG.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:=Str1;
          end;}

        end;
        //OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOAND (БЗ=' + #39 +BZ+ #39 + ')
        If Form1.Vozduh=2 Then
        Begin
          BZ:=Label3BZ.Caption;
          if (Otgr=0 ) Then
          Begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 Label2.Caption + ']=' + #39 + Str + #39 +
                ' WHERE ([Заказ]=' + #39 + Label6.Caption + #39 + ') AND (' +
                FN_NAM + '=' +
                #39 + Label3.Caption + #39 + ') AND (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ')  ', ['[Запуск750]']) then
                Exit;
                Form1.zclrstrngrd1.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:=Str1;
          end;
          if (Pods=0 ) Then     // AND (БЗ=' + #39 +BZ+ #39 + ')
          Begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 Label2.Caption + ']=' + #39 + Str + #39 +
                ' WHERE ([Заказ]=' + #39 + Label6.Caption + #39 + ') AND (' +
                FN_NAM + '=' +
                #39 + Label3.Caption + #39 + ') AND (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') ', ['[Запуск750]']) then
                Exit;
                Form1.zclrstrngrd1.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:=Str1;

          end;
          if Res=0 Then  //Заготовка Запуск   ([Заказ]=' + #39 + Label6.Caption + #39 + ') AND (' +
               // FN_NAM + '=' +
               // #39 + Label3.Caption + #39 + ') AND     AND (БЗ=' + #39 +BZ+ #39 + ')
          Begin
               if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label1.Caption
                + ' ' + Label2.Caption + ']=' + #39 + Str + #39 +
                ' WHERE  (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') ', ['[Запуск750]']) then
                Exit;
                      if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label1.Caption + ' ' + Label2.Caption + ']=' + #39 + Str
                        + #39 +

                        ' WHERE ([Изделие]=' + #39 +
                        Label3.Caption + #39 + ') AND ([Заказ]=' + #39 +
                        Label6.Caption + #39 +
                        ') '  , ['[750]']) then
                        Exit;
                Form1.zclrstrngrd1.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:=Str1;
          end;
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
          if Res1 = 0 then //Заготовка Готовность
          begin
               { }

                if not Form1.mkQuerySelect1(Form1.ADOQuery2,
                        'Select * from %s Where  (Заказ=' + #39 + Label6.Caption + #39 + ') AND (Изделие=' +
                        #39 + Label3.Caption + #39 + ') ',
                        ['[750]']) then
                        exit;


                Zag_S := Form1.ADOQuery2.FieldByName('Заготовка').AsString;
                Kol_Zap := Label7.Caption;
                Dat:=Label10.Caption;
                if Zag_S = '' then
                        Zag := 0
                else
                        Zag := StrToInt(Zag_S);
                if Dat='' Then
                        Zag := Zag - StrToInt(Kol_Zap)
                Else
                        Zag:=Zag;
                if Zag<0 Then
                Zag:=0;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label1.Caption + ' ' + Label2.Caption + ']=' + #39 + Str
                        + #39 +
                        ',Заготовка=' + #39 + IntToStr(Zag) + #39 +
                        ' WHERE ([Изделие]=' + #39 +
                        Label3.Caption + #39 + ') AND ([Заказ]=' + #39 +
                        Label6.Caption + #39 +
                        ') '  , ['[750]']) then
                        Exit;
                        Form1.ADOQuery2.Next;

                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label1.Caption
                + ' ' + Label2.Caption + ']=' + #39 + Str + #39 +
                ' WHERE  (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') AND ([Изделие]=' + #39 +
                        Label3.Caption + #39 + ') AND ([Заказ]=' + #39 +
                        Label6.Caption + #39 +
                        ')  ', ['[Запуск750]']) then
                Exit;
                Form1.zclrstrngrd1.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:=Str1;
          end;
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
          if Res2 = 0 then //Сборка Запуск
          begin
                if not Form1.mkQuerySelect1(Form1.ADOQuery2,
                        'Select * from %s Where  (Заказ=' + #39 + Label6.Caption + #39 +
                        ') AND (Изделие=' + #39 + Label3.Caption + #39 + ') ',
                        ['[750]']) then
                        exit;
                Sbor_S := Form1.ADOQuery2.FieldByName('Сборка').AsString;
                Zap1:=StrToInt(Form1.ADOQuery2.FieldByName('Запуск1').AsString);
                Dat:=Label10.Caption;
                Kol_Zap := Label7.Caption;
                if Sbor_S = '' then
                        Sbor := 0
                else
                        Sbor := StrToInt(Sbor_S);
                if Dat<>'' Then
                        Sbor := Sbor + (0)
                Else
                        Sbor := Sbor + StrToInt(Kol_ZAp);
                if Dat<>'' Then
                        Zap1 := Zap1 + (0)
                Else
                        Zap1 := Zap1 + StrToInt(Kol_ZAp);

                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label1.Caption + ' ' + Label2.Caption + ']=' + #39 + Str
                        + #39 +
                        ',Сборка=' + #39 + IntToStr(Sbor) + #39 +
                        ',Запуск1=' + #39 + IntToStr(Zap1) + #39 +' WHERE  ([Изделие]=' + #39 +
                        Label3.Caption + #39 + ') AND ([Заказ]=' + #39 +
                        Label6.Caption + #39 +
                        ') '  , ['[750]']) then
                        Exit;

              if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label1.Caption
                + ' ' + Label2.Caption + ']=' + #39 + Str + #39 +
                ' WHERE  (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') AND ([Изделие]=' + #39 +
                Label3.Caption + #39 + ') AND ([Заказ]=' + #39 +
                Label6.Caption + #39 +
                ') ', ['[Запуск750]']) then
                Exit;
                Form1.zclrstrngrd1.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:=Str1;
          end;
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
          if Res3 = 0 then //Сборка Готовность
          begin
                if not Form1.mkQuerySelect1(Form1.ADOQuery2,
                        'Select * from %s Where   (Заказ=' + #39 + Label6.Caption
                        + #39 +
                        ') AND (Изделие=' + #39 + Label3.Caption + #39 + ') ',
                        ['[750]']) then
                        exit;
                Sbor_S := Form1.ADOQuery2.FieldByName('Сборка').AsString;
                Zap1:=StrToInt(Form1.ADOQuery2.FieldByName('Запуск1').AsString);
                Dat:=Label10.Caption;
                Kol_Zap := Label7.Caption;
                if Sbor_S = '' then
                        Sbor := 0
                else
                        Sbor := StrToInt(Sbor_S);
                if Dat<>'' Then
                        Sbor:=Sbor
                Else
                        Sbor := Sbor - StrToInt(Kol_Zap);
                if Sbor<0 Then
                        Sbor:=0;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label1.Caption + ' ' + Label2.Caption + ']=' + #39 + Str
                        + #39 +
                        ',Сборка=' + #39 + IntToStr(Sbor) + #39 +
                        ' WHERE ([Изделие]=' + #39 +
                        Label3.Caption + #39 + ') AND ([Заказ]=' + #39 +
                        Label6.Caption + #39 +
                        ') ', ['[750]']) then
                        Exit;
                        //       AND (БЗ=' + #39 +BZ+ #39 + ')
               if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label1.Caption
                + ' ' + Label2.Caption + ']=' + #39 + Str + #39 +
                ' WHERE ([Заказ]=' + #39 + Label6.Caption + #39 + ') AND (' +
                FN_NAM + '=' +
                #39 + Label3.Caption + #39 + ') AND (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ')', ['[Запуск750]']) then
                Exit;
                Form1.zclrstrngrd1.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:=Str1;
          end;
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
          if Res4 = 0 then  //Планирование    //([Заказ]=' + #39 + Label6.Caption + #39 + ') AND (' +
                //FN_NAM + '=' +
                //#39 + Label3.Caption + #39 + ') AND  AND (БЗ=' + #39 +BZ+ #39 + ')
          begin

            S:=FormatDateTime('mm.dd.YYYY',DateTimePicker1.Date);
            if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select  * from %s Where ([Заказ]=' + #39 + Label6.Caption + #39 + ') AND (' +
                FN_NAM + '=' +
                #39 + Label3.Caption + #39 + ') AND (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') ' ,
                ['[Запуск750]']) then
                exit;

            I:=Form1.ADOQuery1.RecordCount;
            Cvet:= Form1.ADOQuery1.FieldByName('Cvet').AsInteger;
        case Cvet of
            0:
            begin

              Cvet1:=1;
            End;
            1:
            begin

              Cvet1:=2;
            End;
            2:
            begin

              Cvet1:=3;
            End;
            3:
            begin

              Cvet1:=1;
            End;

        end;           // AND (БЗ=' + #39 +BZ+ #39 + ')
         if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET Планирование=' + #39 + Str + #39 +
                ' WHERE ([Заказ]=' + #39 + Label6.Caption + #39 + ') AND (' +
                FN_NAM + '=' +
                #39 + Label3.Caption + #39 + ') AND (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ')  ', ['[Запуск750]']) then
                Exit;
                Form1.zclrstrngrd1.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:=Str1;
          end;
        end;

        if Form1.Vozduh=1 Then
        Begin
          if (Otgr=0 ) Then
          Begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 Label2.Caption + ']=' + #39 + Str + #39 +
                ' WHERE ([ID]=' + #39 + Form1.ZCV.Cells[I_FN_KOL_ZAP + 14,Form1.ZCV.Row] + #39 + ') AND  ([Заказ]=' + #39 + Label6.Caption + #39 + ') AND (' +
                FN_NAM + '=' +
                #39 + Label3.Caption + #39 + ') AND (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') ', ['ЗапускВозд']) then
                Exit;
                Form1.ZCV.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:=Str1;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 Label2.Caption + ']=' + #39 + Str + #39 +
                ' WHERE ([Заказ]=' + #39 + Label6.Caption + #39 + ') AND (' +
                FN_NAM + '=' +
                #39 + Label3.Caption + #39 + ')  ', ['KlapanaZap']) then
                Exit;
          end;
          if (Pods=0 ) Then
          Begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 Label2.Caption + ']=' + #39 + Str + #39 +
                ' WHERE  ([ID]=' + #39 + Form1.ZCV.Cells[I_FN_KOL_ZAP + 14,Form1.ZCV.Row] + #39 + ') AND ([Заказ]=' + #39 + Label6.Caption + #39 + ') AND (' +
                FN_NAM + '=' +
                #39 + Label3.Caption + #39 + ') AND (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') ', ['ЗапускВозд']) then
                Exit;
                Form1.ZCV.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:=Str1;

          end;
          if Res=0 Then //Заготовка ЗапускВозд      ([ID]=' + #39 + Form1.ZCV.Cells[I_FN_KOL_ZAP + 14,Form1.ZCV.Row] + #39 + ') AND ([Заказ]=' + #39 + Label6.Caption + #39 + ') AND (' +
                //FN_NAM + '=' +
                //#39 + Label3.Caption + #39 + ') AND'}
          Begin
            if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label1.Caption
                + ' ' + Label2.Caption + ']=' + #39 + Str + #39 +
                ' WHERE  (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') ', ['ЗапускВозд']) then
                Exit;
            if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label1.Caption
                + ' ' + Label2.Caption + ']=' + #39 + Str + #39 +
                ' WHERE  (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') ', ['KlapanaZap']) then
                Exit;
                Form1.ZCV.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:=Str1;
          end;
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
          if Res1 = 0 then //Заготовка Готовность
          begin
               { }

                if not Form1.mkQuerySelect1(Form1.ADOQuery2,
                        'Select * from %s Where  (Заказ=' + #39 + Label6.Caption + #39 + ') AND (Изделие=' + #39 + Label3.Caption + #39 + ') ',
                        ['KlapanaZap']) then
                        exit;

                Zag_S := Form1.ADOQuery2.FieldByName('Заготовка').AsString;
                Kol_Zap := Label7.Caption;
                Dat:=Label10.Caption;
                if Zag_S = '' then
                        Zag := 0
                else
                        Zag := StrToInt(Zag_S);
                if Dat='' Then
                        Zag := Zag - StrToInt(Kol_Zap)
                Else
                        Zag:=Zag;
                if Zag<0 Then
                Zag:=0;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label1.Caption + ' ' + Label2.Caption + ']=' + #39 + Str
                        + #39 +
                        ',Заготовка=' + #39 + IntToStr(Zag) + #39 +
                        ' WHERE ([Изделие]=' + #39 +
                        Label3.Caption + #39 + ') AND ([Заказ]=' + #39 +
                        Label6.Caption + #39 +
                        ') '  , ['KlapanaZap']) then
                        Exit;
                        Form1.ADOQuery2.Next;

                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label1.Caption
                + ' ' + Label2.Caption + ']=' + #39 + Str + #39 +
                ' WHERE  ([ID]=' + #39 + Form1.ZCV.Cells[I_FN_KOL_ZAP + 14,Form1.ZCV.Row] + #39 + ') AND  (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') AND ([Изделие]=' + #39 +
                        Label3.Caption + #39 + ') AND ([Заказ]=' + #39 +
                        Label6.Caption + #39 +
                        ')  ', ['ЗапускВозд']) then
                Exit;

                Form1.ZCV.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:=Str1;
          end;
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
          if Res2 = 0 then //Сборка ЗапускВозд
          begin
                if not Form1.mkQuerySelect1(Form1.ADOQuery2,
                        'Select * from %s Where  (Заказ=' + #39 + Label6.Caption + #39 + ') AND (Изделие=' + #39 + Label3.Caption + #39 + ') ',
                ['KlapanaZap']) then
                exit;
                Sbor_S := Form1.ADOQuery2.FieldByName('Сборка').AsString;
                Zap1:=StrToInt(Form1.ADOQuery2.FieldByName('Запуск1').AsString);
                Dat:=Label10.Caption;
                Kol_Zap := Label7.Caption;
                if Sbor_S = '' then
                        Sbor := 0
                else
                        Sbor := StrToInt(Sbor_S);
                if Dat<>'' Then
                        Sbor := Sbor + (0)
                Else
                        Sbor := Sbor + StrToInt(Kol_ZAp);
                if Dat<>'' Then
                        Zap1 := Zap1 + (0)
                Else
                        Zap1 := Zap1 + StrToInt(Kol_ZAp);

                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label1.Caption + ' ' + Label2.Caption + ']=' + #39 + Str
                        + #39 +
                        ',Сборка=' + #39 + IntToStr(Sbor) + #39 +
                        ',Запуск1=' + #39 + IntToStr(Zap1) + #39 +' WHERE  ([Изделие]=' + #39 +
                        Label3.Caption + #39 + ') AND ([Заказ]=' + #39 +
                        Label6.Caption + #39 +
                        ') '  , ['KlapanaZap']) then
                        Exit;

                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label1.Caption
                + ' ' + Label2.Caption + ']=' + #39 + Str + #39 +
                ' WHERE  ([ID]=' + #39 + Form1.ZCV.Cells[I_FN_KOL_ZAP + 14,Form1.ZCV.Row] + #39 + ') AND  (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') AND ([Изделие]=' + #39 +
                        Label3.Caption + #39 + ') AND ([Заказ]=' + #39 +
                        Label6.Caption + #39 +
                        ')  ', ['ЗапускВозд']) then
                Exit;

                Form1.ZCV.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:=Str1;
          end;
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
          if Res3 = 0 then //Сборка Готовность
          begin
                if not Form1.mkQuerySelect1(Form1.ADOQuery2,
                        'Select * from %s Where   (Заказ=' + #39 + Label6.Caption
                        + #39 +
                        ') AND (Изделие=' + #39 + Label3.Caption + #39 + ') ',
                        ['KlapanaZap']) then
                        exit;
                Sbor_S := Form1.ADOQuery2.FieldByName('Сборка').AsString;
                Zap1:=StrToInt(Form1.ADOQuery2.FieldByName('Запуск1').AsString);
                Dat:=Label10.Caption;
                Kol_Zap := Label7.Caption;
                if Sbor_S = '' then
                        Sbor := 0
                else
                        Sbor := StrToInt(Sbor_S);
                if Dat<>'' Then
                        Sbor:=Sbor
                Else
                        Sbor := Sbor - StrToInt(Kol_Zap);
                if Sbor<0 Then
                        Sbor:=0;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label1.Caption + ' ' + Label2.Caption + ']=' + #39 + Str
                + #39 +',Сборка=' + #39 + IntToStr(Sbor) + #39 +
                ' WHERE   ([Изделие]=' + #39 +
                Label3.Caption + #39 + ') AND ([Заказ]=' + #39 +
                Label6.Caption + #39 +
                ') ', ['KlapanaZap']) then
                Exit;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label1.Caption
                + ' ' + Label2.Caption + ']=' + #39 + Str + #39 +
                ' WHERE  ([ID]=' + #39 + Form1.ZCV.Cells[I_FN_KOL_ZAP + 14,Form1.ZCV.Row] + #39 + ') AND ([Заказ]=' + #39 + Label6.Caption + #39 + ') AND (' +
                FN_NAM + '=' +
                #39 + Label3.Caption + #39 + ') AND (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') ', ['ЗапускВозд']) then
                Exit;
                        Form1.ZCV.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:=Str1;
          end;
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
          if Res4 = 0 then  //([Заказ]=' + #39 + Label6.Caption + #39 + ') AND (' +
                //FN_NAM + '=' +
                //#39 + Label3.Caption + #39 + ') AND
          begin

         S:=FormatDateTime('mm.dd.YYYY',DateTimePicker1.Date);
              res4:=AnsiCompareStr(Plan.Caption,S);
              if Res4=0 then
              Close;
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select  max(планирование) as S from %s Where планирование<'+#39+S+#39 ,
                ['ЗапускВозд']) then
                exit;
        D:= Form1.ADOQuery1.FieldByName('S').AsDateTime;
//+++++++++++++++++++++++++++++++
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select  max(планирование) as S from %s Where планирование<'+#39+S+#39 ,
                ['Запуск']) then
                exit;
        D1:= Form1.ADOQuery1.FieldByName('S').AsDateTime;

        if D>D1 then
        Begin
        S1:=FormatDateTime('mm.dd.YYYY',D);
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select  * from %s Where планирование='+#39+S1+#39 ,
                ['ЗапускВозд']) then
                exit;
        End
        Else
        Begin
        S1:=FormatDateTime('mm.dd.YYYY',D1);
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select  * from %s Where планирование='+#39+S1+#39 ,
                ['Запуск']) then
                exit;
        End;
        I:=Form1.ADOQuery1.RecordCount;
        Cvet:= Form1.ADOQuery1.FieldByName('Cvet').AsInteger;
        case Cvet of
            0:
            begin
              ///Label16.Font.Color:=RGB(255,255,255);//Белый
              //Label17.Font.Color:=RGB(234,239,45); //Желтый
              Cvet1:=1;
            End;
            1:
            begin
              //Label16.Font.Color:=RGB(234,239,45); //Желтый
              //Label17.Font.Color:=RGB(41,243,104);//Зеленый
              Cvet1:=2;
            End;
            2:
            begin
             // Label16.Font.Color:=RGB(41,243,104);//Зеленый
              //Label17.Font.Color:=RGB(122,156,250); //Голубой
              Cvet1:=3;
            End;
            3:
            begin
              //Label16.Font.Color:=RGB(122,156,250); //Голубой
              //Label17.Font.Color:=RGB(234,239,45); //Желтый
              Cvet1:=1;
            End;
        end;
                UOsnova_Main:=Osnova_Main.Create() ;
                fileName := Plan.Caption;//ExtractFileDir(ParamStr(0));//+'\Klapan.EXE';
                Vn_Dat := FormatDateTime('dd.mm.yyyy', DateTimePicker1.Date);
                UOsnova_Main.Flag_Error :=0;
                UOsnova_Main.Osnova( Vn_Dat,'ЗапускВозд','СпецифВозд',#39+Label8.Caption+#39,fileName, 1,Cvet1);
                //++++++++++++++++++++++++++++++++++++++++++++++++++
                if  UOsnova_Main.Flag_Error =2 Then //ДЕТАЛИ НЕ ВСЕ
                //ЗАНЕСТИ В ЗапускВозд И Поставить флаг ФлагЗаготовки и убрать Дату Планирования
                Begin
                        //++++++++++++++++++++++++++++++++++++++++++++++++++++
                        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                        'UPDATE %s SET [ФлагЗаготовки]=' + #39
                        + '0' + #39+ ',Планирование=NULL WHERE [Номер]=' + #39 +
                        (Label8.Caption) + #39, ['ЗапускВозд'])
                        then
                        Exit;
                        Form1.Memo2.Lines.Add('UOsnova_Main.Osnova( ДЕТАЛИ НЕ ВСЕ,ЗапускВозд,СпецифВозд,'+Label8.Caption+','+fileName+' )UOsnova_Main.Flag_Error =2') ;
                        Exit;
                end;
                //++++++++++++++++++++++++++++++++++++++++++++++++++
                if  UOsnova_Main.Flag_Error =1 Then //ДАТА ЛЕВАЯ УДАЛИТЬ Из ЗАПУСКА
                                                //И UPDATE КЛАПАНА НЕ ДЕЛАТЬ
                Begin
                        {if not Form1.mkQueryDelete( Form1.ADOQuery1, 'DELETE FROM %s Where (Номер= '
                        +#39+Label8.Caption+#39+') '  ,
                        ['ЗапускВозд'] )
                        Then
                        Exit; }
                        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                        'UPDATE %s SET Планирование=NULL WHERE [Номер]=' + #39 +
                        (Label8.Caption) + #39, ['ЗапускВозд'])
                        then
                        Exit;
                        Form1.Memo2.Lines.Add('UOsnova_Main.Osnova( ДАТА НЕ ТА - '+Vn_Dat+',ЗапускВозд,СпецифВозд,'+Label8.Caption+','+fileName+' )UOsnova_Main.Flag_Error =1') ;
                        Exit;
                end;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label2.Caption + ']=' + #39 + Str + #39 +',Cvet=' + #39 + IntToStr(Cvet1) + #39 +
                ' WHERE (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') ', ['ЗапускВозд']) then
                Exit;
                        Form1.ZCV.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:=Str1;
          end;

        end;

        Form1.Button55.Click;
        Close;
end;

procedure TForm2.Button1KeyPress(Sender: TObject; var Key: Char);
begin
     If Key=#27 Then
     Close;
end;

procedure TForm2.Button3Click(Sender: TObject);
Var
Col,Otgr:Integer;
begin
        Col:=StrToInt(Label5.Caption);
        if Form1.Vozduh=0 Then
        Begin
        if (Col=I_FN_KOL_ZAP + 11) then
        begin
        if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                         Label2.Caption + ']=NULL' + //#39 + '' + #39 +
                        ' WHERE ([Изделие]=' + #39 +
                        Label3.Caption + #39 + ') AND ([Заказ]=' + #39 +
                        Label6.Caption + #39 +
                        ') ', ['Klapana']) then
                        Exit;
        if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label2.Caption + ']=NULL' + //#39 + '' + #39 +
                ' WHERE  ([ID]=' + #39 + Form1.ZSG.Cells[I_FN_KOL_ZAP + 14,Form1.ZSG.Row] + #39 + ') AND ([Заказ]=' + #39 + Label6.Caption + #39 + ') AND (' +
                FN_NAM + '=' +
                #39 + Label3.Caption + #39 + ') AND (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') ', ['Запуск']) then
                Exit;
        end
        else
        begin
        if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label1.Caption + ' ' + Label2.Caption + ']=NULL' + //#39 + '' + #39 +
                        ' WHERE ([Изделие]=' + #39 +
                        Label3.Caption + #39 + ') AND ([Заказ]=' + #39 +
                        Label6.Caption + #39 +
                        ') ', ['Klapana']) then
                        Exit;
        if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label1.Caption
                + ' ' + Label2.Caption + ']=NULL' + //#39 + '' + #39 +
                ' WHERE  ([ID]=' + #39 + Form1.ZSG.Cells[I_FN_KOL_ZAP + 14,Form1.ZSG.Row] + #39 + ') AND ([Заказ]=' + #39 + Label6.Caption + #39 + ') AND (' +
                FN_NAM + '=' +
                #39 + Label3.Caption + #39 + ') AND (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') ', ['Запуск']) then
                Exit;
        end;
        Form1.ZSG.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:='';
        End;
        ///_____________++++++++++++=================================------
        if Form1.Vozduh=1 Then
        Begin
        if (Col=I_FN_KOL_ZAP + 11) then
        begin
        if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                         Label2.Caption + ']=NULL' + //#39 + '' + #39 +
                        ' WHERE ([Изделие]=' + #39 +
                        Label3.Caption + #39 + ') AND ([Заказ]=' + #39 +
                        Label6.Caption + #39 +
                        ') ', ['KlapanaZap']) then
                        Exit;
        if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label2.Caption + ']=NULL' + //#39 + '' + #39 +
                ' WHERE ([Заказ]=' + #39 + Label6.Caption + #39 + ') AND (' +
                FN_NAM + '=' +
                #39 + Label3.Caption + #39 + ') AND (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') ', ['ЗапускВозд']) then
                Exit;
        end
        else
        begin
        if Label1.Caption='' then
        begin
              if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                         Label2.Caption + ']=NULL' + //#39 + '' + #39 +
                        ' WHERE ([Изделие]=' + #39 +
                        Label3.Caption + #39 + ') AND ([Заказ]=' + #39 +
                        Label6.Caption + #39 +
                        ') ', ['KlapanaZap']) then
                        Exit;
            if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 Label2.Caption + ']=NULL' + //#39 + '' + #39 +
                ' WHERE ([Заказ]=' + #39 + Label6.Caption + #39 + ') AND (' +
                FN_NAM + '=' +
                #39 + Label3.Caption + #39 + ') AND (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') ', ['ЗапускВозд']) then
                Exit;
        end
        else
        begin


            if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        Label1.Caption + ' ' + Label2.Caption + ']=NULL' + //#39 + '' + #39 +
                        ' WHERE ([Изделие]=' + #39 +
                        Label3.Caption + #39 + ') AND ([Заказ]=' + #39 +
                        Label6.Caption + #39 +
                        ') ', ['KlapanaZap']) then
                        Exit;
            if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label1.Caption
                + ' ' + Label2.Caption + ']=NULL' + //#39 + '' + #39 +
                ' WHERE ([Заказ]=' + #39 + Label6.Caption + #39 + ') AND (' +
                FN_NAM + '=' +
                #39 + Label3.Caption + #39 + ') AND (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') ', ['ЗапускВозд']) then
                Exit;
        end;
        end;
        Form1.ZCV.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:='';
        End;

        if Form1.Vozduh=2 Then
        Begin
          Otgr:=AnsiCompareStr('Упаковка',Label2.Caption);
          if (Otgr=0 ) Then
          Begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                 Label2.Caption + ']=NULL  WHERE  (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') AND ([Изделие]=' + #39 +
                        Label3.Caption + #39 + ') AND ([Заказ]=' + #39 +
                        Label6.Caption + #39 +
                        ')  ', ['[Запуск750]']) then
                Exit;
                Form1.zclrstrngrd1.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:='';
          end
          else
          Begin
              if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                Label1.Caption
                + ' ' + Label2.Caption + ']=NULL  WHERE  (' + FN_NOM + '=' + #39 +
                Label8.Caption
                + #39 + ') AND ([Изделие]=' + #39 +
                        Label3.Caption + #39 + ') AND ([Заказ]=' + #39 +
                        Label6.Caption + #39 +
                        ')  ', ['[Запуск750]']) then
                Exit;
                Form1.zclrstrngrd1.Cells[StrToInt(Label5.Caption),StrToInt(Label4.Caption)]:='';
          End;
        End;
        Close;
end;

procedure TForm2.FormKeyPress(Sender: TObject; var Key: Char);
begin
     If Key=#27 Then
     Close;
        if (Key = #13) then
                Button2Click(nil);
end;

procedure TForm2.DateTimePicker1KeyPress(Sender: TObject; var Key: Char);
begin
     If Key=#27 Then
     Close;
        if (Key = #13) then
                Button2Click(nil);
end;

procedure TForm2.Button2KeyPress(Sender: TObject; var Key: Char);
begin
     If Key=#27 Then
     Close;
        if (Key = #13) then
                Button2Click(nil);
end;

procedure TForm2.FormShow(Sender: TObject);
begin
        Form1.Memo2.Lines.Add('Форма');
        if (Form1.FlagDolg = 1) or (Form1.FlagDolg = 7)  or (Form1.FlagDolg = 4) then
        button3.Visible:=True;
end;

end.
