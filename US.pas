unit US;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, JvExControls, JvGradient, StdCtrls, GradientLabel,
  Grids, AdvObj, BaseGrid, AdvGrid, ExtCtrls, ComObj,
  UConnCeh, UConnKlap,UOsnova;

type
  TFS = class(TForm)
    pnl1: TPanel;
    SG21: TAdvStringGrid;
    lbl1: TGradientLabel;
    lbl2: TGradientLabel;
    jvgrdnt1: TJvGradient;
    dtp1: TDateTimePicker;
    lbl3: TLabel;
    lbl4: TLabel;
    lbl5: TLabel;
    lbl6: TLabel;
    lbl7: TLabel;
    btn1: TButton;
    btn2: TButton;
    btn3: TButton;
    btn4: TButton;
    Edit: TEdit;
    btn5: TButton;
    procedure FormShow(Sender: TObject);
    procedure btn3Click(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure dtp1Click(Sender: TObject);
    procedure btn2Click(Sender: TObject);
    procedure SG21SetEditText(Sender: TObject; ACol, ARow: Integer;
      const Value: String);
  private
    { Private declarations }
  public
    { Public declarations }
    UOsnova_Main:Osnova_Main ;
  end;

var
  FS: TFS;

implementation

uses Main;

{$R *.dfm}

procedure TFS.FormShow(Sender: TObject);
var
        I, Kol, Kol_Zap, Max: Integer;
        S,B: string;
begin
        //Btn2.Enabled := False;
        //Enabled_Ok1 := 0;
        //Enabled_Ok2 := 0;
        Lbl1.Caption := '0';
        Lbl4.Caption := '0';
        Lbl6.Caption := '0';
        Dtp1.DateTime:=Now;
        for i:=0 To SG21.RowCount Do
        Begin
            SG21.Rows[i].Clear;
        End;
        if Form1.FlagDolg=1 Then
        Begin
            Btn4.Visible := True;
            Edit.Visible:=True;
            Edit.Enabled:=True;
        end;
        btn1.Click;
end;

procedure TFS.btn3Click(Sender: TObject);
begin
        Close;
end;

procedure TFS.btn1Click(Sender: TObject);
var
  Izdel:String;
  kol,Kol_Zap,I,MAX,Res,res1:Integer;
begin
                SG21.Visible:=True;
                SG21.ColCount := 36;
                SG21.FixedRightCols:=22;
                SG21.Cells[0, 0] := '№';
                SG21.ColWidths[0] := 60;
                SG21.Cells[1, 0] :='Дата' ;
                SG21.Cells[2, 0] := FN_ZAK;
                SG21.Cells[3, 0] := FN_KOL;
                SG21.Cells[4, 0] := 'Свободно';
                SG21.Cells[5, 0] := 'Кол-во на запуск';
                SG21.Cells[6, 0] := 'Изделие';

                SG21.Cells[7, 0] := 'Н\ч РУБКА на единицу';
                SG21.Cells[8, 0] := 'Н\ч РУБКА всего';

                SG21.Cells[9, 0] := 'Н\ч ГИБКА на единицу';
                SG21.Cells[10, 0] := 'Н\ч ГИБКА всего';

                SG21.Cells[11, 0] := 'Н\ч СВАРКА на единицу';
                SG21.Cells[12, 0] := 'Н\ч СВАРКА всего';

                SG21.Cells[13, 0] := 'Н\ч ПОКРАСКА на единицу';
                SG21.Cells[14, 0] := 'Н\ч ПОКРАСКА всего';

                SG21.Cells[15, 0] := 'Н\ч СБОРКА на единицу';
                SG21.Cells[16, 0] := 'Н\ч СБОРКА всего';

                SG21.Cells[17, 0] := 'IdГП';
                SG21.Cells[18, 0] := 'Кол запущенных';
                SG21.Cells[19, 0] := 'Кол Раз';
                SG21.Cells[20, 0] := 'Номер Раз';
                SG21.Cells[21, 0] :='Расчетная дата готовности' ;
                SG21.Cells[22, 0] := 'НачНомер';
                SG21.Cells[23, 0] := 'КонНомер';
                SG21.Cells[24, 0] := 'НачКод';
                SG21.Cells[25, 0] := 'КонКод';//Заводск номер из 20 ячейкй
                SG21.Cells[26, 0] := 'НомерПос';
                SG21.Cells[27, 0] := 'КодПос';
                SG21.Cells[28, 0] := 'Мат';
                SG21.Cells[29, 0] := 'План дата';
                SG21.Cells[30, 0] := 'План неделя';
                SG21.Cells[31, 0] := 'БЗ';
                SG21.Cells[32, 0] := 'Исполнение';
                SG21.Cells[33, 0] := 'Примечания';
                SG21.Cells[34, 0] := 'Клапан';
                SG21.Cells[35, 0] := 'КолКлапан';
                SG21.ColWidths[7] := 0;
                SG21.ColWidths[8] := 0;
                SG21.ColWidths[9] := 0;
                SG21.ColWidths[10] := 0;
                SG21.ColWidths[11] := 0;
                SG21.ColWidths[12] := 0;
                SG21.ColWidths[13] := 0;
                SG21.ColWidths[14] := 0;
                SG21.ColWidths[15] := 0;
                SG21.ColWidths[16] := 0;
                SG21.ColWidths[17] := 0;
                SG21.ColWidths[18] := 0;
                SG21.ColWidths[19] := 0;
                SG21.ColWidths[20] := 0;
                SG21.ColWidths[22] := 0;
                SG21.ColWidths[23] := 0;
                SG21.ColWidths[24] := 0;
                SG21.ColWidths[25] := 0;
                SG21.ColWidths[26] := 0;
                SG21.ColWidths[27] := 0;
    if Form1.Luk=0 then
    begin
                if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s Where'+
                ' ([' + FN_KOL_ZAP + ']<[' + FN_KOL + ']) AND'+
              ' ([' + 'Статус' + ']=' + '1' + ') AND'+
                ' ([Отмена] IS NULL) AND (L=0) AND (ОТК IS NULL) AND (ПР='+#39+'МИ'+#39+') AND (NOT Технолог IS NULL)  '+
                        ' Order By Заказ ', ['СТАМ']) then
                exit;
    end;
     if Form1.Luk=1 then
    begin
                if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s Where'+
                ' ([' + FN_KOL_ZAP + ']<[' + FN_KOL + ']) AND'+
             // ' ([' + 'Статус' + ']=' + '1' + ') AND'+
                ' ([Отмена] IS NULL) AND (L=0) AND (ОТК IS NULL) AND (ПР='+#39+'МИ'+#39+') AND (NOT Технолог IS NULL)  '+
                        ' Order By Заказ ', ['ЛЮК']) then
                exit;
    end;
                SG21.RowCount := Form1.ADOQuery1.RecordCount + 1;

                for I := 0 to Form1.ADOQuery1.RecordCount - 1 do
                begin
                        Kol := Form1.ADOQuery1.FieldByName(FN_KOL).AsInteger;
                        Kol_Zap := Form1.ADOQuery1.FieldByName(FN_KOL_ZAP).AsInteger;
                        SG21.Cells[0, i + 1] :=IntToStr(I+1);
                        SG21.Cells[1, i + 1] :=
                        Form1.ADOQuery1.FieldByName('Дата').AsString;
                        SG21.Cells[2, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_ZAK).AsString;
                        SG21.Cells[3, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_KOL).AsString;

                        Kol := Kol - Kol_Zap;
                        SG21.Cells[4, i + 1] := IntToStr(Kol);
                        Izdel:= Form1.ADOQuery1.FieldByName('Изделие').AsString;
                        Res:=Pos('СТАМ',Izdel);
                        Res1:=Pos('Защита',Izdel);
                        If (Res<>0) and (Res1=0) then
                        Delete(Izdel,1,Res-1);
                        SG21.Cells[6, i + 1] :=Izdel;

                        SG21.Cells[15, i + 1] :=
                        Form1.ADOQuery1.FieldByName('Н\ч Сборка').AsString;
                        SG21.Cells[17, i + 1] :=
                        Form1.ADOQuery1.FieldByName('IdГП').AsString;
                        SG21.Cells[18, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_KOL_ZAP).AsString;
                        SG21.Cells[19, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_KOL_RAZ).AsString;
                        SG21.Cells[20, i + 1] :=
                        Form1.ADOQuery1.FieldByName(FN_NOMER_RAZ).AsString;

                        SG21.Cells[21, i + 1] :=
                        Form1.ADOQuery1.FieldByName('Расчетная дата готовности').AsString;

                        SG21.Cells[22, i + 1] :=
                        Form1.ADOQuery1.FieldByName('НачНомер').AsString;
                        SG21.Cells[23, i + 1] :=
                        Form1.ADOQuery1.FieldByName('КонНомер').AsString;
                        SG21.Cells[24, i + 1] :=
                        Form1.ADOQuery1.FieldByName('НачКод').AsString;
                        SG21.Cells[25, i + 1] :=
                        Form1.ADOQuery1.FieldByName('КонКод').AsString;

                        SG21.Cells[26, i + 1] :=
                        Form1.ADOQuery1.FieldByName('НомерПос').AsString;
                        SG21.Cells[27, i + 1] :=
                        Form1.ADOQuery1.FieldByName('КодПос').AsString;
                        SG21.Cells[28, i + 1] :=
                        Form1.ADOQuery1.FieldByName('Мат').AsString;

                        SG21.Cells[29, i + 1] :=
                        Form1.ADOQuery1.FieldByName('План дата').AsString;
                        SG21.Cells[30, i + 1] :=
                        Form1.ADOQuery1.FieldByName('План неделя').AsString;
                        SG21.Cells[31, i + 1] :=
                        Form1.ADOQuery1.FieldByName('БЗ').AsString;
                        SG21.Cells[32, i + 1] :=
                        Form1.ADOQuery1.FieldByName('Упаковка1').AsString;
                        SG21.Cells[33, i + 1] :=
                        Form1.ADOQuery1.FieldByName('Примечания').AsString;
                        if Form1.Luk=0 then
                        begin

                          SG21.Cells[34, i + 1] :=
                        Form1.ADOQuery1.FieldByName('Клапан').AsString;
                        SG21.Cells[35, i + 1] :=
                        Form1.ADOQuery1.FieldByName('КолКлап').AsString;
                        end;
                        Form1.ADOQuery1.Next;

                end;
    For I:=0 To  SG21.RowCount-1 Do
    begin
             SG21.AutoSizeRow(i,0);
    end;
    if Form1.Luk=0 then
    begin
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select MAX(Номер) as a from %s ',
                ['ЗапускСТАМ']) then
                exit;
    end;

    if Form1.Luk=1 then
    begin
        if not Form1.mkQuerySelect(Form1.ADOQuery1,
                'Select MAX(Номер) as a from %s ',
                ['ЗапускЛЮК']) then
                exit;
    end;
        Max := Form1.ADOQuery1.FieldValues['a'];    // Max
        Lbl1.Caption :=IntToStr(Max+1); //
        Edit.Text :=IntToStr(Max+1);
        // Lbl1.Caption := edit.Text;
        //SG21.ColWidths[17] := 0;
        //SG21.ColWidths[18] := 0;
        //SG21.ColWidths[19] := 0;

end;

procedure TFS.dtp1Click(Sender: TObject);
begin
        btn2.Enabled:=True;
end;

procedure TFS.btn2Click(Sender: TObject);
var Dat_Zak,Plan_Dat,Rasch_Data,Plan_Nedel,Dat1,S_Svar,S_Sbor,IDGP,KTO,PDO,Kol_Raz,Nomer_Raz,God,Mes,Dir,Mat,ss,sss,
Vn_Dat,fileName,BZ:string;
I,NachNom,KonNom,NachKod,KonKod,Kol,Kol_Zap,Kol_Zap_R,NomPos,KodPos:Integer;
NC:Double;
XL2:Variant;
begin
        Dat1 := FormatDateTime('mm.dd.yyyy', DTP1.Date);
        God := FormatDateTime('yyyy', Now);
        Mes := FormatDateTime('mmmm', Now);

        Vn_Dat := FormatDateTime('dd.mm.yyyy', DTP1.Date);
        fileName := ExtractFileDir(ParamStr(0));//+'\Klapan.EXE';
        NC :=0;
        //++++++++++++++++++++++++++++++++++++++++++++++++++
        For I:=0 to SG21.RowCount-1 do
        begin
           if SG21.Cells[5,I+1]<>'' then
           begin
            //if  (UOsnova_Main.Flag_Error =0) {OR (UOsnova_Main.Flag_Error =2)} Then  if A>0 then M:=0
            Begin
                NC :=0;
                Kol:=StrToInt(SG21.Cells[3,I+1]);
                Kol_Zap:=StrToInt(SG21.Cells[5,I+1]);
                S_Sbor:=Form1.ConvertFloat(SG21.Cells[16, i+1 ]);
                NC:=StrToFloat(SG21.Cells[16, i+1 ]); //*Kol_Zap
                SSS:= FloatToStr(NC);
                S_Sbor:=Form1.ConvertFloat(SSS);
                IDGP:=SG21.Cells[17, i+1 ];
                BZ:=SG21.Cells[31, i+1 ];
                Dat_Zak:=Form1.ConvertDat1(SG21.Cells[1, i+1 ]);

                if SG21.Cells[21, i+1 ]<>'' then
                Rasch_data:=#39 +Form1.ConvertDat1(SG21.Cells[21, i+1 ])+#39//Расчетная дата готовности
                else
                Rasch_data:='NULL';
              if Form1.Luk=0 then
              begin
                if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s Where'+
                ' [IDГП]=' + #39 +
                (IDGP) + #39 , ['СТАМ']) then
                exit;
                Kol_Zap_R:=Form1.ADOQuery1.FieldByName('Кол во запущенных').AsInteger;
                Plan_dat:=Form1.ConvertDat1(SG21.Cells[29, i+1 ]);
                Plan_Nedel:=(SG21.Cells[30, i+1 ]);
                Mat:=Form1.ConvertFloat(SG21.Cells[28, i+1 ]);

                //=======================
                if not
                Form1.mkQueryInsert(Form1.ADOQuery2,
                'Insert Into %s ' +
                '([Дата],Заказ,Изделие,[Кол во],'+
                '[Кол во запущенных],[' + FN_RAS_DATA_GOTOVN+ '],Номер,[Н\ч Сварка],'+
                '[Н\ч Сборка],'+
                'НачНомер,КонНомер,НачКод,КонКод,'+
                'IdГП,Мат,[План Дата],[План Неделя],Планирование,БЗ,Исполнение,Примечание,КолКлап,Клапан) ' +
                'Values (%s,%s,%s,' +
                '%s,%s,%s,'+
                '%s,%s ,%s,'+
                '%s,%s,%s,'+
                '%s,%s ,%s,'+
                '%s ,%s ,%s,'+
                '%s,%s,%s,%s,%s)',
                ['ЗапускСТАМ',
                #39 + Dat_Zak + #39,
                #39 + SG21.Cells[2, i+1] +#39,
                #39 + SG21.Cells[6, i+1] +#39,
                #39 + SG21.Cells[3, i+1] +#39,
                #39 + IntToStr(Kol_Zap)+ #39,
                 Rasch_data ,
                #39 + Edit.Text +#39,
                #39 + S_Svar + #39,
                #39 + S_Sbor + #39,
                #39 + SG21.Cells[22, I+1] +#39,
                #39 + SG21.Cells[23, I+1] +#39,
                #39 + SG21.Cells[24, I+1] +#39,
                #39 + SG21.Cells[25, I+1] +#39,
                #39 + IdGP +#39,
                #39 + Mat +#39,
                #39 + Plan_Dat +#39,
                #39 + Plan_Nedel +#39,
                #39 + Dat1 +#39,
                #39 + BZ +#39,
                #39 + SG21.Cells[32, I+1] +#39,
                #39 +SG21.Cells[33, I+1] + #39,
                #39 +SG21.Cells[35, I+1] + #39,
                #39 +SG21.Cells[34, I+1] + #39
                  ] )
                then
                exit;

                UOsnova_Main:=Osnova_Main.Create() ;
                fileName := '';//ExtractFileDir(ParamStr(0));//+'\Klapan.EXE';
                Vn_Dat := FormatDateTime('dd.mm.yyyy', Dtp1.Date);
                UOsnova_Main.Flag_Error :=0;
                UOsnova_Main.Osnova( Vn_Dat,'ЗапускСТАМ','СпецифСТАМ',#39+Edit.Text+#39,fileName,1,0 ) ;
                //then
               //Exit;
                if  UOsnova_Main.Flag_Error =2 Then //ДЕТАЛИ НЕ ВСЕ
                //ЗАНЕСТИ В ЗАПУСК И Поставить флаг ФлагЗаготовки
                Begin
                        //++++++++++++++++++++++++++++++++++++++++++++++++++++
                       { if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                        'UPDATE %s SET [ФлагЗаготовки]=' + #39
                        + '0' + #39+ ',Планирование=NULL WHERE [Номер]=' + #39 +
                        (Label4.Caption) + #39, ['ЗапускВозд'])
                        then
                        Exit; }
                        //Form1.Memo2.Lines.Add('UOsnova_Main.Osnova( ДЕТАЛИ НЕ ВСЕ,ЗапускВозд,СпецифВозд,'+Label4.Caption+','+fileName+' )UOsnova_Main.Flag_Error =2') ;
                end;
                //++++++++++++++++++++++++++++++++++++++++++++++++++
                if  UOsnova_Main.Flag_Error =1 Then //ДАТА ЛЕВАЯ УДАЛИТЬ Из ЗАПУСКА
                                                //И UPDATE КЛАПАНА НЕ ДЕЛАТЬ
                Begin
                      {  if not Form1.mkQueryDelete( Form1.ADOQuery1, 'DELETE FROM %s Where (Номер= '
                        +#39+Edit.Text+#39+') '  ,
                        ['ЗапускСТАМ'] )
                        Then
                        Exit;    }
                        //Form1.Memo2.Lines.Add('UOsnova_Main.Osnova( ДАТА НЕ ТА - '+Vn_Dat+',ЗапускВозд,СпецифВозд,'+Label4.Caption+','+fileName+' )UOsnova_Main.Flag_Error =1') ;
                        //Exit; }
                end;

                if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                'UPDATE %s SET [Кол во запущенных]=' + #39+ IntToStr(Kol_Zap+Kol_Zap_R ) + #39
                + ',[Кол во запущенных пос]=' +#39 + IntToStr(Kol_Zap) + #39 +
                ',Номер=' + #39 +Edit.Text + #39 +
                ',[' + FN_KOL_RAZ + ']=[Кол раз]+'+#39+',' + IntToStr(Kol_Zap) + #39 + ',[' +
                FN_NOMER_RAZ + ']=[Номер раз]+' + #39+',' +Edit.Text +#39 +
                ' WHERE [IDГП]=' + #39 +
                (IDGP) + #39, ['СТАМ'])
                then
                Exit;
              end;

              //////
              if Form1.Luk=1 then
              begin
                if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select * from %s Where'+
                ' [IDГП]=' + #39 +
                (IDGP) + #39 , ['ЛЮК']) then
                exit;
                Kol_Zap_R:=Form1.ADOQuery1.FieldByName('Кол во запущенных').AsInteger;
                Plan_dat:=Form1.ConvertDat1(SG21.Cells[29, i+1 ]);
                Plan_Nedel:=(SG21.Cells[30, i+1 ]);
                Mat:=Form1.ConvertFloat(SG21.Cells[28, i+1 ]);
                //=======================
                if not Form1.mkQueryInsert(Form1.ADOQuery2,
                'Insert Into %s ' +
                '([Дата],Заказ,Изделие,[Кол во],'+
                '[Кол во запущенных],[' + FN_RAS_DATA_GOTOVN+ '],Номер,[Н\ч Сварка],[Н\ч Сборка],'+
                'НачНомер,КонНомер,НачКод,КонКод,IdГП,Мат,[План Дата],[План Неделя],Планирование,БЗ,Исполнение) ' +
                'Values (%s,%s,%s,' +
                '%s,%s,%s,%s,%s ,%s,%s,%s,%s,%s,%s ,%s,%s ,%s ,%s,%s,%s)',
                ['ЗапускЛЮК',
                #39 + Dat_Zak + #39,
                #39 + SG21.Cells[2, i+1] +#39,
                #39 + SG21.Cells[6, i+1] +#39,
                #39 + SG21.Cells[3, i+1] +#39,
                #39 + IntToStr(Kol_Zap)+ #39,
                 Rasch_data ,
                #39 + Edit.Text +#39,
                #39 + S_Svar + #39,
                #39 + S_Sbor + #39,
                #39 + SG21.Cells[22, I+1] +#39,
                #39 + SG21.Cells[23, I+1] +#39,
                #39 + SG21.Cells[24, I+1] +#39,
                #39 + SG21.Cells[25, I+1] +#39,
                #39 + IdGP +#39,
                #39 + Mat +#39,
                #39 + Plan_Dat +#39,
                #39 + Plan_Nedel +#39,
                #39 + Dat1 +#39,
                #39 + BZ +#39,
                #39 + SG21.Cells[32, I+1] +#39  ] )
                then
                exit;
                UOsnova_Main:=Osnova_Main.Create() ;
                fileName := '';//ExtractFileDir(ParamStr(0));//+'\Klapan.EXE';
                Vn_Dat := FormatDateTime('dd.mm.yyyy', Dtp1.Date);
                UOsnova_Main.Flag_Error :=0;
                UOsnova_Main.Osnova( Vn_Dat,'ЗапускЛЮК','СпецифЛЮК',#39+Edit.Text+#39,fileName,1,0 ) ;
                //then
               //Exit;
                if  UOsnova_Main.Flag_Error =2 Then //ДЕТАЛИ НЕ ВСЕ
                //ЗАНЕСТИ В ЗАПУСК И Поставить флаг ФлагЗаготовки
                Begin
                        //++++++++++++++++++++++++++++++++++++++++++++++++++++
                       // if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                        //'UPDATE %s SET [ФлагЗаготовки]=' + #39
                        //+ '0' + #39+ ',Планирование=NULL WHERE [Номер]=' + #39 +
                        //(Label4.Caption) + #39, ['ЗапускВозд'])
                        //then
                       // Exit;
                        //Form1.Memo2.Lines.Add('UOsnova_Main.Osnova( ДЕТАЛИ НЕ ВСЕ,ЗапускВозд,СпецифВозд,'+Label4.Caption+','+fileName+' )UOsnova_Main.Flag_Error =2') ;
                end;
                //++++++++++++++++++++++++++++++++++++++++++++++++++
                if  UOsnova_Main.Flag_Error =1 Then //ДАТА ЛЕВАЯ УДАЛИТЬ Из ЗАПУСКА
                                                //И UPDATE КЛАПАНА НЕ ДЕЛАТЬ
                Begin
                        //if not Form1.mkQueryDelete( Form1.ADOQuery1, 'DELETE FROM %s Where (Номер= '
                        //+#39+Edit.Text+#39+') '  ,
                        //['ЗапускЛЮК'] )
                        //Then
                        //Exit;
                        //Form1.Memo2.Lines.Add('UOsnova_Main.Osnova( ДАТА НЕ ТА - '+Vn_Dat+',ЗапускВозд,СпецифВозд,'+Label4.Caption+','+fileName+' )UOsnova_Main.Flag_Error =1') ;
                        //Exit;
                end;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                'UPDATE %s SET [Кол во запущенных]=' + #39+ IntToStr(Kol_Zap+Kol_Zap_R ) + #39
                + ',[Кол во запущенных пос]=' +#39 + IntToStr(Kol_Zap) + #39 +
                ',Номер=' + #39 +Edit.Text + #39 +
                ',[' + FN_KOL_RAZ + ']=[Кол раз]+'+#39+',' + IntToStr(Kol_Zap) + #39 + ',[' +
                FN_NOMER_RAZ + ']=[Номер раз]+' + #39+',' +Edit.Text +#39 +
                ' WHERE [IDГП]=' + #39 +
                (IDGP) + #39, ['ЛЮК'])
                then
                Exit;
              end;
            End;
           end;
        end;
        Close;

end;

procedure TFS.SG21SetEditText(Sender: TObject; ACol, ARow: Integer;
  const Value: String);
  Var Kol_Sv,Kol_Zap,Kol:Integer;
  nc:Double;
begin
        Kol_Sv:=StrToInt(SG21.Cells[4,ARow]);
        Kol_Zap:=StrToInt(SG21.Cells[5,ARow]);
        //++++++++++++++++++++++++++++++++++++++++
        if Kol_Zap>Kol_Sv then
        begin
            MessageDlg('Нельзя, МНОГО!', mtError, [mbOk], 0);
            Exit;
        end;
        Kol:=Kol_Sv-Kol_Zap;

        if SG21.Cells[15, ARow]='' Then
                NC:=0
        Else
                NC:=StrToFloat(SG21.Cells[15, ARow]);
        SG21.Cells[16, ARow]:=FloatToStr(NC*Kol_Zap);
end;

end.

