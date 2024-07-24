unit UOsnova;

interface
uses SysUtils, Classes, Variants,Forms,
        ADODB, Controls,Dialogs,AdvGrid,Grids,Windows,
  UConnCeh, UConnKlap,ShellAPI,Math;
        Type
        Osnova_Main = class//(TForm)
        SG2:TStringGrid;
        SG3:TStringGrid;
        procedure Clear_StringGrid(StringGrid: TStringGrid);
        function ConvertDat1(StrDat: string): string;
        function Osnova( var Dat : string ;const tab1 : string;tab2 :string ;
                        Nom :string;fileName:String;Flag_Polz,Cvet:Integer) : Boolean;
                        //Cvet- из клапанов 0-белый, 1-Желтый, 2-зеленый, 3-Синий
        function btn1Click(fileName:String):Boolean;
        function btn11Click(Err2:String):Boolean;
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++
        //СТАМ ВРАН ОСА
        //Дата запуска Заготовки= дата заготовки Первой детали с Очередностью=1
        // прибавляя следующую Очередность доходим до Даты Сборки (Планирование)
        function Osnova2( var Dat : string ;const tab1 : string;tab2 :string ;
                        Nom :string;fileName:String;Flag_Polz:Integer) : Boolean;
        function btn2Click(fileName:String):Boolean;
        function btn22Click(Err2:String):Boolean;
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++
        public
                Column: array[0..550, 0..4] of string;
                StrToCol: array[0..100,0..6] of string;
                Smena: array[0..550, 0..4] of string;
                Temp: array[0..550, 0..6] of string;
                Conn_Ceh:Connect_Miass_Ceh;
                Conn_Klap:Connect_Miass_Klap;
                S1,Nomer_Glob:String;
                Dat_Glob,File_Name:String;
                Flag_Error:Integer;
                Flag_Polzov    //0 Технолог   1 Диспетчер
                ,Cvet1:Integer;
                SNCH,Nom_G:String;// возвращает н\ч на каждый станок для каждой даты
                Leg:String;
        End;
implementation

uses
  Main;
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++ Check
//Очистка Грида

procedure Osnova_Main.Clear_StringGrid(StringGrid: TStringGrid);
var
        i, j: integer;
begin
        for i := 0 to StringGrid.ColCount - 1 do
                for j := 1 to StringGrid.RowCount - 2 do
                begin
                        //StringGrid.Objects[i, j] := nil;
                        StringGrid.Cells[I, j] := '';

                end;

end;
function Osnova_Main.ConvertDat1(StrDat: string): string;
var
        G, M, D: string;
        PosM, PosD: integer;
begin
        if StrDat <> '' then
        begin
                PosD := AnsiPos('.', StrDat);
                D := Copy(StrDat, 1, PosD - 1);
                Delete(StrDat, 1, PosD);
                PosM := AnsiPos('.', StrDat);
                M := Copy(StrDat, 1, PosM - 1);
                Delete(StrDat, 1, PosM);
                G := M + '.' + D + '.' + StrDat;
                Result := G;
        end
        else
                Result := '';
end;
//++++++++++++++++++++++++++++++++++++++++++++++++++++++
        //СТАМ ВРАН ОСА
        //Дата запуска Заготовки= дата заготовки Первой детали с Очередностью=1
        // прибавляя следующую Очередность доходим до Даты Сборки (Планирование)
function Osnova_Main.Osnova2( var Dat : string;  const tab1:string;tab2:string;
Nom:string;fileName:String;Flag_Polz:Integer ) : Boolean;
var a,Y,Glob,i,x,Res,T,Res1,L:Integer;
Stanok,Dat_S1,IDGP,Kol_Klap,Izdel,kol_Det,Elem,Obozn,Nomer,BZ,Str1,Str2,Dlit,Kod,
IDKO,Nam_S,ZAKAZ:string;
Dlit_D,Och_D:Double;
Res_Teh,Res_Stan,Memo12_Count,Res_Tab,Polzov,
KOL_RECORD_Tab1,KOL_RECORD_Tab2:Integer;
Memo12,Memo13:TStringList;
Memo12_text,Nam,God,Mes,Dir:String;
h: hwnd;
begin
    Result:=False;
    try
        T:=0;
        Nom_G:=Nom;
        Conn_Ceh := Connect_Miass_Ceh.Create();
        Conn_Klap:=Connect_Miass_Klap.Create();
        Memo12:=TStringList.Create;
        Memo12.Clear;
        Memo13:=TStringList.Create;
        Memo13.Clear;
        Polzov:= Flag_Polz;
        if not Conn_Klap.mkQuerySelect('SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME ='+#39+'СпецифОбщая'+#39
        , []) then
        exit;
        Dat_Glob:= Dat;
        for I:=0 To Conn_Klap.AQuery.RecordCount-1 do
        begin
          Str1:=Conn_Klap.AQuery.FieldByName('COLUMN_NAME').AsString;
          Str2:=Conn_Klap.AQuery.FieldByName('COLUMN_NAME').AsString;
          Res:=Pos('Н\ч',str2);
          if Res<>0 then
          begin
               Delete(str2,1,3);
               Column[T,0]:=Str2;
               Column[T,1]:=Str2+' BIT';
               Column[T,2]:=IntToStr(T+7);
               Inc(T);
          end;
          Conn_Klap.AQuery.Next;
        End;
        //++++++++++++++++++++++++++++
        fillchar(Smena,sizeof(Smena),0);//Очистка массмва
        if not Conn_Klap.mkQuerySelect('Select * From Смены  ' , []) then //Запуск
        exit;
      For I:=0 to Conn_Klap.AQuery.RecordCount-1 Do
      Begin
        Smena[i,0]:=Conn_Klap.AQuery.FieldByName('Передел').AsString;
        Smena[i,1]:=Conn_Klap.AQuery.FieldByName('Смена').AsString;
        Smena[i,2]:=Conn_Klap.AQuery.FieldByName('Часы').AsString;
        Conn_Klap.AQuery.Next;
      End;
          sg2 := TStringGrid.Create(Nil);
          Clear_StringGrid(sg2);
        //sg2.Clear;
        A:=1;
        sg2.ColCount:=11;
        sg2.Cells[0,0]:='№';
        sg2.Cells[1,0]:='Заказ';
        sg2.Cells[2,0]:='Изделие';
        sg2.Cells[3,0]:='Кол во';
        sg2.Cells[4,0]:='Обозначение';
        sg2.Cells[5,0]:='Элемент';
        sg2.Cells[6,0]:='Номер';
        sg2.Cells[7,0]:='БЗ';
        sg2.Cells[8,0]:='IdГП';
        sg2.Cells[9,0]:='Технолог';
        sg2.Cells[10,0]:='IdКО';
        for y := 0 to High(Column) do
        begin
                if Column[y,0]<>'' then
                Begin
                        Stanok:=Column[y,0];
                        SG2.ColCount:=SG2.ColCount+4;
                        sg2.Cells[9+A+1,0]:=Stanok;
                        sg2.Cells[9+A+2,0]:=Stanok+'Оч'; //Очередность
                        sg2.Cells[9+A+3,0]:=Stanok+'Дл';//Длительность
                        sg2.Cells[9+A+4,0]:='Н\ч'+Stanok;//Н\ч
                        //sg2.Cells[9+A+5,0]:=Stanok+'Дата'; //дата
                        A:=A+4;
                end;
        end;
        Glob:=1;//SG2.RowCount;
        //Dat_S1:=FormatDateTime('mm.dd.YYYY',Dtp1.DateTime);
        if Polzov<>4 then
        begin
          if not Conn_Klap.mkQuerySelect('Select * From %s Where (Номер=%s) '+
          'AND (Отмена IS NULL) ' , [tab1,Nom]) then //Запуск
          exit;
          KOL_RECORD_Tab1:= Conn_Klap.AQuery.RecordCount;
        end;
        ///Вентиляторы из Цеха
        if Polzov=4 then
        begin
          if not Conn_Ceh.mkQuerySelect('Select * From %s Where (Номер=%s) '+
          'AND (Отмена IS NULL) ' , [tab1,Nom]) then //Запуск
          exit;
          KOL_RECORD_Tab1:= Conn_Ceh.AQuery.RecordCount;
        end;
        SG2.RowCount:= 2;
        SG2.FixedRows:=1;
        A:=1;
      For I:=0 to KOL_RECORD_Tab1-1 Do
      Begin
        if Polzov<>4 then
        begin
          IDGP:=Conn_Klap.AQuery.FieldByName('IdГП').AsString;
          IDKO:=Conn_Klap.AQuery.FieldByName('IdКО').AsString;
          Izdel:=Conn_Klap.AQuery.FieldByName('Изделие').AsString;
          Kol_Klap:=Conn_Klap.AQuery.FieldByName('Кол во запущенных').AsString;
          Nomer:=Conn_Klap.AQuery.FieldByName('Номер').AsString;
          Nomer_Glob:=Conn_Klap.AQuery.FieldByName('Номер').AsString;
          BZ:=Conn_Klap.AQuery.FieldByName('БЗ').AsString;
          ZAKAZ:=Conn_Klap.AQuery.FieldByName('Заказ').AsString;
        end;
        if Polzov=4 then
        begin
          IDGP:=Conn_Ceh.AQuery.FieldByName('IdГП').AsString;
          IDKO:=Conn_Ceh.AQuery.FieldByName('IdКО').AsString;
          Izdel:=Conn_Ceh.AQuery.FieldByName('Изделие').AsString;
          Kol_Klap:=Conn_Ceh.AQuery.FieldByName('Кол запущенных').AsString;
          Nomer:=Conn_Ceh.AQuery.FieldByName('Номер').AsString;
          Nomer_Glob:=Conn_Ceh.AQuery.FieldByName('Номер').AsString;
          BZ:=Conn_Ceh.AQuery.FieldByName('БЗ').AsString;
          ZAKAZ:=Conn_Ceh.AQuery.FieldByName('Заказ').AsString;
        end;
        if (Tab2='') and (Polzov<>4) Then //Tab2 пустое значит спецификация уже в СпецифОбщая  , ищем по названию изделия СТАМ....
        Begin

                 if not Conn_Klap.mkQuerySelect2(
                'Select * from %s Where (ВидЭлемента='+#39+'Деталь'+#39+') '+
                'AND (([Изделие]='+#39+'Стакан монтажный '+Izdel+#39+') OR ([Изделие]='+#39+'Поддон '+Izdel+#39+'))'
                 , ['СпецифОбщая']) then  //Специф
                 exit;
                 For X:=0 to Conn_Klap.AQuery_Del.RecordCount-1 Do
                Begin
                        Obozn:=Conn_Klap.AQuery_Del.FieldByName('Обозначение').AsString;
                        Elem:=Conn_Klap.AQuery_Del.FieldByName('Элемент').AsString;
                        Kol_Det:=Conn_Klap.AQuery_Del.FieldByName('Количество').AsString;
                   if  Conn_Klap.AQuery_Del.RecordCount<>0 then
                   Begin
                        SG2.RowCount:=SG2.RowCount+ 1;
                        sg2.Cells[0,Glob+1]:=Kol_Klap;//IntToStr(Glob+1);

                        sg2.Cells[1,Glob+1]:=Conn_Klap.AQuery.FieldByName('Заказ').AsString;
                        sg2.Cells[2,Glob+1]:=Conn_Klap.AQuery.FieldByName('Изделие').AsString;

                        sg2.Cells[3,Glob+1]:=IntToStr(StrToInt(Kol_Klap)*StrToInt(Kol_Det));
                        sg2.Cells[4,Glob+1]:=Conn_Klap.AQuery_Del.FieldByName('Обозначение').AsString;
                        sg2.Cells[5,Glob+1]:=Conn_Klap.AQuery_Del.FieldByName('Элемент').AsString;
                        sg2.Cells[6,Glob+1]:=Nom;
                        sg2.Cells[7,Glob+1]:=BZ;
                        sg2.Cells[8,Glob+1]:=IDGP;
                        sg2.Cells[9,Glob+1]:=Conn_Klap.AQuery3.FieldByName('Технолог').AsString;
                        sg2.Cells[10,Glob+1]:=IDKO;
                        Res_Teh:=Pos('rue',sg2.Cells[9,Glob+1]);
                        if (Res_Teh=0)  Then //ТеХНОЛОГ FALSE
                        begin
                                  Memo12.Add('ТЕХНОЛОГ FALSE Заказ '+sg2.Cells[1,Glob+1]+' * '+sg2.Cells[2,Glob+1]+' * Деталь '+sg2.Cells[4,Glob+1]);
                                  Memo13.Add(sg2.Cells[4,Glob+1]);
                        end;
                        A:=1;
                        for y := 0 to High(Column) do
                        begin
                                if Column[y,0]<>'' then
                                Begin
                                        Stanok:=Column[y,0]; //Значение поля

                                        sg2.Cells[9+(A+1),Glob+1]:=Conn_Klap.AQuery_Del.FieldByName(Stanok).AsString;
                                        Res_Stan:=Pos('rue',sg2.Cells[9+(A+1),Glob+1]);
                                        Inc(A);

                                        sg2.Cells[9+(A+1),Glob+1]:=Conn_Klap.AQuery_Del.FieldByName(Stanok+'Оч').AsString; //Очередность
                                        Och_D:=StrToFloat(sg2.Cells[9+(A+1),Glob+1]);
                                        Inc(A);

                                        Dlit:= '0.1';//Conn_Klap.AQuery_Del.FieldByName(Stanok+'Дл').AsString;//Длительность
                                        sg2.Cells[9+(A+1),Glob+1]:='0.1';//Conn_Klap.AQuery_Del.FieldByName(Stanok+'Дл').AsString;//Длительность
                                        Dlit_D:=StrToFloat( sg2.Cells[9+(A+1),Glob+1]);//Длительность
                                        Inc(A);

                                        if (Res_Stan<>0) AND ((Och_D=0) OR (Dlit_D=0))  Then //СТАНОК TRUE
                                                                                             //И Очер=0 или Длит =0 Косяк
                                        begin
                                                Memo12.Add('СТАНОК TRUE И Очер=0 или Длит =0 Заказ '+sg2.Cells[1,Glob+1]+' * '+sg2.Cells[2,Glob+1]+' * Деталь '+sg2.Cells[4,Glob+1]);
                                                Memo13.Add(sg2.Cells[4,Glob+1]);
                                        end;

                                        sg2.Cells[9+(A+1),Glob+1]:=Conn_Klap.AQuery_Del.FieldByName('Н\ч'+Stanok).AsString;//Н\ч
                                        Inc(A);
                                        //sg2.Cells[9+(A+1),Glob+1]:=Conn_Klap.AQuery3.FieldByName(Stanok+'Дата').AsString; //Дата
                                        //Inc(A);
                                End;
                        end;
                        Inc(Glob);
                   end;
                   Conn_Klap.AQuery_Del.next;
                end;
        end
        Else
        Begin
           if (Polzov=4) then //ВРАН в спецификации ID Деталей 3 Запроса
           begin
              Res:=Pos('СпецифОСА',tab2);
             if Res=0 then
             Begin
                //В Tab2 для ВРАНСпециф будет IdВРАН
               if not Conn_Ceh.mkQuerySelect2(
                'SELECT ВРАНСборЕд.Id,Деталь ,Единица, ВРАНСборЕд.kol '+
                'FROM ВРАНСпециф INNER JOIN ВРАНСборЕд ON ВРАНСпециф.IdSb = ВРАНСборЕд.Id '+
                'WHERE (ВРАНСпециф."IdВРАН"='+Tab2+') AND (ВРАНСпециф.IdSb<>'+#39+'0'+#39+') UNION '+

                'SELECT we.Id , Деталь ,  Единица, we.kol '+
                'FROM ВРАНСпециф INNER JOIN '+
	              '(SELECT ВРАНДет.Id,Деталь , Обозначение Единица, ВРАНДет.kol '+
                'from ВРАНДет) '+
	              'we  ON ВРАНСпециф.IdDet = we.Id '+
                'WHERE (ВРАНСпециф.IdВРАН='+Tab2+') AND (ВРАНСпециф.IdSb='+#39+'0'+#39+')'
                 , ['ВРАНСпециф']) then  //Специф
                exit;
             End
             else
             begin
                if not Conn_Ceh.mkQuerySelect2(
                'Select * from %s Where (ВидЭлемента='+#39+'Детали'+#39+
                ') AND ([IdГП]='+#39+IDGP+#39+') '
                 , [Tab2]) then  //Специф
                exit;
             end;
             KOL_RECORD_Tab2:=Conn_Ceh.AQuery_Del.RecordCount;
           end
           else
           begin
                if not Conn_Klap.mkQuerySelect2(
                'Select * from %s Where (ВидЭлемента='+#39+'Детали'+#39+
                ') AND ([IdГП]='+#39+IDGP+#39+') AND ([IdКО]='+#39+IDKO+#39+')'
                 , [Tab2]) then  //Специф
                exit;
                KOL_RECORD_Tab2:=Conn_Klap.AQuery_Del.RecordCount;
           end;

           For X:=0 to KOL_RECORD_Tab2-1 Do
           Begin
             if (Polzov<>4) then
             Begin
                Obozn:=Conn_Klap.AQuery_Del.FieldByName('Обозначение').AsString;
                Elem:=Conn_Klap.AQuery_Del.FieldByName('Элемент').AsString;
                Kol_Det:=Conn_Klap.AQuery_Del.FieldByName('Количество').AsString; // AND (Технолог=1)
              end;
             if (Polzov=4) then
             Begin
               Res:=Pos('СпецифОСА',tab2);
               if Res=0 then
               Begin
                Obozn:=Conn_Ceh.AQuery_Del.FieldByName('Единица').AsString;
                Elem:=Conn_Ceh.AQuery_Del.FieldByName('Деталь').AsString;
                Kol_Det:=Conn_Ceh.AQuery_Del.FieldByName('Kol').AsString; // AND (Технолог=1)
               end
               else
               begin
                 Obozn:=Conn_Ceh.AQuery_Del.FieldByName('Обозначение').AsString;
                Elem:=Conn_Ceh.AQuery_Del.FieldByName('Элемент').AsString;
                Kol_Det:=Conn_Ceh.AQuery_Del.FieldByName('Количество').AsString;
               end;
              end;
                if not Conn_Klap.mkQuerySelect3(
                'Select * from %s Where (Обозначение='+#39+Obozn+#39+') '
                 , ['СпецифОбщая']) then
                exit;
                if  Conn_Klap.AQuery3.RecordCount<>0 then
                Begin
                        SG2.RowCount:=SG2.RowCount+ 1;
                        sg2.Cells[0,Glob+1]:=Kol_Klap;//IntToStr(Glob+1);

                        sg2.Cells[1,Glob+1]:=ZAKAZ;//Conn_Klap.AQuery.FieldByName('Заказ').AsString;
                        sg2.Cells[2,Glob+1]:=Izdel;//Conn_Klap.AQuery.FieldByName('Изделие').AsString;
                        //Добавить поле ТЕХНОЛОГ, если фалш то деталь не обработана
                        sg2.Cells[3,Glob+1]:=IntToStr(StrToInt(Kol_Klap)*StrToInt(Kol_Det));
                        sg2.Cells[4,Glob+1]:=Conn_Klap.AQuery3.FieldByName('Обозначение').AsString;
                        sg2.Cells[5,Glob+1]:=Conn_Klap.AQuery3.FieldByName('Элемент').AsString;
                        sg2.Cells[6,Glob+1]:=Nom;
                        sg2.Cells[7,Glob+1]:=BZ;
                        sg2.Cells[8,Glob+1]:=IDGP;
                        sg2.Cells[9,Glob+1]:=Conn_Klap.AQuery3.FieldByName('Технолог').AsString;
                        sg2.Cells[10,Glob+1]:=IDKO;
                        Res_Teh:=Pos('rue',sg2.Cells[9,Glob+1]);
                        if (Res_Teh=0)  Then //ТеХНОЛОГ FALSE
                        begin
                                  Memo12.Add('ТЕХНОЛОГ FALSE Заказ '+sg2.Cells[1,Glob+1]+' * '+sg2.Cells[2,Glob+1]+' * Деталь '+sg2.Cells[4,Glob+1]);
                                  Memo13.Add(sg2.Cells[4,Glob+1]);
                        end;
                        A:=1;
                        for y := 0 to High(Column) do
                        begin
                                if Column[y,0]<>'' then
                                Begin
                                        Stanok:=Column[y,0]; //Значение поля

                                        sg2.Cells[9+(A+1),Glob+1]:=Conn_Klap.AQuery3.FieldByName(Stanok).AsString;
                                        Res_Stan:=Pos('rue',sg2.Cells[9+(A+1),Glob+1]);
                                        Inc(A);

                                        sg2.Cells[9+(A+1),Glob+1]:=Conn_Klap.AQuery3.FieldByName(Stanok+'Оч').AsString; //Очередность
                                        Och_D:=StrToFloat(sg2.Cells[9+(A+1),Glob+1]);
                                        Inc(A);

                                        sg2.Cells[9+(A+1),Glob+1]:='0.1';//Conn_Klap.AQuery3.FieldByName(Stanok+'Дл').AsString;//Длительность
                                        Dlit_D:=StrToFloat( sg2.Cells[9+(A+1),Glob+1]);//Длительность
                                        Inc(A);

                                        if (Res_Stan<>0) AND ((Och_D=0) OR (Dlit_D=0))  Then //СТАНОК TRUE
                                                                                             //И Очер=0 или Длит =0 Косяк
                                        begin
                                                Memo12.Add('СТАНОК TRUE И Очер=0 или Длит =0 Заказ '+sg2.Cells[1,Glob+1]+' * '+sg2.Cells[2,Glob+1]+' * Деталь '+sg2.Cells[4,Glob+1]);
                                                Memo13.Add(sg2.Cells[4,Glob+1]);
                                        end;

                                        sg2.Cells[9+(A+1),Glob+1]:=Conn_Klap.AQuery3.FieldByName('Н\ч'+Stanok).AsString;//Н\ч
                                        Inc(A);
                                        //sg2.Cells[9+(A+1),Glob+1]:=Conn_Klap.AQuery3.FieldByName(Stanok+'Дата').AsString; //Дата
                                        //Inc(A);
                                End;
                        end;
                        Inc(Glob);
                end;
                if Polzov<>4 then
                  Conn_Klap.AQuery_Del.next;
                if Polzov=4 then
                  Conn_Ceh.AQuery_Del.next;
           end;
        end;
        if Polzov<>4 then
          Conn_Klap.AQuery.next;
        if Polzov=4 then  //ВРАН
          Conn_Ceh.AQuery.next;  //gulyakova-ei
      end;
      Memo12_Count:=Memo12.Count;
      Memo12_Text:= Memo12.Text;
      God := FormatDateTime('yyyy', Now);
      Mes := FormatDateTime('mmmm', Now);
      if Polzov=4 then  //ВРАН
      Begin
        Dir :=Form1.Put_KTO+'\CKlapana\VRAN)\' + God;
        CreateDir(Dir);

        Dir := Form1.Put_KTO+'\CKlapana\VRAN\' + God + '\' + Mes+'\';
        CreateDir(Dir);
      End;
      Dir :=Form1.Put_KTO+'\CKlapana\Накладные(пожар)\' + God;
        CreateDir(Dir);

      Dir := Form1.Put_KTO+'\CKlapana\Накладные(пожар)\' + God + '\' + Mes+'\';
        CreateDir(Dir);

      Dir :=Form1.Put_KTO+'\CKlapana\' + God;
        CreateDir(Dir);
      Dir := Form1.Put_KTO+'\CKlapana\' + God + '\' + Mes+'\';
        CreateDir(Dir);
      if (Polzov=0) AND (Memo12_Count=0)  Then //Технолог
      //Все детали обработаны ,выходим
      Begin
        Flag_Error:=0;
        Memo12.Destroy;
        Exit;
      end;
        if (Polzov=0) AND (Memo12_Count<>0) then  //КОСЯК   Технолог
        Begin
                MessageDlg('Не все детали обработаны! Список в папке с накладной.', mtError,
                        [mbOk], 0);
                Res_Tab:=Pos('Воз',Tab2);
                if Res_Tab=0 Then           // PAnsiChar(P
                Begin
                        Memo12.SaveToFile(Form1.Put_KTO+'\CKlapana\Накладные(пожар)\'+God+'\'+Mes+'\'+Nom+'.txt');
                        Nam :=  Form1.Put_KTO+'\CKlapana\Накладные(пожар)\'+God+'\'+Mes+'\'+Nom+'.txt';
                End;
                if Res_Tab<>0 Then//Воздушные
                Begin
                        Memo12.SaveToFile(Form1.Put_KTO+'\CKlapana\'+God+'\'+Mes+'\'+Nom+'.txt');
                        Nam := Form1.Put_KTO+'\CKlapana\'+God+'\'+Mes+'\'+Nom+'.txt';
                End;
                ShellExecute(0, nil, PChar(Nam), nil, nil, SW_SHOWNORMAL);
                File_Name:= Nam ;
                Flag_Error:=2;
                Memo12.Destroy;
                Result := False;
                Exit;
        End;
        if (Polzov<>0) AND (Memo12_Count<>0) then  //КОСЯК
        Begin
                MessageDlg('Не все детали обработаны! Список в папке с накладной.', mtError,
                        [mbOk], 0);
                Res_Tab:=Pos('Воз',Tab2);
                if (Res_Tab=0) AND (Polzov<>4) Then
                Begin
                        Memo12.SaveToFile(Form1.Put_KTO+'\CKlapana\Накладные(пожар)\'+God+'\'+Mes+'\'+Nom+'.txt');
                        Nam :=  Form1.Put_KTO+'\CKlapana\Накладные(пожар)\'+God+'\'+Mes+'\'+Nom+'.txt';
                End;
                if (Res_Tab<>0) AND (Polzov<>4) Then//Воздушные
                Begin
                        Memo12.SaveToFile(Form1.Put_KTO+'\CKlapana\'+God+'\'+Mes+'\'+Nom+'.txt');
                        Nam :=  Form1.Put_KTO+'\CKlapana\'+God+'\'+Mes+'\'+Nom+'.txt';
                End;
                if  (Polzov=4) Then//VRAN
                Begin
                        Memo12.SaveToFile(Form1.Put_KTO+'\CKlapana\VRAN\'+God+'\'+Mes+'\'+Nom+'.txt');
                        Nam :=  Form1.Put_KTO+'\CKlapana\VRAN\'+God+'\'+Mes+'\'+Nom+'.txt';
                End;
                //ShellExecute(0, nil, PChar(Nam), nil, nil, SW_SHOWNORMAL);
                File_Name:= Nam ;
                Flag_Error:=2;
                Memo12.Destroy;
                Result := False;
                Exit;
        End;
      except

        exit;
    end;
    Memo12.Destroy;
    Memo13.Destroy;
    if not Conn_Klap.mkQueryDelete( 'DELETE FROM %s WHERE (Номер='+Nom+') AND (Заказ='+#39+ZAKAZ+#39+')'
    , ['Заготовка']) then
    exit;

    Result := true;
    btn2Click(fileName) ;
end;
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
function Osnova_Main.btn2Click(fileName:String):Boolean;
        var Dat_S1,IDGP,IDKO,Obozn,Detal,Str,Elem,Stanok,Str1,Str2,Izdel,Kol_Klap,Kol_Det,Zakaz,Nomer,Max:string;
        I,X,Y,A,Glob,Col,Col_Nom,
        Kol_Klap_I,Kol_Det_I:Integer;
        B,D:Boolean;                   // ,S1
        DEF_CatalogName,MSSQL_CONN_STR,S,Razmernost,Nom_Kol,Znach,Str_Stanok,Err,BZ:string;
        TempTableQ:TADOQuery;
        ADOConnection4:TADOConnection;
        Dat_S,Nedely,Kol,Dir,SQLServer,Dat_Z,Nowa_S,Dlit,Dlit_Ob,NCH,fileName1,Kod,Dat_Sbor:string;
        Dat, Dat1,Dat_Zag,nowa,Dat_Sbor_D:TDate;
        Ocheredn,Res_True:Integer;
        Dliteln,
        NC_D:Double;
        //StrToCol:array[0..100,0..5] of string;//[очередь, длительность, Общ длительность, Дата] ДЛЯ ОДНОЙ ДЕТАЛИ
        Kol_Stan:Integer;
        IPs : TStringList;
        Res_Trum:Integer;
        Dl_Trum:Double;
        Memo11:TStringList;
begin
       SyStem.SysUtils.FormatSettings.DecimalSeparator :=('.');
       Memo11:=TStringList.Create;
       Memo11.Clear;
       Glob:=0;
       A:=0;
      //==================================
      S:='';
      Razmernost:='';
      A:=0;
      fileName1:=fileName;
      for y := 0 to High(Column) do
      begin
        if Column[y,0]<>'' then
        Begin
                Str_Stanok:=Str_Stanok+Column[y,0]+','+Column[y,0]+'Оч,'+Column[y,0]+'Дл,[Н\ч'+Column[y,0]+'],'+Column[y,0]+'Дата,'+Column[y,0]+'ДлОбщ,';
                Nom_Kol:=IntToStr(StrToInt(Column[y,2])-2);
                Razmernost:=Razmernost+Column[y,0]+' BIT,'+Column[y,0]+'Оч nvarchar(50),'+Column[y,0]+'Дл nvarchar(50),[Н\ч'+Column[y,0]+'] nvarchar(50),'+Column[y,0]+'Дата DateTime,'+Column[y,0]+'ДлОбщ nvarchar(50),';  //DateTimeГибка bit,ГибкаОЧ nvarchar(50)
        End;
      end;

      S:= SG2.Cells[7,0];
      //++++++++++++++++++++++++++++++++++++
      Dir:=ExtractFileDir(ParamStr(0));
      IPs := TStringList.Create;
      IPs.Clear;
      IPs.LoadFromFile(Dir+'\ConnKlap.ini');
        //Memo1.Lines.LoadFromFile(Dir+'\ConnKlap.ini');

        DEF_CatalogName :=IPs.Strings[0];
        SQLServer:=IPs.Strings[1];
        MSSQL_CONN_STR :=IPs.Strings[4];
       { 'Provider=SQLOLEDB.1;Packet Size = 4096;Password=111;' +
        'Persist Security Info=True;User Id=TestUser;' +
        'Initial Catalog='+ DEF_CatalogName +';Data Source=DINAMO\'+SQLServer; }
        ADOConnection4:=TADOConnection.Create(Nil);
        ADOConnection4.LoginPrompt:=False;
        ADOConnection4.Connected:=False;
        ADOConnection4.ConnectionString := MSSQL_CONN_STR;
        ADOConnection4.Connected:=True;
         //Запятая лишняя
        Y:=Length(Razmernost);
        Delete(Razmernost,Y,1);
        //Создание временной таблицы
        TempTableQ:=TADOQuery.Create(nil);
        TempTableQ.Connection:=ADOConnection4;
        //Добавить поле ТЕХНОЛОГ, если фалш то деталь не обработана
        TempTableQ.SQL.Text:='CREATE TABLE #ClientToDBFL (Заказ nvarchar(100),Изделие nvarchar(100),Номер nvarchar(100),'+
        'Обозначение nvarchar(100) ,Элемент nvarchar(100),Количество nvarchar(100),БЗ nvarchar(100),IdГП nvarchar(100),IdКО nvarchar(100),КолКлап nvarchar(100),'+Razmernost+')';
        try
                TempTableQ.ExecSQL;
        except
                Err:='TempTableQ.ExecSQL;';
        end;
        //------------------------------------------------
      //Запятая лишняя
      Y:=Length(Str_Stanok);
      Delete(Str_Stanok,Y,1);

      //-------------------------------------------------
        for Y:=0 to SG2.RowCount-1 do
        begin
         if Y >0 then
         begin                        //  AND (SG2.Cells[2,Y]='')
                if (SG2.Cells[1,Y]<>'') then
                begin
                   //------------------------------------------------
                   S1:='';
                   Dliteln:=0;
                   Ocheredn:=0;
                   Col_Nom:=0;
                   A :=StrToInt(Column[0,2])+4;
                   i :=0;
                   Kol_Stan:=0;
                   fillchar(StrToCol,sizeof(StrToCol),0);//Очистка массмва
                   for X := 0 to High(Column) do  //Строка номеров колонок
                   begin
                        if Column[X,2]<>'' then
                        Begin   //'Пила,ПилаОч,ПилаДл,[Н\чПила],
                                //'False','0','0','0',
                                i :=0;
                                Col_Nom:=StrToInt(Column[X,2])-1;
                               // S1:=S1+#39+SG2.Cells[A,Y]+#39+','; //True  0
                                StrToCol[X, I ]:=SG2.Cells[A,Y];
                                Inc(A);
                                Inc(i);
                                //
                               // S1:=S1+#39+SG2.Cells[A,Y]+#39+','; //Ocheredn 1
                                StrToCol[X,i]:=SG2.Cells[A,Y];
                                Inc(A);
                                Inc(i);
                                //
                               // S1:=S1+#39+SG2.Cells[A,Y]+#39+','; //Dliteln 2
                                StrToCol[X,i]:=SG2.Cells[A,Y];
                                Inc(A);
                                Inc(i);
                                //
                                //S1:=S1+#39+SG2.Cells[A,Y]+#39+','; //Н\ч  3
                                StrToCol[X,i]:=SG2.Cells[A,Y];
                                Inc(A);
                                Inc(i);
                                //
                                //S1:=S1+#39+'0'+#39+',';  // Дата
                                StrToCol[X,i]:='0';
                                Inc(i);
                                //
                                StrToCol[X,i]:=Column[X,0];
                                Inc(i);
                                //
                        End;
                   end;
                   Kol_Klap:=SG2.Cells[0,Y];
                   if Kol_Klap='' Then
                   Kol_klap_I:=0
                   Else
                   Kol_Klap_I:=StrToInt(Kol_Klap);

                   ///()))))))))))))))))))))))(((((((((((((((((((((
                   Btn22Click(fileName1); ///()))))))))))))))))))))))(((((((((((((((((((((
                   ///()))))))))))))))))))))))(((((((((((((((((((((
                   //Запятая лишняя
                   i:=Length(S1);
                   Delete(S1,i,1);

                   Zakaz:=SG2.Cells[1,Y];
                   Obozn:=SG2.Cells[4,y];
                   Izdel:=SG2.Cells[2,y];
                   Kol_Det:=SG2.Cells[3,y];
                   Elem:=SG2.Cells[5,y];
                   Nomer:=SG2.Cells[6,y];
                   BZ:=SG2.Cells[7,y];
                   IDGP:=SG2.Cells[8,y];
                   IDKO:=SG2.Cells[10,y];
                   //Кол-во деталей уже умножено на кол клапанов!!!!
                   TempTableQ.SQL.Text:='INSERT INTO #ClientToDBFL '+
                        '(Заказ,Изделие,Обозначение,Элемент,'+
                        'Количество,Номер,БЗ,IdГП, IdКО,КолКлап,'+Str_Stanok+') Values ('
                        +#39+Zakaz+#39+
                        ','+#39+Izdel+#39+
                        ','+#39+Obozn+#39+
                        ','+#39+Elem+#39+
                        ','+#39+Kol_Det+#39+
                        ','+Nomer+
                        ','+#39+BZ+#39+
                        ','+#39+IDGP+#39+ ','+#39+IDKO+#39+
                        ','+#39+Kol_Klap+#39+
                        ','+S1+
                        ')';

                        try
                        TempTableQ.ExecSQL;
                        except
                        Err:='1222';
                        end;
                end;

         end;
        end;
        //Группировка по TRUE
        S1:='';
        A:=1;
        Glob:=1;
        sg3 := TStringGrid.Create(Nil );
        Clear_StringGrid(sg3);
    SG3.ColCount:=17;
    sg3.Cells[0,0]:='КолКлап';
    sg3.Cells[1,0]:='Заказ';
    sg3.Cells[2,0]:='Изделие';
    sg3.Cells[3,0]:='Номер';

    sg3.Cells[4,0]:='Обозначение';
    sg3.Cells[5,0]:='Элемент';
    sg3.Cells[6,0]:='Количество';
    sg3.Cells[7,0]:='Очередность'; //Очередность
    sg3.Cells[8,0]:='Длительность';//Длительность
    sg3.Cells[9,0]:='Н\ч';//Н\ч
    sg3.Cells[10,0]:='БЗ';
    sg3.Cells[11,0]:='IdГП';
    sg3.Cells[12,0]:='Станок';
    sg3.Cells[13,0]:='Планирование';
    sg3.Cells[14,0]:='Заготовка';
    sg3.Cells[15,0]:='Длит Общая';
    sg3.Cells[16,0]:='IdКО';
    //Memo12:=TStringList.Create;
    //Memo12.Clear;
        for Y:=0 to High(Column) do  //Строка номеров колонок
        begin
                if Column[y,0]<>'' then
                Begin
                        TempTableQ.SQL.Text:='SELECT *  FROM #ClientToDBFL'+
                        ' WHERE '+Column[y,0]+'='+#39+'True'+#39 ;
                        TempTableQ.Open;
                        //+++++++++++++++++++++++++++++++++++++++
                        TempTableQ.First;
                        if TempTableQ.RecordCount<>0 Then
                        Begin
                                sg3.RowCount:= sg3.RowCount+1 ;
                                sg3.Cells[0,A]:=Column[y,0];
                                Stanok:=Column[y,0]; //Значение поля
                                A:=A+1;

                          for I := 0 to TempTableQ.RecordCount - 1 do
                          begin
                                sg3.RowCount:= sg3.RowCount+1 ;
                                Str:=sg3.Cells[0,A];
                                sg3.Cells[0,A]:=TempTableQ.FieldByName('КолКлап').AsString;//IntToStr(Glob);
                                sg3.Cells[1,A]:=TempTableQ.FieldByName('Заказ').AsString;
                                sg3.Cells[2,A]:=TempTableQ.FieldByName('Изделие').AsString;
                                sg3.Cells[3,A]:=TempTableQ.FieldByName('Номер').AsString;

                                sg3.Cells[4,A]:=TempTableQ.FieldByName('Обозначение').AsString;
                                sg3.Cells[5,A]:=TempTableQ.FieldByName('Элемент').AsString;
                                Kol_Det:=TempTableQ.FieldByName('Количество').AsString;
                                sg3.Cells[6,A]:= Kol_Det;
                                if Kol_Det='' Then
                                Kol_Det_I:=0
                                Else
                                Kol_Det_I:=StrToInt(Kol_Det);

                                sg3.Cells[7,A]:=TempTableQ.FieldByName(Stanok+'Оч').AsString; //Очередность
                                Ocheredn:=StrToInt( sg3.Cells[7,A]);
                                sg3.Cells[8,A]:=TempTableQ.FieldByName(Stanok+'Дл').AsString;//Длительность
                                Dliteln:=StrToFloat(StringReplace(sg3.Cells[8,A],',','.',[rfReplaceAll]));
                                sg3.Cells[9,A]:=TempTableQ.FieldByName('Н\ч'+Stanok).AsString;//Н\ч
                                NCH:=StringReplace(sg3.Cells[9,A],',','.',[rfReplaceAll]);
                                if NCh='' Then
                                NC_D:=0
                                Else
                                NC_D:=StrToFloat(NCH);
                                sg3.Cells[10,A]:=TempTableQ.FieldByName('БЗ').AsString;
                                sg3.Cells[11,A]:=TempTableQ.FieldByName('IdГП').AsString;
                                sg3.Cells[12,A]:=Column[y,0];
                                sg3.Cells[13,A]:=Dat_Glob;

                                Dat_Z :=TempTableQ.FieldByName(Stanok+'Дата').AsString;
                                Dat_Zag:= StrToDate(Dat_Z);
                                Nowa_S:=FormatDateTime('dd.mm.YYYY',Now);
                                Nowa:=StrToDate(Nowa_s);
                                {if Dat_Zag<=Nowa then //Дата начала заготовки МЕНЬШЕ сегодняшней
                                begin
                                        MessageDlg(Nom_G+'  Дата начала заготовки МЕНЬШЕ сегодняшней!' + #10#13 +
                                        'Необходимо сместить Дату Планирования на ' + FloatToStr(Nowa-Dat_Zag)+' Дней!' ,mtInformation, [mbOk], 0);
                                        Flag_Error:=1;
                                        Exit;
                                end; }
                                Dat_Sbor:=TempTableQ.FieldByName(Stanok+'Дата').AsString;
                                Dat_Sbor_D:=StrToDate(Dat_Sbor)+1;
                                sg3.Cells[13,A]:=FormatDateTime('dd.mm.YYYY',Dat_Sbor_D);
                                sg3.Cells[14,A]:=TempTableQ.FieldByName(Stanok+'Дата').AsString;
                                sg3.Cells[15,A]:=TempTableQ.FieldByName(Stanok+'ДлОбщ').AsString;
                                sg3.Cells[16,A]:=TempTableQ.FieldByName('IdКО').AsString;
                                Dlit:=StringReplace(sg3.Cells[8,A],',','.',[rfReplaceAll]);
                                Dlit_Ob:=StringReplace(sg3.Cells[15,A],',','.',[rfReplaceAll]);

                                //++++++++++++++++++++++++++++++++++++++++++++++++++
                                if sg3.Cells[14,A]='' then
                                        Dat_S1:='NULL'
                                else
                                        DaT_S1:=#39+ConvertDat1(SG3.Cells[14,A])+#39;
                                                        //   AND (Технолог=1)
                                if not Conn_Klap.mkQuerySelect3(
                                'Select * from %s Where (Обозначение='+#39+Obozn+#39+')'
                                , ['СпецифОбщая']) then
                                exit;

                            if  (Conn_Klap.AQuery3.RecordCount<>0)  then
                            Begin
                                NCH:=StringReplace(FloatToStr(NC_D*Kol_Det_I),',','.',[rfReplaceAll]);
                                if not Conn_Klap.mkQueryInsert(
                                'INSERT INTO Заготовка '+
                                '(Заказ,Изделие,Номер,Обозначение,Элемент,'+
                                'Количество,Очередность,Длительность,НЧ,БЗ,IdГП,Станок,Планирование,Заготовка,ДлительностьОбщ,КолКлап,IdКО) Values ('
                                +#39+sg3.Cells[1,A]+#39+
                                ','+#39+sg3.Cells[2,A]+#39+
                                ','+#39+sg3.Cells[3,A]+#39+
                                ','+#39+sg3.Cells[4,A]+#39+
                                ','+#39+sg3.Cells[5,A]+#39+
                                ','+#39+sg3.Cells[6,A]+#39+
                                ','+#39+sg3.Cells[7,A]+#39+
                                ','+#39+Dlit+#39+
                                ','+#39+NCH+#39+
                                ','+#39+sg3.Cells[10,A]+#39+
                                ','+#39+sg3.Cells[11,A]+#39+
                                ','+#39+sg3.Cells[12,A]+#39+
                                ','+#39+ConvertDat1(sg3.Cells[13,A])+#39+
                                ','+Dat_S1+
                                ','+#39+Dlit_Ob+#39+
                                ','+#39+sg3.Cells[0,A]+#39+
                                ','+#39+sg3.Cells[16,A]+#39+
                                ')',[] )then
                                Exit;
                                //+++++++++++++++++++++++++++++++++++++++++++
                                A:=A+1;
                                Glob:=Glob+1;
                            End;
                                TempTableQ.Next;

                          end;
                        end;

                end;
        end;
        TempTableQ.Close;
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        if not Conn_Klap.mkQuerySelect('Select Заготовка From Заготовка Where (Номер='+
        #39+Nomer_Glob+#39+') Group By Заготовка',[]) Then
        Exit;
        For I:=0 To  Conn_Klap.AQuery.RecordCount-1 Do
        Begin
              Dat_S:= ConvertDat1(Conn_Klap.AQuery.FieldByName('Заготовка').AsString);

              Memo11.Add('');
              Memo11.Add(Dat_S);
              if not Conn_Klap.mkQuerySelect3('Select Станок, SUM(НЧ) As SNH From Заготовка Where (Заготовка='+
                 #39+Dat_S+#39+') AND (Станок<>'+#39+'Канбан'+#39+') AND (Станок<>'+#39+'Trumph'+#39+') Group by Станок',[]) Then
                Exit;

              For Y:=0 To  Conn_Klap.AQuery3.RecordCount-1 Do
              Begin
                   Memo11.Add(Conn_Klap.AQuery3.FieldByName('Станок').AsString+'  Н\ч: '+Conn_Klap.AQuery3.FieldByName('SNH').AsString);
                   Inc(A);
                   Conn_Klap.AQuery3.Next;
              end;
              Conn_Klap.AQuery.Next;
        end;
        SNCH:=Memo11.Text;
        //================================================
        TempTableQ.SQL.Text:='DROP TABLE #ClientToDBFL';
        try
                TempTableQ.ExecSQL;
        except
        end;

   End;
//++++++++++++++++++++++++++++++++++++++++++++++++++
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function Osnova_Main.btn22Click(Err2:String):Boolean;
var //StrToCol:array[0..100,0..6] of string;
//[TRUE,очередь, длительность, н/Ч, Дата,Stanok] ДЛЯ ОДНОЙ ДЕТАЛИ
I,X,J,y,A,Res:Integer;
Res_True,min_Col,Col,Col_1,
Col_Grid,Res1:Integer;
Izdel,Nowa_S:string;
Dat_Plan,Dat_Zag,Nowa:TDate;
Dat_P,Dat_Z:string;
Dlit,Dlit_Min,Dlit_Obsh,Dlit_Max,
Min,Och_Min,Och_Max,Dlit_ObshO,Ocher,Ocher_Pred:Double;
//SREZ1: array[0..10,0..100] of string;

SKPU :array[0..10,0..100] of string;

//Sort: TStringList;
Stanok,Stanok_Posl:String;
Kol_Smen,Max,F1,Den,Den_pred,Celoe,Res_Trum:Integer;
Chas,H1,H2,H3,Den1,Dlit_Trum:Double;
SREZ1:TAdvStringGrid;
//AD1:TAdvStringGrid;
Chet:Integer;
begin
     X:=0;
     J:=0;
     Dlit:=0;
     Dlit_Obsh:=0;
     Col_Grid:=97;
     Col_1:=0;
     Max:=0;
     SyStem.SysUtils.FormatSettings.DecimalSeparator :=('.');
     //Sort:=TStringList.Create;
     SREZ1 := TAdvStringGrid.Create(Nil );
     //AD1 := TAdvStringGrid.Create(Nil );
     fillchar(Temp,sizeof(Temp),0);//Очистка массмва
     fillchar(SKPU,sizeof(SKPU),0);//Очистка массмва
     Form1.Memo13.Clear;
     //++Добавим к длительности Трумпф 1 день
    { for I := 0 to High(StrToCol) do
     begin
        Res_True:=Pos('rue',StrToCol[i,0]);
        Res_Trum:=Pos('Trumph',StrToCol[i,5]);
        if (Res_Trum<>0) AND (Res_True<>0) then  //ХУЙ
        begin
                Dlit_Min:=StrToFloat(StrToCol[i,2]);
                StrToCol[i,2]:=FloatToStr(Dlit_Min+1);
        end;
     end; }
     Dlit_Min:=0;
     //++++++++++++++++++++++++++++++++Ищем Минимальное по ОЧЕРЕДНОСТИ
     for I := 0 to High(StrToCol) do
     begin
        Res_True:=Pos('rue',StrToCol[i,0]);
        if Res_True<>0 then
        begin
                min := StrToInt(StrToCol[i,1]); // Первый не нулевой
                Och_Min:= StrToInt(StrToCol[i,1]);
                Col:=I;
                Dlit_Min:=StrToFloat(StrToCol[i,2]);
                Break;
        end;
     end;
     X:=0;
     A:=0;
     Y:=0;
     SREZ1.ColCount:=8;//AD   SREZ
     SREZ1.RowCount:=1;
     //Form1.SKPU.ColCount:=10; //AD1
     //Form1.SKPU.RowCount:=1;

     Dlit_Obsh:=0;
     for I := 0 to High(StrToCol) do
     begin
        Res_True:=Pos('rue',StrToCol[i,0]);
        if Res_True<>0 then
      begin
        Stanok_Posl:=StrToCol[i,5];
        For y:=0 To High(Smena) do
        Begin
          Stanok:=Smena[y,0];
          Res:=AnsiCompareStr(Stanok_Posl,Stanok);
          if Res=0 Then
          Begin
                Kol_Smen:=StrToInt(Smena[y,1]);
                Chas:=StrToFloat(Smena[y,2]);
                SREZ1.Cells[0,A]:=StrToCol[i,1]; //Ocher
                Ocher:= StrToInt(StrToCol[i,1]);

                SREZ1.Cells[1,A]:=Stanok_Posl;   //Stanok
                SREZ1.Cells[2,A]:=StrToCol[i,2];//Dlit
                Dlit_Min:=StrToFloat(StrToCol[i,2]);//*8;// 0,5*8=4 - пол смены
                                                    //1*8=8 полная смена
                Dlit_Obsh:=Dlit_Obsh+ (StrToFloat(StrToCol[i,2]));//В ЧАСАХ

                SREZ1.Cells[3,A]:=Smena[y,1];//Kol_Smen
                Kol_Smen:=StrToInt(Smena[y,1]);

                SREZ1.Cells[4,A]:=Smena[y,2];//Chas
                SREZ1.Cells[5,A]:='0';//Den
                SREZ1.Cells[6,A]:=StrToCol[I,3];//Н\Ч
                SREZ1.Cells[7,A]:=StrToCol[I,0];//TRUE
                SREZ1.RowCount:=SREZ1.RowCount+1;
                Inc(A);
          end;
        End;
      End;
     End;
     F1:=0;
     SREZ1.Cells[5,0]:='1';
     //AD.Sort(0,sdAscending);//sdDescending По убыванию
     SREZ1.SortByColumn(0);
     Dat_Plan:=StrToDate(Dat_Glob);
     A:=0;
     For I:=0 To SREZ1.RowCount-1 Do
     Begin
       if SREZ1.Cells[0,i]<> '' Then
       Begin
           {Form1.SKPU.Cells[0,A]:=Form1.SREZ1.Cells[0,i]; //Ocher
           Form1.SKPU.Cells[1,A]:=Form1.SREZ1.Cells[1,i];   //Stanok
           Form1.SKPU.Cells[2,A]:=Form1.SREZ1.Cells[2,i];//Dlit
           Form1.SKPU.Cells[3,A]:=Form1.SREZ1.Cells[3,i];//Kol_Smen
           Form1.SKPU.Cells[4,A]:=Form1.SREZ1.Cells[4,i];//Chas
           Form1.SKPU.Cells[5,A]:='0';//Den
           Form1.SKPU.Cells[6,A]:=Form1.SREZ1.Cells[6,i];//Н\Ч
           Form1.SKPU.Cells[7,A]:=Form1.SREZ1.Cells[7,i];//TRUE
           Form1.SKPU.RowCount:=Form1.SKPU.RowCount+1;}
           SKPU[0,a]:=SREZ1.Cells[0,i]; //Ocher
           SKPU[1,a]:=SREZ1.Cells[1,i];   //Stanok
           SKPU[2,a]:=SREZ1.Cells[2,i];//Dlit
           SKPU[3,a]:=SREZ1.Cells[3,i];//Kol_Smen
           SKPU[4,a]:=SREZ1.Cells[4,i];//Chas
           SKPU[5,a]:='0';//Den
           SKPU[6,a]:=SREZ1.Cells[6,i];//Н\Ч
           SKPU[7,a]:=SREZ1.Cells[7,i];//TRUE
           Inc(A);
       end;
     end;

     //For I:=0 To Form1.SKPU.RowCount-1 Do
     For i:=0 To High(SKPU) do
     Begin
       //if Form1.SKPU.Cells[0,I]<> '' Then
       if SKPU[0,I]<> '' Then
       Begin

            Ocher:= StrToInt(SKPU[0,I]);//Ocher
            Stanok_Posl:=SKPU[1,i];   //Stanok
            Dlit_Min:=StrToFloat(SKPU[2,I]);//Dlit

            Kol_Smen:=StrToInt(SKPU[3,I]);//Kol_Smen
            Chas:=StrToFloat(SKPU[4,I]);
            Den:= StrToInt(SKPU[5,I]);
            Celoe:=Pos(',',SKPU[2,I]);
            if Celoe<>0 Then
            Begin
                if I=0 Then
                Begin
                        H1:=(1-Dlit_Min);
                        SKPU[5,0]:=FloatToStr(1);
                        Form1.Memo13.Lines.Add('1');
                End;
                if I=1 Then
                Begin
                 if (H1=0.5) Then //Остаток часов от первого передела
                 Begin
                   SKPU[5,1]:='1';
                   Form1.Memo13.Lines.Add('1');
                   H2:=1;
                  End;
                 if (H1=0) OR (H1<0) Then    //ХУЙ   OR (H1<0)
                 Begin
                   SKPU[5,1]:='2';   // переносим на второй день
                   Form1.Memo13.Lines.Add('2');
                   H2:=2;
                  End;
                  H1:=(1-Dlit_Min);
                end;
                if I>1 Then //
                Begin
                 Chet:=I MOD 2;//Остаток от деления
                 if (Chet<>0) AND (Dlit_Min=0.5) Then //НЕ Четное
                 Begin
                    H1:=Ocher-Dlit_Min-H2;
                    Den1:= SimpleRoundTo(H1,0);
                    SKPU[5,I]:=(FloatToStr(Den1));
                    Form1.Memo13.Lines.Add(FloatToStr(Den1));
                 end;
                 if (Chet=0) AND (Dlit_Min=0.5) Then //Четное
                 Begin
                     H2:=Ocher-Dlit_Min-H1;
                     Den1:= SimpleRoundTo(H2,0);
                     SKPU[5,I]:=(FloatToStr(Den1));
                     Form1.Memo13.Lines.Add(FloatToStr(Den1));
                 end;
                end;
            End
            Else
            Begin
                  SKPU[5,I]:=(FloatToStr(Ocher));
                  Form1.Memo13.Lines.Add(FloatToStr(Ocher));
            end;
            {Den:= StrToInt(AD1.Cells[5,I]);
            Dat_Plan:=StrToDate(Dat_Glob);
            Dat_Z:=FormatDateTime('mm.dd.YYYY',Dat_Plan-Den);
            AD1.Cells[8,i]:=Dat_Z;  }
       end;
     end;
     A:= Form1.Memo13.Lines.Count-1 ;
     For i:=0 To High(SKPU) do
     Begin
       if Form1.Memo13.Lines.Strings[a]<> '' Then
       Begin
           SKPU[9,I]:=Form1.Memo13.Lines.Strings[a];
       end;
       A:=A-1;
     End;
     For i:=0 To High(SKPU) do
     Begin
       if SKPU[9,I]<> '' Then
       Begin
         Den:= StrToInt(SKPU[5,I]);//В Данном случае Дат_Глоб- это Дата начала ЗАГОТОВКИ
         Dat_Plan:=StrToDate(Dat_Glob);  //
         Dat_Z:=FormatDateTime('mm.dd.YYYY',Dat_Plan+Den-1);//ПЛЮС 1ДЕНЬ
         SKPU[8,i]:=Dat_Z;

       end;
     End;

     for I := 0 to High(StrToCol) do
     begin
          if StrToCol[i,5]<>'' Then
          Begin
            For a:=0 To High(SKPU) do
            Begin
                Res:=AnsiCompareStr(StrToCol[i,5],SKPU[1,A]);
                if Res=0 Then
                Begin
                    StrToCol[i,4]:=FloatToStr(Dlit_Obsh);
                    StrToCol[i,6]:=SKPU[8,A];//Дата заготовка
                end;
            End;
          End;
     end;

     for I := 0 to High(StrToCol) do
     begin
        if StrToCol[i,5]<>'' then
        begin
        S1:=S1+#39+StrToCol[I,0]+#39+','; //True  0
        S1:=S1+#39+StrToCol[I,1]+#39+','; //Очередность 1
        S1:=S1+#39+StrToCol[I,2]+#39+','; //Длитель  2
        S1:=S1+#39+StrToCol[I,3]+#39+','; //Н\ч  3
        if StrToCol[I,2]='' then
          S1:=S1+'NULL,';
        //
        if (StrToCol[I,2]<>'') then
        Begin
            S1:=S1+#39+StrToCol[I,6]+#39+',';
        End;
          //
          S1:=S1+#39+StrToCol[I,4]+#39+',';
          //Дата Заготовки  5 будем искать во временной базе
        end;
     end;
     FreeAndNil(Srez1);
     //FreeAndNil(SKPU);
end;

                                                //Запуск                 //Tab2 Спецификация
function Osnova_Main.Osnova( var Dat : string;  const tab1:string;tab2:string;
Nom:string;fileName:String;Flag_Polz,Cvet:Integer ) : Boolean;

var a,Y,Glob,i,x,Res,T,Res1,L:Integer;
Stanok,Dat_S1,IDGP,Kol_Klap,Izdel,kol_Det,Elem,Obozn,Nomer,BZ,Str1,Str2,Dlit,Kod,
IDKO,Nam_S,ZAKAZ,Dl,Sh,DL_R,Sh_R,Mat,Diam,Mass,Kol_G,Sh1,Sh2,Ie:string;
Dlit_D,Och_D:Double;
Res_Teh,Res_Stan,Memo12_Count,Res_Tab,Polzov,Zavesa,//
KOL_RECORD_Tab1,KOL_RECORD_Tab2:Integer;
Memo12,Memo13:TStringList;
Memo12_text,Nam,God,Mes,Dir:String;
Leg:String;
h: hwnd;
begin
        Result:=False;
        Zavesa:=0;//НЕ ЗАВЕСА
        SyStem.SysUtils.FormatSettings.DecimalSeparator :=('.');
    try
        T:=0;
        Nom_G:=Nom;
        Conn_Ceh := Connect_Miass_Ceh.Create();
        Conn_Klap:=Connect_Miass_Klap.Create();
        Memo12:=TStringList.Create;
        Memo12.Clear;
        Memo13:=TStringList.Create;
        Memo13.Clear;
        Polzov:= Flag_Polz;
        if Polzov=7 then
        Begin
           Zavesa:=1;//ЗАВЕСА
           Polzov:=6;
        End;
        CVet1:=Cvet;
        if not Conn_Klap.mkQuerySelect('SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME ='+#39+'СпецифОбщая'+#39
        , []) then
        exit;
        Dat_Glob:= Dat;
        for I:=0 To Conn_Klap.AQuery.RecordCount-1 do
        begin
          Str1:=Conn_Klap.AQuery.FieldByName('COLUMN_NAME').AsString;
          Str2:=Conn_Klap.AQuery.FieldByName('COLUMN_NAME').AsString;
          Res:=Pos('Н\ч',str2);
          if Res<>0 then
          begin
               Delete(str2,1,3);
               Column[T,0]:=Str2;
               Column[T,1]:=Str2+' BIT';
               Column[T,2]:=IntToStr(T+7);
               Inc(T);
          end;
          Conn_Klap.AQuery.Next;
        End;
        //++++++++++++++++++++++++++++
        fillchar(Smena,sizeof(Smena),0);//Очистка массмва
        if not Conn_Klap.mkQuerySelect('Select * From Смены  ' , []) then //Запуск
        exit;
      For I:=0 to Conn_Klap.AQuery.RecordCount-1 Do
      Begin
        Smena[i,0]:=Conn_Klap.AQuery.FieldByName('Передел').AsString;
        Smena[i,1]:=Conn_Klap.AQuery.FieldByName('Смена').AsString;
        Smena[i,2]:=Conn_Klap.AQuery.FieldByName('Часы').AsString;
        Conn_Klap.AQuery.Next;
      End;
          sg2 := TStringGrid.Create(Nil);
          Clear_StringGrid(sg2);
        //sg2.Clear;
        A:=1;
        sg2.ColCount:=23;
        sg2.Cells[0,0]:='№';
        sg2.Cells[1,0]:='Заказ';
        sg2.Cells[2,0]:='Изделие';
        sg2.Cells[3,0]:='Кол во';
        sg2.Cells[4,0]:='Обозначение';
        sg2.Cells[5,0]:='Элемент';
        sg2.Cells[6,0]:='Номер';
        sg2.Cells[7,0]:='БЗ';
        sg2.Cells[8,0]:='IdГП';
        sg2.Cells[9,0]:='Технолог';
        sg2.Cells[10,0]:='IdКО';
        sg2.Cells[11,0]:='Легион';
        //Dl,Sh,DL_R,Sh_R,Mat,Diam,Mass,Kol_G,Sh1,Sh2,Ie
        sg2.Cells[12,0]:='Длина';
        sg2.Cells[13,0]:='Ширина';
        sg2.Cells[14,0]:='ДлинаРазв';
        sg2.Cells[15,0]:='ШиринаРазв';
        sg2.Cells[16,0]:='Диаметр';
        sg2.Cells[17,0]:='Масса';
        sg2.Cells[18,0]:='КолГибов';
        sg2.Cells[19,0]:='ШиринаПолки1';
        sg2.Cells[20,0]:='ШиринаПолки2';
        sg2.Cells[21,0]:='ЕИ';
        sg2.Cells[22,0]:='Материал';
        for y := 0 to High(Column) do
        begin
                if Column[y,0]<>'' then
                Begin
                        Stanok:=Column[y,0];
                        SG2.ColCount:=SG2.ColCount+4;
                        sg2.Cells[10+A+1,0]:=Stanok;
                        sg2.Cells[10+A+2,0]:=Stanok+'Оч'; //Очередность
                        sg2.Cells[10+A+3,0]:=Stanok+'Дл';//Длительность
                        sg2.Cells[10+A+4,0]:='Н\ч'+Stanok;//Н\ч
                        //sg2.Cells[9+A+5,0]:=Stanok+'Дата'; //дата
                        A:=A+4;
                end;
        end;
        Glob:=1;//SG2.RowCount;
        //Dat_S1:=FormatDateTime('mm.dd.YYYY',Dtp1.DateTime);
        if Polzov<4 then
        begin
          Res1:=Pos('-2',Nom);
          Res:=Pos('-3',Nom);
          if (CVet1>10) and (Res1<>0) then
          begin
          if not Conn_Klap.mkQuerySelect('Select * From %s Where (IDГП='+#39+'%s'+#39+') ', ['СТАМ',IntToStr(CVet1)]) then //Запуск
          exit;
          end;
          if (CVet1>10) and (Res<>0) then
          begin
          if not Conn_Klap.mkQuerySelect('Select * From %s Where (IDГП='+#39+'%s'+#39+') ', ['ЛЮК',IntToStr(CVet1)]) then //Запуск
          exit;
          end;
          if (Res=0) and (Res1=0) then
          if not Conn_Klap.mkQuerySelect('Select * From %s Where (Номер=%s) '+
          'AND (Отмена IS NULL) ' , [tab1,Nom]) then //Запуск
          exit;
          KOL_RECORD_Tab1:= Conn_Klap.AQuery.RecordCount;
        end;
        ///Вентиляторы из Цеха
        if (Polzov=4) or (Polzov=6) then
        begin
          Res:=Pos('-1',Nom);
          if (CVet1>10) and (Res<>0) then
          begin
          if not Conn_Ceh.mkQuerySelect('Select * From %s Where (IDГП='+#39+'%s'+#39+') ', ['Завесы',IntToStr(CVet1)]) then //Запуск
          exit;
          end

          else
          if not Conn_Ceh.mkQuerySelect('Select * From %s Where (Номер=%s) '+
          'AND (Отмена IS NULL) ' , [tab1,Nom]) then //Запуск
          exit;
          KOL_RECORD_Tab1:= Conn_Ceh.AQuery.RecordCount;
        end;
        ///КЦ из Цеха    вместо номера Брикет
        if Polzov=5 then
        begin
          if not Conn_Ceh.mkQuerySelect('Select * From %s Where (Брикет=%s) '+
          'AND (Отмена IS NULL) AND (Статус=1) AND ПускПлан='+#39+ConvertDat1(Dat)+#39 , [tab1,Nom]) then //КЦ
          exit;
          KOL_RECORD_Tab1:= Conn_Ceh.AQuery.RecordCount;
        end;
        //
        if Polzov=55 then
        begin
        Res:=Pos('-55',Nom);
          if (CVet1>10) and (Res<>0) then
          begin
          if not Conn_Ceh.mkQuerySelect('Select * From %s Where (IDГП='+#39+'%s'+#39+') AND (IDКО='+#39+'%s'+#39+') '
          , ['Комп',IntToStr(CVet1),(fileName)]) then //CVet1=GP fileName  =KO
          exit;
          fileName:='';
          end

          else             //  AND (Статус=1)
          if not Conn_Ceh.mkQuerySelect('Select * From %s Where ([№ бланк заказа]=%s) '+
          'AND (Отмена IS NULL)  AND ПускПлан='+#39+ConvertDat1(Dat)+#39 , [tab1,Nom]) then //КЦ
          exit;
          KOL_RECORD_Tab1:= Conn_Ceh.AQuery.RecordCount;
        end;
        SG2.RowCount:= 2;
        SG2.FixedRows:=1;
        A:=1;
      For I:=0 to KOL_RECORD_Tab1-1 Do
      Begin
        if Polzov<4 then
        begin
          IDGP:=Conn_Klap.AQuery.FieldByName('IdГП').AsString;
          IDKO:=Conn_Klap.AQuery.FieldByName('IdКО').AsString;
          Izdel:=Conn_Klap.AQuery.FieldByName('Изделие').AsString;
          Kol_Klap:=Conn_Klap.AQuery.FieldByName('Кол во запущенных').AsString;
          Nomer:=Conn_Klap.AQuery.FieldByName('Номер').AsString;
          Nomer_Glob:=Conn_Klap.AQuery.FieldByName('Номер').AsString;
          BZ:=Conn_Klap.AQuery.FieldByName('БЗ').AsString;
          ZAKAZ:=Conn_Klap.AQuery.FieldByName('Заказ').AsString;
          Leg:=Conn_Klap.AQuery.FieldByName('Легион').AsString;

        end;
        if Polzov=4 then
        begin
          IDGP:=Conn_Ceh.AQuery.FieldByName('IdГП').AsString;
          IDKO:=Conn_Ceh.AQuery.FieldByName('IdКО').AsString;
          Izdel:=Conn_Ceh.AQuery.FieldByName('Изделие').AsString;
          Kol_Klap:=Conn_Ceh.AQuery.FieldByName('Кол запущенных').AsString;
          Nomer:=Conn_Ceh.AQuery.FieldByName('Номер').AsString;
          Nomer_Glob:=Conn_Ceh.AQuery.FieldByName('Номер').AsString;
          BZ:=Conn_Ceh.AQuery.FieldByName('БЗ').AsString;
          ZAKAZ:=Conn_Ceh.AQuery.FieldByName('Заказ').AsString;
          //Leg:=Conn_Klap.AQuery.FieldByName('Легион').AsBoolean;
        end;
        if (Polzov=5)  then  //ВЕВР
        begin
          IDGP:=Conn_Ceh.AQuery.FieldByName('IdГП').AsString;
          IDKO:=Conn_Ceh.AQuery.FieldByName('IdКО').AsString;
          Izdel:=Conn_Ceh.AQuery.FieldByName('Наименование').AsString;
          Kol_Klap:=Conn_Ceh.AQuery.FieldByName('Кол-во').AsString;
          Nomer:=Conn_Ceh.AQuery.FieldByName('Брикет').AsString;
          Nomer_Glob:=Conn_Ceh.AQuery.FieldByName('Брикет').AsString;
          BZ:=Conn_Ceh.AQuery.FieldByName('№ бланк заказа').AsString;
          ZAKAZ:=Conn_Ceh.AQuery.FieldByName('Заказ').AsString;
        end;
        if  (Polzov=55) then  //

        begin
          IDGP:=Conn_Ceh.AQuery.FieldByName('IdГП').AsString;
          IDKO:=Conn_Ceh.AQuery.FieldByName('IdКО').AsString;
          Izdel:=Conn_Ceh.AQuery.FieldByName('Наименование').AsString;
          Kol_Klap:=Conn_Ceh.AQuery.FieldByName('Кол-во').AsString;
          if Kol_Klap='0' then
          Kol_Klap:='1';
          Nomer:=Conn_Ceh.AQuery.FieldByName('№ бланк заказа').AsString;
          Nomer_Glob:=Conn_Ceh.AQuery.FieldByName('№ бланк заказа').AsString;
          BZ:=Conn_Ceh.AQuery.FieldByName('№ бланк заказа').AsString;
          ZAKAZ:=Conn_Ceh.AQuery.FieldByName('Заказ').AsString;
        end;

        if Polzov=6 then   //ЭКО
        begin
          IDGP:=Conn_Ceh.AQuery.FieldByName('IdГП').AsString;
          IDKO:=Conn_Ceh.AQuery.FieldByName('IdКО').AsString;
          res:=Pos('Заве',tab1);
          if Res=0 then
          Begin
            Izdel:=Conn_Ceh.AQuery.FieldByName('Наименование').AsString ;
            Kol_Klap:=Conn_Ceh.AQuery.FieldByName('Кол-во').AsString;
            BZ:=Conn_Ceh.AQuery.FieldByName('№ бланк заказа').AsString;
          End
          else
          Begin
            Izdel:=Conn_Ceh.AQuery.FieldByName('Изделие').AsString; //Завесы
            Kol_Klap:=Conn_Ceh.AQuery.FieldByName('Кол запущенных').AsString;
            BZ:=Conn_Ceh.AQuery.FieldByName('БЗ').AsString;
          End;
          Nomer:=Conn_Ceh.AQuery.FieldByName('Номер').AsString;
          Nomer_Glob:=Conn_Ceh.AQuery.FieldByName('Номер').AsString;
          ZAKAZ:=Conn_Ceh.AQuery.FieldByName('Заказ').AsString;
        end;
        if Tab2='' Then //Tab2 пустое значит спецификация уже в СпецифОбщая  , ищем по названию изделия СТАМ....
        Begin

                 if not Conn_Klap.mkQuerySelect2(
                'Select * from %s Where (ВидЭлемента='+#39+'Деталь'+#39+') '+
                'AND (([Изделие]='+#39+'Стакан монтажный '+Izdel+#39+') OR ([Изделие]='+#39+'Поддон '+Izdel+#39+'))'
                 , ['СпецифОбщая']) then  //Специф
                 exit;
                 For X:=0 to Conn_Klap.AQuery_Del.RecordCount-1 Do
                Begin
                        Obozn:=Conn_Klap.AQuery_Del.FieldByName('Обозначение').AsString;
                        Elem:=Conn_Klap.AQuery_Del.FieldByName('Элемент').AsString;
                        Kol_Det:=Conn_Klap.AQuery_Del.FieldByName('Количество').AsString;
                   if  Conn_Klap.AQuery_Del.RecordCount<>0 then
                   Begin
                        SG2.RowCount:=SG2.RowCount+ 1;
                        sg2.Cells[0,Glob+1]:=Kol_Klap;//IntToStr(Glob+1);

                        sg2.Cells[1,Glob+1]:=Conn_Klap.AQuery.FieldByName('Заказ').AsString;
                        sg2.Cells[2,Glob+1]:=Conn_Klap.AQuery.FieldByName('Изделие').AsString;

                        sg2.Cells[3,Glob+1]:=IntToStr(StrToInt(Kol_Klap)*StrToInt(Kol_Det));
                        sg2.Cells[4,Glob+1]:=Conn_Klap.AQuery_Del.FieldByName('Обозначение').AsString;
                        sg2.Cells[5,Glob+1]:=Conn_Klap.AQuery_Del.FieldByName('Элемент').AsString;
                        sg2.Cells[6,Glob+1]:=Nom;
                        sg2.Cells[7,Glob+1]:=BZ;
                        sg2.Cells[8,Glob+1]:=IDGP;
                        sg2.Cells[9,Glob+1]:=Conn_Klap.AQuery3.FieldByName('Технолог').AsString;
                        sg2.Cells[10,Glob+1]:=IDKO;
                        sg2.Cells[11,Glob+1]:=(Leg);

                        Res_Teh:=Pos('rue',sg2.Cells[9,Glob+1]);
                        if (Res_Teh=0)  Then //ТеХНОЛОГ FALSE
                        begin
                                  Memo12.Add('ТЕХНОЛОГ FALSE Заказ '+sg2.Cells[1,Glob+1]+' * '+sg2.Cells[2,Glob+1]+' * Деталь '+sg2.Cells[4,Glob+1]);
                                  Memo13.Add(sg2.Cells[4,Glob+1]);
                        end;
                        A:=1;
                        for y := 0 to High(Column) do
                        begin
                                if Column[y,0]<>'' then
                                Begin
                                        Stanok:=Column[y,0]; //Значение поля

                                        sg2.Cells[10+(A+1),Glob+1]:=Conn_Klap.AQuery_Del.FieldByName(Stanok).AsString;
                                        Res_Stan:=Pos('rue',sg2.Cells[10+(A+1),Glob+1]);
                                        Inc(A);

                                        sg2.Cells[10+(A+1),Glob+1]:=Conn_Klap.AQuery_Del.FieldByName(Stanok+'Оч').AsString; //Очередность
                                        Och_D:=StrToFloat(sg2.Cells[10+(A+1),Glob+1]);
                                        Inc(A);

                                        Dlit:= '0.1';//Conn_Klap.AQuery_Del.FieldByName(Stanok+'Дл').AsString;//Длительность
                                        sg2.Cells[10+(A+1),Glob+1]:='0.1';//Conn_Klap.AQuery_Del.FieldByName(Stanok+'Дл').AsString;//Длительность
                                        Dlit_D:=StrToFloat( sg2.Cells[10+(A+1),Glob+1]);//Длительность
                                        Inc(A);

                                        if (Res_Stan<>0) AND ((Och_D=0) OR (Dlit_D=0))  Then //СТАНОК TRUE
                                                                                             //И Очер=0 или Длит =0 Косяк
                                        begin
                                                Memo12.Add('СТАНОК TRUE И Очер=0 или Длит =0 Заказ '+sg2.Cells[1,Glob+1]+' * '+sg2.Cells[2,Glob+1]+' * Деталь '+sg2.Cells[4,Glob+1]);
                                                Memo13.Add(sg2.Cells[4,Glob+1]);
                                        end;

                                        sg2.Cells[10+(A+1),Glob+1]:=Conn_Klap.AQuery_Del.FieldByName('Н\ч'+Stanok).AsString;//Н\ч
                                        Inc(A);
                                        //sg2.Cells[9+(A+1),Glob+1]:=Conn_Klap.AQuery3.FieldByName(Stanok+'Дата').AsString; //Дата
                                        //Inc(A);
                                End;
                        end;
                        Inc(Glob);
                   end;
                   Conn_Klap.AQuery_Del.next;
                end;
        end
        Else
        Begin
           if (Polzov=4) then //ВРАН в спецификации ID Деталей 3 Запроса
           begin
                //В Tab2 для ВРАНСпециф будет IdВРАН
               if not Conn_Ceh.mkQuerySelect2(
                'SELECT ВРАНСборЕд.Id,Деталь ,Единица, ВРАНСборЕд.kol '+
                'FROM ВРАНСпециф INNER JOIN ВРАНСборЕд ON ВРАНСпециф.IdSb = ВРАНСборЕд.Id '+
                'WHERE (ВРАНСпециф."IdВРАН"='+#39+Tab2+#39+') AND (ВРАНСпециф.IdSb<>'+#39+'0'+#39+') UNION '+

                'SELECT we.Id , Деталь ,  Единица, we.kol '+
                'FROM ВРАНСпециф INNER JOIN '+
	              '(SELECT ВРАНДет.Id,Деталь , Обозначение Единица, ВРАНДет.kol '+
                'from ВРАНДет) '+
	              'we  ON ВРАНСпециф.IdDet = we.Id '+
                'WHERE (ВРАНСпециф.IdВРАН='+#39+Tab2+#39+') AND (ВРАНСпециф.IdSb='+#39+'0'+#39+')'
                 , ['ВРАНСпециф']) then  //Специф
                exit;
                KOL_RECORD_Tab2:=Conn_Ceh.AQuery_Del.RecordCount;
           end;
           if (Polzov<4) then
           begin
                if not Conn_Klap.mkQuerySelect2(
                'Select * from %s Where (ВидЭлемента='+#39+'Детали'+#39+
                ') AND ([IdГП]='+#39+IDGP+#39+') AND (([IdКО]='+#39+IDKO+#39+') OR ([IdКО]='+#39+''+#39+'))'
                 , [Tab2]) then  //Специф
                exit;
                KOL_RECORD_Tab2:=Conn_Klap.AQuery_Del.RecordCount;
           end;
           if (Polzov=5) or (Polzov=6) then
           begin         //AND ([Комплект]<>'+#39+'99'+#39+') Сам ставлю 99 на ОСА в составе Завесы
                         //AND ([Комплект]<>'+#39+'55'+#39+') Сам ставлю 55 на ЭКО в составе Завесы
                if not Conn_Ceh.mkQuerySelect2(
                'Select * from %s Where (ВидЭлемента='+#39+'Детали'+#39+
                ') AND ([IdГП]='+#39+IDGP+#39+') AND ([Комплект]<>'+#39+'99'+#39+
                ') AND ([Комплект]<>'+#39+'77'+#39+') AND ([Комплект]<>'+#39+'88'+#39+') AND ([Комплект]<>'+#39+'55'+#39+') AND (([IdКО]='+#39+IDKO+
                #39+') or ([IdКО]='+#39+#39+'))'
                 , [Tab2]) then  //Специф
                exit;
                KOL_RECORD_Tab2:=Conn_Ceh.AQuery_Del.RecordCount;
           end;
            if (Polzov=55) then
           begin
                if not Conn_Ceh.mkQuerySelect2(
                'Select * from %s Where (ВидЭлемента='+#39+'Детали'+#39+
                ') AND ([IdГП]='+#39+IDGP+#39+')  AND (([IdКО]='+#39+IDKO+
                #39+') or ([IdКО]='+#39+#39+'))'
                 , [Tab2]) then  //Специф
                exit;
                KOL_RECORD_Tab2:=Conn_Ceh.AQuery_Del.RecordCount;
           end;
           For X:=0 to KOL_RECORD_Tab2-1 Do
           Begin
             if (Polzov<4) then
             Begin
                Obozn:=Conn_Klap.AQuery_Del.FieldByName('Обозначение').AsString;
                Elem:=Conn_Klap.AQuery_Del.FieldByName('Элемент').AsString;
                Kol_Det:=Conn_Klap.AQuery_Del.FieldByName('Количество').AsString; // AND (Технолог=1)
                                                //Dl,Sh,DL_R,Sh_R,Mat,Diam,Mass,Kol_G,Sh1,Sh2,Ie
                Dl :=Conn_Klap.AQuery_Del.FieldByName('Длина').AsString;
                Sh :=Conn_Klap.AQuery_Del.FieldByName('Ширина').AsString;
                DL_R  :=Conn_Klap.AQuery_Del.FieldByName('ДлинаРазв').AsString;
                Sh_R  :=Conn_Klap.AQuery_Del.FieldByName('ШиринаРазв').AsString;

                Diam  :=Conn_Klap.AQuery_Del.FieldByName('Диаметр').AsString;
                Mass  :=Conn_Klap.AQuery_Del.FieldByName('Масса').AsString;
                Kol_G :=Conn_Klap.AQuery_Del.FieldByName('КолГибов').AsString;
                Sh1   :=Conn_Klap.AQuery_Del.FieldByName('ШиринаПолки1').AsString;
                Sh2   :=Conn_Klap.AQuery_Del.FieldByName('ШиринаПолки2').AsString;
                Ie    :=Conn_Klap.AQuery_Del.FieldByName('ЕИ').AsString;
                Mat   :=Conn_Klap.AQuery_Del.FieldByName('Материал').AsString;

              end;
             if (Polzov=4) then   //VRAN
             Begin
                Obozn:=Conn_Ceh.AQuery_Del.FieldByName('Единица').AsString;
                Elem:=Conn_Ceh.AQuery_Del.FieldByName('Деталь').AsString;
                Kol_Det:=Conn_Ceh.AQuery_Del.FieldByName('Kol').AsString; // AND (Технолог=1)
              end;

              if (Polzov=5) or (Polzov=6) or (Polzov=55) then  //В/А +ЭКО
             Begin
                Obozn:=Conn_Ceh.AQuery_Del.FieldByName('Обозначение').AsString;
                Elem:=Conn_Ceh.AQuery_Del.FieldByName('Элемент').AsString;
                Kol_Det:=Conn_Ceh.AQuery_Del.FieldByName('Количество').AsString; // AND (Технолог=1)
                Leg:=Conn_Ceh.AQuery_Del.FieldByName('КомплВед').AsString; //ЭКО
                //Dl,Sh,DL_R,Sh_R,Mat,Diam,Mass,Kol_G,Sh1,Sh2,Ie
                Dl :=Conn_Ceh.AQuery_Del.FieldByName('Длина').AsString;
                Sh :=Conn_Ceh.AQuery_Del.FieldByName('Ширина').AsString;
                DL_R  :=Conn_Ceh.AQuery_Del.FieldByName('ДлинаРазв').AsString;
                Sh_R  :=Conn_Ceh.AQuery_Del.FieldByName('ШиринаРазв').AsString;

                Diam  :=Conn_Ceh.AQuery_Del.FieldByName('Диаметр').AsString;
                Mass  :=Conn_Ceh.AQuery_Del.FieldByName('Масса').AsString;
                Kol_G :=Conn_Ceh.AQuery_Del.FieldByName('КолГибов').AsString;
                Sh1   :=Conn_Ceh.AQuery_Del.FieldByName('ШиринаПолки1').AsString;
                Sh2   :=Conn_Ceh.AQuery_Del.FieldByName('ШиринаПолки2').AsString;
                Ie    :=Conn_Ceh.AQuery_Del.FieldByName('ЕИ').AsString;
                Mat   :=Conn_Ceh.AQuery_Del.FieldByName('Материал').AsString;
              end;

                if not Conn_Klap.mkQuerySelect3(
                'Select * from %s Where (Обозначение='+#39+Obozn+#39+') AND (Примечание=0)  '
                 , ['СпецифОбщая']) then
                exit;
                if  Conn_Klap.AQuery3.RecordCount<>0 then
                Begin
                        SG2.RowCount:=SG2.RowCount+ 1;
                        sg2.Cells[0,Glob+1]:=Kol_Klap;//IntToStr(Glob+1);

                        sg2.Cells[1,Glob+1]:=ZAKAZ;//Conn_Klap.AQuery.FieldByName('Заказ').AsString;
                        sg2.Cells[2,Glob+1]:=Izdel;//Conn_Klap.AQuery.FieldByName('Изделие').AsString;
                        //Добавить поле ТЕХНОЛОГ, если фалш то деталь не обработана
                        Conn_Klap.Check('Begin1', Kol_Klap+'  '+Kol_Det, 0);
                        sg2.Cells[3,Glob+1]:=FloatToStr(StrToInt(Kol_Klap)*StrToFloat(Kol_Det));
                        Conn_Klap.Check('End1', Kol_Klap+'  '+Kol_Det, 0);
                        sg2.Cells[4,Glob+1]:=Conn_Klap.AQuery3.FieldByName('Обозначение').AsString;
                        sg2.Cells[5,Glob+1]:=Conn_Klap.AQuery3.FieldByName('Элемент').AsString;
                        sg2.Cells[6,Glob+1]:=Nom;
                        sg2.Cells[7,Glob+1]:=BZ;
                        sg2.Cells[8,Glob+1]:=IDGP;
                        sg2.Cells[9,Glob+1]:=Conn_Klap.AQuery3.FieldByName('Технолог').AsString;
                        sg2.Cells[10,Glob+1]:=IDKO;
                        sg2.Cells[11,Glob+1]:=Leg;
                        sg2.Cells[12,Glob+1]:=Dl;
                        sg2.Cells[13,Glob+1]:=Sh;
                        sg2.Cells[14,Glob+1]:=DL_R;
                        sg2.Cells[15,Glob+1]:=Sh_R;
                        sg2.Cells[16,Glob+1]:=Diam;
                        sg2.Cells[17,Glob+1]:=Mass;
                        sg2.Cells[18,Glob+1]:=Kol_G;
                        sg2.Cells[19,Glob+1]:=Sh1;
                        sg2.Cells[20,Glob+1]:=Sh2;
                        sg2.Cells[21,Glob+1]:=Ie;
                        sg2.Cells[22,Glob+1]:=Mat;
                        Res_Teh:=Pos('rue',sg2.Cells[9,Glob+1]);
                        if (Res_Teh=0)  Then //ТеХНОЛОГ FALSE
                        begin
                                  Memo12.Add('ТЕХНОЛОГ FALSE Заказ '+sg2.Cells[1,Glob+1]+' * '+sg2.Cells[2,Glob+1]+' * Деталь '+sg2.Cells[4,Glob+1]);
                                  Memo13.Add(sg2.Cells[4,Glob+1]);
                        end;
                        A:=13;
                        for y := 0 to High(Column) do
                        begin
                                if Column[y,0]<>'' then
                                Begin
                                        Stanok:=Column[y,0]; //Значение поля

                                        sg2.Cells[9+(A+1),Glob+1]:=Conn_Klap.AQuery3.FieldByName(Stanok).AsString;
                                        Res_Stan:=Pos('rue',sg2.Cells[9+(A+1),Glob+1]);
                                        Inc(A);

                                        sg2.Cells[9+(A+1),Glob+1]:=Conn_Klap.AQuery3.FieldByName(Stanok+'Оч').AsString; //Очередность
                                        Och_D:=StrToFloat(sg2.Cells[9+(A+1),Glob+1]);
                                        Inc(A);

                                        sg2.Cells[9+(A+1),Glob+1]:='0.1';//Conn_Klap.AQuery3.FieldByName(Stanok+'Дл').AsString;//Длительность
                                        Dlit_D:=StrToFloat( sg2.Cells[9+(A+1),Glob+1]);//Длительность
                                        Inc(A);

                                        if (Res_Stan<>0) AND ((Och_D=0) OR (Dlit_D=0))  Then //СТАНОК TRUE
                                                                                             //И Очер=0 или Длит =0 Косяк
                                        begin
                                                Memo12.Add('СТАНОК TRUE И Очер=0 или Длит =0 Заказ '+sg2.Cells[1,Glob+1]+' * '+sg2.Cells[2,Glob+1]+' * Деталь '+sg2.Cells[4,Glob+1]);
                                                Memo13.Add(sg2.Cells[4,Glob+1]);
                                        end;

                                        sg2.Cells[9+(A+1),Glob+1]:=Conn_Klap.AQuery3.FieldByName('Н\ч'+Stanok).AsString;//Н\ч
                                        Inc(A);
                                        //sg2.Cells[9+(A+1),Glob+1]:=Conn_Klap.AQuery3.FieldByName(Stanok+'Дата').AsString; //Дата
                                        //Inc(A);
                                End;
                        end;
                        Inc(Glob);
                end;
                if Polzov<4 then
                  Conn_Klap.AQuery_Del.next;
                if (Polzov=4) or (Polzov=5) or (Polzov=55) or (Polzov=6) then
                  Conn_Ceh.AQuery_Del.next;
           end;
        end;
        if Polzov<4 then
          Conn_Klap.AQuery.next;
        if  (Polzov=4) or (Polzov=5) or (Polzov=55) or (Polzov=6) then
          Conn_Ceh.AQuery.next;  //gulyakova-ei
      end;
      Memo12_Count:=Memo12.Count;
      Memo12_Text:= Memo12.Text;
      God := FormatDateTime('yyyy', Now);
      Mes := FormatDateTime('mmmm', Now);
      if ( (Polzov=4) or (Polzov=5) or (Polzov=55) or (Polzov=6) ) AND (Memo12_Count<>0) then  //КОСЯК ВРАН
      begin
        Dir :=Form1.Put_KTO+'\CKlapana\VRAN\' + God;
        CreateDir(Dir);

        Dir := Form1.Put_KTO+'\CKlapana\VRAN\' + God + '\' + Mes+'\';
        CreateDir(Dir);
      end;
      Dir :=Form1.Put_KTO+'\CKlapana\Накладные(пожар)\' + God;
        CreateDir(Dir);

      Dir := Form1.Put_KTO+'\CKlapana\Накладные(пожар)\' + God + '\' + Mes+'\';
        CreateDir(Dir);

      Dir :=Form1.Put_KTO+'\CKlapana\' + God;
        CreateDir(Dir);
      Dir := Form1.Put_KTO+'\CKlapana\' + God + '\' + Mes+'\';
        CreateDir(Dir);
      if (Polzov=0) AND (Memo12_Count=0)  Then //Технолог
      //Все детали обработаны ,выходим
      Begin
        Flag_Error:=0;
        Memo12.Destroy;
        Exit;
      end;
        if (Polzov=0) AND (Memo12_Count<>0) then  //КОСЯК   Технолог
        Begin
                MessageDlg('Не все детали обработаны! Список в папке с накладной.', mtError,
                        [mbOk], 0);
                Res_Tab:=Pos('Воз',Tab2);
                if Res_Tab=0 Then           // PAnsiChar(P
                Begin
                        Memo12.SaveToFile(Form1.Put_KTO+'\CKlapana\Накладные(пожар)\'+God+'\'+Mes+'\'+Nom+'.txt');
                        Nam :=  Form1.Put_KTO+'\CKlapana\Накладные(пожар)\'+God+'\'+Mes+'\'+Nom+'.txt';
                End;
                if Res_Tab<>0 Then//Воздушные
                Begin
                        Memo12.SaveToFile(Form1.Put_KTO+'\CKlapana\'+God+'\'+Mes+'\'+Nom+'.txt');
                        Nam := Form1.Put_KTO+'\CKlapana\'+God+'\'+Mes+'\'+Nom+'.txt';
                End;
                ShellExecute(0, nil, PChar(Nam), nil, nil, SW_SHOWNORMAL);
                File_Name:= Nam ;
                Flag_Error:=2;
                Memo12.Destroy;
                if fileName<>'' then
                begin
                    if not Conn_Klap.mkQueryDelete( 'DELETE FROM %s WHERE (Номер='+Nom+') AND (Планирование='+#39+fileName+#39+')'
                                , ['Заготовка']) then
                    exit;
                end
                Else
                begin
                    if not Conn_Klap.mkQueryDelete( 'DELETE FROM %s WHERE (Номер='+Nom+') AND (Планирование='+#39+ConvertDat1(Dat_Glob)+#39+')'
                                , ['Заготовка']) then
                    exit;
                end;
                Result := False;
                Exit;
        End;
        if ((Polzov<4) AND (Polzov<>0)) AND (Memo12_Count<>0) then  //КОСЯК
        Begin
                MessageDlg('Не все детали обработаны! Список в папке с накладной.', mtError,
                        [mbOk], 0);
                Res_Tab:=Pos('Воз',Tab2);
                if Res_Tab=0 Then
                Begin
                        Memo12.SaveToFile(Form1.Put_KTO+'\CKlapana\Накладные(пожар)\'+God+'\'+Mes+'\'+Nom+'.txt');
                        Nam :=  Form1.Put_KTO+'\CKlapana\Накладные(пожар)\'+God+'\'+Mes+'\'+Nom+'.txt';
                End;
                if Res_Tab<>0 Then//Воздушные
                Begin
                        Memo12.SaveToFile(Form1.Put_KTO+'\CKlapana\'+God+'\'+Mes+'\'+Nom+'.txt');
                        Nam :=  Form1.Put_KTO+'\CKlapana\'+God+'\'+Mes+'\'+Nom+'.txt';
                End;
                ShellExecute(0, nil, PChar(Nam), nil, nil, SW_SHOWNORMAL);
                File_Name:= Nam ;
                Flag_Error:=2;
                Memo12.Destroy;
                if fileName<>'' then
                begin
                    if not Conn_Klap.mkQueryDelete( 'DELETE FROM %s WHERE (Номер='+Nom+') AND (Планирование='+#39+fileName+#39+')'
                                , ['Заготовка']) then
                    exit;
                end
                Else
                begin
                    if not Conn_Klap.mkQueryDelete( 'DELETE FROM %s WHERE (Номер='+Nom+') AND (Планирование='+#39+ConvertDat1(Dat_Glob)+#39+')'
                                , ['Заготовка']) then
                    exit;
                end;
                Result := False;
                Exit;
        End;

        if ( (Polzov=4) or (Polzov=5) or (Polzov=55) or (Polzov=6) ) AND (Memo12_Count<>0) then  //КОСЯК ВРАН
        Begin
                MessageDlg('Не все детали обработаны! Список в папке с накладной.', mtError,
                        [mbOk], 0);

                Memo12.SaveToFile(Form1.Put_KTO+'\CKlapana\VRAN\'+God+'\'+Mes+'\'+Nom+'.txt');
                Nam :=  Form1.Put_KTO+'\CKlapana\VRAN\'+God+'\'+Mes+'\'+Nom+'.txt';

                ShellExecute(0, nil, PChar(Nam), nil, nil, SW_SHOWNORMAL);
                File_Name:= Nam ;
                Flag_Error:=2;
                Memo12.Destroy;
                if fileName<>'' then
                begin
                    if not Conn_Klap.mkQueryDelete( 'DELETE FROM %s WHERE (Номер='+Nom+') AND (Планирование='+#39+fileName+#39+')'
                                , ['Заготовка']) then
                    exit;
                end
                Else
                begin
                    if not Conn_Klap.mkQueryDelete( 'DELETE FROM %s WHERE (Номер='+Nom+') AND (Планирование='+#39+ConvertDat1(Dat_Glob)+#39+')'
                                , ['Заготовка']) then
                    exit;
                end;
                Result := False;
                Exit;
        End;
      except

        exit;
    end;
    Memo12.Destroy;
    Memo13.Destroy;
    if fileName<>'' then
    begin
       if not Conn_Klap.mkQueryDelete( 'DELETE FROM %s WHERE (Номер='+Nom+') AND (Планирование='+#39+fileName+#39+')'
                                , ['Заготовка']) then
    exit;
    end
    Else
    begin
       if not Conn_Klap.mkQueryDelete( 'DELETE FROM %s WHERE (Номер='+Nom+') AND (Планирование='+#39+ConvertDat1(Dat_Glob)+#39+')'
                                , ['Заготовка']) then
    exit;
    end;

    Result := true;
    btn1Click(fileName) ;
end;
function Osnova_Main.btn1Click(fileName:String):Boolean;
        var Dat_S1,IDGP,IDKO,Obozn,Detal,Str,Elem,Stanok,Str1,Str2,Izdel,Kol_Klap,Kol_Det,Zakaz,Nomer,Max:string;
        I,X,Y,A,Glob,Col,Col_Nom,
        Kol_Klap_I,Kol_Det_I,Res1,res2,Res:Integer;
        B,D:Boolean;                   // ,S1
        DEF_CatalogName,MSSQL_CONN_STR,S,Razmernost,Nom_Kol,Znach,Str_Stanok,Err,BZ:string;
        TempTableQ:TADOQuery;
        ADOConnection4:TADOConnection;
        Dat_S,Nedely,Kol,Dir,SQLServer,Dat_Z,Nowa_S,Dlit,Dlit_Ob,NCH,fileName1,Kod,Dl,Sh,DL_R,Sh_R,Mat,Diam,Mass,Kol_G,Sh1,Sh2,Ie:string;
        Dat, Dat1,Dat_Zag,nowa:TDate;
        Ocheredn,Res_True:Integer;
        Dliteln,
        NC_D:Double;
        //StrToCol:array[0..100,0..5] of string;//[очередь, длительность, Общ длительность, Дата] ДЛЯ ОДНОЙ ДЕТАЛИ
        Kol_Stan:Integer;
        IPs : TStringList;
        Res_Trum,II:Integer;
        Dl_Trum,DD:Double;
        Memo11:TStringList;
        leg:String;
   begin
       SyStem.SysUtils.FormatSettings.DecimalSeparator :=('.');
       Memo11:=TStringList.Create;
       Memo11.Clear;
       Glob:=0;
       A:=0;
      //==================================
      S:='';
      Razmernost:='';
      A:=0;
      fileName1:=fileName;
      //Dat:=dtp1.Date;
      //Dat_S:=FormatDateTime('dd.mm.YYYY',Dat);
      //Dat_S1:=FormatDateTime('dd.mm.YYYY',Dat-1);
      for y := 0 to High(Column) do
      begin
        if Column[y,0]<>'' then
        Begin
                Str_Stanok:=Str_Stanok+Column[y,0]+','+Column[y,0]+'Оч,'+Column[y,0]+'Дл,[Н\ч'+Column[y,0]+'],'+Column[y,0]+'Дата,'+Column[y,0]+'ДлОбщ,';
                Nom_Kol:=IntToStr(StrToInt(Column[y,2])-2);
                Razmernost:=Razmernost+Column[y,0]+' BIT,'+Column[y,0]+'Оч nvarchar(50),'+Column[y,0]+'Дл nvarchar(50),[Н\ч'+Column[y,0]+'] nvarchar(50),'+Column[y,0]+'Дата DateTime,'+Column[y,0]+'ДлОбщ nvarchar(50),';  //DateTimeГибка bit,ГибкаОЧ nvarchar(50)
        End;
      end;
      //sg2.Group(4);
      //sg2.Group(5);
      //SG2.GroupSum(4);
     // SG2.ClearAll;
      S:= SG2.Cells[7,0];
      //++++++++++++++++++++++++++++++++++++
      Dir:=ExtractFileDir(ParamStr(0));
      IPs := TStringList.Create;
      IPs.Clear;
      IPs.LoadFromFile(Dir+'\ConnKlap.ini');
        //Memo1.Lines.LoadFromFile(Dir+'\ConnKlap.ini');

        DEF_CatalogName :=IPs.Strings[0];
        SQLServer:=IPs.Strings[1];
        MSSQL_CONN_STR :=IPs.Strings[4];
       { 'Provider=SQLOLEDB.1;Packet Size = 4096;Password=111;' +
        'Persist Security Info=True;User Id=TestUser;' +
        'Initial Catalog='+ DEF_CatalogName +';Data Source=DINAMO\'+SQLServer; }
        ADOConnection4:=TADOConnection.Create(Nil);
        ADOConnection4.LoginPrompt:=False;
        ADOConnection4.Connected:=False;
        ADOConnection4.ConnectionString := MSSQL_CONN_STR;
        ADOConnection4.Connected:=True;
         //Запятая лишняя
        Y:=Length(Razmernost);
        Delete(Razmernost,Y,1);
        //Создание временной таблицы
        TempTableQ:=TADOQuery.Create(nil);
        TempTableQ.Connection:=ADOConnection4;
        //Добавить поле ТЕХНОЛОГ, если фалш то деталь не обработана
        TempTableQ.SQL.Text:='CREATE TABLE #ClientToDBFL (Заказ nvarchar(100),Изделие nvarchar(100),Номер nvarchar(100),'+
        'Обозначение nvarchar(100) ,Элемент nvarchar(100),Количество nvarchar(100),БЗ nvarchar(100),'+
        'IdГП nvarchar(100),IdКО nvarchar(100),КолКлап nvarchar(100),Leg nvarchar(100),'+
        'Длина nvarchar(100),Ширина nvarchar(100),ДлинаРазв nvarchar(100),ШиринаРазв nvarchar(100),'+
        'Диаметр nvarchar(100),Масса nvarchar(100),КолГибов nvarchar(100),ШиринаПолки1 nvarchar(100),'+
        'ШиринаПолки2 nvarchar(100),ЕИ nvarchar(100),Материал nvarchar(100),'+Razmernost+')';
                //Dl,Sh,DL_R,Sh_R,Mat,Diam,Mass,Kol_G,Sh1,Sh2,Ie
        {sg2.Cells[12,0]:='Длина';
        sg2.Cells[13,0]:='Ширина';
        sg2.Cells[14,0]:='ДлинаРазв';
        sg2.Cells[15,0]:='ШиринаРазв';

        sg2.Cells[16,0]:='Диаметр';
        sg2.Cells[17,0]:='Масса';
        sg2.Cells[18,0]:='КолГибов';
        sg2.Cells[19,0]:='ШиринаПолки1';
        sg2.Cells[20,0]:='ШиринаПолки2';
        sg2.Cells[21,0]:='ЕИ';
        sg2.Cells[22,0]:='Материал'; }
        try
                TempTableQ.ExecSQL;
        except
                Err:='TempTableQ.ExecSQL;';
        end;
        //------------------------------------------------
      //Запятая лишняя
      Y:=Length(Str_Stanok);
      Delete(Str_Stanok,Y,1);

      //-------------------------------------------------
        for Y:=0 to SG2.RowCount-1 do
        begin
         if Y >0 then
         begin                        //  AND (SG2.Cells[2,Y]='')
                if (SG2.Cells[1,Y]<>'') then
                begin
                   //------------------------------------------------
                   S1:='';
                   Dliteln:=0;
                   Ocheredn:=0;
                   Col_Nom:=0;
                   A :=StrToInt(Column[0,2])+16;
                   i :=0;
                   Kol_Stan:=0;
                   fillchar(StrToCol,sizeof(StrToCol),0);//Очистка массмва
                   for X := 0 to High(Column) do  //Строка номеров колонок
                   begin
                        if Column[X,2]<>'' then
                        Begin   //'Пила,ПилаОч,ПилаДл,[Н\чПила],
                                //'False','0','0','0',
                                i :=0;
                                Col_Nom:=StrToInt(Column[X,2])-1;
                               // S1:=S1+#39+SG2.Cells[A,Y]+#39+','; //True  0
                                StrToCol[X, I ]:=SG2.Cells[A,Y];
                                Inc(A);
                                Inc(i);
                                //
                               // S1:=S1+#39+SG2.Cells[A,Y]+#39+','; //Ocheredn 1
                                StrToCol[X,i]:=SG2.Cells[A,Y];
                                Inc(A);
                                Inc(i);
                                //
                               // S1:=S1+#39+SG2.Cells[A,Y]+#39+','; //Dliteln 2
                                StrToCol[X,i]:=SG2.Cells[A,Y];
                                Inc(A);
                                Inc(i);
                                //
                                //S1:=S1+#39+SG2.Cells[A,Y]+#39+','; //Н\ч  3
                                StrToCol[X,i]:=SG2.Cells[A,Y];
                                Inc(A);
                                Inc(i);
                                //
                                //S1:=S1+#39+'0'+#39+',';  // Дата
                                StrToCol[X,i]:='0';
                                Inc(i);
                                //
                                StrToCol[X,i]:=Column[X,0];
                                Inc(i);
                                //
                        End;
                   end;
                   Kol_Klap:=SG2.Cells[0,Y];
                   if Kol_Klap='' Then
                   Kol_klap_I:=0
                   Else
                   Kol_Klap_I:=StrToInt(Kol_Klap);

                   ///()))))))))))))))))))))))(((((((((((((((((((((
                   Btn11Click(fileName1); ///()))))))))))))))))))))))(((((((((((((((((((((
                   ///()))))))))))))))))))))))(((((((((((((((((((((
                   //Запятая лишняя
                   i:=Length(S1);
                   Delete(S1,i,1);

                   Zakaz:=SG2.Cells[1,Y];
                   Obozn:=SG2.Cells[4,y];
                   Izdel:=SG2.Cells[2,y];
                   res:=Pos('ЛА-2510-1-3',Obozn);
                   if Res<>0 then
                      Kol_Det:='1';

                                                   {if not Conn_Klap.mkQuerySelect6('Select *  From %s Where (Обозначение='+#39+'%s'+#39+
                                ') ', ['Заготовка',Obozn]) then //Специф
                                exit;   }
                                //ВГ 410.00.01.002-400 меняем на ВГ 410.00.01.001-400  Гапова 21.01.2019
                                {Res1:=Pos('КЭД-', Izdel);
                                Res2:=Pos('-2*ф-', Izdel);
                                res:=Pos('ВГ 410.00.01.002-',Obozn);
                                if (Res<>0) and (Res1<>0) AND (res2<>0) then
                                begin
                                  //StringReplace(Dlin_Raz_S,'.',',',[rfReplaceAll]);
                                  Res:=Pos('.002-',Obozn);
                                  if Res<>0 then
                                  begin
                                    Delete(Obozn,Res,5);
                                    S:=  '.001-';
                                    Insert(S,Obozn,Res);
                                  end;
                                end;
                                // ВГ 463.00.00.006 еняем на ВГ 463.00.00.006-01   Гапова 21.01.2019
                                res:=Pos('ВГ 463.00.00.006',Obozn);
                                if  (Res<>0) and (Res1<>0) AND (res2<>0) then
                                begin
                                     Obozn:='ВГ 463.00.00.006-01';
                                end;                  //Dl,Sh,DL_R,Sh_R,Mat,Diam,Mass,Kol_G,Sh1,Sh2,Ie
}
                   Kol_Det:=SG2.Cells[3,y];
                   Elem:=SG2.Cells[5,y];
                   Nomer:=SG2.Cells[6,y];
                   BZ:=SG2.Cells[7,y];
                   IDGP:=SG2.Cells[8,y];
                   IDKO:=SG2.Cells[10,y];
                   Leg:=(SG2.Cells[11,y]);

                   Dl :=SG2.Cells[12,y];
                   Sh :=SG2.Cells[13,y];
                   DL_R :=SG2.Cells[14,y];
                   Sh_R :=SG2.Cells[15,y];
                   Diam :=SG2.Cells[16,y];
                   Mass :=SG2.Cells[17,y];
                   Kol_G:=SG2.Cells[18,y];
                   Sh1  :=SG2.Cells[19,y];
                   Sh2  :=SG2.Cells[20,y];
                   Ie   :=SG2.Cells[21,y];
                   Mat  :=SG2.Cells[22,y];
                   //Кол-во деталей уже умножено на кол клапанов!!!!

                   TempTableQ.SQL.Text:='INSERT INTO #ClientToDBFL '+
                        '(Заказ,Изделие,Обозначение,Элемент,'+
                        'Количество,Номер,БЗ,IdГП, IdКО,КолКлап,Leg,'+
                        'Длина,Ширина,ДлинаРазв,ШиринаРазв,Диаметр,Масса,КолГибов,ШиринаПолки1,ШиринаПолки2,ЕИ,Материал,'

                        +Str_Stanok+') Values ('
                        +#39+Zakaz+#39+
                        ','+#39+Izdel+#39+
                        ','+#39+Obozn+#39+
                        ','+#39+Elem+#39+
                        ','+#39+Kol_Det+#39+
                        ','+Nomer+
                        ','+#39+BZ+#39+
                        ','+#39+IDGP+#39+ ','+#39+IDKO+#39+
                        ','+#39+Kol_Klap+#39+
                        ','+#39+(Leg)+#39+

                        ','+#39+Dl+#39+
                        ','+#39+Sh+#39+
                        ','+#39+DL_R+#39+
                        ','+#39+Sh_R+#39+
                        ','+#39+Diam+#39+
                        ','+#39+Mass+#39+
                        ','+#39+Kol_G+#39+
                        ','+#39+Sh1+#39+
                        ','+#39+Sh2+#39+
                        ','+#39+Ie+#39+
                        ','+#39+Mat+#39+
                            ','+S1+
                        ')';

                        try
                        TempTableQ.ExecSQL;
                        except
                        Err:='1222';
                        end;
                end;

         end;
        end;
        //Группировка по TRUE
        S1:='';
        A:=1;
        Glob:=1;
        sg3 := TStringGrid.Create(Nil );
        Clear_StringGrid(sg3);
    SG3.ColCount:=29;
    sg3.Cells[0,0]:='КолКлап';
    sg3.Cells[1,0]:='Заказ';
    sg3.Cells[2,0]:='Изделие';
    sg3.Cells[3,0]:='Номер';

    sg3.Cells[4,0]:='Обозначение';
    sg3.Cells[5,0]:='Элемент';
    sg3.Cells[6,0]:='Количество';
    sg3.Cells[7,0]:='Очередность'; //Очередность
    sg3.Cells[8,0]:='Длительность';//Длительность
    sg3.Cells[9,0]:='Н\ч';//Н\ч
    sg3.Cells[10,0]:='БЗ';
    sg3.Cells[11,0]:='IdГП';
    sg3.Cells[12,0]:='Станок';
    sg3.Cells[13,0]:='Планирование';
    sg3.Cells[14,0]:='Заготовка';
    sg3.Cells[15,0]:='Длит Общая';
    sg3.Cells[16,0]:='IdКО';
    sg3.Cells[17,0]:='Leg';
            //Dl,Sh,DL_R,Sh_R,Mat,Diam,Mass,Kol_G,Sh1,Sh2,Ie
        sg3.Cells[18,0]:='Длина';
        sg3.Cells[19,0]:='Ширина';
        sg3.Cells[20,0]:='ДлинаРазв';
        sg3.Cells[21,0]:='ШиринаРазв';
        sg3.Cells[22,0]:='Диаметр';
        sg3.Cells[23,0]:='Масса';
        sg3.Cells[24,0]:='КолГибов';
        sg3.Cells[25,0]:='ШиринаПолки1';
        sg3.Cells[26,0]:='ШиринаПолки2';
        sg3.Cells[27,0]:='ЕИ';
        sg3.Cells[28,0]:='Материал';
    //Memo12:=TStringList.Create;
    //Memo12.Clear;
        for Y:=0 to High(Column) do  //Строка номеров колонок
        begin
                if Column[y,0]<>'' then
                Begin
                        TempTableQ.SQL.Text:='SELECT *  FROM #ClientToDBFL'+
                        ' WHERE '+Column[y,0]+'='+#39+'True'+#39 ;
                        TempTableQ.Open;
                        //+++++++++++++++++++++++++++++++++++++++
                        TempTableQ.First;
                        if TempTableQ.RecordCount<>0 Then
                        Begin
                                sg3.RowCount:= sg3.RowCount+1 ;
                                sg3.Cells[0,A]:=Column[y,0];
                                Stanok:=Column[y,0]; //Значение поля
                                A:=A+1;

                          for I := 0 to TempTableQ.RecordCount - 1 do
                          begin
                                sg3.RowCount:= sg3.RowCount+1 ;
                                Str:=sg3.Cells[0,A];
                                sg3.Cells[0,A]:=TempTableQ.FieldByName('КолКлап').AsString;//IntToStr(Glob);
                                sg3.Cells[1,A]:=TempTableQ.FieldByName('Заказ').AsString;
                                sg3.Cells[2,A]:=TempTableQ.FieldByName('Изделие').AsString;
                                sg3.Cells[3,A]:=TempTableQ.FieldByName('Номер').AsString;

                                sg3.Cells[4,A]:=TempTableQ.FieldByName('Обозначение').AsString;
                                sg3.Cells[5,A]:=TempTableQ.FieldByName('Элемент').AsString;
                                Kol_Det:=TempTableQ.FieldByName('Количество').AsString;
                                II:=Round(StrToFloat(Kol_Det));
                                Kol_Det:=IntToStr(II);
                                sg3.Cells[6,A]:= Kol_Det;
                                if Kol_Det='' Then
                                Kol_Det_I:=0
                                Else
                                Kol_Det_I:=StrToInt(Kol_Det);

                                sg3.Cells[7,A]:=TempTableQ.FieldByName(Stanok+'Оч').AsString; //Очередность
                                Ocheredn:=StrToInt( sg3.Cells[7,A]);
                                sg3.Cells[8,A]:=TempTableQ.FieldByName(Stanok+'Дл').AsString;//Длительность
                                Dliteln:=StrToFloat( StringReplace(sg3.Cells[8,A],',','.',[rfReplaceAll]));
                                sg3.Cells[9,A]:=TempTableQ.FieldByName('Н\ч'+Stanok).AsString;//Н\ч
                                NCH:=StringReplace(sg3.Cells[9,A],',','.',[rfReplaceAll]);
                                if NCh='' Then
                                NC_D:=0
                                Else
                                NC_D:=StrToFloat(NCH);
                                sg3.Cells[10,A]:=TempTableQ.FieldByName('БЗ').AsString;
                                sg3.Cells[11,A]:=TempTableQ.FieldByName('IdГП').AsString;
                                sg3.Cells[12,A]:=Column[y,0];
                                sg3.Cells[13,A]:=Dat_Glob;

                                Dat_Z :=TempTableQ.FieldByName(Stanok+'Дата').AsString;
                                Dat_Zag:= StrToDate(Dat_Z);
                                Nowa_S:=FormatDateTime('dd.mm.YYYY',Now);
                                Nowa:=StrToDate(Nowa_s);
                                //28.05.2018 Отменили проверку
                               { if Dat_Zag<=Nowa then //Дата начала заготовки МЕНЬШЕ сегодняшней
                                begin
                                        MessageDlg(Nom_G+'  Дата начала заготовки МЕНЬШЕ сегодняшней!' + #10#13 +
                                        'Необходимо сместить Дату Планирования на ' + FloatToStr(Nowa-Dat_Zag)+' Дней!' ,mtInformation, [mbOk], 0);
                                        Flag_Error:=1;
                                        Exit;
                                end; }
                                sg3.Cells[14,A]:=TempTableQ.FieldByName(Stanok+'Дата').AsString;
                                sg3.Cells[15,A]:=TempTableQ.FieldByName(Stanok+'ДлОбщ').AsString;
                                sg3.Cells[16,A]:=TempTableQ.FieldByName('IdКО').AsString;
                                sg3.Cells[17,A]:=TempTableQ.FieldByName('Leg').AsString;
                                                //Dl,Sh,DL_R,Sh_R,Mat,Diam,Mass,Kol_G,Sh1,Sh2,Ie
                                sg3.Cells[18,A] :=TempTableQ.FieldByName('Длина').AsString;
                                sg3.Cells[19,A] :=TempTableQ.FieldByName('Ширина').AsString;
                                sg3.Cells[20,A]  :=TempTableQ.FieldByName('ДлинаРазв').AsString;
                                sg3.Cells[21,A]  :=TempTableQ.FieldByName('ШиринаРазв').AsString;

                                sg3.Cells[22,A]  :=TempTableQ.FieldByName('Диаметр').AsString;
                                sg3.Cells[23,A]  :=TempTableQ.FieldByName('Масса').AsString;
                                sg3.Cells[24,A] :=TempTableQ.FieldByName('КолГибов').AsString;
                                sg3.Cells[25,A]   :=TempTableQ.FieldByName('ШиринаПолки1').AsString;
                                sg3.Cells[26,A]   :=TempTableQ.FieldByName('ШиринаПолки2').AsString;
                                sg3.Cells[27,A]    :=TempTableQ.FieldByName('ЕИ').AsString;
                                sg3.Cells[28,A]   :=TempTableQ.FieldByName('Материал').AsString;
                                Dlit:=StringReplace(sg3.Cells[8,A],',','.',[rfReplaceAll]);
                                Dlit_Ob:=StringReplace(sg3.Cells[15,A],',','.',[rfReplaceAll]);

                                //++++++++++++++++++++++++++++++++++++++++++++++++++
                                if sg3.Cells[14,A]='' then
                                        Dat_S1:='NULL'
                                else
                                        DaT_S1:=#39+ConvertDat1(SG3.Cells[14,A])+#39;
                                                        //   AND (Технолог=1)
                                if not Conn_Klap.mkQuerySelect3(
                                'Select * from %s Where (Обозначение='+#39+Obozn+#39+')'
                                , ['СпецифОбщая']) then
                                exit;
                                //Проверка на дубль
                                {if not Conn_Klap.mkQuerySelect4(
                                'Select * from %s Where (Заказ='+#39+sg3.Cells[1,A]+#39+') AND '+
                                '(Изделие='+#39+sg3.Cells[2,A]+#39+') AND (Номер='+#39+sg3.Cells[3,A]+#39+') AND '+
                                '(Обозначение='+#39+sg3.Cells[4,A]+#39+') AND (Элемент='+#39+sg3.Cells[5,A]+#39+') AND (IdГП='+#39+sg3.Cells[11,A]+#39+')'+
                                ' AND (Станок='+#39+sg3.Cells[12,A]+#39+')'
                                , ['Заготовка']) then
                                exit;
                                if  (Conn_Klap.AQuery4.RecordCount<>0) Then
                                Begin
                                  if not Conn_Klap.mkQueryDelete( 'DELETE FROM %s WHERE (Заказ='+#39+sg3.Cells[1,A]+#39+') AND '+
                                '(Изделие='+#39+sg3.Cells[2,A]+#39+') AND (Номер='+#39+sg3.Cells[3,A]+#39+') AND '+
                                '(Обозначение='+#39+sg3.Cells[4,A]+#39+') AND (Элемент='+#39+sg3.Cells[5,A]+#39+') AND (IdГП='+#39+sg3.Cells[11,A]+#39+')'+
                                ' AND (Станок='+#39+sg3.Cells[12,A]+#39+')'
                                , ['Заготовка']) then
                                        exit;
                                end; }
                            if  (Conn_Klap.AQuery3.RecordCount<>0)  then
                            Begin
                                NCH:=StringReplace(FloatToStr(NC_D*Kol_Det_I),',','.',[rfReplaceAll]);

                                if not Conn_Klap.mkQueryInsert(
                                'INSERT INTO Заготовка '+
                                '(Заказ,Изделие,Номер,Обозначение,Элемент,'+
                                'Количество,Очередность,Длительность,НЧ,БЗ,IdГП,Станок,Планирование,Заготовка,ДлительностьОбщ,КолКлап,IdКО,'+
                                'Цвет,Легион,КомплВед,Длина,Ширина,ДлинаРазв,ШиринаРазв,Диаметр,Масса,КолГибов,ШиринаПолки1,ШиринаПолки2,ЕИ,Материал) Values ('
                                +#39+sg3.Cells[1,A]+#39+
                                ','+#39+sg3.Cells[2,A]+#39+
                                ','+#39+sg3.Cells[3,A]+#39+
                                ','+#39+sg3.Cells[4,A]+#39+
                                ','+#39+sg3.Cells[5,A]+#39+
                                ','+#39+sg3.Cells[6,A]+#39+
                                ','+#39+sg3.Cells[7,A]+#39+
                                ','+#39+Dlit+#39+
                                ','+#39+NCH+#39+
                                ','+#39+sg3.Cells[10,A]+#39+
                                ','+#39+sg3.Cells[11,A]+#39+
                                ','+#39+sg3.Cells[12,A]+#39+
                                ','+#39+ConvertDat1(sg3.Cells[13,A])+#39+
                                ','+Dat_S1+
                                ','+#39+Dlit_Ob+#39+
                                ','+#39+sg3.Cells[0,A]+#39+
                                ','+#39+sg3.Cells[16,A]+#39+
                                ','+#39+IntToStr(Cvet1)+#39+
                                ','+#39+'0'+#39+
                                ','+#39+sg3.Cells[17,A]+#39+

                                ','+#39+sg3.Cells[18,A]+#39+
                                ','+#39+sg3.Cells[19,A]+#39+
                                ','+#39+sg3.Cells[20,A]+#39+
                                ','+#39+sg3.Cells[21,A]+#39+
                                ','+#39+sg3.Cells[22,A]+#39+
                                ','+#39+sg3.Cells[23,A]+#39+
                                ','+#39+sg3.Cells[24,A]+#39+
                                ','+#39+sg3.Cells[25,A]+#39+
                                ','+#39+sg3.Cells[26,A]+#39+
                                ','+#39+sg3.Cells[27,A]+#39+
                                ','+#39+sg3.Cells[28,A]+#39+
                                ')',[] )then
                                Exit;
                                //+++++++++++++++++++++++++++++++++++++++++++
                                A:=A+1;
                                Glob:=Glob+1;
                            End;
                                TempTableQ.Next;

                          end;
                        end;

                end;
        end;
        TempTableQ.Close;
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        if not Conn_Klap.mkQuerySelect('Select Заготовка From Заготовка Where (Номер='+
        #39+Nomer_Glob+#39+') Group By Заготовка',[]) Then
        Exit;
        For I:=0 To  Conn_Klap.AQuery.RecordCount-1 Do
        Begin
              Dat_S:= ConvertDat1(Conn_Klap.AQuery.FieldByName('Заготовка').AsString);

              Memo11.Add('');
              Memo11.Add(Dat_S);
              if not Conn_Klap.mkQuerySelect3('Select Станок, SUM(НЧ) As SNH From Заготовка Where (Заготовка='+
                 #39+Dat_S+#39+') AND (Станок<>'+#39+'Канбан'+#39+') AND (Станок<>'+#39+'Trumph'+#39+') Group by Станок',[]) Then
                Exit;

              For Y:=0 To  Conn_Klap.AQuery3.RecordCount-1 Do
              Begin
                   Memo11.Add(Conn_Klap.AQuery3.FieldByName('Станок').AsString+'  Н\ч: '+Conn_Klap.AQuery3.FieldByName('SNH').AsString);
                   Inc(A);
                   Conn_Klap.AQuery3.Next;
              end;
              Conn_Klap.AQuery.Next;
        end;
        SNCH:=Memo11.Text;
        //================================================
        TempTableQ.SQL.Text:='DROP TABLE #ClientToDBFL';
        try
                TempTableQ.ExecSQL;
        except
        end;

   End;
//++++++++++++++++++++++++++++++++++++++++++++++++++
Function Osnova_Main.btn11Click(Err2:String):Boolean;
var //StrToCol:array[0..100,0..6] of string;
//[TRUE,очередь, длительность, н/Ч, Дата,Stanok] ДЛЯ ОДНОЙ ДЕТАЛИ
I,X,J,y,A,Res:Integer;
Res_True,min_Col,Col,Col_1,
Col_Grid,Res1:Integer;
Izdel,Nowa_S:string;
Dat_Plan,Dat_Zag,Nowa:TDate;
Dat_P,Dat_Z:string;
Dlit,Dlit_Min,Dlit_Obsh,Dlit_Max,
Min,Och_Min,Och_Max,Dlit_ObshO,Ocher,Ocher_Pred:Double;
//SREZ1: array[0..10,0..100] of string;

SKPU :array[0..10,0..100] of string;

//Sort: TStringList;
Stanok,Stanok_Posl:String;
Kol_Smen,Max,F1,Den,Den_pred,Celoe,Res_Trum:Integer;
Chas,H1,H2,H3,Den1,Dlit_Trum:Double;
SREZ1:TAdvStringGrid;
//AD1:TAdvStringGrid;
Chet:Integer;
begin
     X:=0;
     J:=0;
     Dlit:=0;
     Dlit_Obsh:=0;
     Col_Grid:=97;
     Col_1:=0;
     Max:=0;
     //Sort:=TStringList.Create;
     SREZ1 := TAdvStringGrid.Create(Nil );
     //AD1 := TAdvStringGrid.Create(Nil );
     fillchar(Temp,sizeof(Temp),0);//Очистка массмва
     fillchar(SKPU,sizeof(SKPU),0);//Очистка массмва
     Form1.Memo13.Clear;
     //++Добавим к длительности Трумпф 1 день
    { for I := 0 to High(StrToCol) do
     begin
        Res_True:=Pos('rue',StrToCol[i,0]);
        Res_Trum:=Pos('Trumph',StrToCol[i,5]);
        if (Res_Trum<>0) AND (Res_True<>0) then  //ХУЙ
        begin
                Dlit_Min:=StrToFloat(StrToCol[i,2]);
                StrToCol[i,2]:=FloatToStr(Dlit_Min+1);
        end;
     end; }
     Dlit_Min:=0;
     //++++++++++++++++++++++++++++++++Ищем Минимальное по ОЧЕРЕДНОСТИ
     for I := 0 to High(StrToCol) do
     begin
        Res_True:=Pos('rue',StrToCol[i,0]);
        if Res_True<>0 then
        begin
                min := StrToInt(StrToCol[i,1]); // Первый не нулевой
                Och_Min:= StrToInt(StrToCol[i,1]);
                Col:=I;
                Dlit_Min:=StrToFloat(StrToCol[i,2]);
                Break;
        end;
     end;
     X:=0;
     A:=0;
     Y:=0;
     SREZ1.ColCount:=8;//AD   SREZ
     SREZ1.RowCount:=1;
     //Form1.SKPU.ColCount:=10; //AD1
     //Form1.SKPU.RowCount:=1;

     Dlit_Obsh:=0;
     for I := 0 to High(StrToCol) do
     begin
        Res_True:=Pos('rue',StrToCol[i,0]);
        if Res_True<>0 then
      begin
        Stanok_Posl:=StrToCol[i,5];
        For y:=0 To High(Smena) do
        Begin
          Stanok:=Smena[y,0];
          Res:=AnsiCompareStr(Stanok_Posl,Stanok);
          if Res=0 Then
          Begin
                Kol_Smen:=StrToInt(Smena[y,1]);
                Chas:=StrToFloat(Smena[y,2]);
                SREZ1.Cells[0,A]:=StrToCol[i,1]; //Ocher
                Ocher:= StrToInt(StrToCol[i,1]);

                SREZ1.Cells[1,A]:=Stanok_Posl;   //Stanok
                SREZ1.Cells[2,A]:=StrToCol[i,2];//Dlit
                Dlit_Min:=StrToFloat(StrToCol[i,2]);//*8;// 0,5*8=4 - пол смены
                                                    //1*8=8 полная смена
                Dlit_Obsh:=Dlit_Obsh+ (StrToFloat(StrToCol[i,2]));//В ЧАСАХ

                SREZ1.Cells[3,A]:=Smena[y,1];//Kol_Smen
                Kol_Smen:=StrToInt(Smena[y,1]);

                SREZ1.Cells[4,A]:=Smena[y,2];//Chas
                SREZ1.Cells[5,A]:='0';//Den
                SREZ1.Cells[6,A]:=StrToCol[I,3];//Н\Ч
                SREZ1.Cells[7,A]:=StrToCol[I,0];//TRUE
                SREZ1.RowCount:=SREZ1.RowCount+1;
                Inc(A);
          end;
        End;
      End;
     End;
     F1:=0;
     SREZ1.Cells[5,0]:='1';
     //AD.Sort(0,sdAscending);//sdDescending По убыванию
     SREZ1.SortByColumn(0);
     Dat_Plan:=StrToDate(Dat_Glob);
     A:=0;
     For I:=0 To SREZ1.RowCount-1 Do
     Begin
       if SREZ1.Cells[0,i]<> '' Then
       Begin
           {Form1.SKPU.Cells[0,A]:=Form1.SREZ1.Cells[0,i]; //Ocher
           Form1.SKPU.Cells[1,A]:=Form1.SREZ1.Cells[1,i];   //Stanok
           Form1.SKPU.Cells[2,A]:=Form1.SREZ1.Cells[2,i];//Dlit
           Form1.SKPU.Cells[3,A]:=Form1.SREZ1.Cells[3,i];//Kol_Smen
           Form1.SKPU.Cells[4,A]:=Form1.SREZ1.Cells[4,i];//Chas
           Form1.SKPU.Cells[5,A]:='0';//Den
           Form1.SKPU.Cells[6,A]:=Form1.SREZ1.Cells[6,i];//Н\Ч
           Form1.SKPU.Cells[7,A]:=Form1.SREZ1.Cells[7,i];//TRUE
           Form1.SKPU.RowCount:=Form1.SKPU.RowCount+1;}
           SKPU[0,a]:=SREZ1.Cells[0,i]; //Ocher
           SKPU[1,a]:=SREZ1.Cells[1,i];   //Stanok
           SKPU[2,a]:=SREZ1.Cells[2,i];//Dlit
           SKPU[3,a]:=SREZ1.Cells[3,i];//Kol_Smen
           SKPU[4,a]:=SREZ1.Cells[4,i];//Chas
           SKPU[5,a]:='0';//Den
           SKPU[6,a]:=SREZ1.Cells[6,i];//Н\Ч
           SKPU[7,a]:=SREZ1.Cells[7,i];//TRUE
           Inc(A);
       end;
     end;

     //For I:=0 To Form1.SKPU.RowCount-1 Do
     For i:=0 To High(SKPU) do
     Begin
       //if Form1.SKPU.Cells[0,I]<> '' Then
       if SKPU[0,I]<> '' Then
       Begin

            Ocher:= StrToInt(SKPU[0,I]);//Ocher
            Stanok_Posl:=SKPU[1,i];   //Stanok
            Dlit_Min:=StrToFloat(SKPU[2,I]);//Dlit

            Kol_Smen:=StrToInt(SKPU[3,I]);//Kol_Smen
            Chas:=StrToFloat(SKPU[4,I]);
            Den:= StrToInt(SKPU[5,I]);
            Celoe:=Pos(',',SKPU[2,I]);
            if Celoe<>0 Then
            Begin
                if I=0 Then
                Begin
                        H1:=(1-Dlit_Min);
                        SKPU[5,0]:=FloatToStr(1);
                        Form1.Memo13.Lines.Add('1');
                End;
                if I=1 Then
                Begin
                 if (H1=0.5) Then //Остаток часов от первого передела
                 Begin
                   SKPU[5,1]:='1';
                   Form1.Memo13.Lines.Add('1');
                   H2:=1;
                  End;
                 if (H1=0) OR (H1<0) Then    //ХУЙ   OR (H1<0)
                 Begin
                   SKPU[5,1]:='2';   // переносим на второй день
                   Form1.Memo13.Lines.Add('2');
                   H2:=2;
                  End;
                  H1:=(1-Dlit_Min);
                end;
                if I>1 Then //
                Begin
                 Chet:=I MOD 2;//Остаток от деления
                 if (Chet<>0) AND (Dlit_Min=0.5) Then //НЕ Четное
                 Begin
                    H1:=Ocher-Dlit_Min-H2;
                    Den1:= SimpleRoundTo(H1,0);
                    SKPU[5,I]:=(FloatToStr(Den1));
                    Form1.Memo13.Lines.Add(FloatToStr(Den1));
                 end;
                 if (Chet=0) AND (Dlit_Min=0.5) Then //Четное
                 Begin
                     H2:=Ocher-Dlit_Min-H1;
                     Den1:= SimpleRoundTo(H2,0);
                     SKPU[5,I]:=(FloatToStr(Den1));
                     Form1.Memo13.Lines.Add(FloatToStr(Den1));
                 end;
                end;
            End
            Else
            Begin
                  SKPU[5,I]:=(FloatToStr(Ocher));
                  Form1.Memo13.Lines.Add(FloatToStr(Ocher));
            end;
            {Den:= StrToInt(AD1.Cells[5,I]);
            Dat_Plan:=StrToDate(Dat_Glob);
            Dat_Z:=FormatDateTime('mm.dd.YYYY',Dat_Plan-Den);
            AD1.Cells[8,i]:=Dat_Z;  }
       end;
     end;
     A:= Form1.Memo13.Lines.Count-1 ;
     For i:=0 To High(SKPU) do
     Begin
       if Form1.Memo13.Lines.Strings[a]<> '' Then
       Begin
           SKPU[9,I]:=Form1.Memo13.Lines.Strings[a];
       end;
       A:=A-1;
     End;
     For i:=0 To High(SKPU) do
     Begin
       if SKPU[9,I]<> '' Then
       Begin
         Den:= StrToInt(SKPU[9,I]);
         Dat_Plan:=StrToDate(Dat_Glob);            //Конев 18.07.2017
         //Dat_Z:=FormatDateTime('mm.dd.YYYY',Dat_Plan-Den-1);//МИНУс 1ДЕНЬ
         //Яковлев, заготовка одним пакетом
         Dat_Z:=FormatDateTime('mm.dd.YYYY',Dat_Plan);//МИНУс 1ДЕНЬ
         SKPU[8,i]:=Dat_Z;

       end;
     End;

     for I := 0 to High(StrToCol) do
     begin
          if StrToCol[i,5]<>'' Then
          Begin
            For a:=0 To High(SKPU) do
            Begin
                Res:=AnsiCompareStr(StrToCol[i,5],SKPU[1,A]);
                if Res=0 Then
                Begin
                    StrToCol[i,4]:=FloatToStr(Dlit_Obsh);
                    StrToCol[i,6]:=SKPU[8,A];//Дата заготовка
                end;
            End;
          End;
     end;

     for I := 0 to High(StrToCol) do
     begin
        if StrToCol[i,5]<>'' then
        begin
        S1:=S1+#39+StrToCol[I,0]+#39+','; //True  0
        S1:=S1+#39+StrToCol[I,1]+#39+','; //Очередность 1
        S1:=S1+#39+StrToCol[I,2]+#39+','; //Длитель  2
        S1:=S1+#39+StrToCol[I,3]+#39+','; //Н\ч  3
        if StrToCol[I,2]='' then
          S1:=S1+'NULL,';
        //
        if (StrToCol[I,2]<>'') then
        Begin
            S1:=S1+#39+StrToCol[I,6]+#39+',';
        End;
        //
        S1:=S1+#39+StrToCol[I,4]+#39+',';
        //Дата Заготовки  5 будем искать во временной базе
        end;
     end;
     FreeAndNil(Srez1);
     //FreeAndNil(SKPU);
end;
end.
