unit USborAll;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Vcl.ExtCtrls, Vcl.Buttons, Vcl.ComCtrls;

type
  TFSborAll = class(TForm)
    ComboBox1: TComboBox;
    Button1: TButton;
    Button2: TButton;
    Label1: TLabel;
    CBB1: TComboBox;
    CBB2: TComboBox;
    CBB3: TComboBox;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    btn1: TBitBtn;
    btn2: TBitBtn;
    btn3: TBitBtn;
    dtp1: TDateTimePicker;
    dtp2: TDateTimePicker;
    dtp3: TDateTimePicker;
    dtp4: TDateTimePicker;
    Label6: TLabel;
    Label7: TLabel;
    procedure Button2KeyPress(Sender: TObject; var Key: Char);
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure ComboBox1KeyPress(Sender: TObject; var Key: Char);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btn1Click(Sender: TObject);
    procedure btn2Click(Sender: TObject);
    procedure btn3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    All,Sbor,Sbor1:Integer;
    KO,GP:String;
  end;

var
  FSborAll: TFSborAll;

implementation

uses Main;

{$R *.dfm}

procedure TFSborAll.Button2KeyPress(Sender: TObject; var Key: Char);
begin
     If Key=#27 Then
     Close;
end;

procedure TFSborAll.FormShow(Sender: TObject);
Var
I,Ger,du:integer;
S:string;
begin

        if not Form1.mkQuerySelect(Form1.ADOQuery2,
                'Select * from %s WHERE МаркаАвто='+#39+'Клапана'+#39+' Order by Фамилия',
                ['Сборщик']) then
                exit;
        ComboBox1.Items.Add('');
        ComboBox1.DropDownCount := 30;
        CBB1.Items.Add('');
        CBB1.DropDownCount := 30;
        CBB2.Items.Add('');
        CBB2.DropDownCount := 30;
        CBB3.Items.Add('');
        CBB3.DropDownCount := 30;
        for I := 0 to Form1.ADOQuery2.RecordCount - 1 do
        begin
                ComboBox1.Items.Add(Form1.ADOQuery2.FieldByName('Фамилия').AsString);
                CBB1.Items.Add(Form1.ADOQuery2.FieldByName('Фамилия').AsString);
                CBB2.Items.Add(Form1.ADOQuery2.FieldByName('Фамилия').AsString);
                CBB3.Items.Add(Form1.ADOQuery2.FieldByName('Фамилия').AsString);
                Form1.ADOQuery2.Next;
        end;
        ComboBox1.SetFocus;
         if Form1.Vozduh=0 Then
        Begin
                if not Form1.mkQuerySelect(Form1.ADOQuery2,
                'Select * from %s  Where Номер='+#39+Label1.Caption+#39,
                ['Запуск']) then
                exit;
                S:=Form1.ADOQuery2.FieldByName('Сборка Готовность').AsString;
                if S<>'' then
                dtp1.DateTime:=Form1.ADOQuery2.FieldByName('Сборка Готовность').AsDateTime
                else
                 dtp1.DateTime:=Now;
                //
                S:=Form1.ADOQuery2.FieldByName('Подсборка').AsString;
                if S<>'' then
                dtp2.DateTime:=Form1.ADOQuery2.FieldByName('Подсборка').AsDateTime
                else
                 dtp2.DateTime:=Now;
                //
                S:=Form1.ADOQuery2.FieldByName('ЛопаткаДат').AsString;
                if S<>'' then
                dtp3.DateTime:=Form1.ADOQuery2.FieldByName('ЛопаткаДат').AsDateTime
                else
                 dtp3.DateTime:=Now;
                //
                S:=Form1.ADOQuery2.FieldByName('КорпусДат').AsString;
                if S<>'' then
                dtp4.DateTime:=Form1.ADOQuery2.FieldByName('КорпусДат').AsDateTime
                else
                 dtp4.DateTime:=Now;
                //
                ComboBox1.Text:=Form1.ZSG.Cells[I_FN_KOL_ZAP + 12, Form1.ZSG.Row] ;
                CBB1.Text:=Form1.ZSG.Cells[I_FN_KOL_ZAP + 25, Form1.ZSG.Row] ;
                CBB2.Text:=Form1.ZSG.Cells[I_FN_KOL_ZAP + 33, Form1.ZSG.Row]  ;
                CBB3.Text:=Form1.ZSG.Cells[I_FN_KOL_ZAP + 35, Form1.ZSG.Row] ;
        End;
        if Form1.Vozduh=1 Then
        Begin
            if not Form1.mkQuerySelect(Form1.ADOQuery2,
                'Select * from %s  Where Номер='+#39+Label1.Caption+#39,
                ['ЗапускВозд']) then
                exit;
                S:=Form1.ADOQuery2.FieldByName('Сборка Готовность').AsString;
                if S<>'' then
                dtp1.DateTime:=Form1.ADOQuery2.FieldByName('Сборка Готовность').AsDateTime
                else
                 dtp1.DateTime:=Now;
                //
                S:=Form1.ADOQuery2.FieldByName('Подсборка').AsString;
                if S<>'' then
                dtp2.DateTime:=Form1.ADOQuery2.FieldByName('Подсборка').AsDateTime
                else
                 dtp2.DateTime:=Now;
                du:=Pos('ГЕРМИК-ДУ-',Label7.caption);
                if DU<>0 then
                Begin
                  //
                  FSborAll.CBB2.Visible:=True;
                  FSborAll.CBB3.Visible:=True;
                  S:=Form1.ADOQuery2.FieldByName('ЛопаткаДат').AsString;
                  if S<>'' then
                    dtp3.DateTime:=Form1.ADOQuery2.FieldByName('ЛопаткаДат').AsDateTime
                  else
                    dtp3.DateTime:=Now;
                  //
                  S:=Form1.ADOQuery2.FieldByName('ТягаДат').AsString;
                  if S<>'' then
                    dtp4.DateTime:=Form1.ADOQuery2.FieldByName('ТягаДат').AsDateTime
                  else
                    dtp4.DateTime:=Now;
                End;
                ComboBox1.Text:=Form1.ZCV.Cells[17, Form1.ZCV.Row] ;
                CBB1.Text:=Form1.ZCV.Cells[43, Form1.ZCV.Row] ;
                CBB2.Text:=Form1.ZCV.Cells[44, Form1.ZCV.Row]  ;
                CBB3.Text:=Form1.ZCV.Cells[45, Form1.ZCV.Row] ;
        End;
        if CBB2.Visible=True then
        begin
            CBB2.Enabled:=True;
        end;
        if CBB2.Visible=False then
        begin
            CBB2.Visible:=True;
            CBB2.Enabled:=False;
        end;

        if CBB3.Visible=True then
        begin
            CBB3.Enabled:=True;
        end;
        if CBB3.Visible=False then
        begin
            CBB3.Visible:=True;
            CBB3.Enabled:=False;
        end;
end;

procedure TFSborAll.btn1Click(Sender: TObject);
begin
     if CBB1.Text='' then
     begin
        CBB1.Text:= ComboBox1.Text;
        dtp2.DateTime:=dtp1.DateTime;
     end

     else
      CBB1.Text:='';
end;

procedure TFSborAll.btn2Click(Sender: TObject);
begin
     if CBB2.Text='' then
      begin
        CBB2.Text:= ComboBox1.Text;
        dtp3.DateTime:=dtp1.DateTime;
     end
     else
      CBB2.Text:='';
end;

procedure TFSborAll.btn3Click(Sender: TObject);
begin
     if CBB3.Text='' then
      begin
        CBB3.Text:= ComboBox1.Text;
        dtp4.DateTime:=dtp1.DateTime;
     end
     else
      CBB3.Text:='';
end;

procedure TFSborAll.Button1Click(Sender: TObject);
Var I,Nom,R:Integer;
Str,Str1,Str2,Str3,Str4:string;
begin
      if ComboBox1.Text<>'' then
                Str1:=#39+FormatDateTime('mm.dd.YYYY',dtp1.DateTime)+#39
      else
                Str1:='NULL';
                //
      if CBB1.Text<>'' then
                Str2:=#39+FormatDateTime('mm.dd.YYYY',dtp2.DateTime)+#39
      else
                Str2:='NULL';
                //
      if CBB2.Text<>'' then
                Str3:=#39+FormatDateTime('mm.dd.YYYY',dtp3.DateTime)+#39
      else
                Str3:='NULL';
                //
      if CBB3.Text<>'' then
                Str4:=#39+FormatDateTime('mm.dd.YYYY',dtp4.DateTime)+#39
      else
                Str4:='NULL';
        if Form1.Vozduh=0 Then
        Begin
          // if Sbor1=0 then
           begin
            if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                'UPDATE %s SET [%s]=%s WHERE ([Номер]=%s) ',
                ['Klapana', 'Сборщик', #39 + ComboBox1.Text
                + #39, #39 + Label1.Caption + #39]) then
                Exit;

            if not Form1.mkQueryUpdate(Form1.ADOQuery1,          //[%s]=%s,'Сборка Готовность',Str1,    ,[%s]=%s,[%s]=%s,[%s]=%s
                'UPDATE %s SET [%s]=%s,[%s]=%s,[%s]=%s,'+
                ' [%s]=%s,[%s]=%s,[%s]=%s,[%s]=%s WHERE ([Номер]=%s)',
                ['Запуск', 'Сборщик', #39 + ComboBox1.Text
                + #39,
                 'Подсборка1', #39 + CBB1.Text
                + #39,'Подсборка',Str2,
                 'СборЛопатка', #39 + CBB2.Text
                + #39,'ЛопаткаДат',Str3,
                 'СборТяга', #39 + CBB3.Text
                + #39,'КорпусДат',Str4,

                #39 + Label1.Caption + #39]) then
                Exit;
            For I:=2 to Form1.ZSG.RowCount+1 Do
            Begin
                Try
                Nom:=StrToInt(Form1.ZSG.Cells[0, i]);
                except
               Continue;
                end;
                if Nom=StrToInt(Label1.Caption) Then
                begin
                Form1.ZSG.Cells[I_FN_KOL_ZAP + 12, i] :=
                ComboBox1.Text;
                Form1.ZSG.Cells[I_FN_KOL_ZAP + 25, i] :=
                CBB1.Text;
                Form1.ZSG.Cells[I_FN_KOL_ZAP + 33, i] :=
                CBB2.Text;
                Form1.ZSG.Cells[I_FN_KOL_ZAP + 35, i] :=
                CBB3.Text;
                end;

            end;
           end;


        end;
        //+++++++++++++++++++++++
          //ZCV.Cells[43, 1] := 'СборЛопатка';
          //ZCV.Cells[44, 1] := 'СборТяга';
          //ZCV.Cells[45, 1] := 'Подсборка1';
        if Form1.Vozduh=1 Then
        Begin
            if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                'UPDATE %s SET [%s]=%s WHERE ([Номер]=%s) ',
                ['KlapanaZap', 'Сборщик', #39 + ComboBox1.Text
                + #39, #39 + Label1.Caption + #39]) then
                Exit;
            r:=Pos('ГЕРМИК-ДУ-',Label7.caption);
                if r<>0 then
                Begin
                     if not Form1.mkQueryUpdate(Form1.ADOQuery1,        //[%s]=%s,  'Сборка Готовность',Str1,
                'UPDATE %s SET [%s]=%s,[%s]=%s,[%s]=%s,[%s]=%s,[%s]=%s,[%s]=%s,[%s]=%s WHERE ([Номер]=%s)',
                ['ЗапускВозд', 'Сборщик', #39 + ComboBox1.Text
                + #39,
                 'Подсборка1', #39 + CBB1.Text
                + #39,'Подсборка',Str2,
                'СборЛопатка', #39 + CBB2.Text
                + #39,'ЛопаткаДат',Str3,'СборТяга', #39 + CBB3.Text
                + #39,'ТягаДат',Str4, #39 + Label1.Caption + #39]) then
                  Exit;
                End
                else
                Begin

                  if not Form1.mkQueryUpdate(Form1.ADOQuery1,        //[%s]=%s,  'Сборка Готовность',Str1,
                'UPDATE %s SET [%s]=%s,[%s]=%s,[%s]=%s,[%s]=%s,[%s]=%s WHERE ([Номер]=%s)',
                ['ЗапускВозд', 'Сборщик', #39 + ComboBox1.Text
                + #39,
                 'Подсборка1', #39 + CBB1.Text
                + #39,'Подсборка',Str2,
                'СборЛопатка', #39 + CBB2.Text
                + #39,'СборТяга', #39 + CBB3.Text
                + #39, #39 + Label1.Caption + #39]) then
                  Exit;
                  End;

            For I:=1 to Form1.ZCV.RowCount Do
            Begin
                Try
                Nom:=StrToInt(Form1.ZCV.Cells[0, i]);
                except
                Continue;
                end;
                if Nom=StrToInt(Label1.Caption) Then
                Begin
                 {       Form1.ZCV.Cells[17, i] :=
                  ComboBox1.Text;

                  Form1.ZCV.Cells[18, i] :=
                  CBB1.Text;  }

                  Form1.ZCV.Cells[17, i] :=
                  ComboBox1.Text;
                  Form1.ZCV.Cells[43, i] :=
                  CBB1.Text;
                  Form1.ZCV.Cells[44, i] :=
                  CBB2.Text;
                  Form1.ZCV.Cells[45, i] :=
                  CBB3.Text;
                End;
            end;

        end;
                //+++++++++++++++++++++++
        if Form1.Vozduh=2 Then
        Begin
        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                'UPDATE %s SET [%s]=%s WHERE ([Номер]=%s) ',
                ['[750]', 'Сборщик', #39 + ComboBox1.Items[ComboBox1.ItemIndex]
                + #39, #39 + Label1.Caption + #39]) then
                Exit;
        if not Form1.mkQueryUpdate(Form1.ADOQuery1,
                'UPDATE %s SET [%s]=%s WHERE ([Номер]=%s)',
                ['[Запуск750]', 'Сборщик', #39 + ComboBox1.Items[ComboBox1.ItemIndex]
                + #39, #39 + Label1.Caption + #39]) then
                Exit;
        For I:=1 to Form1.zclrstrngrd1.RowCount Do
        Begin
                Try
                Nom:=StrToInt(Form1.zclrstrngrd1.Cells[0, i]);
                except
                Continue;
                end;
                if Nom=StrToInt(Label1.Caption) Then
                        Form1.zclrstrngrd1.Cells[I_FN_KOL_ZAP + 12, i] :=
                ComboBox1.Items[ComboBox1.ItemIndex];
        end;

        end;
        CBB1.Visible:=False;
        CBB2.Visible:=False;
        CBB3.Visible:=False;

        Close;
end;

procedure TFSborAll.Button2Click(Sender: TObject);
begin
        Close;
end;

procedure TFSborAll.ComboBox1KeyPress(Sender: TObject; var Key: Char);
begin
     If Key=#27 Then
     Close;
        if (Key = #13) then
                Button1Click(nil);
end;

procedure TFSborAll.FormClose(Sender: TObject; var Action: TCloseAction);
begin
        CBB1.Visible:=False;
        CBB2.Visible:=False;
        CBB3.Visible:=False;
end;

procedure TFSborAll.FormKeyPress(Sender: TObject; var Key: Char);
begin
     If Key=#27 Then
     Close;
        if (Key = #13) then
                Button1Click(nil);
end;

end.
