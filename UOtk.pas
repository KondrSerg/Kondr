unit UOtk;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, Vcl.ExtCtrls;

type
  TFOTK = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    ME1: TEdit;
    DateTimePicker1: TDateTimePicker;
    Label9: TLabel;
    Label10: TLabel;
    Button3: TButton;
    LabGP: TLabel;
    LabKO: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    rg1: TRadioGroup;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button1KeyPress(Sender: TObject; var Key: Char);
    procedure FormShow(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure ME1KeyPress(Sender: TObject; var Key: Char);
    procedure DateTimePicker1KeyPress(Sender: TObject; var Key: Char);
    procedure Button2KeyPress(Sender: TObject; var Key: Char);
    procedure Button3Click(Sender: TObject);
    procedure rg1Click(Sender: TObject);
  private
    { Private declarations }
    Fam:string;
  public
    { Public declarations }
  end;

var
  FOTK: TFOTK;

implementation

uses Main, Reg, mainBRAK;

{$R *.dfm}

procedure TFOTK.Button1Click(Sender: TObject);
begin
        Close;
end;

procedure TFOTK.Button2Click(Sender: TObject);
var
        I, Res, Kol_ZAP, Kol,Kol_Ran,Kol_O,KO: Integer;
        Str, Zak,Nom, Nam, Dat,Dat1,ID,IDGP,IDKO: string;
        Dat_Got:TDate;
begin

        Nom := Label2.Caption;
        Nam := Label3.Caption;
        Zak := Label4.Caption;

        ID  := (Label6.Caption);
        IDGP:= LabGP.Caption;
        IDKO:= LabKO.Caption;
        Kol:=StrToInt(ME1.Text);
        Dat:=FormatDateTime('mm.dd.yyyy', DateTimePicker1.Date);
        Dat1:=FormatDateTime('dd.mm.yyyy', DateTimePicker1.Date);
        Kol_O:=0;

        if Kol=0 Then
        Begin
                MessageDlg('Поставте кол-во!', mtError,
                        [mbOk], 0);
                Exit;
        end;
        //СТАМ----------------------------------------------
        if Form1.Vozduh=3 Then
        Begin
            if Fam='' then
              Begin
                MessageDlg('Выберите фамилию!', mtError,
                        [mbOk], 0);
                Exit;
              end;
              if Form1.Luk=0 then
              begin
                     Kol_Ran := StrToInt(Label1.Caption);
                     if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        FN_KOL_OTK + ']=' + #39 + IntToStr(Kol+Kol_Ran) + #39 + ',[' + FN_PR_OTK + ']=' + #39 + Dat + #39 +
                        ' WHERE  ([IDГП]=' + #39 + IDGP + #39 +') ', ['СТАМ']) then
                        Exit;
                     if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        FN_KOL_OTK + ']=' + #39 + ME1.Text + #39 +',[' + FN_PR_OTK + ']=' + #39 + Dat + #39 +
                        ',[ОТКколраз]=ОТКколраз+' + #39 +','+ ME1.Text + #39 +
                        ',[ОТКдата]=ОТКдата+' + #39 +','+ Dat + #39 +',ОТКФам=' + #39 + Fam + #39 +
                        ' WHERE ([IDГП]=' + #39 + IDGP + #39 +') AND ([' + FN_NOM + ']=' + #39 + Nom + #39 +
                        ')', ['ЗапускСТАМ']) then
                        Exit;
                     Form1.Btn5.Click;
                     Close;
              end;

              if Form1.Luk=1 then
              begin
                     Kol_Ran := 0;//StrToInt(Label1.Caption);
                     if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        FN_KOL_OTK + ']=' + #39 + IntToStr(Kol+Kol_Ran) + #39 + ',[' + FN_PR_OTK + ']=' + #39 + Dat + #39 +
                        ' WHERE  ([IDГП]=' + #39 + IDGP + #39 +') ', ['ЛЮК']) then
                        Exit;
                     if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        FN_KOL_OTK + ']=' + #39 + ME1.Text + #39 +',[' + FN_PR_OTK + ']=' + #39 + Dat + #39 +
                        ',[ОТКколраз]=ОТКколраз+' + #39 +','+ ME1.Text + #39 +
                        ',[ОТКдата]=ОТКдата+' + #39 +','+ Dat + #39 +',ОТКФам=' + #39 + Fam + #39 +
                        ' WHERE ([IDГП]=' + #39 + IDGP + #39 +') AND ([' + FN_NOM + ']=' + #39 + Nom + #39 +
                        ')', ['ЗапускЛЮК']) then
                        Exit;
                     Form1.Btn46.Click;
                     Close;
              end;

        end;
        Kol_Zap := StrToInt(Label5.Caption);
        //-----------------------------------------------
        if Kol>Kol_Zap Then
        Begin
                MessageDlg('Было запущенно всего '+IntToStr(Kol_Zap)+' по этой накладной!', mtError,
                        [mbOk], 0);
                Exit;
        end;
        {Dat_Got:=StrToDate(Label10.Caption);
        if Dat_Got>DateTimePicker1.Date Then
        Begin
                MessageDlg('Дата приемки не соответствует дате изготовления!', mtError,
                        [mbOk], 0);
                Exit;
        end; }
        if Form1.Vozduh=0 Then
        Begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        FN_KOL_OTK + ']=' + #39 + ME1.Text + #39 + ',[' + FN_PR_OTK + ']=' + #39 + Dat + #39 +
                        ' WHERE ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ')', ['Klapana']) then
                        Exit;
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        FN_KOL_OTK + ']=' + #39 + ME1.Text + #39 +',[' + FN_PR_OTK + ']=' + #39 + Dat + #39 +
                        ',[ОТКколраз]=ОТКколраз+' + #39 +','+ ME1.Text + #39 +
                        ',[ОТКдата]=ОТКдата+' + #39 +','+ Dat + #39 +',ОТКФам=' + #39 + Fam + #39 +
                        ' WHERE ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ') AND ([ID]=' + #39 + Id + #39 +')  ',
                         ['Запуск']) then
                        Exit;
                if Form1.Edit4.Text<>'' Then
                        Form1.Button36.Click;
        End;
        if Form1.Vozduh=1 Then
        Begin
                    // Fam:=Regist.Famili;
                     if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        FN_KOL_OTK + ']=' + #39 + ME1.Text + #39 + ',[' + FN_PR_OTK + ']=' + #39 + Dat + #39 +
                        ' WHERE ([IDКО]=' + #39 + IDKO + #39 +') AND ([IDГП]=' + #39 + IDGP + #39 +') AND ([' + FN_NAM + ']=' + #39 + Nam + #39 +') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ')', ['KlapanaZap']) then
                        Exit;
                     if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        FN_KOL_OTK + ']=' + #39 + ME1.Text + #39 +',[' + FN_PR_OTK + ']=' + #39 + Dat + #39 +
                        ',[ОТКколраз]=ОТКколраз+' + #39 +','+ ME1.Text + #39 +
                        ',[ОТКдата]=ОТКдата+' + #39 +','+ Dat + #39 + ',ОТКФам=' + #39+ Fam + #39 +
                        ' WHERE ([IDКО]=' + #39 + IDKO + #39 +') AND ([IDГП]=' + #39 + IDGP + #39 +') AND ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ') AND ([ID]=' + #39 + Id + #39 +')',
                        ['ЗапускВозд']) then
                        Exit;
                     Form1.Button46.Click;
        end;
        if Form1.Vozduh=2 Then
        Begin
                    // Fam:=Regist.Famili;

                     if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE [%s] SET [' +
                        FN_KOL_OTK + ']=' + #39 + ME1.Text + #39 +',[' + FN_PR_OTK + ']=' + #39 + Dat + #39 +
                        ',[ОТКколраз]=ОТКколраз+' + #39 +','+ ME1.Text + #39 +
                        ',[ОТКдата]=ОТКдата+' + #39 +','+ Dat + #39 + ',ОТКФам=' + #39+ Fam + #39 +
                        ' WHERE ([IDКО]=' + #39 + IDKO + #39 +') AND ([IDГП]=' + #39 + IDGP + #39 +') AND ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ') AND ([ID]=' + #39 + Id + #39 +')',
                        ['Запуск750']) then
                        Exit;
                    if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select [Кол принятых] from [%s] WHERE ([IDКО]=' + #39 + IDKO + #39 +') AND ([IDГП]=' + #39 + IDGP + #39 +')', ['Запуск750']) then
                    exit;
                    for I := 0 to Form1.ADOQuery1.RecordCount-1 do
                    Begin
                      Kol_O:=Kol_o+ Form1.ADOQuery1.FieldByName('Кол принятых').AsInteger;
                      Form1.ADOQuery1.Next;
                    End;
                    Label1.Caption:=Form1.ADOQuery1.FieldByName('Кол принятых').AsString;
                     if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE [%s] SET [' +
                        FN_KOL_OTK + ']=' + #39 + IntToStr(Kol_O) + #39 + ',[' + FN_PR_OTK + ']=' + #39 + Dat + #39 +
                        ' WHERE ([IDКО]=' + #39 + IDKO + #39 +') AND ([IDГП]=' + #39 + IDGP + #39 +') AND ([' + FN_NAM + ']=' + #39 + Nam + #39 +') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ')', ['750']) then
                        Exit;

                     Form1.Button71.Click;
        end;
              if not FBRAK.mkQueryUpdate(FBRAK.ADOQuery1,
                'UPDATE %s SET  ДатаОТК='+#39+ Dat +#39 +
                ' WHERE ([IdGP]=' + #39 + IDGP + #39 + ') and ([IdKO]=' + #39 + IDKO + #39 + ')', ['Брак']) then
                Exit;


        Close;
end;

procedure TFOTK.Button1KeyPress(Sender: TObject; var Key: Char);
begin
     If Key=#27 Then
     Close;
end;

procedure TFOTK.FormShow(Sender: TObject);
begin
ME1.SetFocus;
ME1.AutoSelect:=True;
ME1.SelectAll;
Fam:='';
  rg1.Items.Clear;
  rg1.Items.LoadFromFile( 'V:\ОИТ\Cklapana2\2013\OTK.ini');
  rg1.ItemIndex:=-1;
        rg1.Visible:=true;
        if Form1.Vozduh=3 Then
        Begin

          if not Form1.mkQuerySelect(Form1.ADOQuery1, 'Select [Кол принятых] from %s WHERE (IDГП='+#39+LabGP.Caption+#39+') AND (L=0) ', ['СТАМ']) then
                exit;
          Label1.Caption:=Form1.ADOQuery1.FieldByName('Кол принятых').AsString;
        end;
end;

procedure TFOTK.FormKeyPress(Sender: TObject; var Key: Char);
begin
     If Key=#27 Then
     Close;
        if (Key = #13) then
                Button2Click(nil);
end;

procedure TFOTK.ME1KeyPress(Sender: TObject; var Key: Char);
begin
     If Key=#27 Then
     Close;
        if (Key = #13) then
                Button2Click(nil);
end;

procedure TFOTK.rg1Click(Sender: TObject);
begin
  Fam:=rg1.Items[rg1.ItemIndex];
end;

procedure TFOTK.DateTimePicker1KeyPress(Sender: TObject; var Key: Char);
begin
     If Key=#27 Then
     Close;
        if (Key = #13) then
                Button2Click(nil);
end;

procedure TFOTK.Button2KeyPress(Sender: TObject; var Key: Char);
begin
     If Key=#27 Then
     Close;
        if (Key = #13) then
                Button2Click(nil);
end;

procedure TFOTK.Button3Click(Sender: TObject);
var
        I, Res, Kol_ZAP, Kol: Integer;
        Str, Zak,Nom, Nam, Dat,ID,IDGP,IDKO,Fam: string;
        Dat_Got:TDate;
begin

        Nom := Label2.Caption;
        Nam := Label3.Caption;
        Zak := Label4.Caption;
        //Fam:=Regist.Famili;
        Kol_Zap := StrToInt(Label5.Caption);
        ID  := (Label6.Caption);
        IDGP:= LabGP.Caption;
        IDKO:= LabKO.Caption;
        if Form1.Vozduh=0 Then
        Begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        FN_KOL_OTK + ']=' + #39 + '0' + #39 + ',[' + FN_PR_OTK + ']=NULL' +
                        ' WHERE ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ')', ['Klapana']) then
                        Exit;

                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        FN_KOL_OTK + ']=' + #39 + '0' + #39 +
                        ',[' + FN_PR_OTK + ']=NULL,ОТКФам=' + #39 +Fam + #39 +'  WHERE ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ') AND ([ID]=' + #39 + Id + #39 +')  ',
                        ['Запуск']) then
                        Exit;
        end;
        if Form1.Vozduh=1 Then
        Begin
                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        FN_KOL_OTK + ']=' + #39 + '0' + #39 + ',[' + FN_PR_OTK + ']=NULL' +
                        ' WHERE ([IDКО]=' + #39 + IDKO + #39 +') AND ([IDГП]=' + #39 + IDGP + #39 +') AND  ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ')', ['KlapanaZap']) then
                        Exit;

                if not Form1.mkQueryUpdate(Form1.ADOQuery1, 'UPDATE %s SET [' +
                        FN_KOL_OTK + ']=' + #39 + '0' + #39 +
                        ',[' + FN_PR_OTK + ']=NULL,ОТКФам=' + #39 + Fam + #39 +'  WHERE ([IDКО]=' + #39 + IDKO + #39 +') AND ([IDГП]=' + #39 + IDGP + #39 +') AND  ([' + FN_NAM + ']=' + #39 + Nam + #39 +
                        ') AND ([' + FN_ZAK + ']='
                        + #39 + ZAK + #39 + ') AND ([ID]=' + #39 + Id + #39 +')  ',
                        ['ЗапускВозд']) then
                        Exit;
        end;
        Close;
end;

end.
