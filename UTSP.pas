unit UTSP;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, WinSock,
                uSockets, StdCtrls;

type
  TFTSP = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Memo1: TMemo;
    ListBox1: TListBox;
    Memo2: TMemo;
    Edit1: TEdit;
    Edit2: TEdit;
                  procedure Send(Str_Line: string);

    procedure Button1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    function GetLocalIPs: TStringList;
    procedure FormShow(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
                  Server: TTCPServer;
                ClientCounter: Integer;
                Client: TTCPClientConnection;
                procedure ClientConnect(Client: TTCPConnection);
                procedure ClientDisconnect(Client: TTCPConnection);
                procedure ClientExecute(Thread: TTCPConnectionThread);
                procedure ClientError(Socket: TTCPSocket);
    { Private declarations }
  public
  Client_Str:String;
    { Public declarations }
  end;

var
  FTSP: TFTSP;

implementation

uses Reg;

{$R *.dfm}
function TFTSP.GetLocalIPs: TStringList;
type
        TaPInAddr = array[0..10] of PInAddr;
        PaPInAddr = ^TaPInAddr;
var
        phe: PHostEnt;
        pptr: PaPInAddr;
        Buffer: array[0..63] of AnsiChar;
        I: Integer;
        GInitData: TWSAData;
        IPs: TStringList;
begin
        IPs := TStringList.Create;
        WSAStartup($101, GInitData);
        GetHostName(Buffer, SizeOf(Buffer));
        phe := GetHostByName(buffer);
        if phe = nil then
                IPs.Add('No IP found')
        else
        begin
                pPtr := PaPInAddr(phe^.h_addr_list);
                I := 0;
                while pPtr^[I] <> nil do
                begin
                        IPs.Add(inet_ntoa(pptr^[I]^));
                        Memo2.Lines.Add(inet_ntoa(pptr^[I]^));
                        Inc(I);
                end;
        end;
        WSACleanup;

        Result := IPs;
end;
procedure TFTSP.Send(Str_Line: string);
var
        II: Integer;
        Vn_Dat:String;
begin
        Vn_Dat := FormatDateTime('mm.dd.yyyy', Now);
        if Client.Connected then
        begin
                Str_Line:=Vn_Dat+' '+Regist.Famili+' '+Str_Line;
                //Memo1.Lines.Add(Format('> %s', [Str_Line]));
                Client.WriteLn(Str_Line);

        end;
end;


procedure TFTSP.ClientConnect(Client: TTCPConnection);
begin
        
        Memo1.Lines.Add('Connected');
        //Button2.Caption := 'Disconnect';
        Client.Detach;
end;

procedure TFTSP.ClientDisconnect(Client: TTCPConnection);
begin
        Memo1.Lines.Add('Disconnected');
        //Button2.Caption := 'Connect';
end;

procedure TFTSP.ClientExecute(Thread: TTCPConnectionThread);

begin

end;

procedure TFTSP.ClientError(Socket: TTCPSocket);
begin
        if Socket.LastError <> WSAECONNRESET then
                Memo1.Lines.Add(Format('Error #%d: %s', [Socket.LastError,
                        Socket.LastErrorMessage]));
end;
procedure TFTSP.Button1Click(Sender: TObject);
begin
        //ClientCounter := 0;

        Client := TTCPClientConnection.Create;
        with Client do
        begin
                OnConnect := ClientConnect;
                OnDisconnect := ClientDisconnect;
                OnExecute := ClientExecute;
                OnError := ClientError;
        if Client.Connected then
    Client.Disconnect
  else
    Client.Connect(Edit2.Text);
    end;

end;

procedure TFTSP.Button3Click(Sender: TObject);
begin
        Send(Edit1.Text);
end;

procedure TFTSP.FormDestroy(Sender: TObject);
begin

        //Client.Disconnect;
        //Client.Free;

end;

procedure TFTSP.FormShow(Sender: TObject);
begin
//GetLocalIPs;
end;

procedure TFTSP.Button2Click(Sender: TObject);
begin
        Client.Disconnect;
        Client.Free;
end;

end.
