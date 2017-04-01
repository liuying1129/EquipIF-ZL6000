unit UfrmMain;

interface

uses
  Windows, Messages, SysUtils, Classes, Controls, Forms,
  LYTray, Menus, StdCtrls, Buttons, ADODB,
  ActnList, AppEvnts, ComCtrls, ToolWin, ExtCtrls,
  registry,inifiles,Dialogs,
  StrUtils, DB,ComObj,Variants,Math;

type
  TfrmMain = class(TForm)
    LYTray1: TLYTray;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    ADOConnection1: TADOConnection;
    ApplicationEvents1: TApplicationEvents;
    CoolBar1: TCoolBar;
    ToolBar1: TToolBar;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    ActionList1: TActionList;
    editpass: TAction;
    about: TAction;
    stop: TAction;
    ToolButton2: TToolButton;
    ToolButton5: TToolButton;
    ToolButton9: TToolButton;
    OpenDialog1: TOpenDialog;
    ADOConn_BS: TADOConnection;
    BitBtn3: TBitBtn;
    DateTimePicker1: TDateTimePicker;
    Label1: TLabel;
    Label2: TLabel;
    procedure N3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure ApplicationEvents1Activate(Sender: TObject);
    procedure ToolButton7Click(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure ToolButton5Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
  private
    { Private declarations }
    procedure WMSyscommand(var message:TWMMouse);message WM_SYSCOMMAND;
    procedure UpdateConfig;{�����ļ���Ч}
    function LoadInputPassDll:boolean;
    function MakeDBConn:boolean;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

uses ucommfunction;

const
  CR=#$D+#$A;
  STX=#$2;ETX=#$3;ACK=#$6;NAK=#$15;
  sCryptSeed='lc';//�ӽ�������
  //SEPARATOR=#$1C;
  sCONNECTDEVELOP='����!���뿪������ϵ!' ;
  IniSection='Setup';

var
  ConnectString:string;
  GroupName:string;//
  SpecType:string ;//
  SpecStatus:string ;//
  CombinID:string;//
  LisFormCaption:string;//
  QuaContSpecNoG:string;
  QuaContSpecNo:string;
  QuaContSpecNoD:string;
  EquipChar:string;
  MrConnStr:string;
  ifConnSucc:boolean;

  hnd:integer;
  bRegister:boolean;

{$R *.dfm}

function ifRegister:boolean;
var
  HDSn,RegisterNum,EnHDSn:string;
  configini:tinifile;
  pEnHDSn:Pchar;
begin
  result:=false;
  
  HDSn:=GetHDSn('C:\')+'-'+GetHDSn('D:\')+'-'+ChangeFileExt(ExtractFileName(Application.ExeName),'');

  CONFIGINI:=TINIFILE.Create(ChangeFileExt(Application.ExeName,'.ini'));
  RegisterNum:=CONFIGINI.ReadString(IniSection,'RegisterNum','');
  CONFIGINI.Free;
  pEnHDSn:=EnCryptStr(Pchar(HDSn),sCryptSeed);
  EnHDSn:=StrPas(pEnHDSn);

  if Uppercase(EnHDSn)=Uppercase(RegisterNum) then result:=true;

  if not result then messagedlg('�Բ���,��û��ע���ע�������,��ע��!',mtinformation,[mbok],0);
end;

function GetConnectString:string;
var
  Ini:tinifile;
  userid, password, datasource, initialcatalog: string;
  ifIntegrated:boolean;//�Ƿ񼯳ɵ�¼ģʽ

  pInStr,pDeStr:Pchar;
  i:integer;
begin
  result:='';
  
  Ini := tinifile.Create(ChangeFileExt(Application.ExeName,'.INI'));
  datasource := Ini.ReadString('�������ݿ�', '������', '');
  initialcatalog := Ini.ReadString('�������ݿ�', '���ݿ�', '');
  ifIntegrated:=ini.ReadBool('�������ݿ�','���ɵ�¼ģʽ',false);
  userid := Ini.ReadString('�������ݿ�', '�û�', '');
  password := Ini.ReadString('�������ݿ�', '����', '107DFC967CDCFAAF');
  Ini.Free;
  //======����password
  pInStr:=pchar(password);
  pDeStr:=DeCryptStr(pInStr,sCryptSeed);
  setlength(password,length(pDeStr));
  for i :=1  to length(pDeStr) do password[i]:=pDeStr[i-1];
  //==========

  result := result + 'user id=' + UserID + ';';
  result := result + 'password=' + Password + ';';
  result := result + 'data source=' + datasource + ';';
  result := result + 'Initial Catalog=' + initialcatalog + ';';
  result := result + 'provider=' + 'SQLOLEDB.1' + ';';
  //Persist Security Info,��ʾADO�����ݿ����ӳɹ����Ƿ񱣴�������Ϣ
  //ADOȱʡΪTrue,ADO.netȱʡΪFalse
  //�����лᴫADOConnection��Ϣ��TADOLYQuery,������ΪTrue
  result := result + 'Persist Security Info=True;';
  if ifIntegrated then
    result := result + 'Integrated Security=SSPI;';
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var
  ctext        :string;
  reg          :tregistry;
begin
  ConnectString:=GetConnectString;
  
  UpdateConfig;
  DateTimePicker1.DateTime:=now;
  if ifRegister then bRegister:=true else bRegister:=false;  

  lytray1.Hint:='���ݽ��շ���'+ExtractFileName(Application.ExeName);

//=============================��ʼ������=====================================//
    reg:=tregistry.Create;
    reg.RootKey:=HKEY_CURRENT_USER;
    reg.OpenKey('\sunyear',true);
    ctext:=reg.ReadString('pass');
    if ctext='' then
    begin
        reg:=tregistry.Create;
        reg.RootKey:=HKEY_CURRENT_USER;
        reg.OpenKey('\sunyear',true);
        reg.WriteString('pass','JIHONM{');
        //MessageBox(application.Handle,pchar('��л��ʹ�����ܼ��ϵͳ��'+chr(13)+'���ס��ʼ�����룺'+'lc'),
        //            'ϵͳ��ʾ',MB_OK+MB_ICONinformation);     //WARNING
    end;
    reg.CloseKey;
    reg.Free;
//============================================================================//
end;

procedure TfrmMain.N3Click(Sender: TObject);
begin
    if not LoadInputPassDll then exit;
    application.Terminate;
end;

procedure TfrmMain.N1Click(Sender: TObject);
begin
  show;
end;

procedure TfrmMain.ApplicationEvents1Activate(Sender: TObject);
begin
  hide;
end;

procedure TfrmMain.WMSyscommand(var message: TWMMouse);
begin
  inherited;
  if message.Keys=SC_MINIMIZE then hide;
  message.Result:=-1;
end;

procedure TfrmMain.ToolButton7Click(Sender: TObject);
begin
  if MakeDBConn then ConnectString:=GetConnectString;
end;

procedure TfrmMain.UpdateConfig;
var
  INI:tinifile;
  autorun:boolean;
begin
  ini:=TINIFILE.Create(ChangeFileExt(Application.ExeName,'.ini'));

  autorun:=ini.readBool(IniSection,'�����Զ�����',false);

  GroupName:=trim(ini.ReadString(IniSection,'���',''));
  EquipChar:=trim(uppercase(ini.ReadString(IniSection,'������ĸ','')));//�������Ǵ�д������һʧ��
  SpecType:=ini.ReadString(IniSection,'Ĭ����������','');
  SpecStatus:=ini.ReadString(IniSection,'Ĭ������״̬','');
  CombinID:=ini.ReadString(IniSection,'�����Ŀ����','');

  LisFormCaption:=ini.ReadString(IniSection,'����ϵͳ�������','');

  QuaContSpecNoG:=ini.ReadString(IniSection,'��ֵ�ʿ�������','9999');
  QuaContSpecNo:=ini.ReadString(IniSection,'��ֵ�ʿ�������','9998');
  QuaContSpecNoD:=ini.ReadString(IniSection,'��ֵ�ʿ�������','9997');

  MrConnStr:=ini.ReadString(IniSection,'�����������ݿ�','');

  ini.Free;

  OperateLinkFile(application.ExeName,'\'+ChangeFileExt(ExtractFileName(Application.ExeName),'.lnk'),15,autorun);

  try
    ADOConn_BS.Connected := false;
    ADOConn_BS.ConnectionString := MrConnStr;
    ADOConn_BS.Connected := true;
    ifConnSucc:=true;
  except
    ifConnSucc:=false;
    showmessage('�����������ݿ�ʧ��!');
  end;
end;

function TfrmMain.LoadInputPassDll: boolean;
TYPE
    TDLLFUNC=FUNCTION:boolean;
VAR
    HLIB:THANDLE;
    DLLFUNC:TDLLFUNC;
    PassFlag:boolean;
begin
    result:=false;
    HLIB:=LOADLIBRARY('OnOffLogin.dll');
    IF HLIB=0 THEN BEGIN SHOWMESSAGE(sCONNECTDEVELOP);EXIT; END;
    DLLFUNC:=TDLLFUNC(GETPROCADDRESS(HLIB,'showfrmonofflogin'));
    IF @DLLFUNC=NIL THEN BEGIN SHOWMESSAGE(sCONNECTDEVELOP);EXIT; END;
    PassFlag:=DLLFUNC;
    FREELIBRARY(HLIB);
    result:=passflag;
end;

function TfrmMain.MakeDBConn:boolean;
var
  newconnstr,ss: string;
  Label labReadIni;
begin
  result:=false;

  labReadIni:
  newconnstr := GetConnectString;
  
  try
    ADOConnection1.Connected := false;
    ADOConnection1.ConnectionString := newconnstr;
    ADOConnection1.Connected := true;
    result:=true;
  except
  end;
  if not result then
  begin
    ss:='������'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '���ݿ�'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '���ɵ�¼ģʽ'+#2+'CheckListBox'+#2+#2+'0'+#2+#2+#3+
        '�û�'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '����'+#2+'Edit'+#2+#2+'0'+#2+#2+'1';
    if ShowOptionForm('�������ݿ�','�������ݿ�',Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
      goto labReadIni else application.Terminate;
  end;
end;

procedure TfrmMain.ToolButton2Click(Sender: TObject);
var
  ss:string;
begin
  if LoadInputPassDll then
  begin
    ss:='�����������ݿ�'+#2+'DBConn'+#2+#2+'1'+#2+#2+#3+
      '���'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '������ĸ'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '����ϵͳ�������'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      'Ĭ����������'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      'Ĭ������״̬'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '�����Ŀ����'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '�����Զ�����'+#2+'CheckListBox'+#2+#2+'1'+#2+#2+#3+
      '��ֵ�ʿ�������'+#2+'Edit'+#2+#2+'2'+#2+#2+#3+
      '��ֵ�ʿ�������'+#2+'Edit'+#2+#2+'2'+#2+#2+#3+
      '��ֵ�ʿ�������'+#2+'Edit'+#2+#2+'2'+#2+#2+#3;

  if ShowOptionForm('',Pchar(IniSection),Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
	  UpdateConfig;
  end;
end;

procedure TfrmMain.ToolButton5Click(Sender: TObject);
var
  ss:string;
begin
  ss:='RegisterNum'+#2+'Edit'+#2+#2+'0'+#2+'���ô���������ϵ��ַ�������������,�Ի�ȡע����'+#2;
  if bRegister then exit;
  if ShowOptionForm(Pchar('ע��:'+GetHDSn('C:\')+'-'+GetHDSn('D:\')+'-'+ChangeFileExt(ExtractFileName(Application.ExeName),'')),Pchar(IniSection),Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
    if ifRegister then bRegister:=true else bRegister:=false;
end;

procedure TfrmMain.BitBtn3Click(Sender: TObject);
VAR
  adotemp22,adotemp,adotemp33:tadoquery;
  SamNo:string;
  ReceiveItemInfo:OleVariant;
  FInts:OleVariant;
  sName,sSex,sAge,sKB,sBQ,sBLH,sBedNo,sLCZD,sSJYS,sJYYS:String;
  i,RecNum:integer;
begin
  if not ifConnSucc then
  begin
    showmessage('�����������ݿ�ʧ��,���ܷ���!');
    exit;
  end;
  
  (sender as TBitBtn).Enabled:=false;  

  adotemp22:=tadoquery.Create(nil);
  adotemp22.Connection:=ADOConn_BS;
  adotemp22.Close;
  adotemp22.SQL.Clear;
  adotemp22.SQL.Text:='select TestDataID,���,����,�Ա�,����,�Ʊ�,����,������,����,�ٴ����,����,ʱ��,�ͼ�ҽ��,����ҽ��,'+
                      'ȫѪճ��,Ѫ��ճ��,ѹ��,Ѫ��,Ѫ����������,Ѫ����������ʱ��,ȫѪ�������ָ��,ȫѪ�������ָ��,'+
                      'Ѫ������Kֵ,��ϸ���ۼ�ָ��,��ϸ���ۼ�ϵ��,��ϸ������ָ��,ȫѪ���л�ԭճ��,ȫѪ���л�ԭճ��,'+
                      '��ϸ������ָ��TK,��ϸ������ָ��,����ճ��,Ѫ�쵰��,��ϸ����ճ��,��������,��������,��������,��ά����ԭ,Ѫ���̴�,'+
                      '������֬,����֬����,Ѫ��,ѪС��ճ����,����Ѫ˨����,��ϸ����Ӿ,ѪС��ۼ���,'+
                      '����Ѫ˨����,�������,ȫѪ���л�ԭճ��,����Ӧ��,��ϸ����Ӿָ��,ȫѪ�������ָ��,��ϸ������ '+
                      ' from TestData '+
                      ' where format(����,''YYYY-MM-DD'')='''+FormatDateTime('YYYY-MM-DD',DateTimePicker1.Date)+''' ';
  adotemp22.Open;
  while not adotemp22.Eof do
  begin
    adotemp33:=tadoquery.Create(nil);
    adotemp33.Connection:=ADOConn_BS;
    adotemp33.Close;
    adotemp33.SQL.Clear;
    adotemp33.SQL.Text:='select count(*) as RecNum from Visc where TestDataID='+adotemp22.fieldbyname('TestDataID').AsString;
    adotemp33.Open;
    RecNum:=adotemp33.fieldbyname('RecNum').AsInteger;
    adotemp33.Free;
  
    ReceiveItemInfo:=VarArrayCreate([0,38+RecNum-1],varVariant);
    
    adotemp:=tadoquery.Create(nil);
    adotemp.Connection:=ADOConn_BS;
    adotemp.Close;
    adotemp.SQL.Clear;
    adotemp.SQL.Text:='select ShearRate,Visc from Visc where TestDataID='+adotemp22.fieldbyname('TestDataID').AsString;
    adotemp.Open;
    i:=0;
    while not adotemp.Eof do
    begin
      ReceiveItemInfo[i]:=VarArrayof([adotemp.fieldbyname('ShearRate').AsString,adotemp.fieldbyname('Visc').AsString,'','']);
      inc(i);
      adotemp.Next;
    end;
    adotemp.Free;

    SamNo:=adotemp22.fieldbyname('���').AsString;
    sName:=adotemp22.fieldbyname('����').AsString;
    sSex:=ifThen(uppercase(adotemp22.fieldbyname('�Ա�').AsString)='TRUE','��','Ů');
    sAge:=adotemp22.fieldbyname('����').AsString;
    sKB:=adotemp22.fieldbyname('�Ʊ�').AsString;
    sBQ:=adotemp22.fieldbyname('����').AsString;
    sBLH:=adotemp22.fieldbyname('������').AsString;
    sBedNo:=adotemp22.fieldbyname('����').AsString;
    sLCZD:=adotemp22.fieldbyname('�ٴ����').AsString;
    sSJYS:=adotemp22.fieldbyname('�ͼ�ҽ��').AsString;
    sJYYS:=adotemp22.fieldbyname('����ҽ��').AsString;
      
    ReceiveItemInfo[0+i]:=VarArrayof(['ȫѪճ��',adotemp22.fieldbyname('ȫѪճ��').AsString,'','']);
    ReceiveItemInfo[1+i]:=VarArrayof(['Ѫ��ճ��',adotemp22.fieldbyname('Ѫ��ճ��').AsString,'','']);
    ReceiveItemInfo[2+i]:=VarArrayof(['ѹ��',adotemp22.fieldbyname('ѹ��').AsString,'','']);
    ReceiveItemInfo[3+i]:=VarArrayof(['Ѫ��',adotemp22.fieldbyname('Ѫ��').AsString,'','']);
    ReceiveItemInfo[4+i]:=VarArrayof(['Ѫ����������',adotemp22.fieldbyname('Ѫ����������').AsString,'','']);
    ReceiveItemInfo[5+i]:=VarArrayof(['Ѫ����������ʱ��',adotemp22.fieldbyname('Ѫ����������ʱ��').AsString,'','']);
    ReceiveItemInfo[6+i]:=VarArrayof(['ȫѪ�������ָ��',adotemp22.fieldbyname('ȫѪ�������ָ��').AsString,'','']);
    ReceiveItemInfo[7+i]:=VarArrayof(['ȫѪ�������ָ��',adotemp22.fieldbyname('ȫѪ�������ָ��').AsString,'','']);
    ReceiveItemInfo[8+i]:=VarArrayof(['Ѫ������Kֵ',adotemp22.fieldbyname('Ѫ������Kֵ').AsString,'','']);
    ReceiveItemInfo[9+i]:=VarArrayof(['��ϸ���ۼ�ָ��',adotemp22.fieldbyname('��ϸ���ۼ�ָ��').AsString,'','']);
    ReceiveItemInfo[10+i]:=VarArrayof(['��ϸ���ۼ�ϵ��',adotemp22.fieldbyname('��ϸ���ۼ�ϵ��').AsString,'','']);
    ReceiveItemInfo[11+i]:=VarArrayof(['��ϸ������ָ��',adotemp22.fieldbyname('��ϸ������ָ��').AsString,'','']);
    ReceiveItemInfo[12+i]:=VarArrayof(['ȫѪ���л�ԭճ��',adotemp22.fieldbyname('ȫѪ���л�ԭճ��').AsString,'','']);
    ReceiveItemInfo[13+i]:=VarArrayof(['ȫѪ���л�ԭճ��',adotemp22.fieldbyname('ȫѪ���л�ԭճ��').AsString,'','']);
    ReceiveItemInfo[14+i]:=VarArrayof(['��ϸ������ָ��TK',adotemp22.fieldbyname('��ϸ������ָ��TK').AsString,'','']);
    ReceiveItemInfo[15+i]:=VarArrayof(['��ϸ������ָ��',adotemp22.fieldbyname('��ϸ������ָ��').AsString,'','']);
    ReceiveItemInfo[16+i]:=VarArrayof(['����ճ��',adotemp22.fieldbyname('����ճ��').AsString,'','']);
    ReceiveItemInfo[17+i]:=VarArrayof(['Ѫ�쵰��',adotemp22.fieldbyname('Ѫ�쵰��').AsString,'','']);
    ReceiveItemInfo[18+i]:=VarArrayof(['��ϸ����ճ��',adotemp22.fieldbyname('��ϸ����ճ��').AsString,'','']);
    ReceiveItemInfo[19+i]:=VarArrayof(['��������',adotemp22.fieldbyname('��������').AsString,'','']);
    ReceiveItemInfo[20+i]:=VarArrayof(['��������',adotemp22.fieldbyname('��������').AsString,'','']);
    ReceiveItemInfo[21+i]:=VarArrayof(['��������',adotemp22.fieldbyname('��������').AsString,'','']);
    ReceiveItemInfo[22+i]:=VarArrayof(['��ά����ԭ',adotemp22.fieldbyname('��ά����ԭ').AsString,'','']);
    ReceiveItemInfo[23+i]:=VarArrayof(['Ѫ���̴�',adotemp22.fieldbyname('Ѫ���̴�').AsString,'','']);
    ReceiveItemInfo[24+i]:=VarArrayof(['������֬',adotemp22.fieldbyname('������֬').AsString,'','']);
    ReceiveItemInfo[25+i]:=VarArrayof(['����֬����',adotemp22.fieldbyname('����֬����').AsString,'','']);
    ReceiveItemInfo[26+i]:=VarArrayof(['Ѫ��',adotemp22.fieldbyname('Ѫ��').AsString,'','']);
    ReceiveItemInfo[27+i]:=VarArrayof(['ѪС��ճ����',adotemp22.fieldbyname('ѪС��ճ����').AsString,'','']);
    ReceiveItemInfo[28+i]:=VarArrayof(['����Ѫ˨����',adotemp22.fieldbyname('����Ѫ˨����').AsString,'','']);
    ReceiveItemInfo[29+i]:=VarArrayof(['��ϸ����Ӿ',adotemp22.fieldbyname('��ϸ����Ӿ').AsString,'','']);
    ReceiveItemInfo[30+i]:=VarArrayof(['ѪС��ۼ���',adotemp22.fieldbyname('ѪС��ۼ���').AsString,'','']);
    ReceiveItemInfo[31+i]:=VarArrayof(['����Ѫ˨����',adotemp22.fieldbyname('����Ѫ˨����').AsString,'','']);
    ReceiveItemInfo[32+i]:=VarArrayof(['�������',adotemp22.fieldbyname('�������').AsString,'','']);
    ReceiveItemInfo[33+i]:=VarArrayof(['ȫѪ���л�ԭճ��',adotemp22.fieldbyname('ȫѪ���л�ԭճ��').AsString,'','']);
    ReceiveItemInfo[34+i]:=VarArrayof(['����Ӧ��',adotemp22.fieldbyname('����Ӧ��').AsString,'','']);
    ReceiveItemInfo[35+i]:=VarArrayof(['��ϸ����Ӿָ��',adotemp22.fieldbyname('��ϸ����Ӿָ��').AsString,'','']);
    ReceiveItemInfo[36+i]:=VarArrayof(['ȫѪ�������ָ��',adotemp22.fieldbyname('ȫѪ�������ָ��').AsString,'','']);
    ReceiveItemInfo[37+i]:=VarArrayof(['��ϸ������',adotemp22.fieldbyname('��ϸ������').AsString,'','']);

    if bRegister then
    begin
      FInts :=CreateOleObject('Data2LisSvr.Data2Lis');
      FInts.fData2Lis(ReceiveItemInfo,rightstr('0000'+SamNo,4),
        FormatDateTime('YYYY-MM-DD',DateTimePicker1.Date)+' '+FormatDateTime('hh:nn:ss',adotemp22.fieldbyname('ʱ��').AsDateTime),
        (GroupName),(SpecType),(SpecStatus),(EquipChar),
        (CombinID),
        sName+'{!@#}'+sSex+'{!@#}{!@#}'+sAge+'{!@#}'+sBLH+'{!@#}'+sKB+'{!@#}'+sSJYS+'{!@#}'+sBedNo+'{!@#}'+sLCZD+'{!@#}{!@#}'+sJYYS,
        (LisFormCaption),(ConnectString),
        (QuaContSpecNoG),(QuaContSpecNo),(QuaContSpecNoD),'',
        true,true,'����');
      if not VarIsEmpty(FInts) then FInts:= unAssigned;
    end;

    adotemp22.Next;
  end;
  adotemp22.Free;
  
  (sender as TBitBtn).Enabled:=true;
end;

initialization
    hnd := CreateMutex(nil, True, Pchar(ExtractFileName(Application.ExeName)));
    if GetLastError = ERROR_ALREADY_EXISTS then
    begin
        MessageBox(application.Handle,pchar('�ó������������У�'),
                    'ϵͳ��ʾ',MB_OK+MB_ICONinformation);
        Halt;
    end;

finalization
    if hnd <> 0 then CloseHandle(hnd);

end.
