unit FrmUtils;

interface

uses
  Windows,Classes,CommonUtils,System.SysUtils,Graphics,Variants, forms,
  Controls,Dialogs,ExtCtrls,StdCtrls,StrUtils,System.UITypes,IniFiles;

const
  ob='[';
  cb=']';

  {Conversion and size unit}
  cInch=2.54;
  cPPU=96;         {Pixels in 1 inch}
  cCMU=37.795275;  {Pixels in 1 cm}
  cMaxDigit=7;

resourcestring
  rsFYI='For your information...';
  rsErrConvertColor='Error converting Color';
  rsWantSave='Would you like to save changes';

type
  TSetUnits=(uNotSet,uCm,uInch,uPixel);
  TWeight=(LB,KG,KL);

  TBorderRec=record
  Public
    Color: TColor;
    Thickness: Integer;
    Style: TPenStyle;
    procedure Assign(r: TBorderRec);
    procedure Reset;
  end;
  PBorderRec=^TBorderRec;

  RHTMLCom=record
    sName       : string;
    sDBField    : string;
    sTable      : string;
    sType       : string;
    sCaption    : string;
    sSize       : string;
    sValue      : string;
    sHelp       : string;
    sClass      : string;
    sOptions    : array of string;
    bFlag       : boolean;
    bAutoSubmit : boolean;
    bReadOnly   : boolean;
    sSubmit     : string;
    sHotIdx     : string;
  end;

procedure AdjustToDesktop(form: Tform);
procedure LoadPosition(ini: TCustomIniFile; sSection:String; aform: Tform);
procedure SavePosition(ini: TCustomIniFile; sSection:String; aform: Tform);
procedure BroadcastToAllforms(aMessage: Cardinal;
                              aWParam,aLParam: Longint;
                              Excludeform: Tform;
                              aSync: Boolean=True);
function  CheckToSave(b1: Boolean; sFileName:String=''): Integer;
function  CleanDesc(sDesc: String): String;
function  Cleanpx(sValue: String): Integer;
function  ColorToHTML(ColorCode: TColor): String;
function  ColorToHTML2(ColorCode: TColor): String;
function  DeleteFileRecover(sFileName: String; bRecoverable: boolean): boolean;
function  ExPos(const SubStr,Str: UnicodeString; Offset: Integer=0): Integer;
function  getAtDbHot(sMacro:string):string;
function  GetDirectories(sDir: String): TStringList;
function  GetFieldDetail(sFilename,sTable,sField: String; iField,iType: integer): RHTMLCom;
function  GetFieldDetail3(sFilename,sTable,sField: String; iField,iType: Integer): RHTMLCom;
function  GetFieldDetail2(sFilename,sTable,sField,sDN,sCgi: String; iField,iType: Integer): String;
function  GetFieldsAndDesc(sFilename,sTable: String): String;
function  GetFieldsAndDesc2(sFilename,sTable,sOrder,sType: String): String;
function  GetFieldsFromFile(sFilename,sTable: String;
                            iField: Integer): String;
function  GetOpenWithFile: String;
function  GetParam(aParam: String; bLowerCase: Boolean=False): String;
function  GetPrm(sName,sList: String): String;
function  GetTablesFromFile(sFileName: string): String;
function  GetX(angle: Double; x,y: Integer): Integer;
function  GetY(Angle: Double; x,y: Integer): Integer;
function  HTMLToColor(sHtmlColor: String): TColor;
function  HTMLToColor2(sHtmlColor: String): TColor;
function  Inch2Dec(sInch:string): String;
function  MeasureToMetric(vValue: Variant; uUnit: TSetUnits): Variant;
function  MultRect(r1: TRect; z1: Double): TRect;
function  NumFloatVal(sValue: String): Double;
function  NumVal(aString: String; Sender: TObject): Double;
function  PageSizeCmToInch(sValue: String): String;
function  PageSizeInchToCm(sValue: String): String;
function  PixelToMetric(iValue: Integer; uUnit: TSetUnits): Double; inline;
function  ProcessExecute2(CommandLine: String; StopExec: Boolean): Integer;
procedure ReplaceTxt(Var str: TStringList; sOld,sNew: String);
function  SwapSlash(sStr: String): String; inline;
procedure UpdateException(Self: TObject;
                          E: Exception;
                          aName: String;
                          Extra: String='');  inline; overload;
procedure UpdateException(Self: String;
                          E: Exception;
                          aName: String;
                          Extra: String=''); inline; overload;
function  VerifyMultiTabs(sTable: String): String;

procedure GetResourceVersionNumbers(out AMajor,AMinor,ARelease,ABuild: Integer; Out aPrivate: Boolean);
procedure GetFileVersionNumbers(out AMajor,AMinor,ARelease,ABuild: Integer);

function ProtectMailSlotUrl(sMsg:string):string;
function SetRelativePath(sPath,sFile:string):string;
function cleanHiddenSpace(sHtml:string):string;
procedure populateMultiRecTabs(ini : TIniFile);
function revPos(Const subStr, str:String; iPos: Integer): Integer;
var
  aMultiRecTabs: array[0..20] of String;

implementation

uses
  Math,Types,jpeg,GIFImg,pngimage,ShellAPI;

{ *****************************************************************************}
{if a form is away from the desktop,this function will move the form back on it}
{Use this function to validate positioning when coordinates are read from configuration}
{in case the hardware changes or some invalid value was given that pushes the form out}
{of the visible range.}
procedure AdjustToDesktop(form: Tform);
begin
  with screen do
   begin
     if form.left+form.width>DesktopLeft+DesktopWidth then
       form.left:=DesktopLeft+DesktopWidth-form.width;
     if form.left<DesktopLeft then
       form.left:=DesktopLeft;
     if form.top+form.height>DesktopTop+DesktopHeight then
       form.top:=DesktopTop+DesktopHeight-form.Height;
     if form.top<DesktopTop then
       form.top:=DesktopTop;
   end;
end;

{***************************************************************}
procedure SavePosition(ini: TCustomIniFile; sSection:String; aform: Tform);
begin
  if aform.WindowState=wsNormal then
   begin
   //windowlayout
     Ini.WriteString(sSection,'maximized','0');
     Ini.WriteInteger(sSection,'mainleft',aform.Left);
     Ini.WriteInteger(sSection,'maintop',aform.Top);
     Ini.WriteInteger(sSection,'mainwidth',aform.Width);
     Ini.WriteInteger(sSection,'mainheight',aform.Height);
   end
  else
   begin
     Ini.WriteString(sSection,'maximized','1');
     Ini.WriteInteger(sSection,'mainleftmax',aform.Left);
     Ini.WriteInteger(sSection,'maintopmax',aform.Top);
   end;
end;

procedure LoadPosition(ini: TCustomIniFile; sSection:String; aform: Tform);
begin
  aform.WindowState:=wsNormal;
  aform.Width:=Ini.ReadInteger(sSection,'mainwidth',600);
  aform.Height:=Ini.ReadInteger(sSection,'mainheight',500);
  if (Ini.ReadString(sSection,'maximized','1')='1') then
   begin
    aform.Left:=Ini.ReadInteger(sSection,'mainleftmax',100);
    aform.Top:=Ini.ReadInteger(sSection,'maintopmax',100);
    aform.WindowState:=wsMaximized;
  end
  else
   begin
     aform.Left:=Ini.ReadInteger(sSection,'mainleft',100);
     aform.Top:=Ini.ReadInteger(sSection,'maintop',100);
   end;
  AdjustToDesktop(aform);
end;

procedure BroadcastToAllforms(aMessage: Cardinal;
                              aWParam,aLParam: Longint;
                              Excludeform: Tform;
                              aSync: Boolean=True);
var
  I1: Integer;
begin
  for I1:=0 to Screen.formCount-1 do
   if Screen.forms[I1]<>Excludeform then
    if Async then
      PostMessage(Screen.forms[I1].Handle,aMessage,aWParam,aLParam)
    else
      Screen.forms[I1].Perform(aMessage,aWParam,aLParam);
end;

function CheckToSave(b1: Boolean; sFileName:String=''): Integer;
var
  s1: String;
begin
  Result:=0;
  if b1 then
   begin
     s1:=rsWantSave;
     if sFileName<>'' then s1:=s1+' on '+sFileName+' file';
     Result:=MessageDlg(s1+'?',mtConfirmation,[mbYes,mbNo,mbCancel],0);
   end;
end;

function CleanDesc(sDesc: String): String;
var
  i1: Integer;
  s1: String;
begin
  for i1:=1 to Length(sDesc) do
   begin
     s1:=sDesc[i1];
     if not CharInSet(sDesc[i1],['A'..'Z']) and not CharInSet(sDesc[i1],['a'..'z']) and (i1>0) then
       sDesc[i1]:='_';
   end;
  Result:=sDesc;
end;

function Cleanpx(sValue: String): Integer;
var
  i1:    Integer;
  s1:    String;
  iCode: Integer;
begin
  cleanpx:=0;
  s1:='';
  {loop through String}
  for i1:=1 to length(sValue) do
    if (sValue[i1] >= '0') and (sValue[i1] <= '9') then
      s1:=s1+sValue[i1];

  Val(s1,i1,iCode);
  if (iCode<>0) then
    s1:='0'
  else
    Result:=i1;
end;

function CmToInch(sValue: String): String;
var
  rValue: Extended;
  buffer: array[0..64] of Char;
  nCount: Integer;
begin
  try
    rValue:=AnyStrToFloat(sValue)/2.54;
    nCount:=FloatToText(Buffer,rValue,fvExtended,ffFixed,7,2);
    Buffer[nCount]:=#0;
    Result:=Buffer;
  except
    Result:='0'; {in case the value can not be interpreted}
  end;
end;

function ColorToHTML(ColorCode: TColor): String;
var
  R,G,B: String;
begin
  try
    case ColorCode of
      clwhite   : Result:='FFFFFF';
      clBlack   : Result:='000000';
      clRed     : Result:='FF0000';
      clGreen   : Result:='008000';
      clBlue    : Result:='0000ff';
      clNavy    : Result:='00007F';
      clGray    : Result:='7F7F7F';
      clSilver  : Result:='BFBFBF';
      clYellow  : Result:='FFFF00';
      clLime    : Result:='00FF00';
      clAqua    : Result:='00FFFF';
      clMaroon  : Result:='7F0000';
      clOlive   : Result:='7F7F00';
      clPurple  : Result:='7F007F';
      clFuchsia : Result:='FF00FF';
      clTeal    : Result:='007F7F';
      clWindowText : Result:='000000';
    else
      R:=Copy(colorToString(colorcode),8,2);
      G:=Copy(colorToString(colorcode),6,2);
      B:=Copy(colorToString(colorcode),4,2);
      Result:=R+G+B;
    end;
  except
  end;
end;

function colorToHTML2(ColorCode: TColor): String;
begin
  try
    case ColorCode of
      clwhite   : Result:='FFFFFF';
      clBlack   : Result:='000000';
      clRed     : Result:='FF0000';
      clGreen   : Result:='008000';
      clBlue    : Result:='0000ff';
      clNavy    : Result:='00007F';
      clGray    : Result:='7F7F7F';
      clSilver  : Result:='BFBFBF';
      clYellow  : Result:='FFFF00';
      clLime    : Result:='00FF00';
      clAqua    : Result:='00FFFF';
      clMaroon  : Result:='7F0000';
      clOlive   : Result:='7F7F00';
      clPurple  : Result:='7F007F';
      clFuchsia : Result:='FF00FF';
      clTeal    : Result:='007F7F';
      clWindowText : Result:='000000';
    else
      Result:=ColorToString(ColorCode);
    end;
  except
    MessageDlg(rsErrConvertColor,mtError,[mbOk],0);
  end;
end;

function getFieldDetail3(sFilename,sTable,sField: String; iField,iType: Integer): RHTMLCom;
var
  dBackup: TextFile;
  sBackup: String;
  s1:      String;
  rField:  RHTMLCom;
  sTmpTab: String;
begin
  try
    if Pos('_lup',LowerCase(sTable))>0 then
      sTmpTab:=Copy(sTable,1,length(sTable)-3)+'TAB'
    else
      sTmpTab:=sTable;

    AssignFile(dBackup,sFilename);
    {$I-} Reset(dBackup); {$I+}
    if IOResult=0 then
     begin
       while not eof(dBackup) do
        begin
          ReadLn(dBackup,sBackup);
          s1:=Copy(sBackup,2,Length(sBackup)-2);
          s1:=ComaParam(s1,iField);

          if (Pos(UpperCase(sTmpTab),UpperCase(sBackup))>0) and
             (LowerCase(sField)=LowerCase(s1)) then
           begin
             sBackup:=copy(sBackup,2,length(sBackup)-2);
             rField.sDBField :=ComaParam(sBackup,2);

             rField.sTable   :=sTable;
             rField.sName    :=sTable+ob+rField.sDBField+cb;
             rField.sCaption :=ComaParam(sBackup,3);
             rField.sType    :=ComaParam(sBackup,7);

             if (iType=1) then
               rField.sValue   :='@dbRow(Q2,'''+ComaParam(sBackup,2)+''''
             else
               rField.sValue   :='@dbHot('+sTable+',' +ComaParam(sBackup,2);

             if (LowerCase(rField.sType)='dtinteger') then
              begin
                rField.sType:='qty';
                rField.sValue:=rField.sValue+',Z#)';
                rField.sClass:='EntNum';
              end
             else
             if (LowerCase(rField.sType)='dtdouble') then
              begin
                rField.sType:='num';
                rField.sValue:=rField.sValue+',Z#)';
                rField.sClass:='EntNum';
              end
             else
             if (LowerCase(rField.sType)='dtcurrency') then
              begin
                rField.sType:='cur';
                rField.sValue:=rField.sValue+',Z$)';
                rField.sClass:='EntCur';
              end
             else
             if (LowerCase(rField.sType)='dtdatetime') then
              begin
                rField.sType:='date';
                rField.sValue:=rField.sValue+',DWDF)';
                rField.sClass:='EntDat';
              end
             else
              begin
                if (LowerCase(rField.sType)='dtstring') then
                  rField.sType:='text';
                rField.sValue:=rField.sValue+')';
                rField.sClass:='EntTxt';
              end;
           end;
        end;
       CloseFile(dBackup);
     end;
  finally
    Result:=rField;
  end;
end;

function DeleteFileRecover(sFileName: String; bRecoverable: boolean): boolean;
var
  SHFileOpStruct: TSHFileOpStruct;
begin
  SHFileOpStruct.Wnd:=0;
  SHFileOpStruct.wFunc:=FO_DELETE;
  SHFileOpStruct.pFrom:=PChar(sFileName+#0+#0);
  SHFileOpStruct.pTo:=nil;
  SHFileOpStruct.fFlags:=FOF_SILENT or FOF_NOCONFIRMATION;
  if bRecoverable then
    SHFileOpStruct.fFlags:=SHFileOpStruct.fFlags or FOF_ALLOWUNDO;
  SHFileOpStruct.hNameMappings:=nil;
  SHFileOpStruct.lpszProgressTitle:=nil;
  Result:=SHFileOperation(SHFileOpStruct)=0;
end;

{ ExcludePos
Search a caracter that is not included in a delimited String
Ex.: ExPos('(','1,2,3,4,'')'',5)') will return the position of the ) at the end
not the one in '')''}
function ExPos(const SubStr,Str: UnicodeString; Offset: Integer=0): Integer;
var
  b1: Boolean;
  i1: Integer;
  iLen: Integer;  {Lenght}
begin
  i1:=Max(1,Offset);
  b1:=False;
  iLen:=Length(Str);
  Result:=0;

  while i1 <= iLen do
   begin
     if (str[i1]='''') or (str[i1]='"') then
      b1:=not b1
     else
     if not b1 and (str[i1]=SubStr) then
      begin
        Result:=i1;
        break;
      end;

     Inc(i1);
   end;
end;

// *******************************************  Get @dbHot  ********************
function getAtDbHot(sMacro:string):string;
begin
  Result:=Copy(sMacro,Pos('@dbhot(',LowerCase(sMacro)),Length(sMacro));
  Result:=Copy(Result,1,Pos(')',Result));
end;

function getDirectories(sDir: String): TStringList;
var
  asRec:    TSearchRec;
  DosError: Integer;
  s1:       TStringList;
begin
  s1:=TStringList.create;
  DosError:=FindFirst(sDir,FaDirectory,asRec);
  while DosError=0 do
   begin
     if (asRec.name<>'.') and (asRec.name<>'..') then
       s1.Add(asRec.name);
     DosError:=FindNext(asRec);
   end;
  GetDirectories:=s1;
end;

function GetFieldsAndDesc(sFilename,sTable: String): String;
var
  dBackup: TextFile;
  sBackup: String;
  s1:      TStringList;
begin
  s1:=TStringList.Create;
  try
    s1.text:='';
    AssignFile(dBackup,sFilename);
    {$I-} Reset(dBackup); {$I+}
    if IOResult=0 then
     begin
       while not EOF(dBackup) do
        begin
          ReadLn(dBackup,sBackup);
          if Pos(UpperCase(sTable),UpperCase(sBackup))>0 then
           begin
             sBackup:=Copy(sBackup,2,Length(sBackup)-2);
             s1.Add(ComaParam(sBackup,2)+'='+ComaParam(sBackup,3));
           end;
        end;
       CloseFile(dBackup);
     end;
  finally
    Result:=s1.Text;
    s1.Free;
  end;
end;

function GetFieldsAndDesc2(sFilename,sTable,sOrder,sType: String): String;
var
  dBackup: TextFile;
  sBackup: String;
  s1:      TStringList;
begin
  s1:=TStringList.create;
  try
    s1.text:='';
    AssignFile(dBackup,sFilename);
    {$I-} Reset(dBackup); {$I+}
    if IOResult=0 then
     begin
       while not EOF(dBackup) do
        begin
          ReadLn(dBackup,sBackup);
          if Pos(''''+UpperCase(sTable)+'''',UpperCase(sBackup))>0 then
           begin
             sBackup:=Copy(sBackup,2,Length(sBackup)-2);

             if (sType='') or (UpperCase(sType)=UpperCase(ComaParam(sBackup,7))) then
              if sOrder='' then
                s1.Add(ComaParam(sBackup,2)+'='+ComaParam(sBackup,3))
              else
                s1.Add(ComaParam(sBackup,3)+'='+ComaParam(sBackup,2));
           end;
        end;
       CloseFile(dBackup);
     end;
  finally
    Result:=s1.text;
    s1.Free;
  end;
end;

function GetFieldDetail(sFilename,sTable,sField: String; iField,iType: Integer): RHTMLCom;
var
  dBackup: TextFile;
  sBackup: String;
  s1:      String;
  rField:  RHTMLCom;
begin
  AssignFile(dBackup,sFilename);
  {$I-} Reset(dBackup); {$I+}
  if IOResult=0 then
   begin
     while not EOF(dBackup) do
      begin
        ReadLn(dBackup,sBackup);
        s1:=Copy(sBackup,2,Length(sBackup)-2);
        s1:=ComaParam(s1,iField);

        if (Pos(UpperCase(sTable),UpperCase(sBackup))>0) and
           (LowerCase(sField)=LowerCase(s1)) then
         begin
             sBackup:=Copy(sBackup,2,Length(sBackup)-2);
             rField.sDBField:=ComaParam(sBackup,2);
             rField.sTable  :=sTable;
             rField.sName   :=sTable+ob+rField.sDBField+cb;
             rField.sCaption:=ComaParam(sBackup,3);
             rField.sType   :=ComaParam(sBackup,7);


             if (iType=1) then
               rField.sValue:='@dbRow(Q2,'''+ComaParam(sBackup,2)+''''
             else
               rField.sValue:='@dbHot('+sTable+',' +sTable+ob+ComaParam(sBackup,2)+cb;

             if (LowerCase(rField.sType)='dtinteger') then
              begin
                rField.sType:='qty';
                rField.sValue:=rField.sValue+',Z#)';
                rField.sClass:='EntNum';
              end
             else
             if (LowerCase(rField.sType)='dtdouble') then
              begin
                rField.sType:='num';
                rField.sValue:=rField.sValue+',Z#)';
                rField.sClass:='EntNum';
              end
             else
             if (LowerCase(rField.sType)='dtcurrency') then
              begin
                rField.sType:='cur';
                rField.sValue:=rField.sValue+',Z$)';
                rField.sClass:='EntCur';
              end
             else
             if (LowerCase(rField.sType)='dtdatetime') then
              begin
                rField.sType:='date';
                rField.sValue:=rField.sValue+')';
                rField.sClass:='EntDat';
              end
             else
             if (lowercase(rField.sType)='dtstring') then
             begin
               rField.sType:='text';
               rField.sValue:=rField.sValue+',L)';
               rField.sClass:='EntTxt';
             end;
         end;
      end;
     CloseFile(dBackup);
   end;

  Result:=rField;
end;

function GetFieldDetail2(sFilename,sTable,sField,sDN,sCgi: String; iField,iType: Integer): String;
var
  dBackup:  TextFile;
  sBackup:  String;
  s1:       String;
  sResult:  String;
  sDbField: String;
  sName:    String;
  sType:    String;
  sValue:   String;
  sClass:   String;
begin
  AssignFile(dBackup,sFilename);
  {$I-} Reset(dBackup); {$I+}
  if IOResult=0 then
   begin
     while not EOF(dBackup) do
      begin
        Readln(dBackup,sBackup);
        s1:=Copy(sBackup,2,Length(sBackup)-2);
        s1:=ComaParam(s1,iField);
        if (Pos(UpperCase(sTable),UpperCase(sBackup))>0) and
           (LowerCase(sField)=LowerCase(s1)) then
         begin
           sBackup:=Copy(sBackup,2,Length(sBackup)-2);
           sDBField:=ComaParam(sBackup,2);
           sName:=sTable+ob+sDBField+cb;
           sType:=LowerCase(ComaParam(sBackup,7));
           sTable:=VerifyMultiTabs(sTable);
           sValue:='@dbHot('+sTable+','+ComaParam(sBackup,2);
           if (LowerCase(sType)='dtinteger') or ((sType)='dtdouble') then
            begin
              sValue:=sValue+',#)';
              if (iType<>0) and (iType<>4) then sClass:='EntNum' else sClass:='EntValNum';
            end
           else
           if LowerCase(sType)='dtcurrency' then
            begin
              sValue:=sValue+',#:2)';
              if (iType<>0) and (iType<>4) then sClass:='EntNum' else sClass:='EntValNum';
            end
           else
           if (LowerCase(sType)='dtdatetime') then
            begin
              sValue:=sValue+',DWDF)';
              if (iType<>0) and (iType<>4) then sClass:='EntTxt' else sClass:='EntValTxt';
            end
           else
            begin
              sValue:=sValue+')';
              if iType=1 then
                sClass:='EntChk'
              else
              if (iType<>0) and (iType<>4) then
                sClass:='EntTxt'
              else
                sClass:='EntValTxt';
            end;
         end;
      end;
     CloseFile(dBackup);

     if iType=4 then
      begin
        if Pos('num',LowerCase(sClass))>0 then
          sResult:='<input class="'+sClass+'" value="@fmt(r10:space,'+sValue+')" type="text" readonly="readonly"/>'
        else
          sResult:='<input class="'+sClass+'" value="'+sValue+'" type="text" readonly="readonly"/>';
      end
     else
     if iType=0 then
       sResult:='<input class="'+sClass+'" value="'+sValue+'" type="text" readonly="readonly"/>'
     else
     if iType=1 then
      begin
        if AnsiSameText(sDN,'PDA') then
          sResult:='<input class="'+sClass+'" onblur="EntBlur(this)" onkeydown="EntKeyDown(this)" '+
              'onfocus="EntFocus(this)" onclick="EntClick(this,''Http:/Scripts/trs.exe?cn=@CNC&amp;tn=@TER&amp;dn=PDA&amp;cgi='+
              sCgi+'&amp;CP=0:0:0'');" value="'+sValue+'" name="'+sName+'" type="checkbox" @FMT(CMP,'''+sValue+'=1'',checked) />'
        else
          sResult:='<input class="'+sClass+'" onclick="CheckboxSubmit(this);" value="'+sValue+'" name="'+sName+'" type="checkbox"/>';
      end
     else
      begin
        if AnsiSameText(sDN,'RFC') then
          sResult:='<input class="'+sClass+'" value="'+sValue+'" name="'+sName+'" type="text"/>'

        else
        if AnsiSameText(sDN,'PDA') then
          sResult:='<input class="'+sClass+'" onblur="EntBlur(this)" onkeydown="EntKeyDown(this)" onfocus="EntFocus(this)" value="' +
              sValue+'" name="'+sName+'" type="text"/>'
        else
          sResult:='<input class="'+sClass+'" datasrc="@attrib(E)" value="'+sValue+'" name="'+sName+'" type="text"/>';
      end;
   end;
  Result:=sResult;
end;

function getFieldsFromFile(sFilename,sTable: String; iField: Integer): String;
var
  dBackup: TextFile;
  sBackup: String;
  s1:      TStringList;
begin
  getFieldsFromFile:='';
  try
    s1:=TStringList.Create;
    try
      s1.text:='';
      AssignFile(dBackup,sFilename);
      {$I-} Reset(dBackup); {$I+}
      if IOResult=0 then
       begin
         while not EOF(dBackup) do
          begin
            ReadLn(dBackup,sBackup);
            if Pos(UpperCase(sTable),UpperCase(sBackup))>0 then
             begin
               sBackup:=Copy(sBackup,2,Length(sBackup)-2);
               s1.add(ComaParam(sBackup,iField));
             end;
          end;
         CloseFile(dBackup);
       end;
    finally
      Result:=s1.text;
      s1.Free;
    end;
  except
  end;
end;

function getOpenWithFile: String;
var
  i1: Integer;
  s1,s2: String;
begin
  s1:='';
  try
    i1:=0;

    while i1 <= paramcount do
     begin
       s2:=ParamStr(i1);
       if (FileExists(s2)) and (Pos('.exe',LowerCase(ParamStr(i1)))=0) then
        begin
          s1:=ParamStr(i1);
          i1:=ParamCount+1;
        end
       else
         Inc(i1);
     end;
  finally
    Result:=s1;
  end;
end;

function GetParam(aParam: String; bLowerCase: Boolean=False): String;
var
  i1: Integer;
begin
  Result:='';
  for i1:=0 to ParamCount do
   begin
    if Pos(AnsiLowerCase(aParam),AnsiLowerCase(ParamStr(i1)))>0 then
     begin
       Result:=Copy(ParamStr(i1),Pos('=',ParamStr(i1))+1,Length(ParamStr(i1)));
       if bLowerCase then Result:=AnsiLowerCase(Result);
     end;
   end;
end;

function getPrm(sName,sList: String): String;
var
  i1: Integer;
  s1: String;
begin
  s1:='';
  try
    i1:=Pos(LowerCase(sName)+'=',LowerCase(sList));
    if i1>0 then
     begin
       s1:=Copy(sList,i1+Length(sName)+1,Length(sList));
       i1:=Pos(',',s1);
       if i1>0 then
        s1:=Copy(s1,1,i1-1);
     end;
  finally
    Result:=s1;
  end;
end;

function GetTablesFromFile(sFileName: String): String;
var
  dBackup: TextFile;
  sBackup: String;
  s1:      TStringList;
begin
  s1:=TStringList.Create;
  try
    s1.text:='';
    AssignFile(dBackup,sFileName);
    {$I-} Reset(dBackup); {$I+}
    if IOResult=0 then
     begin
       while not EOF(dBackup) do
        begin
          ReadLn(dBackup,sBackup);
          if Copy(sBackup,0,1)='(' then
           begin
             sBackup:=Copy(sBackup,2,Length(sBackup)-2);
             s1.Add(UpperCase(ComaParam(sBackup,1)));
           end;
        end;
       CloseFile(dBackup);
     end;
  finally
    Result:=s1.Text;
    FreeAndNil(s1);
  end;
end;

function GetX(angle: Double; x,y: Integer): Integer;
begin
  Result:=Floor(x*Cos(angle)-y*sin(angle));
end;

function GetY(Angle: Double; x,y: Integer): Integer;
begin
  Result:=Floor(y*Cos(angle)+x*sin(angle));
end;

function HTMLToColor(sHtmlColor: String): TColor;
var
  r,g,b: String;
begin
  HTMLToColor:=$00FFFFFF;
  try
    if Length(sHTMLColor)=0 then Exit;
    sHTMLColor:=Copy(sHTMLColor,2,255);
    r:=copy(sHTMLColor,1,2);
    g:=copy(sHTMLColor,3,2);
    b:=copy(sHTMLColor,5,2);
    Result:=StringToColor('$'+b+g+r);
  except
  end;
end;

function HTMLToColor2(sHtmlColor: String): TColor;
begin
  HtmlToColor2:=$00FFFFFF;
  try
    if Length(sHTMLColor)=0 then Exit;
    sHTMLColor:=Copy(sHTMLColor,2,255);
    Result:=StringToColor('$'+sHtmlcolor+'00');
  except
  end;
end;

function InchToCm(sValue: String): String; overload;
var
  rValue: Extended;
  buffer: array[0..64] of Char;
  nCount: Integer;
begin
  try
    rValue:=AnyStrToFloat(sValue)*2.54;
    nCount:=FloatToText(buffer,rValue,fvExtended,ffFixed,7,2);
    Buffer[nCount]:=#0;
    Result:=Buffer;
  except
    Result:='0'; {in case the value cannot be interpreted}
  end;
end;

function InchToCm(sValue: Double): String;  Overload;
var
  rValue: Extended;
begin
  try
    rValue:=sValue*2.54;
    Result:=FloatToStr(rValue);
  except
    Result:='0'; {in case the value cannot be interpreted}
  end;
end;

function Inch2Dec(sInch: String): String;
Var
  iSpace:  Integer;
  iSlash:  Integer;
  sResult: String;
  iNum:    Integer;
  sNum:    String;
  iDom:    Integer;
  sDom:    String;
  rResult: Double;
begin
  try
    sResult:=sInch;
    iSlash:=Pos('/',sInch);

    if (iSlash>0) then
     begin
       {get complete inchs}
       iSpace:=Pos(' ',sInch);
       sResult:=Copy(sInch,1,iSpace-1)+'.';

       {get fruction}
       sNum:=Trim(Copy(sInch,iSpace+1,iSlash-iSpace-1));
       sDom:=Trim(Copy(sInch,iSlash+1,Length(sInch)));
       if (sNum='') Or (sDom='') then Exit;
       iNum:=StrToInt(sNum);
       iDom:=StrToInt(sDom);

       {caluclation}
       rResult:=(iNum/iDom)+StrToFloat(sResult);
       sResult:=FloatToStrF(rResult,ffFixed,4,cMaxDigit);
     end;

    Result:=sResult;
  except
  end;
end;

function MeasureToMetric(vValue: Variant; uUnit: TSetUnits): Variant;
begin
  if VarIsType(vValue,VarInteger) then
   begin
     if uUnit=uCM then
       Result:=vValue/cCMu
     else
     if uUnit=uPixel then
       Result:=vValue/cPPU
     else
       Result:=vValue/cPPU;
   end
  else
  if VarIsType(vValue,VarUString) then
   begin
     if uUnit=uCM then
       Result:=Double(vValue)/cInch
     else
     if uUnit=uPixel then
       Result:=Integer(vValue)/cPPU
     else
       Result:=vValue;
   end
  else
   if uUnit=uCM then
     Result:=vValue*cInch
   else
   if uUnit=uPixel then
     Result:=Double(vValue)*cPPU
   else
     Result:=vValue;
end;

function MultRect(r1: TRect; z1: Double): TRect;
begin
  if z1=1.0 then
    Result:=r1
  else
   begin
     Result.Left:=Round(r1.Left*z1);
     Result.Top:=Round(r1.Top*z1);
     Result.Width:=Round(r1.Width*z1);
     Result.Height:=Round(r1.Height*z1);
   end;
end;

function NumFloatVal(sValue: String): Double;
var
  iCode: Integer;
  r:     Double;
begin
  Val(sValue,r,iCode);
  if iCode=0 then
    Result:=r
  else
    Result:=-100;
end;

function NumVal(aString: String; Sender: TObject): Double;
var
  i: Integer;
  r: Double;
begin
  NumVal:=0;
  try
    Val(aString,r,i);
    if i=0 then
      Result:=r
    else
    if (Sender is TEdit) then
     begin
       (Sender as TEdit).Text:=aString;
       (Sender as TEdit).SetFocus;
     end;
  except
    on E: Exception do
     begin
       UpdateException('FrmUtils',E,'HandleMoved');
       Raise;
     end;
  end;
end;

function PageSizeCmToInch(sValue: String): String;
var
  sWidth: String;
  sHeight: String;
begin
  sWidth:=Copy(sValue,0,Pos('X',UpperCase(sValue))-1);
  sHeight:=Copy(sValue,Length(sWidth)+2,Length(sValue));
  sWidth:=CmToInch(sWidth);
  sHeight:=CmToInch(sHeight);
  Result:=sWidth+' X '+sHeight;
end;

function PageSizeInchToCm(sValue: String): String;
var
  sWidth: String;
  sHeight: String;
begin
  sWidth:=Copy(sValue,0,Pos('X',UpperCase(sValue))-1);
  sHeight:=Copy(sValue,Length(sWidth)+2,Length(sValue));
  sWidth:=InchToCm(sWidth);
  sHeight:=InchToCm(sHeight);
  Result:=sWidth+' X '+sHeight;
end;

function PixelToMetric(iValue: Integer; uUnit: TSetUnits): Double;
begin
  case uUnit of
    uCm: Result:=iValue/cCMU;
    uInch: Result :=iValue/cPPU;
  else
    Result:=iValue;
  end;
end;

{*** CREATE PROCESS IN A NORMAL WINDOW *************************************}
function ProcessExecute2(CommandLine: String; StopExec: Boolean): Integer;
var
  Rslt: LongBool;
  StartupInfo: TStartupInfo;
  ProcessInfo: TProcessInformation;
  WindowsResult: DWord;
begin
  FillChar(StartupInfo,SizeOf(TStartupInfo),0);
  with StartupInfo do
   begin
     cb:=SizeOf(TStartupInfo);
     dwFlags:=STARTF_USESHOWWINDOW;
     wShowWindow:=SW_HIDE;
   end;

  Rslt:=CreateProcessW(nil,@CommandLine[1],nil,nil,False,
          CREATE_NEW_CONSOLE or NORMAL_PRIORITY_CLASS,nil,nil,StartupInfo,ProcessInfo);

  if Rslt then
   begin
     if StopExec then
       WaitforSingleObject(ProcessInfo.hProcess,INFINITE)
     else
       WaitforInputIdle(ProcessInfo.hProcess,INFINITE);
     GetExitCodeProcess(ProcessInfo.hProcess,WindowsResult);

     CloseHandle(ProcessInfo.hProcess);
     CloseHandle(ProcessInfo.hThread);
   end
  else
    WindowsResult:=0;
  Result:=WindowsResult;
end;

procedure ReplaceTxt(Var Str: TStringList; sOld,sNew: String);
var
  i1: Integer;
begin
  i1:=0;
  sOld:=LowerCase(sOld);
  while (i1<str.count) do
   begin
     str[i1]:=ReplaceText(str[i1],sOld,sNew);
     Inc(i1);
   end;
end;

function SwapAmp(sStr: String): String;
var
  i1: Integer;
begin
  for i1:=0 to Length(sStr) Do
   if sStr[i1]='&' then
    sStr[i1]:=',';
  Result:=sStr;
end;

function SwapSlash(sStr: String): String;
begin
  Result:=ReplaceText(sStr,'\','/');
end;

procedure UpdateException(Self: TObject;
                          E: Exception;
                          aName: String;
                          Extra: String='');
begin
  E.Message:='Unit Name: '+Self.UnitName +
               ': Object: '+Self.ClassName+'-'+aName+';'#13#10+E.Message;

  if Extra<>'' then E.Message:=E.Message+': Variables: '+Extra;
end;

procedure UpdateException(Self: String;
                          E: Exception;
                          aName: String;
                          Extra: String='');
begin
  E.Message:='Unit Name: '+Self +
               ': Object: '+Self+'-'+aName+';'#13#10+E.Message;

  if Extra<>'' then E.Message:=E.Message+': Variables: '+Extra;
end;

function VerifyMultiTabs(sTable: String): String;
var
  i1: Integer;
  s1: String;
begin
  s1:=sTable;
  try
    i1:=0;
    while (i1<20) and (s1=sTable) do
     if Pos(LowerCase(aMultiRecTabs[i1]),LowerCase(sTable))>0 then
      s1:=aMultiRecTabs[i1]+'_LUP'
     else
      Inc(i1);
  finally
    Result:=s1;
  end;
end;

procedure GetResourceVersionNumbers(out AMajor,AMinor,ARelease,ABuild: Integer; Out aPrivate: Boolean);
{ Get the module version information from the embedded resource.
Called upon Initialization }
var
  HResource: TResourceHandle;
  HResData: THandle;
  PRes: Pointer;
  InfoSize: DWORD;
  FileInfo: PVSFixedFileInfo;
  FileInfoSize: DWORD;
begin
 aMajor:=0;
 aMinor:=0;
 aRelease:=0;
 aBuild:=0;

 HResource:=FindResource(HInstance,MakeIntResource(VS_VERSION_INFO),RT_VERSION);
 if HResource<>0 then
  begin
    HResData:=LoadResource(HInstance,HResource);
    if HResData<>0 then
     begin
       PRes:=LockResource(HResData);
       if Assigned(PRes) then
        begin
          InfoSize:=SizeofResource(HInstance,HResource);
          if InfoSize=0 then
           raise Exception.Create('Can''t get version information.');
          VerQueryValue(PRes,'\',Pointer(FileInfo),FileInfoSize);
          {Now fill in the version information}
          AMajor:=FileInfo.dwFileVersionMS shr 16;
          AMinor:=FileInfo.dwFileVersionMS and $FFFF;
          ARelease:=FileInfo.dwFileVersionLS shr 16;
          ABuild:=FileInfo.dwFileVersionLS and $FFFF;
          aPrivate:=Boolean(FileInfo.dwFileFlags and VS_FF_PRIVATEBUILD);
        end;
     end;
  end
 else
   RaiseLastOSError;
end;

procedure GetFileVersionNumbers(out AMajor,AMinor,ARelease,ABuild: Integer);
{ Get the file version information. Called upon Initialization }
var
  PInfo: Pointer;
  InfoSize: DWORD;
  FileInfo: PVSFixedFileInfo;
  FileInfoSize: DWORD;
  FileName: string;
  Tmp: DWORD;
begin
  aMajor:=0;
  aMinor:=0;
  aRelease:=0;
  aBuild:=0;

  FileName:=ParamStr(0);
  {Get the size of the FileVersionInformatioin}
  InfoSize:=GetFileVersionInfoSize(PChar(FileName),Tmp);
  if InfoSize=0 then Exit;

  GetMem(PInfo,InfoSize);
  try
    GetFileVersionInfo(PChar(FileName),0,InfoSize,PInfo);
    VerQueryValue(PInfo,'\',Pointer(FileInfo),FileInfoSize);
    AMajor:=FileInfo.dwFileVersionMS shr 16;
    AMinor:=FileInfo.dwFileVersionMS and $FFFF;
    ARelease:=FileInfo.dwFileVersionLS shr 16;
    ABuild:=FileInfo.dwFileVersionLS and $FFFF;
  finally
    FreeMem(PInfo);
  end;
end;



procedure TBorderRec.Assign(r: TBorderRec);
begin
  Color:=r.Color;
  Thickness:=r.Thickness;
  Style:=r.Style;
end;

procedure TBorderRec.Reset;
begin
  Color:=clBlack;
  Thickness:=0;
  Style:=psSolid;
end;

// ****************************************   Protect url for Mailslot  ********
function ProtectMailSlotUrl(sMsg:string):string;
var
  i1: integer;
  sParam: string;
  sMail: string;
  sData: string;
begin
  try
    {Protect string}
    i1:=pos('&',sMsg);
    sMail:='';
    while i1>0 do
     begin
       sParam:=DosParamSplit(sData,copy(sMsg,1,i1-1));
       sMail:=sMail+sParam+'='+ComaProtect(sData)+',';
       Delete(sMsg,1,i1);
       i1:=pos('&',sMsg);
     end;
    sParam:=DosParamSplit(sData,sMsg);
    sMail:=sMail+sParam+'='+ComaProtect(sData)+',';
    Delete(sMail,length(sMail),1);
  finally
    result := sMail;
  end;
end;

{************************************  Set Relative Path  ********************}
function SetRelativePath(sPath,sFile:string):string;
begin
  SetRelativePath := sFile;
  try
    if lowercase(copy(sFile,1,length(sPath))) = lowercase(sPath) then
      SetRelativePath := '../../'+copy(sFile,length(sPath)+1,length(sFile));
  except
  end;
end;

{****************************************  Clean space between input tags  ***}
function cleanHiddenSpace(sHtml:string):string;
var
  i1: integer;
  s1: string;
begin
  s1:='';
  try
    {loop through string}
    for i1 := 1 to length(sHtml) do
     if (sHtml[i1]<>#9)  and
        (sHtml[i1]<>#10) and
        (sHtml[i1]<>#13)  then
      s1 := s1 + sHtml[i1];
  finally
    result := s1;
  end;
end;

{******************************  Populate Multi Record Table Array  **********}
procedure populateMultiRecTabs(ini : TIniFile);
var
  i1: integer;
begin
  i1:=0;
  while (i1<20) and (ini.ValueExists('MultiRecTabs','Tab'+intToStr(i1+1))) do
   begin
     aMultiRecTabs[i1] := ini.ReadString('MultiRecTabs','Tab'+intToStr(i1+1),'');
     inc(i1);
   end;
end;

function revPos(Const subStr, str:String; iPos: Integer): Integer;
var
  arr: array of Integer;
  i1: Integer;
begin
  Result:=1;
  if iPos<=1 then exit;
  i1:=1;
  repeat
    i1:=Pos(SubStr, str, i1);
    if i1<>0 then
     begin
       SetLength(arr,Length(arr)+1);
       arr[Length(arr)-1]:=i1;
       inc(i1);
     end;
  until i1=0;

  for i1:=1 to high(arr) do
   if (arr[i1]<iPos) then
    result:=arr[i1];
end;
end.
