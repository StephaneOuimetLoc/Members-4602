//TASK3594.01 StephaneO 2019/11/20 New macro : @WizAppend(<parameter>=<value>)

unit MacPrm;

interface

Uses System.SysUtils,System.AnsiStrings;

Const
  cMaxRptParam=500;
  cMsgSeparator:String='----------------------------------------------------------------------';

Type
  TRptParam = Record
    Key: String;
    Data: String;
  end;
  aRptParam = array[1..cMaxRptParam] of TRptParam;

  TMacPrm = class(TObject)

    {* parameters *}
    RptParam: aRptParam;
    iRptParamCount:Integer;
    iRptParamCur:Integer;

    {* errors *}
    dErrFile:TextFile;
    sErrFileName:String;

    {* parameters *}

    Constructor Create; Reintroduce;
    Destructor Destroy; Override;

    Procedure ResetParam;
    Procedure ClrParam(Skey:String);
    Procedure ClrParamAll(sKey:String);
    Procedure ClrParamFrom(iFrom:Integer);
    Procedure ClrAllBut(sPrms:String);

    Procedure SetParam(sKey, sData: String);
    Procedure RplParam(skey,sData:String; sAppend:String='');
//TASK3594.01 {
    Procedure AppendParam(sKey,sAppend:String);
//TASK3594.01 }
    Function  GetParam(Skey:String):String;
    Function  IsParam(sKey:String):Boolean;
    Function  IsParamAll(sKey:String):Boolean;

    Procedure LineToParam(sLine:String);
    Procedure LineRplParam(sLine:String);
    Procedure LineClrParam(sLine:String);

    Procedure HttpToParam(sLine:String; bReplace:Byte);
    Procedure MailToParam(sLine:String; bReplace:Byte);

    Function  LoadParam(sName:String):Boolean;
    Procedure SaveParam(sName:String);
    Procedure DelParamFile(sName:String);
    Function  LoadParamToLine(sName:String):String;

    Procedure CopyAll(var oPrmTmp: TMacPrm);
    Procedure AddAll(var oPrmTmp: TMacPrm; bReplace:Boolean; sFilter:String='');
    Function  ParamCount:Integer;

    Function  ParamToLine:String;
    Function  ParamNameToLine:String;
    Function  ParamToNodes:String;
    Function  ParamToInput:String;
    Function  ParamToComa:String;
    Function  ParamToUrl:String;
    Function  ParamToLink:String;

    {* errors *}

    Function  fNewError(sErrFile:String):Boolean;
    procedure pAddError(sMsg:String);
    procedure pCloseError;
  end;

implementation

Uses CommonUtils,Math
     {$IFDEF TRS}, SMSUtils, HighUtil {$ENDIF};

{*** DONE PARAM VAR *******************************************************}

Constructor TMacPrm.Create;
Begin
  inherited;
  ResetParam;
End;

Destructor TMacPrm.Destroy;
Var
  i1:Integer;
Begin
  For i1:=1 to cMaxRptParam do
   begin
     Finalize(RptParam[i1]);
     FillChar(RptParam[i1], SizeOf(RptParam[i1]), 0);
   end;
  iRptParamCount:=0;
  inherited;
End;

{*** RESET PARAM POOL *******************************************************}

Procedure TMacPrm.ResetParam;
Var
  i1:Integer;
begin
   For i1:=1 to cMaxRptParam do
    begin
      RptParam[i1].Key:='';
      RptParam[i1].Data:='';
    end;
   iRptParamCount:=0;
end;

{*** ClEAR PARAM IN THE POOL **************************************************}

Procedure TMacPrm.ClrParam(Skey:String);
Var
  IParam:Integer;
  i1:Integer;
begin
  IParam:=1;
  SKey:=UpperCase(SKey);
  While (IParam<=iRptParamCount) and
        (SKey<>RptParam[IParam].Key) Do Inc(IParam);

  if (SKey=RptParam[IParam].Key) and (IParam<=iRptParamCount) then
   begin
     For i1:=IParam to iRptParamCount-1 do
      begin
        RptParam[I1].Key:=RptParam[I1+1].Key;
        RptParam[I1].Data:=RptParam[I1+1].Data;
      end;
     RptParam[iRptParamCount].Key:='';
     RptParam[iRptParamCount].Data:='';
     Dec(iRptParamCount);
   end;
end;

{*** ClEAR ALL PARAM IN THE POOL **************************************************}

Procedure TMacPrm.ClrParamAll(sKey:String);
Var
  iPrm,iRpl,iLen:Integer;
begin
  iRpl:=1;
  iLen:=length(sKey);
  sKey:=UpperCase(sKey);

  {Clean parameters}
  For iPrm:=1 to iRptParamCount do
   begin
     if (copy(RptParam[iPrm].Key,1,iLen)<>sKey) and (RptParam[iPrm].Key<>'') then
      if iRpl=iPrm then inc(iRpl) else
       begin
         RptParam[iRpl].Key:=RptParam[iPrm].Key;
         RptParam[iRpl].Data:=RptParam[iPrm].Data;
         Inc(iRpl);
       end;
   end;

  {Clear remaining parameters}
  For iPrm:=iRpl to iRptParamCount do
   begin
     RptParam[iPrm].Key:='';
     RptParam[iPrm].Data:='';
   end;

  iRptParamCount:=iRpl-1;
end;

{*** Clear remaining parameters ******************************************}

Procedure TMacPrm.ClrParamFrom(iFrom:Integer);
var
  iPrm:Integer;
begin
  For iPrm:=iFrom to iRptParamCount do
   begin
     RptParam[iPrm].Key:='';
     RptParam[iPrm].Data:='';
   end;
  iRptParamCount:=iFrom-1;
end;

Procedure TMacPrm.ClrAllBut(sPrms:String);
Var
  iPrm,iRpl:Integer;
begin
  iRpl:=1;

  sPrms:=','+ReplaceStr(sPrms,';',',')+',';

  {Clean parameters}
  For iPrm:=1 to iRptParamCount do
   begin
     if (RptParam[iPrm].Key<>'') and
        (PosNoCase(','+RptParam[iPrm].Key+',',sPrms)<>0) then
      if iRpl=iPrm then inc(iRpl) else
       begin
         RptParam[iRpl].Key:=RptParam[iPrm].Key;
         RptParam[iRpl].Data:=RptParam[iPrm].Data;
         Inc(iRpl);
       end;
   end;

  {Clear remaining parameters}
  For iPrm:=iRpl to iRptParamCount do
   begin
     RptParam[iPrm].Key:='';
     RptParam[iPrm].Data:='';
   end;

  iRptParamCount:=iRpl-1;
end;

{*** SET PARAM IN THE POOL **************************************************}

Procedure TMacPrm.SetParam(sKey,sData:String);
Var
  iParam:Integer;
begin
  if sKey='' then Exit;
  iParam:=1;
  sKey:=UpperCase(sKey);

  While (iParam<=iRptParamCount) and
        (sKey<>RptParam[iParam].Key) Do Inc(iParam);

  if (iParam>iRptParamCount) or (sKey<>RptParam[iParam].Key) then
   begin
     if iRptParamCount<cMaxRptParam then inc(iRptParamCount);
     RptParam[iRptParamCount].Key:=sKey;
     RptParam[iRptParamCount].Data:=sData
   end;
end;

{*** ADD/REPLACE PARAM IN THE POOL *******************************************}

Procedure TMacPrm.RplParam(skey,sData,sAppend: String);
Var
  iParam:Integer;
begin
  if sKey='' then Exit;
  iParam:=1;
  sKey:=UpperCase(sKey);

  While (iParam<=iRptParamCount) and
        (sKey<>RptParam[iParam].Key) Do Inc(iParam);

  if (iParam<=iRptParamCount) and (sKey=RptParam[iParam].Key) then
   begin
     if sAppend='' then
      RptParam[iParam].Data:=sData else
      RptParam[iParam].Data:=RptParam[iParam].Data+sAppend+sData;
   end
  else
   begin
     if iRptParamCount<cMaxRptParam then inc(iRptParamCount);
     RptParam[iRptParamCount].Key:=sKey;
     RptParam[iRptParamCount].Data:=sData
   end;
end;
//TASK3594.01 {

{*** APPEND VALUE TO PARAM IN THE POOL *********************************}

Procedure TMacPrm.AppendParam(skey,sAppend: String);
Var
  iParam:Integer;
begin
  if sKey='' then Exit;
  iParam:=1;
  sKey:=UpperCase(sKey);

  While (iParam<=iRptParamCount) and
        (sKey<>RptParam[iParam].Key) Do Inc(iParam);

  if (iParam<=iRptParamCount) and (sKey=RptParam[iParam].Key) then
   begin
     RptParam[iParam].Data:=RptParam[iParam].Data+sAppend;
   end
  else
   begin
     if iRptParamCount<cMaxRptParam then inc(iRptParamCount);
     RptParam[iRptParamCount].Key:=sKey;
     RptParam[iRptParamCount].Data:=sAppend;
   end;
end;
//TASK3594.01 }

{*** HTTP MESSAGE TO PARAMETER POOL *********************************}

Procedure TMacPrm.HttpToParam(sLine:String; bReplace:Byte);
Var
  iDot,iPos,iKey,iData,iLen:Integer;
  sDot,sKey,sData:String;
  sError:String;
  i1:Integer;

begin
  sError:='';

  {Remove URL prefix}
  iPos:=pos('?',sLine);
  if iPos<>0 then
   sLine:=copy(sLine,iPos+1,length(sLine)) else
   begin
     iPos:=pos(':',sLine);
     if iPos<>0 then sLine:=copy(sLine,iPos+1,length(sLine));
   end;

  {Replace enter by comma}
  sLine:=ReplaceStr(sLine,#13#10,' ');
  sLine:=ReplaceStr(sLine,#13,' ');
  sLine:=ReplaceStr(sLine,#10,' ');

  {LOOP THROUGH ALL CHARACTERS}

  iPos:=1;
  iLen:=length(sLine);
  While iPos<=iLen do
   begin
     iKey:=iPos;

     {Search = separator}
     iPos:=Pos('=',sLine,iKey);
     if iPos<2 then Raise Exception.Create('Invalid Url:'+copy(sLine,iKey,200));

     sKey:=Trim(URLStrip(Copy(sLine,iKey,iPos-iKey)));

     Inc(iPos);
     iData:=iPos;

     {look for next separator}
     While (iPos<=iLen) and (sLine[iPos]<>'&') do Inc(iPos);
     {look further for more separator}
     i1:=iPos;
     While (i1<=iLen) and (sLine[i1]<>'=') do
      begin
        if sLine[i1]='&' then iPos:=i1;
        Inc(i1);
      end;

     {Get data}
     sData:=URLStrip(copy(sLine,iData,iPos-iData));
     {Remove quotes}
     if (length(sData)>1) and (sData[1]='''') and (sData[length(sData)]='''') then
      sData:=copy(sData,2,length(sData)-2);

     {Optional manipulation}
     iDot:=Pos('$',sKey ,1);
     if iDot<>0 then
      begin
        sDot:=UpperCase(Copy(sKey,iDot+1,255));
        if sDot<>'' then
         begin
           Delete(sKey,iDot,255);
           {$IFDEF TRS}
           sError:=ParamSufix(sDot,sData,Self);
           {$ENDIF}
         end;
      end;

     {Replace parameters}
     if bReplace=2 then
      ClrParam(sKey) else
     if bReplace=1 then
      RplParam(sKey,sData) else
      SetParam(sKey,sData);

     Inc(iPos);
   end;

  if sError<>'' then raise EConvertError.Create(sError);
end;

{*** MAILSLOT MESSAGE TO PARAMETER POOL *********************************}

Procedure TMacPrm.MailToParam(sLine:String; bReplace:Byte);
Var
  iPos,iKey,iData,iLen,iDot:Integer;
  sDot,sKey,sData:String;
  sError:String;
  bQuote:Boolean;
begin
  sError:='';
  sLine:=ReplaceStr(sLine,#13#10,' ');
  sLine:=ReplaceStr(sLine,#13,' ');
  sLine:=ReplaceStr(sLine,#10,' ');

  {LOOP THROUGH ALL CHARACTERS}

  iPos:=1;
  iLen:=length(sLine);
  While iPos<=iLen do
   begin
     iKey:=iPos;

     iPos:=Pos('=', sLine,iKey);
     if iPos=0 then iPos:=length(sLine)+1;
     sKey:=Trim(Copy(sLine,iKey,iPos-iKey));
     if (length(sKey) <> 0) And CharInSet(sKey[1], ['-','/',',']) then Delete(sKey, 1, 1);

     bQuote:=False;
     Inc(iPos);
     iData:=iPos;

     {look for next separator}
     While (iPos<=iLen) and ((bQuote) or
           ((sLine[iPos]<>',') and (sLine[iPos]<>'&') and (sLine[iPos]<>' '))) do
      begin
        if sLine[iPos]='''' then
         begin
           if (iPos<iLen) and (sLine[iPos+1]='''') and (bQuote) then Inc(iPos,1)
           else if bQuote then bQuote:=False else bQuote:=True;
         end;
        Inc(iPos);
      end;

     {Get data}
     sData:=copy(sLine,iData,iPos-iData);
     {Remove quotes}
     if (length(sData)>1) and (sData[1]='''') and (sData[length(sData)]='''') then
       sData:=ApostStrip(copy(sData,2,length(sData)-2));

     {Optional manipulation}
     iDot:=Pos('$',sKey,1);
     if iDot<>0 then
      begin
        sDot:=UpperCase(Copy(sKey,iDot+1,255));
        if sDot<>'' then
         begin
           Delete(sKey,iDot,255);
           {$IFDEF TRS}
           sError:=ParamSufix(sDot,sData,Self);
           {$ENDIF}
         end;
      end;

     {Replace parameters}
     if bReplace=2 then
      ClrParam(sKey) else
     if bReplace=1 then
      RplParam(sKey,sData) else
      SetParam(sKey,sData);

     Inc(iPos);
   end;

  if sError<>'' then raise EConvertError.Create(sError);
end;

{*** PARSE LINE ADD TO PARAMETER POOL *********************************}

Procedure TMacPrm.LineToParam(sLine:String);
begin
  if sLine='' then exit;
  if (UpperCase(copy(sLine,1,5))='HTTP:') or
     (UpperCase(copy(sLine,1,6))='HTTPS:') then
   HttpToParam(sLine,0) else
   MailToParam(sLine,0);
end;

{*** PARSE LINE REPLACE PARAMETER POOL *********************************}

Procedure TMacPrm.LineRplParam(sLine:String);
begin
  if sLine='' then exit;
  if (UpperCase(copy(sLine,1,5))='HTTP:') or
     (UpperCase(copy(sLine,1,6))='HTTPS:') then
   HttpToParam(sLine,1) else
   MailToParam(sLine,1);
end;

{*** PARSE LINE CLEAR PARAMETER POOL *********************************}

Procedure TMacPrm.LineClrParam(sLine:String);
begin
  if sLine='' then exit;
  if (UpperCase(copy(sLine,1,5))='HTTP:') or
     (UpperCase(copy(sLine,1,6))='HTTPS:') then
   HttpToParam(sLine,2) else
   MailToParam(sLine,2);
end;

{*** DISPOSE PARAMETERS ON A LINE *********************************}

Function TMacPrm.ParamToLine:String;
Var
  iPrm:Integer;
begin
  Result:='';
  For iPrm:=1 to iRptParamCount do
   if RptParam[iPrm].Key<>'' then
    Result:=Result+' /'+RptParam[iPrm].Key+'='+
     ComaProtect(ApostProtect(RptParam[iPrm].Data));
  Delete(Result,1,1);
end;

{*** DISPOSE PARAMETERS NAME ON A LINE *********************************}

Function TMacPrm.ParamNameToLine:String;
Var
  iPrm:Integer;
begin
  Result:=',';
  For iPrm:=1 to iRptParamCount do
   if RptParam[iPrm].Key<>'' then
    Result:=Result+RptParam[iPrm].Key+',';
end;

{*** LOAD PARAMETER FROM FILE *********************************************}

Function TMacPrm.LoadParam(sName:String):Boolean;
Var
  dPrm:TextFile;
  sPrm,sData:String;
  sGroup:String;
  iGroup:Integer;
  
begin
  {Extract optional group name}
  iGroup:=pos('[',sName);
  if iGroup=0 then sGroup:='' else
   begin
     sGroup:=UpperCase(Trim(copy(sName,iGroup,length(sName))));
     sName:=copy(sName,1,iGroup-1);
   end;

  if pos('.',sName)=0 then sName:=sName+'.ini';

  {open ini file}
  AssignFile(dPrm,sName);
  {$I-} Reset(dPrm); {$I+}
  if IOResult=0 then
   begin

     {load only a section}
     if sGroup<>'' then
      begin
        sPrm:='';
        While (not Eof(dPrm)) and
              (UpperCase(Trim(sPrm))<>sGroup) do Readln(dPrm,sPrm);

        sPrm:='';
        While (not eof(dPrm)) and (copy(trim(sPrm),1,1)<>'[') do
         begin
           Readln(dPrm,sPrm);
           if copy(trim(sPrm),1,1)<>'[' then
            begin
              sPrm:=ComaParamSplit(sData,sPrm);
              if copy(sPrm,1,1)<>';' then SetParam(sPrm,sData);
            end;
         end;
      end

     {load complete ini file}
     else
      begin
        While not eof(dPrm) do
         begin
           Readln(dPrm,sPrm);
           sPrm:=DosParamSplit(sData,sPrm);
           if copy(sPrm,1,1)<>';' then SetParam(sPrm,sData);
         end;
      end;

     CloseFile(dPrm);
     LoadParam:=True;
   end
  else LoadParam:=False;
end;


{*** DELETE PARAMETER FILE *********************************************}

Procedure TMacPrm.DelParamFile(sName:String);
begin
  DeleteFile(sName);
end;

{*** SAVE PARAMETER IN A FILE *********************************************}

Procedure TMacPrm.SaveParam(sName:String);
Var
  dPrm:TextFile;
  iPrm:Integer;
begin
  AssignFile(dPrm,sName);
  Rewrite(dPrm);
  For iPrm:=1 to iRptParamCount do
   Writeln(dPrm,RptParam[iPrm].Key+'='+RptParam[iPrm].Data);
  CloseFile(dPrm);
end;

{*** GET PARAM FROM POOL ************************************************}

Function TMacPrm.GetParam(sKey:String):String;
Var
  iParam:Integer;
begin
  sKey:=UpperCase(sKey);
  GetParam:='';
  For iParam:=1 to iRptParamCount do
   begin
     if sKey=RptParam[iParam].Key then
      begin
        GetParam:=RptParam[iParam].Data;
        Break;
      end;
   end;
end;

{*** CHECK IF PARAM EXIST ************************************************}

Function  TMacPrm.IsParam(sKey:String):Boolean;
Var
  iParam:Integer;
begin
  sKey:=UpperCase(sKey);
  Result:=False;
  For iParam:=1 to iRptParamCount do
   begin
     if sKey=RptParam[iParam].Key then
      begin
        Result:=True;
        Break;
      end;
   end;
end;

{*** CHECK IF PARAM EXIST ************************************************}

Function  TMacPrm.IsParamAll(sKey:String):Boolean;
Var
  iParam:Integer;
  iLen:Integer;
begin
  sKey:=UpperCase(sKey);
  iLen:=length(sKey);
  Result:=False;
  For iParam:=1 to iRptParamCount do
   begin
     if sKey=copy(RptParam[iParam].Key,1,iLen) then
      begin
        Result:=True;
        Break;
      end;
   end;
end;

{*** LOAD PARAM FROM FILE to STRING ******************************************}

Function TMacPrm.LoadParamToLine(sName:String):String;
Var
  SavRptParam:aRptParam;
  iSavRptParamCount:Integer;
  iSavRptParamCur:Integer;
begin
   {save parameter}
   iSavRptParamCount:=iRptParamCount;
   iSavRptParamCur:=iRptParamCur;
   SavRptParam:=RptParam;
   ResetParam;

   {Load new param}
   LoadParam(sName);
   LoadParamToLine:=ParamToLine;

   {restore Param}
   iRptParamCount:=iSavRptParamCount;
   iRptParamCur:=iSavRptParamCur;
   RptParam:=SavRptParam;
end;

{*** PARAM TO HTML INPUT ***********************************************}

Function FilterParam(sKey:String):Boolean;
begin
  Result:=(sKey<>'OUTPUT') and
          (sKey<>'OLD_HTT') and
          (sKey<>'LIVEENTRY') and
          (sKey<>'FUNCTION') and
          (sKey<>'ENTRY_PROTECT') and
          (sKey<>'MESSAGE') and
          (sKey<>'MESSAGE_HELP') and
          (sKey<>'MESSAGE_URL');
end;

Function TMacPrm.ParamToInput:String;
var
  i1:Integer;
begin
  Result:='';
  For i1:=1 to iRptParamCount do
   if RptParam[i1].Key<>'' then
    begin
      if RptParam[i1].Key='FUNCTION' then
       Result:=Result+'<input type="hidden" name="FCT" value="'+QuoteProtect(RptParam[i1].Data)+'">'
      else
      if FilterParam(RptParam[i1].Key) then 
       Result:=Result+'<input type="hidden" name="'+RptParam[i1].Key+'" value="'+QuoteProtect(RptParam[i1].Data)+'">';
    end;
end;

{*** PARAM TO NODES *******************************************************}

Function TMacPrm.ParamToNodes:String;
var
  i1:Integer;

  {Checks if the XML Tag is XML Friendly or not}
  Function XmlFriendly (str: string): boolean;
  Var
    i: integer;
  Begin
    Result := False;
    if str='' then Exit;
    for i := 1 to Length(str) do
     if not (str[i] in ['a'..'z', 'A'..'Z', '0'..'9', '-', '_']) then
      Exit;
    Result:=True;
  End;

begin
  Result:='';
  For i1:=1 to iRptParamCount do
   if RptParam[i1].Key<>'' then
    begin
      if FilterParam(RptParam[i1].Key) then
       if XmlFriendly(RptParam[i1].Key) then
        Result:=Result+'<'+RptParam[i1].Key+'>'+HtmlProtect(RptParam[i1].Data)+'</'+RptParam[i1].Key+'>';
    end;
end;

{*** PARAM TO COMA *******************************************************}

Function TMacPrm.ParamToComa:String;
var
  i1:Integer;
begin
  Result:='';
  For i1:=1 to iRptParamCount do
   if RptParam[i1].Key<>'' then
    begin
      if RptParam[i1].Key='FUNCTION' then
       Result:=Result+',FCT='+RptParam[i1].Data
      else
      if FilterParam(RptParam[i1].Key) then
       Result:=Result+','+RptParam[i1].Key+'='+RptParam[i1].Data;
    end;
  Delete(Result,1,1);
end;

{*** PARAM TO URL CMD *******************************************************}

Function TMacPrm.ParamToUrl:String;
var
  i1:Integer;
begin
  Result:='';
  For i1:=1 to iRptParamCount do
   if RptParam[i1].Key<>'' then
    Result:=Result+'&'+UrlProtect(RptParam[i1].Key)+'='+UrlProtect(RptParam[i1].Data);
  Delete(Result,1,1);
end;

{*** PARAM TO LINK ****************************************************}

Function TMacPrm.ParamToLink:String;
var
  i1:Integer;
  sCgi:String;
begin
  sCgi:=GetParam('PROMPTEXECCGI');
  if sCgi='' then
   Result:='' else
   Result:='&amp;CGI='+sCgi;

  For i1:=1 to iRptParamCount do
   if RptParam[i1].Key<>'' then
    begin
      if RptParam[i1].Key='OUTPUT' then else
      if (RptParam[i1].Key='FUNCTION') and (sCgi='') then
       Result:=Result+'&amp;FCT='+RptParam[i1].Data
      else
      if FilterParam(RptParam[i1].Key) then
       Result:=Result+'&amp;'+RptParam[i1].Key+'='+HtmlProtect(RptParam[i1].Data);
    end;
  Delete(Result,1,5);
end;

{*** BACKUP PARAMETERS ****************************************************}

procedure TMacPrm.CopyAll(var oPrmTmp: TMacPrm);
var
  iPrm:Integer;
begin
  for iPrm:=1 to iRptParamCount do
   begin
     oPrmTmp.RptParam[iPrm].Key:=RptParam[iPrm].Key;
     oPrmTmp.RptParam[iPrm].Data:=RptParam[iPrm].Data;
   end;
  oPrmTmp.iRptParamCount:=iRptParamCount;
  oPrmTmp.iRptParamCur:=iRptParamCur;
  {Clear rest}
  For iPrm:=iRptParamCount+1 to cMaxRptParam do
   begin
     oPrmTmp.RptParam[iPrm].Key:='';
     oPrmTmp.RptParam[iPrm].Data:='';
   end;
end;

{*** ADD PARAMETERS ****************************************************}

procedure TMacPrm.AddAll(var oPrmTmp:TMacPrm; bReplace:Boolean; sFilter:String='');
var
  iPrm:Integer;
  iLen:Integer;
begin
  iLen:=length(sFilter);
  sFilter:=UpperCase(sFilter);
  for iPrm:=1 to iRptParamCount do
   if copy(RptParam[iPrm].Key,1,iLen)=sFilter then
    if bReplace then
     oPrmTmp.RplParam(RptParam[iPrm].Key,RptParam[iPrm].Data) else
     oPrmTmp.SetParam(RptParam[iPrm].Key,RptParam[iPrm].Data);
end;


Function TMacPrm.ParamCount:Integer;
begin
  Result:=iRptParamCOunt;
end;

{/////////////////////////////////////////////////////////////////////////}

{*** ERROR MANAGEMENT ****************************************************}

Function TMacPrm.FNewError(sErrFile:String):Boolean;
begin
  sErrFileName:=sErrFile;
  AssignFile(dErrFile,sErrFileName);
  {$I-} Append(dErrFile); {$I+}
  if IOResult=0 then FNewError:=False else
   begin
     Rewrite(dErrFile);
     FNewError:=True;
   end;
end;

procedure TMacPrm.PAddError(sMsg:String);
begin
  writeln(dErrFile,sMsg);
end;

procedure TMacPrm.PCloseError;
begin
  CloseFile(dErrFile);
  {$WARN SYMBOL_PLATFORM OFF}
  FileSetAttr(sErrFileName,0);
  {$WARN SYMBOL_PLATFORM ON}
end;

end.


