unit u7zip;

interface

uses
  Classes, StrUtils, IdHashMessageDigest, IdHash, IdGlobal, funcoes;

function Compress7z(FileOrDirName, NameOfDestinationFile, Password: string; ComprimirComo7z: Boolean = True): Integer;

implementation

uses
  SysUtils, Windows, Forms, uSimpleHttp, menu, Dialogs;

function RunAndGetDosOutput(CommandLine: string; Work: string = 'C:\'; OutputContent: TSTringList = nil): string;
var
  SecAtrrs: TSecurityAttributes;
  StartupInfo: TStartupInfo;
  ProcessInfo: TProcessInformation;
  StdOutPipeRead, StdOutPipeWrite: THandle;
  WasOK: Boolean;
  pCommandLine: array[0..255] of AnsiChar;
  BytesRead: Cardinal;
  WorkDir: string;
  Handle: Boolean;
begin
  Result := '';
  with SecAtrrs do
  begin
    nLength := SizeOf(SecAtrrs);
    bInheritHandle := True;
    lpSecurityDescriptor := nil;
  end;
  CreatePipe(StdOutPipeRead, StdOutPipeWrite, @SecAtrrs, 0);
  try
    with StartupInfo do
    begin
      FillChar(StartupInfo, SizeOf(StartupInfo), 0);
      cb := SizeOf(StartupInfo);
      dwFlags := STARTF_USESHOWWINDOW or STARTF_USESTDHANDLES;
      wShowWindow := SW_HIDE;
      hStdInput := GetStdHandle(STD_INPUT_HANDLE); // don't redirect stdin
      hStdOutput := StdOutPipeWrite;
      hStdError := StdOutPipeWrite;
    end;
    WorkDir := Work;
    Handle := CreateProcess(nil, PChar('cmd.exe /C ' + CommandLine), nil, nil, True, 0, nil, PChar(WorkDir), StartupInfo, ProcessInfo);
    CloseHandle(StdOutPipeWrite);
    if Handle then
    try
      repeat
        WasOK := ReadFile(StdOutPipeRead, pCommandLine, 255, BytesRead, nil);
        if BytesRead > 0 then
        begin
          pCommandLine[BytesRead] := #0;
          Result := Result + pCommandLine;
          if OutputContent <> nil then
            OutputContent.Add(Result);
        end;
      until not WasOK or (BytesRead = 0);
      WaitForSingleObject(ProcessInfo.hProcess, INFINITE);
    finally
      CloseHandle(ProcessInfo.hThread);
      CloseHandle(ProcessInfo.hProcess);
    end;
  finally
    CloseHandle(StdOutPipeRead);
  end;
end;


procedure Download7z;
begin
  FmMenu.MostrarTEladeProcesso('Aguarde.. Realizando Downlaod 7z.dll');
  uSimpleHttp.Get('https://github.com/JB-Software/7zip/releases/download/unico/7z.dll', nil, 'C:\SisECF\7z.dll');
  FmMenu.MostrarTEladeProcesso('Aguarde.. Realizando Downlaod 7z.exe');
  uSimpleHttp.Get('https://github.com/JB-Software/7zip/releases/download/unico/7z.exe', nil, 'C:\SisECF\7z.exe');
end;





function Compress7z(FileOrDirName, NameOfDestinationFile, Password: string; ComprimirComo7z: Boolean = True): Integer;
var
  Command, ComplementDir, ComplementDir2: string;
  I, CountLines: Integer;
  md5Exe, Md5Dll, Complemento: string;
  OutputContent: TStringList;
begin
  md5Exe := 'A1EFCEDC97C76B356F7FFA7CF909D733';
  Md5Dll := 'AECEF77725F3EE0B84B6B8046EFE5AC0';
  if (not FileExists('C:\SisECF\7z.exe')) or (not FileExists('C:\SisECF\7z.exe')) then
     Begin
       //apagar os que existirem na pasta  arquivos para baixar novamente
        If FileExists('C:\SisECF\7z.exe') then
           DeleteFile('C:\SisECF\7z.exe');
        if FileExists('C:\SisECF\7z.dll') then
           DeleteFile('C:\SisECF\7z.dll');
       Download7z;
     end
  else
  begin
    if (UpperCase(GErar_MD5_ARQ('C:\SisECF\7z.exe')) <> md5Exe) or (UpperCase(GErar_MD5_ARQ('C:\SisECF\7z.dll')) <> Md5Dll) then
    begin
      DeleteFile('C:\SisECF\7z.exe');
      DeleteFile('C:\SisECF\7z.dll');
      Download7z;
    end;
  end;

  FmMenu.MostrarTEladeProcesso('Aguarde.. Compactando Arquivos');
  ComplementDir := '';
  ComplementDir2 := '';
  if Copy(FileOrDirName, Length(FileOrDirName)) = '\' then
    FileOrDirName := Copy(FileOrDirName, 1, Length(FileOrDirName) - 1);

  if DirectoryExists(FileOrDirName) then
  begin
    ComplementDir := '\*.*';
    ComplementDir2 := '-r';
  end;
  if FileExists(NameOfDestinationFile) then
    DeleteFile(PChar(NameOfDestinationFile));
  CountLines := 0;

  if DirectoryExists(FileOrDirName) or FileExists(FileOrDirName) then
  begin
    if ComprimirComo7z then
    begin
      if Password <> '' then
        Command := Format('C:\SisECF\7z.exe a -mx=9 -p%s %s "%s.7z" "%s%s"', [Password, ComplementDir2, Trim(StringReplace(NameOfDestinationFile, '.7z', '', [rfIgnoreCase, rfReplaceAll])), Trim(FileOrDirName), ComplementDir])
      else
        Command := Format('C:\SisECF\7z.exe a -mx=9 %s "%s.7z" "%s%s"', [ComplementDir2, Trim(StringReplace(NameOfDestinationFile, '.7z', '', [rfIgnoreCase, rfReplaceAll])), Trim(FileOrDirName), ComplementDir])
    end
    else
    begin
      Command := Format('C:\SisECF\7z.exe a -tzip %s "%s.zip" "%s%s"', [ComplementDir2, Trim(StringReplace(NameOfDestinationFile, '.zip', '', [rfIgnoreCase, rfReplaceAll])), Trim(FileOrDirName), ComplementDir])
    end;
  end
  else
    raise Exception.Create('O arquivo ou diretório não existe em: ' + FileOrDirName + '-' + NameOfDestinationFile);
  OutputContent := TStringList.Create;
  OutputContent.Text := RunAndGetDosOutput(Command, ExtractFilePath(Application.ExeName));
  for I := 0 to OutputContent.Count - 1 do
  begin
    if StartsText('Compressing', OutputContent.Strings[I]) then
      Inc(CountLines);
  end;
  if (ContainsText(OutputContent.Text, 'Error:')) or (ContainsText(OutputContent.Text, 'WARNING:')) then
    raise Exception.Create(OutputContent.Text);

  Result := CountLines;
  FreeAndNil(OutputContent);
end;

end.

