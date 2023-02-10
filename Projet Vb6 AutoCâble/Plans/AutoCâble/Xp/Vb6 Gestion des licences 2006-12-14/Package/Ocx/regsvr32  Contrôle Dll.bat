mkdir "C:\Program Files\Dll Ocx Visual Studio"
copy "\\10.30.0.5\production\Cablage-production\AutoCable\Package\ocx\RdSmtp.ocx" "C:\Program Files\Dll Ocx Visual Studio\"


regsvr32 "C:\Program Files\Dll Ocx Visual Studio\RdSmtp.ocx"

pause

