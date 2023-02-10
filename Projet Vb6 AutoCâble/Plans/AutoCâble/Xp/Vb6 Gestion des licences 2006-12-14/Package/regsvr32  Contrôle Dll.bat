mkdir "C:\Program Files\Dll Ocx Visual Studio"

copy "\\10.30.0.5\production\Cablage-production\AutoCable\Package\MSOWC.DLL" "C:\Program Files\Dll Ocx Visual Studio\"
copy "\\10.30.0.5\production\Cablage-production\AutoCable\Package\TABCTL32.OCX" "C:\Program Files\Dll Ocx Visual Studio\"
copy "\\10.30.0.5\production\Cablage-production\AutoCable\Package\Msflxgrd.ocx" "C:\Program Files\Dll Ocx Visual Studio\"
copy "\\10.30.0.5\production\Cablage-production\AutoCable\Package\MSADODC.OCX" "C:\Program Files\Dll Ocx Visual Studio\"
copy "\\10.30.0.5\production\Cablage-production\AutoCable\Package\MSDATGRD.OCX" "C:\Program Files\Dll Ocx Visual Studio\"

regsvr32 "C:\Program Files\Dll Ocx Visual Studio\MSOWC.DLL"
regsvr32 "C:\Program Files\Dll Ocx Visual Studio\TABCTL32.OCX"
regsvr32 "C:\Program Files\Dll Ocx Visual Studio\Msflxgrd.ocx"
regsvr32 "C:\Program Files\Dll Ocx Visual Studio\MSADODC.OCX"
regsvr32 "C:\Program Files\Dll Ocx Visual Studio\MSDATGRD.OCX"

pause

