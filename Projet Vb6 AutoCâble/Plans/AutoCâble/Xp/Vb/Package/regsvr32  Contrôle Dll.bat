mkdir "C:\Program Files\Dll Ocx Visual Studio"
copy MSOWC.DLL "C:\Program Files\Dll Ocx Visual Studio\MSOWC.DLL"
copy TABCTL32.OCX "C:\Program Files\Dll Ocx Visual Studio\TABCTL32.OCX"
regsvr32 "C:\Program Files\Dll Ocx Visual Studio\MSOWC.DLL"
regsvr32 "C:\Program Files\Dll Ocx Visual Studio\TABCTL32.OCX"

pause

