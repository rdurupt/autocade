@echo on
rem X:
rem cd X:\Utilitaires\FLEXLM
set LM_LICENSE_FILE=27000@10.30.0.1
X:\Utilitaires\FLEXLM\lmutil lmstat -s 10.30.0.1 -f > "c:\AutocableLicenceAcad\Licence.txt"
