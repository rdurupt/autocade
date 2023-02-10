Attribute VB_Name = "ExporterXml"
' ***********************************************************************
'   Last Update : January 6th 2005
'   Author      : Alexandre VASQUES
'   E-Mail      : avq@ds-fr.com
'   Tel         : + 33 1 55 49 84 74
' ***********************************************************************


Sub Create(Fichier As String, extension As String, Chemin As String, MyWorkbook As Workbook)
'If Right(UCase(Fichier), 4) <> ".XML" Then Fichier = Fichier & ".XML"
If Fichier = "" Or (UCase(extension) <> "XML" And UCase(extension) <> UCase("iXFElec")) Or Chemin = "" Then
MsgBox "You have to specify the :" & vbCrLf & "- System Name" & vbCrLf & "- File extension (xml or iXFElec)" & vbCrLf & "- Destionation path.", vbExclamation, "iXFElecGeneratorV9"
Exit Sub
Else
SetFile Fichier, extension, Chemin, MyWorkbook
End If


End Sub

Public Sub SetFile(Filename As String, extension As String, Chemin As String, MyWorkbook As Workbook)

Dim Txt As String

Dim iRow As Integer
Dim SicId As String
Dim numFile As Long
numFile = FreeFile
Dim splitFilename
splitFilename = split(Filename, "\")

'If Right(Chemin, 1) <> "\" Then Chemin = Chemin & "\"

Open Chemin & "." & extension For Output As #numFile
Print #numFile, "<SOAP_ENV:Envelope xmlns:NS1 = "; Chr(34); " CATIA/V5/Electrical/1.0"; Chr(34); " xmlns:NS2 = "; Chr(34); "http://www.ixfstd.org/std/ns/core/classBehaviors/links/1.0"; Chr(34); " xmlns:ixf = "; Chr(34); "http://www.ixfstd.org/std/ns/core/1.0"; Chr(34); " xmlns:tns = "; Chr(34); "IXF_Schema.xsd"; Chr(34); " xmlns:xsi = "; Chr(34); "http://www.w3.org/2001/XMLSchema-instance"; Chr(34); " xmlns:SOAP_ENV = "; Chr(34); "http://schemas.xmlsoap.org/soap/envelope/"; Chr(34); " xsi:schemaLocation = "; Chr(34); "IXF_Schema.xsd ElectricalSchema.xsd"; Chr(34); "><SOAP_ENV:Body><ixf:object id = "; Chr(34); "" & splitFilename(UBound(splitFilename)); Chr(34); " xsi:type = "; Chr(34); "tns:Harness"; Chr(34); "><tns:Name>"; "" & splitFilename(UBound(splitFilename)); "</tns:Name>"
Print #numFile, "</ixf:object>"

'************************************************************
'SIC
'************************************************************

MyWorkbook.Sheets("SIC-TERM").Select
Txt = ""
For iRow = 2 To MyWorkbook.Sheets("SIC-TERM").Range("A1").CurrentRegion.Rows.Count
    If MyWorkbook.Sheets("SIC-TERM").Cells(iRow, 1) <> 0 Then
    SicId = Replace(MyWorkbook.Sheets("SIC-TERM").Cells(iRow, 3), ".", "*")
        Print #numFile, "<ixf:object id = "; Chr(34); Replace(MyWorkbook.Sheets("SIC-TERM").Cells(iRow, 3), ".", "*"); Chr(34); " xsi:type = "; Chr(34); "tns:Connector"; Chr(34); "><tns:Name>"; Replace(MyWorkbook.Sheets("SIC-TERM").Cells(iRow, 2), ".", "*"); "</tns:Name>"
            If MyWorkbook.Sheets("SIC-TERM").Cells(iRow, 4) <> "" Then
        Print #numFile, "<NS1:Connector>"
        Print #numFile, "<NS1:MatingConnector>"; MyWorkbook.Sheets("SIC-TERM").Cells(iRow, 4); "</NS1:MatingConnector>"
        Print #numFile, "</NS1:Connector>"
            Else
        Print #numFile, "<NS1:Connector/>"
            End If
        Print #numFile, "<NS1:Product><NS1:PartNumber>"; MyWorkbook.Sheets("SIC-TERM").Cells(iRow, 1); "</NS1:PartNumber>"
        Print #numFile, "</NS1:Product>"
        Print #numFile, "</ixf:object>"
    
    End If
    
    
    If MyWorkbook.Sheets("SIC-TERM").Cells(iRow, 5) <> 0 Then
        Print #numFile, "<ixf:object id = "; Chr(34); "UID" & iRow; Chr(34) & " xsi:type = "; Chr(34); "tns:DeviceLink"; Chr(34); "><NS2:link><NS2:object1 href = "; Chr(34); ; "#"; SicId & "." & MyWorkbook.Sheets("SIC-TERM").Cells(iRow, 5); Chr(34); "/>"
        Print #numFile, "<NS2:object2 href = "; Chr(34); ; "#"; SicId; Chr(34); " />"
        Print #numFile, "</NS2:link>"
        Print #numFile, "</ixf:object>"
        Print #numFile, " <ixf:object id = "; Chr(34); SicId & "." & MyWorkbook.Sheets("SIC-TERM").Cells(iRow, 5); Chr(34); " xsi:type = "; Chr(34); "tns:Pin"; Chr(34); "><tns:Name>"; MyWorkbook.Sheets("SIC-TERM").Cells(iRow, 5); "</tns:Name>"
        Print #numFile, " </ixf:object>"
        
   
        
    End If
Next

'************************************************************
'CONTACT
'************************************************************

MyWorkbook.Sheets("SIC-CONT").Select
For iRow = 2 To MyWorkbook.Sheets("SIC-CONT").Range("A1").CurrentRegion.Rows.Count
    If MyWorkbook.Sheets("SIC-CONT").Cells(iRow, 1) <> 0 Then
    SicId = Replace(MyWorkbook.Sheets("SIC-CONT").Cells(iRow, 3), ".", "*")
        Print #numFile, "<ixf:object id = "; Chr(34); Replace(MyWorkbook.Sheets("SIC-CONT").Cells(iRow, 3), ".", "*"); Chr(34); " xsi:type = "; Chr(34); "tns:Connector"; Chr(34); "><tns:Name>"; Replace(Cells(iRow, 2), ".", "*"); "</tns:Name>"
              If Cells(iRow, 4) <> "" Then
        Print #numFile, "<NS1:Connector>"
        Print #numFile, "<NS1:MatingConnector>"; MyWorkbook.Sheets("SIC-CONT").Cells(iRow, 4); "</NS1:MatingConnector>"
        Print #numFile, "</NS1:Connector>"
            Else
        Print #numFile, "<NS1:Connector/>"
            End If
        Print #numFile, "<NS1:Product><NS1:PartNumber>"; MyWorkbook.Sheets("SIC-CONT").Cells(iRow, 1); "</NS1:PartNumber>"
        Print #numFile, "</NS1:Product>"
        Print #numFile, "</ixf:object>"
    
    End If
    
    
    If MyWorkbook.Sheets("SIC-CONT").Cells(iRow, 5) <> 0 Then
        Print #numFile, "<ixf:object id = "; Chr(34); "UID" & iRow; Chr(34) & " xsi:type = "; Chr(34); "tns:DeviceLink"; Chr(34); "><NS2:link><NS2:object1 href = "; Chr(34); ; "#"; SicId & "." & MyWorkbook.Sheets("SIC-CONT").Cells(iRow, 5); Chr(34); "/>"
        Print #numFile, "<NS2:object2 href = "; Chr(34); ; "#"; SicId; Chr(34); " />"
        Print #numFile, "</NS2:link>"
        Print #numFile, "</ixf:object>"
        
        Print #numFile, " <ixf:object id = "; Chr(34); SicId & "." & MyWorkbook.Sheets("SIC-CONT").Cells(iRow, 5); Chr(34); " xsi:type = "; Chr(34); "tns:Cavity"; Chr(34); ">"
        Print #numFile, "<tns:Name>"; MyWorkbook.Sheets("SIC-CONT").Cells(iRow, 5); "</tns:Name>"
        Print #numFile, " </ixf:object>"
        
        Print #numFile, "<ixf:object id = "; Chr(34); "UID" & iRow; Chr(34) & " xsi:type = "; Chr(34); "tns:DeviceLink"; Chr(34); "><NS2:link><NS2:object1 href = "; Chr(34); ; "#"; SicId & "." & MyWorkbook.Sheets("SIC-CONT").Cells(iRow, 5) & "." & MyWorkbook.Sheets("SIC-CONT").Cells(iRow, 6); Chr(34); "/>"
        Print #numFile, "<NS2:object2 href = "; Chr(34); ; "#"; SicId & "." & MyWorkbook.Sheets("SIC-CONT").Cells(iRow, 5); Chr(34); " />"
        Print #numFile, "</NS2:link>"
        Print #numFile, "</ixf:object>"
        
        Print #numFile, " <ixf:object id = "; Chr(34); SicId & "." & MyWorkbook.Sheets("SIC-CONT").Cells(iRow, 5) & "." & MyWorkbook.Sheets("SIC-CONT").Cells(iRow, 6); Chr(34); " xsi:type = "; Chr(34); "tns:Pin"; Chr(34); ">"
        Print #numFile, "<tns:Name>"; SicId & "." & MyWorkbook.Sheets("SIC-CONT").Cells(iRow, 5) & "." & MyWorkbook.Sheets("SIC-CONT").Cells(iRow, 6); "</tns:Name>"
        Print #numFile, "<NS1:Product><NS1:PartNumber>"; MyWorkbook.Sheets("SIC-CONT").Cells(iRow, 6); "</NS1:PartNumber></NS1:Product>"
        Print #numFile, " </ixf:object>"
        
    End If
Next


'************************************************************
'IS
'************************************************************

MyWorkbook.Sheets("IS").Select

For iRow = 2 To MyWorkbook.Sheets("IS").Range("A1").CurrentRegion.Rows.Count
    If MyWorkbook.Sheets("IS").Cells(iRow, 1) <> 0 Then
    ISid = Replace(MyWorkbook.Sheets("IS").Cells(iRow, 3), ".", "*")
        Print #numFile, "<ixf:object id = "; Chr(34); Replace(MyWorkbook.Sheets("IS").Cells(iRow, 3), ".", "*"); Chr(34); " xsi:type = "; Chr(34); "tns:Splice"; Chr(34); "><tns:Name>"; Replace(MyWorkbook.Sheets("IS").Cells(iRow, 2), ".", "*"); "</tns:Name>"
        Print #numFile, "<NS1:Splice>"
        Print #numFile, "<NS1:SubType></NS1:SubType>"
        Print #numFile, "</NS1:Splice>"
        Print #numFile, "<NS1:Product><NS1:PartNumber>"; MyWorkbook.Sheets("IS").Cells(iRow, 1); "</NS1:PartNumber>"
        Print #numFile, "</NS1:Product>"
        Print #numFile, "</ixf:object>"
    
    End If
    
    
    If MyWorkbook.Sheets("IS").Cells(iRow, 4) <> 0 Then
        Print #numFile, "<ixf:object id = "; Chr(34); "UID" & iRow; Chr(34) & " xsi:type = "; Chr(34); "tns:DeviceLink"; Chr(34); "><NS2:link><NS2:object1 href = "; Chr(34); ; "#"; ISid & "." & MyWorkbook.Sheets("IS").Cells(iRow, 4); Chr(34); "/>"
        Print #numFile, "<NS2:object2 href = "; Chr(34); ; "#"; ISid; Chr(34); " />"
        Print #numFile, "</NS2:link>"
        Print #numFile, "</ixf:object>"
        Print #numFile, " <ixf:object id = "; Chr(34); ISid & "." & MyWorkbook.Sheets("IS").Cells(iRow, 4); Chr(34); " xsi:type = "; Chr(34); "tns:Pin"; Chr(34); "><tns:Name>"; MyWorkbook.Sheets("IS").Cells(iRow, 4); "</tns:Name>"
        Print #numFile, " </ixf:object>"
        
    End If
Next

'************************************************************
'Shell
'************************************************************

MyWorkbook.Sheets("Shell").Select
 
For iRow = 2 To MyWorkbook.Sheets("Shell").Range("A1").CurrentRegion.Rows.Count
    If MyWorkbook.Sheets("Shell").Cells(iRow, 1) <> 0 Then
    ShellID = Replace(Cells(iRow, 3), ".", "*")
        Print #numFile, "<ixf:object id = "; Chr(34); ShellID; Chr(34); " xsi:type = "; Chr(34); "tns:ConnectorShell"; Chr(34); "><tns:Name>"; Replace(MyWorkbook.Sheets("Shell").Cells(iRow, 2), ".", "*"); "</tns:Name>"
        Print #numFile, "<NS1:Product><NS1:PartNumber>"; MyWorkbook.Sheets("Shell").Cells(iRow, 1); "</NS1:PartNumber>"
        Print #numFile, "</NS1:Product>"
        Print #numFile, "</ixf:object>"
    
    End If
    
    
    If Cells(iRow, 4) <> 0 Then
    SicId = Cells(iRow, 6)
        Print #numFile, "<ixf:object id = "; Chr(34); "UID" & iRow & Chr(34); " xsi:type = "; Chr(34); "tns:DeviceLink"; Chr(34); "><NS2:link><NS2:object1 href = "; Chr(34); ; "#"; MyWorkbook.Sheets("Shell").Cells(iRow, 6); Chr(34); "/>"
        Print #numFile, "<NS2:object2 href = "; Chr(34); ; "#"; ShellID; Chr(34); " />"
        Print #numFile, "</NS2:link>"
        Print #numFile, "</ixf:object>"
        Print #numFile, " <ixf:object id = "; Chr(34); MyWorkbook.Sheets("Shell").Cells(iRow, 6); Chr(34); " xsi:type = "; Chr(34); "tns:Connector"; Chr(34); "><tns:Name>"; MyWorkbook.Sheets("Shell").Cells(iRow, 4); "</tns:Name><NS1:Connector/>"
        Print #numFile, " <NS1:Product><NS1:PartNumber>"; Cells(iRow, 5); "</NS1:PartNumber>"
        Print #numFile, "</NS1:Product>"
        Print #numFile, " </ixf:object>"
    End If
    
    
    If Cells(iRow, 7) <> 0 Then
        Print #numFile, "<ixf:object id = "; Chr(34); " UID " & iRow; Chr(34) & " xsi:type = "; Chr(34); "tns:DeviceLink"; Chr(34); "><NS2:link><NS2:object1 href = "; Chr(34); ; "#"; SicId & "." & MyWorkbook.Sheets("Shell").Cells(iRow, 7); Chr(34); "/>"
        Print #numFile, "<NS2:object2 href = "; Chr(34); ; "#"; SicId; Chr(34); " />"
        Print #numFile, "</NS2:link>"
        Print #numFile, "</ixf:object>"
        Print #numFile, " <ixf:object id = "; Chr(34); SicId & "." & MyWorkbook.Sheets("Shell").Cells(iRow, 7); Chr(34); " xsi:type = "; Chr(34); "tns:Pin"; Chr(34); "><tns:Name>"; MyWorkbook.Sheets("Shell").Cells(iRow, 7); "</tns:Name>"
        Print #numFile, " </ixf:object>"
    End If
    
Next


'************************************************************
'EQT
'************************************************************

MyWorkbook.Sheets("EQT").Select

For iRow = 2 To MyWorkbook.Sheets("EQT").Range("A1").CurrentRegion.Rows.Count
    If MyWorkbook.Sheets("EQT").Cells(iRow, 1) <> 0 Then
    EQTID = Replace(MyWorkbook.Sheets("EQT").Cells(iRow, 3), ".", "*")
        Print #numFile, "<ixf:object id = "; Chr(34); EQTID; Chr(34); " xsi:type = "; Chr(34); "tns:Equipment"; Chr(34); "><tns:Name>"; Replace(MyWorkbook.Sheets("EQT").Cells(iRow, 2), ".", "*"); "</tns:Name>"
        Print #numFile, "<NS3:Function>"
        Print #numFile, "<NS3:System_Type>"; MyWorkbook.Sheets("EQT").Cells(iRow, 8); "</NS3:System_Type>"
        Print #numFile, "<NS3:Description>"; MyWorkbook.Sheets("EQT").Cells(iRow, 9); "</NS3:Description>"
        Print #numFile, "<NS3:Localisation>"; MyWorkbook.Sheets("EQT").Cells(iRow, 10); "</NS3:Localisation>"
        Print #numFile, "</NS3:Function>"
        Print #numFile, "<NS1:Product><NS1:PartNumber>"; MyWorkbook.Sheets("EQT").Cells(iRow, 1); "</NS1:PartNumber>"
        Print #numFile, "</NS1:Product>"
        Print #numFile, "</ixf:object>"
    
    End If
    
    
    If Cells(iRow, 4) <> 0 Then
    SicId = MyWorkbook.Sheets("EQT").Cells(iRow, 6)
        Print #numFile, "<ixf:object id = "; Chr(34); "UID" & iRow & Chr(34); " xsi:type = "; Chr(34); "tns:DeviceLink"; Chr(34); "><NS2:link><NS2:object1 href = "; Chr(34); ; "#"; MyWorkbook.Sheets("EQT").Cells(iRow, 6); Chr(34); "/>"
        Print #numFile, "<NS2:object2 href = "; Chr(34); ; "#"; EQTID; Chr(34); " />"
        Print #numFile, "</NS2:link>"
        Print #numFile, "</ixf:object>"
        Print #numFile, " <ixf:object id = "; Chr(34); MyWorkbook.Sheets("EQT").Cells(iRow, 6); Chr(34); " xsi:type = "; Chr(34); "tns:Connector"; Chr(34); "><tns:Name>"; MyWorkbook.Sheets("EQT").Cells(iRow, 4); "</tns:Name><NS1:Connector/>"
        Print #numFile, " <NS1:Product><NS1:PartNumber>"; MyWorkbook.Sheets("EQT").Cells(iRow, 5); "</NS1:PartNumber>"
        Print #numFile, "</NS1:Product>"
        Print #numFile, " </ixf:object>"
    End If
    
    
    If MyWorkbook.Sheets("EQT").Cells(iRow, 7) <> 0 Then
        Print #numFile, "<ixf:object id = "; Chr(34); " UID " & iRow; Chr(34) & " xsi:type = "; Chr(34); "tns:DeviceLink"; Chr(34); "><NS2:link><NS2:object1 href = "; Chr(34); ; "#"; SicId & MyWorkbook.Sheets("EQT").Cells(iRow, 7); Chr(34); "/>"
        Print #numFile, "<NS2:object2 href = "; Chr(34); ; "#"; SicId; Chr(34); " />"
        Print #numFile, "</NS2:link>"
        Print #numFile, "</ixf:object>"
        Print #numFile, " <ixf:object id = "; Chr(34); SicId & MyWorkbook.Sheets("EQT").Cells(iRow, 7); Chr(34); " xsi:type = "; Chr(34); "tns:Pin"; Chr(34); "><tns:Name>"; MyWorkbook.Sheets("EQT").Cells(iRow, 7); "</tns:Name>"
        Print #numFile, " </ixf:object>"
    End If
    
Next

'************************************************************
'Wire Group
'************************************************************

MyWorkbook.Sheets("WireGroup").Select

Dim WGID As String
Dim J As Integer

For iRow = 2 To MyWorkbook.Sheets("WireGroup").Range("A1").CurrentRegion.Rows.Count

If Cells(iRow, 1) = 0 And Cells(iRow, 6) = 0 Then
    J = 0
End If

If MyWorkbook.Sheets("WireGroup").Cells(iRow, 1) <> 0 Then
        WGID = MyWorkbook.Sheets("WireGroup").Cells(iRow, 1)
        J = J + 1
        
        If J = 1 Then
            Print #numFile, "<ixf:object xsi:type="; Chr(34); "tns:HarnessLink"; Chr(34); " id="; Chr(34); "Link_"; MyWorkbook.Sheets("WireGroup").Cells(iRow, 1); "_Harness"; Chr(34); ">"
            Print #numFile, "<NS2:link>"
            Print #numFile, "<NS2:object1 href="; Chr(34); "#"; MyWorkbook.Sheets("WireGroup").Cells(iRow, 1); Chr(34); "/>"
            Print #numFile, "<NS2:object2 href ="; Chr(34); "#"; "" & splitFilename(UBound(splitFilename)); Chr(34); "/>"
            Print #numFile, "</NS2:link>"
            Print #numFile, "</ixf:object>"
        Else
            Print #numFile, "<ixf:object xsi:type="; Chr(34); "tns:WireGroupLink"; Chr(34); " id="; Chr(34); "Link_"; MyWorkbook.Sheets("WireGroup").Cells(iRow - J + 1, 1); "_"; MyWorkbook.Sheets("WireGroup").Cells(iRow, 1); Chr(34); "> "
            Print #numFile, "<NS2:link>"
            Print #numFile, "<NS2:object1 href="; Chr(34); "#"; MyWorkbook.Sheets("WireGroup").Cells(iRow - J + 1, 1); Chr(34); "/>"
            Print #numFile, "<NS2:object2 href="; Chr(34); "#"; MyWorkbook.Sheets("WireGroup").Cells(iRow, 1); Chr(34); "/>"
            Print #numFile, "</NS2:link>"
            Print #numFile, "</ixf:object>"
        End If

        Print #numFile, "<ixf:object xsi:type="; Chr(34); "tns:WireGroup"; Chr(34); " id="; Chr(34); MyWorkbook.Sheets("WireGroup").Cells(iRow, 1); Chr(34); ">"
        Print #numFile, "<tns:Name>"; Cells(iRow, 1); "</tns:Name>"
        Print #numFile, "<NS1:Product>"
        Print #numFile, "<NS1:PartNumber>"; Replace(MyWorkbook.Sheets("WireGroup").Cells(iRow, 2), ".", "*"); "</NS1:PartNumber>"
        Print #numFile, "</NS1:Product>"
        Print #numFile, "<NS1:WireGroup>"
        Print #numFile, "<NS1:Diameter>"; Replace(MyWorkbook.Sheets("WireGroup").Cells(iRow, 3), ".", "*"); "</NS1:Diameter>"
        Print #numFile, "<NS1:WireLengthCoeff>"; MyWorkbook.Sheets("WireGroup").Cells(iRow, 4); "</NS1:WireLengthCoeff>"
        Print #numFile, "<NS1:BendRadius>"; MyWorkbook.Sheets("WireGroup").Cells(iRow, 5); "</NS1:BendRadius>"
        Print #numFile, "</NS1:WireGroup>"
        Print #numFile, "</ixf:object>"
End If

If MyWorkbook.Sheets("WireGroup").Cells(iRow, 6) <> 0 Then
        J = J + 1
        Print #numFile, "<ixf:object id ="; Chr(34); MyWorkbook.Sheets("WireGroup").Cells(iRow, 6); Chr(34); " xsi:type="; Chr(34); "tns:Wire"; Chr(34); "><tns:Name>" & MyWorkbook.Sheets("WireGroup").Cells(iRow, 6) & "</tns:Name>"
        Print #numFile, "<NS1:Product><NS1:PartNumber>" & MyWorkbook.Sheets("WireGroup").Cells(iRow, 7) & "</NS1:PartNumber>"
        Print #numFile, "</NS1:Product>"
        Print #numFile, "<NS1:Wire>"
        Print #numFile, "<NS1:OuterDiameter>" & MyWorkbook.Sheets("WireGroup").Cells(iRow, 8) & "</NS1:OuterDiameter>"
        Print #numFile, "<NS1:BendRadius>" & MyWorkbook.Sheets("WireGroup").Cells(iRow, 9) & "</NS1:BendRadius>"
        Print #numFile, "<NS1:LinearMass>" & MyWorkbook.Sheets("WireGroup").Cells(iRow, 10) & "</NS1:LinearMass>"
        Print #numFile, "<NS1:Color>" & MyWorkbook.Sheets("WireGroup").Cells(iRow, 11) & "</NS1:Color>"
        Print #numFile, "</NS1:Wire>"
        Print #numFile, "</ixf:object>"
        
        Print #numFile, "<ixf:object id = "; Chr(34); "UI" & iRow * 2 & Chr(34); " xsi:type = "; Chr(34); "tns:WireLink"; Chr(34); " ><NS2:link><NS2:object1 href = "; Chr(34); "#" & MyWorkbook.Sheets("WireGroup").Cells(iRow, 12) & Chr(34); "/>"
        Print #numFile, "<NS2:object2 href = "; Chr(34); "#" & MyWorkbook.Sheets("WireGroup").Cells(iRow, 6); Chr(34); "/>"
        Print #numFile, "</NS2:link>"
        Print #numFile, "</ixf:object>"
        Print #numFile, "<ixf:object id = "; Chr(34); " UI" & iRow * 2 & Chr(34); " xsi:type = "; Chr(34); "tns:WireLink"; Chr(34); "><NS2:link><NS2:object1 href = "; Chr(34); "#" & MyWorkbook.Sheets("WireGroup").Cells(iRow, 13) & Chr(34); "/>"
        Print #numFile, "<NS2:object2 href = "; Chr(34); "#" & MyWorkbook.Sheets("WireGroup").Cells(iRow, 6); Chr(34); "/>"
        Print #numFile, "</NS2:link>"
        Print #numFile, "</ixf:object>"
        
        Print #numFile, "<ixf:object xsi:type="; Chr(34); "tns:WireGroupLink"; Chr(34); " id="; Chr(34); "Link_"; WGID; "_"; MyWorkbook.Sheets("WireGroup").Cells(iRow, 6); Chr(34); "> "
        Print #numFile, "<NS2:link>"
        Print #numFile, "<NS2:object1 href="; Chr(34); "#"; WGID; Chr(34); "/>"
        Print #numFile, "<NS2:object2 href="; Chr(34); "#"; Cells(iRow, 6); Chr(34); "/>"
        Print #numFile, "</NS2:link>"
        Print #numFile, "</ixf:object>"
        
End If
Next


'************************************************************
'Wires
'************************************************************

MyWorkbook.Sheets("Wire").Select

For iRow = 2 To MyWorkbook.Sheets("Wire").Range("A1").CurrentRegion.Rows.Count

If MyWorkbook.Sheets("Wire").Cells(iRow, 1) <> 0 Then

        Print #numFile, "<ixf:object id = "; Chr(34); MyWorkbook.Sheets("Wire").Cells(iRow, 1); Chr(34); " xsi:type = "; Chr(34); "tns:Wire"; Chr(34); "><tns:Name>" & MyWorkbook.Sheets("Wire").Cells(iRow, 1) & "</tns:Name>"
        Print #numFile, "<NS1:Wire>"
        Print #numFile, "<NS1:OuterDiameter>" & Replace(MyWorkbook.Sheets("Wire").Cells(iRow, 3), ".", "*") & "</NS1:OuterDiameter>"
        Print #numFile, "<NS1:BendRadius>" & MyWorkbook.Sheets("Wire").Cells(iRow, 4) & "</NS1:BendRadius>"
        Print #numFile, "<NS1:LinearMass>" & MyWorkbook.Sheets("Wire").Cells(iRow, 5) & "</NS1:LinearMass>"
        Print #numFile, "<NS1:Color>" & MyWorkbook.Sheets("Wire").Cells(iRow, 6) & "</NS1:Color>"
        
        Print #numFile, "</NS1:Wire>"
        Print #numFile, "<NS1:Product><NS1:PartNumber>" & Replace(MyWorkbook.Sheets("Wire").Cells(iRow, 2), ".", "*") & "</NS1:PartNumber>"
        Print #numFile, "</NS1:Product>"
        Print #numFile, "</ixf:object>"
        Print #numFile, "<ixf:object id = "; Chr(34); " UI" & iRow & Chr(34); " xsi:type = "; Chr(34); "tns:HarnessLink"; Chr(34); "><NS2:link><NS2:object1 href = "; Chr(34); "#" & "" & splitFilename(UBound(splitFilename)); Chr(34); "/>"
        Print #numFile, "<NS2:object2 href = "; Chr(34); "#" & MyWorkbook.Sheets("Wire").Cells(iRow, 1); Chr(34); "/>"
        Print #numFile, "</NS2:link>"
        Print #numFile, "</ixf:object>"
        Print #numFile, "<ixf:object id = "; Chr(34); " UI" & iRow * 2 & Chr(34); " xsi:type = "; Chr(34); "tns:WireLink"; Chr(34); " ><NS2:link><NS2:object1 href = "; Chr(34); "#" & MyWorkbook.Sheets("Wire").Cells(iRow, 7) & Chr(34); "/>"
        Print #numFile, "<NS2:object2 href = "; Chr(34); "#" & MyWorkbook.Sheets("Wire").Cells(iRow, 1); Chr(34); "/>"
        Print #numFile, "</NS2:link>"
        Print #numFile, "</ixf:object>"
        Print #numFile, "<ixf:object id = "; Chr(34); " UI" & iRow * 2 & Chr(34); " xsi:type = "; Chr(34); "tns:WireLink"; Chr(34); "><NS2:link><NS2:object1 href = "; Chr(34); "#" & MyWorkbook.Sheets("Wire").Cells(iRow, 8) & Chr(34); "/>"
        Print #numFile, "<NS2:object2 href = "; Chr(34); "#" & MyWorkbook.Sheets("Wire").Cells(iRow, 1); Chr(34); "/>"
        Print #numFile, "</NS2:link>"
        Print #numFile, "</ixf:object>"
End If
Next


Print #numFile, "</SOAP_ENV:Body></SOAP_ENV:Envelope>"
Close #numFile

MyWorkbook.Sheets("Create").Select


End Sub


Sub clearAll(MyWorkbook As Workbook)
Dim Answer As Integer
Answer = MsgBox("Are you sure you want to clear all?", vbYesNo)
If Answer = 6 Then

MyWorkbook.Sheets("SIC-TERM").Select
Detruire MyWorkbook.Sheets("SIC-TERM")
MyWorkbook.Sheets("SIC-CONT").Select
Detruire MyWorkbook.Sheets("SIC-CONT")
MyWorkbook.Sheets("IS").Select
Detruire MyWorkbook.Sheets("IS")
MyWorkbook.Sheets("Shell").Select
Detruire MyWorkbook.Sheets("Shell")
MyWorkbook.Sheets("EQT").Select
Detruire MyWorkbook.Sheets("EQT")
MyWorkbook.Sheets("Wire").Select
Detruire yWorkbook.Sheets("Wire")
MyWorkbook.Sheets("WireGroup").Select
Detruire yWorkbook.Sheets("Wire")
MyWorkbook.Sheets("create").Select

Else
Exit Sub
End If
End Sub

Sub Detruire(MySeet As Worksheet)
    Rows("2:500").Select
    Selection.Delete Shift:=xlUp
    Cells(2, 1).Select
    
End Sub

Sub Clear_WorkSheet(MySeet As Worksheet)
Answer = MsgBox("Are you sure you want to clear the active WorkSheet?", vbYesNo)
If Answer = 6 Then
Detruire MySeet
End If
End Sub


Sub AddIS(MySeet As Worksheet)
MySeet.Select
For iRow = 2 To MySeet.Range("A1").CurrentRegion.Rows.Count
MySeet.Cells(iRow, 1).Select
    If (MySeet.Cells(iRow, 1) <> 0 And MySeet.Cells(iRow, 2) <> 0 And MySeet.Cells(iRow, 3) <> 0) Then
        jRow = iRow + 1
            Do While (MySeet.Cells(jRow, 4) <> 0)
                jRow = jRow + 1
            Loop
            If (MySeet.Cells(jRow, 4) = 0) Then
                MySeet.Range("A" & jRow, "F" & jRow).Interior.ColorIndex = 16
                
                If MySeet.Name = "IS" Then
                MySeet.Cells(jRow + 1, 1).Select
                End If
                
            End If
    Else
    If (((MySeet.Cells(iRow, 1) = 0 And MySeet.Cells(iRow, 2) <> 0 And MySeet.Cells(iRow, 3) <> 0)) Or ((MySeet.Cells(iRow, 1) <> 0 And MySeet.Cells(iRow, 2) = 0 And MySeet.Cells(iRow, 3) <> 0)) Or ((MySeet.Cells(iRow, 1) <> 0 And MySeet.Cells(iRow, 2) <> 0 And MySeet.Cells(iRow, 3) = 0)) Or ((MySeet.Cells(iRow, 1) <> 0 And MySeet.Cells(iRow, 2) = 0 And MySeet.Cells(iRow, 3) = 0)) Or ((MySeet.Cells(iRow, 1) = 0 And MySeet.Cells(iRow, 2) <> 0 And MySeet.Cells(iRow, 3) = 0)) Or ((MySeet.Cells(iRow, 1) = 0 And MySeet.Cells(iRow, 2) = 0 And MySeet.Cells(iRow, 3) <> 0))) Then
               MsgBox "There is an error on this line: " & iRow & vbCrLf & "At least one fill is missing.", vbInformation
        End If
    End If
Next
End Sub
Sub AddShell(MySeet As Worksheet)
MySeet.Select
Status = 1
For iRow = 2 To MySeet.Range("A1").CurrentRegion.Rows.Count
MySeet.Cells(iRow, 1).Select
    If (MySeet.Cells(iRow, 1) <> 0 And MySeet.Cells(iRow, 2) <> 0 And MySeet.Cells(iRow, 3) <> 0) Then
        jRow = iRow + 1
            Do While ((MySeet.Cells(jRow, 4) <> 0 Or MySeet.Cells(jRow, 5) <> 0 Or MySeet.Cells(jRow, 6) <> 0) Or MySeet.Cells(jRow, 7) <> 0)
                If (((MySeet.Cells(jRow, 4) = 0 And MySeet.Cells(jRow, 5) <> 0 And MySeet.Cells(jRow, 6) <> 0)) Or ((MySeet.Cells(jRow, 4) <> 0 And MySeet.Cells(jRow, 5) = 0 And MySeet.Cells(jRow, 6) <> 0)) Or ((MySeet.Cells(jRow, 4) <> 0 And MySeet.Cells(jRow, 5) <> 0 And MySeet.Cells(jRow, 6) = 0)) Or ((MySeet.Cells(jRow, 4) <> 0 And MySeet.Cells(jRow, 5) = 0 And MySeet.Cells(jRow, 6) = 0)) Or ((MySeet.Cells(jRow, 4) = 0 And MySeet.Cells(jRow, 5) <> 0 And MySeet.Cells(jRow, 6) = 0)) Or ((MySeet.Cells(jRow, 4) = 0 And MySeet.Cells(jRow, 5) = 0 And MySeet.Cells(jRow, 6) <> 0))) Then
                    MsgBox "There is an error on this line: " & jRow & vbCrLf & "At least one fill is missing.", vbInformation
                    Status = 0
                End If
                
                jRow = jRow + 1
            Loop

            If (Status = 1) Then
                MySeet.Range("A" & jRow, "I" & jRow).Interior.ColorIndex = 16
                
                If MySeet.Name = "Shell" Then
                MySeet.Cells(jRow + 1, 1).Select
                End If
                
            End If
    Else
        If (((MySeet.Cells(iRow, 1) = 0 And MySeet.Cells(iRow, 2) <> 0 And MySeet.Cells(iRow, 3) <> 0)) Or ((MySeet.Cells(iRow, 1) <> 0 And MySeet.Cells(iRow, 2) = 0 And MySeet.Cells(iRow, 3) <> 0)) Or ((MySeet.Cells(iRow, 1) <> 0 And MySeet.Cells(iRow, 2) <> 0 And MySeet.Cells(iRow, 3) = 0)) Or ((MySeet.Cells(iRow, 1) <> 0 And MySeet.Cells(iRow, 2) = 0 And MySeet.Cells(iRow, 3) = 0)) Or ((MySeet.Cells(iRow, 1) = 0 And MySeet.Cells(iRow, 2) <> 0 And MySeet.Cells(iRow, 3) = 0)) Or ((MySeet.Cells(iRow, 1) = 0 And MySeet.Cells(iRow, 2) = 0 And MySeet.Cells(iRow, 3) <> 0))) Then
        MsgBox "There is an error on this line: " & iRow & vbCrLf & "At least one fill is missing.", vbInformation
        End If
    End If
Next
End Sub

Sub UpdateIS(MySeet As Worksheet)


MySeet.Select

Dim Formule As String

For iRow = 2 To MySeet.Range("A1").CurrentRegion.Rows.Count
MySeet.Cells(iRow, 1).Select
    If MySeet.Cells(iRow, 1) <> 0 Then
        ISid = MySeet.Cells(iRow, 3)
    End If
    If MySeet.Cells(iRow, 4) <> 0 Then
        MySeet.Cells(iRow, 5) = ISid & "." & MySeet.Cells(iRow, 4)
        Formule = "=IF(ISERROR(VLOOKUP(E" & iRow & ",Wire!G:G,1,0)),IF(ISERROR(VLOOKUP(E" & iRow & ",Wire!H:H,1,0)),IF(ISERROR(VLOOKUP(E" & iRow & ",WireGroup!L:L,1,0)),IF(ISERROR(VLOOKUP(E" & iRow & ",WireGroup!M:M,1,0)),""N"",""Y""),""Y""),""Y""),""Y"")"
        MySeet.Cells(iRow, 6).Formula = Formule
        numIS = iRow
    End If
    If MySeet.Cells(iRow, 1) <> 0 Then
        MySeet.Cells(iRow, 5) = ISid
        Formule = "=IF(ISERROR(VLOOKUP(E" & iRow & ",Wire!G:G,1,0)),IF(ISERROR(VLOOKUP(E" & iRow & ",Wire!H:H,1,0)),IF(ISERROR(VLOOKUP(E" & iRow & ",WireGroup!L:L,1,0)),IF(ISERROR(VLOOKUP(E" & iRow & ",WireGroup!M:M,1,0)),""N"",""Y""),""Y""),""Y""),""Y"")"
        MySeet.Cells(iRow, 6).Formula = Formule
        numIS = iRow
    End If

Next

End Sub
Sub UpdateShell(MySeet As Worksheet)
MySeet.Select
Dim Formule As String

For iRow = 2 To MySeet.Range("A1").CurrentRegion.Rows.Count
MySeet.Cells(iRow, 1).Select
    If MySeet.Cells(iRow, 6) <> 0 Then
        SicId = MySeet.Cells(iRow, 6)
    End If
    If MySeet.Cells(iRow, 7) <> 0 Then
        MySeet.Cells(iRow, 8) = SicId & "." & MySeet.Cells(iRow, 7)
        Formule = "=IF(ISERROR(VLOOKUP(H" & iRow & ",Wire!G:G,1,0)),IF(ISERROR(VLOOKUP(H" & iRow & ",Wire!H:H,1,0)),IF(ISERROR(VLOOKUP(H" & iRow & ",WireGroup!L:L,1,0)),IF(ISERROR(VLOOKUP(H" & iRow & ",WireGroup!M:M,1,0)),""N"",""Y""),""Y""),""Y""),""Y"")"
        MySeet.Cells(iRow, 9).Formula = Formule
        numS = iRow
    End If
    If MySeet.Cells(iRow, 6) <> 0 Then
        MySeet.Cells(iRow, 8) = SicId
        Formule = "=IF(ISERROR(VLOOKUP(H" & iRow & ",Wire!G:G,1,0)),IF(ISERROR(VLOOKUP(H" & iRow & ",Wire!H:H,1,0)),IF(ISERROR(VLOOKUP(H" & iRow & ",WireGroup!L:L,1,0)),IF(ISERROR(VLOOKUP(H" & iRow & ",WireGroup!M:M,1,0)),""N"",""Y""),""Y""),""Y""),""Y"")"
        MySeet.Cells(iRow, 9).Formula = Formule
        numS = iRow
    End If
    If MySeet.Cells(iRow, 3) <> 0 Then
        numS = iRow
    End If

Next

End Sub

Sub AddEQT(MySeet As Worksheet)
'MySeet.Select
Status = 1
For iRow = 2 To MySeet.Range("A1").CurrentRegion.Rows.Count
MySeet.Cells(iRow, 1).Select
    If (MySeet.Cells(iRow, 1) <> 0 And MySeet.Cells(iRow, 2) <> 0 And MySeet.Cells(iRow, 3) <> 0) Then
        jRow = iRow + 1
            Do While ((MySeet.Cells(jRow, 4) <> 0 Or MySeet.Cells(jRow, 5) <> 0 Or MySeet.Cells(jRow, 6) <> 0) Or MySeet.Cells(jRow, 7) <> 0)
                If (((MySeet.Cells(jRow, 4) = 0 And MySeet.Cells(jRow, 5) <> 0 And MySeet.Cells(jRow, 6) <> 0)) Or ((MySeet.Cells(jRow, 4) <> 0 And MySeet.Cells(jRow, 5) = 0 And MySeet.Cells(jRow, 6) <> 0)) Or ((MySeet.Cells(jRow, 4) <> 0 And MySeet.Cells(jRow, 5) <> 0 And MySeet.Cells(jRow, 6) = 0)) Or ((MySeet.Cells(jRow, 4) <> 0 And MySeet.Cells(jRow, 5) = 0 And MySeet.Cells(jRow, 6) = 0)) Or ((MySeet.Cells(jRow, 4) = 0 And MySeet.Cells(jRow, 5) <> 0 And MySeet.Cells(jRow, 6) = 0)) Or ((MySeet.Cells(jRow, 4) = 0 And MySeet.Cells(jRow, 5) = 0 And MySeet.Cells(jRow, 6) <> 0))) Then
                    MsgBox "There is an error on this line: " & jRow & vbCrLf & "At least one fill is missing.", vbInformation
                    Status = 0
                End If
                
                jRow = jRow + 1
            Loop

            If (Status = 1) Then
                MySeet.Range("A" & jRow, "I" & jRow).Interior.ColorIndex = 16
                
                If MySeet.Name = "EQT" Then
                MySeet.Cells(jRow + 1, 1).Select
                End If
            End If
    Else
        If (((MySeet.Cells(iRow, 1) = 0 And MySeet.Cells(iRow, 2) <> 0 And MySeet.Cells(iRow, 3) <> 0)) Or ((MySeet.Cells(iRow, 1) <> 0 And MySeet.Cells(iRow, 2) = 0 And MySeet.Cells(iRow, 3) <> 0)) Or ((MySeet.Cells(iRow, 1) <> 0 And MySeet.Cells(iRow, 2) <> 0 And MySeet.Cells(iRow, 3) = 0)) Or ((MySeet.Cells(iRow, 1) <> 0 And MySeet.Cells(iRow, 2) = 0 And MySeet.Cells(iRow, 3) = 0)) Or ((MySeet.Cells(iRow, 1) = 0 And MySeet.Cells(iRow, 2) <> 0 And MySeet.Cells(iRow, 3) = 0)) Or ((MySeet.Cells(iRow, 1) = 0 And MySeet.Cells(iRow, 2) = 0 And MySeet.Cells(iRow, 3) <> 0))) Then
        MsgBox "There is an error on this line: " & iRow & vbCrLf & "At least one fill is missing.", vbInformation
        End If
    End If
Next
End Sub
Sub AddWGFather(MySeet As Worksheet)
'MySeet.Select
For iRow = 2 To MySeet.Range("A1").CurrentRegion.Rows.Count
MySeet.Cells(iRow, 1).Select
    Do While (MySeet.Cells(iRow, 1) <> "" And MySeet.Cells(iRow, 2) <> "" And MySeet.Cells(iRow, 3) <> "" And MySeet.Cells(iRow, 4) <> "" And MySeet.Cells(iRow, 5) <> "")
    iRow = iRow + 1
    Loop
    Do While (MySeet.Cells(iRow, 6) <> "" And MySeet.Cells(iRow, 7) <> "" And MySeet.Cells(iRow, 8) <> "" And MySeet.Cells(iRow, 9) <> "" And MySeet.Cells(iRow, 10) <> "" And MySeet.Cells(iRow, 11) <> "")
    iRow = iRow + 1
    Loop
If ((iRow <> 2 And MySeet.Cells(iRow, 1) = "" And MySeet.Cells(iRow, 2) = "" And MySeet.Cells(iRow, 3) = "" And MySeet.Cells(iRow, 4) = "" And MySeet.Cells(iRow, 5) = "" And MySeet.Cells(iRow, 6) = "" And MySeet.Cells(iRow, 7) = "" And MySeet.Cells(iRow, 8) = "" And MySeet.Cells(iRow, 9) = "" And MySeet.Cells(iRow, 10) = "" And MySeet.Cells(iRow, 11) = "" And MySeet.Cells(iRow, 12) = "" And MySeet.Cells(iRow, 13) = "") And _
(MySeet.Cells(iRow - 1, 6) <> "" And MySeet.Cells(iRow - 1, 7) <> "" And MySeet.Cells(iRow - 1, 8) <> "" And MySeet.Cells(iRow - 1, 9) <> "" And MySeet.Cells(iRow - 1, 10) <> "" And MySeet.Cells(iRow - 1, 11) <> "")) Then
MySeet.Range("A" & iRow, "M" & iRow).Interior.ColorIndex = 16

If MySeet.Name = "WireGroup" Then
MySeet.Cells(iRow + 1, 1).Select
End If

End If
Next
     
End Sub

Public Sub CheckI(MyClasseur As Workbook)
Dim I As Long
Dim J As Long

    numST = 0
    Tocount = 0
    NBERRORS = 0
    NBERRORSW = 0

    MyClasseur.Sheets("Create").Cells(7, 2).Value = "Prepring to analyse"
    
    MyClasseur.Sheets("SIC-TERM").Range("F2:F500").Interior.ColorIndex = xlNone
    AddSic MyClasseur.Sheets("SIC-TERM")
    UpdateSIC MyClasseur.Sheets("SIC-TERM")
    
        MyClasseur.Sheets("Create").Cells(7, 2).Value = 1 / 66
    MyClasseur.Sheets("SIC-CONT").Range("h2:h500").Interior.ColorIndex = xlNone
    AddContacts MyClasseur.Sheets("SIC-CONT")
    UpdateContacts MyClasseur.Sheets("SIC-CONT")

     MyClasseur.Sheets("IS").Range("e2:e500").Interior.ColorIndex = xlNone
    AddIS MyClasseur.Sheets("IS")
    UpdateIS MyClasseur.Sheets("IS")
        MyClasseur.Sheets("Create").Cells(7, 2).Value = 3 / 66
    MyClasseur.Sheets("Shell").Range("H2:H500").Interior.ColorIndex = xlNone
    MyClasseur.Sheets("Shell").Range("c2:c500").Interior.ColorIndex = xlNone
    AddShell MyClasseur.Sheets("Shell")
    UpdateShell MyClasseur.Sheets("Shell")
        MyClasseur.Sheets("Create").Cells(7, 2).Value = 4 / 66
    MyClasseur.Sheets("EQT").Range("c2:c500").Interior.ColorIndex = xlNone
    MyClasseur.Sheets("EQT").Range("f2:g500").Interior.ColorIndex = xlNone
    AddEQT MyClasseur.Sheets("EQT")

    MyClasseur.Sheets("wire").Range("a2:a500").Interior.ColorIndex = xlNone
    MyClasseur.Sheets("wire").Range("g2:g500").Interior.ColorIndex = xlNone
    MyClasseur.Sheets("wire").Range("h2:h500").Interior.ColorIndex = xlNone
        MyClasseur.Sheets("Create").Cells(7, 2).Value = 6 / 66
    MyClasseur.Sheets("wireGroup").Range("a2:a500").Interior.ColorIndex = xlNone
    MyClasseur.Sheets("wireGroup").Range("f2:f500").Interior.ColorIndex = xlNone
    MyClasseur.Sheets("wireGroup").Range("l2:l500").Interior.ColorIndex = xlNone
    MyClasseur.Sheets("wireGroup").Range("m2:m500").Interior.ColorIndex = xlNone
    AddWGFather MyClasseur.Sheets("wireGroup")
    
    MyClasseur.Sheets("Create").Cells(7, 2).Value = 9 / 66
   
   If numST = Empty Then ' on est dans le cas ou il n'y a pas de SIC et qu'on veut copmpter le nbre de composants
   numST = 2
   Tocount = 1
   End If
   
   On Error GoTo Fin
        For I = 2 To numST
           DoEvents
            If MyClasseur.Sheets("SIC-TERM").Cells(I, 6).Value <> "" Or I = 2 Then
                        
            refdesvalue = MyClasseur.Sheets("SIC-TERM").Cells(I, 6)
            
                'SIC-TERM DEBUT
                    For J = 2 To numST
                        DoEvents
                        If refdesvalue = MyClasseur.Sheets("SIC-TERM").Cells(J, 6) And J <> I Then
                            MyClasseur.Sheets("SIC-TERM").Cells(I, 6).Interior.ColorIndex = 3
                            MyClasseur.Sheets("SIC-TERM").Cells(J, 6).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'SIC-TERM FIN
                'SIC-CONT DEBUT
                    For J = 2 To numSC
                       DoEvents
                        If refdesvalue = MyClasseur.Sheets("SIC-CONT").Cells(J, 8) Then
                            MyClasseur.Sheets("SIC-TERM").Cells(I, 6).Interior.ColorIndex = 3
                            MyClasseur.Sheets("SIC-CONT").Cells(J, 8).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'SIC-CONT FIN
                'IS DEBUT
                    For J = 2 To numIS
                       DoEvents
                        If refdesvalue = MyClasseur.Sheets("IS").Cells(J, 5) Then
                            MyClasseur.Sheets("SIC-TERM").Cells(I, 6).Interior.ColorIndex = 3
                            MyClasseur.Sheets("IS").Cells(J, 5).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'IS FIN
                'SHELL DEBUT
                    For J = 2 To numS
                       DoEvents
                        If refdesvalue = MyClasseur.Sheets("Shell").Cells(J, 8) Then
                            MyClasseur.Sheets("SIC-TERM").Cells(I, 6).Interior.ColorIndex = 3
                            MyClasseur.Sheets("Shell").Cells(J, 8).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                        If refdesvalue = MyClasseur.Sheets("Shell").Cells(J, 3) Then
                            MyClasseur.Sheets("SIC-TERM").Cells(I, 6).Interior.ColorIndex = 3
                            MyClasseur.Sheets("Shell").Cells(J, 3).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If

                    Next
                'SHELLFIN
                'EQT DEBUT Sheets("EQT")
                    For J = 2 To MyClasseur.Sheets("EQT").Range("A1").CurrentRegion.Rows.Count
                           DoEvents
                        If MyClasseur.Sheets("EQT").Cells(I, 1).Value <> "" Or MyClasseur.Sheets("EQT").Cells(I, 6).Value <> "" Or MyClasseur.Sheets("EQT").Cells(I, 7).Value <> "" Then
                            numE = J
                        End If
                    
                    
                        If refdesvalue = MyClasseur.Sheets("EQT").Cells(J, 3) And Tocount <> 1 Then
                            MyClasseur.Sheets("SIC-TERM").Cells(I, 6).Interior.ColorIndex = 3
                            MyClasseur.Sheets("EQT").Cells(J, 3).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                        If refdesvalue = MyClasseur.Sheets("EQT").Cells(J, 6) And Tocount <> 1 Then
                            MyClasseur.Sheets("SIC-TERM").Cells(I, 6).Interior.ColorIndex = 3
                            MyClasseur.Sheets("EQT").Cells(J, 6).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'EQT FIN
                 'WIRE DEBUT
                    For J = 2 To MyClasseur.Sheets("Wire").Range("A1").CurrentRegion.Rows.Count
                       DoEvents
                        If MyClasseur.Sheets("Wire").Cells(J, 1).Value <> "" Then
                            numW = J
                        End If
                        If refdesvalue = MyClasseur.Sheets("Wire").Cells(J, 1) And Tocount <> 1 Then
                            MyClasseur.Sheets("SIC-TERM").Cells(I, 6).Interior.ColorIndex = 3
                            MyClasseur.Sheets("Wire").Cells(J, 1).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'WIRE FIN
                'WIRE GROUP DEBUT
                    For J = 2 To MyClasseur.Sheets("WireGroup").Range("A1").CurrentRegion.Rows.Count
                           DoEvents
                        If MyClasseur.Sheets("WireGroup").Cells(J, 6).Value <> "" Or MyClasseur.Sheets("WireGroup").Cells(J, 1).Value <> "" Then
                            numWG = J
                        End If
                    
                        If refdesvalue = MyClasseur.Sheets("WireGroup").Cells(J, 1) And Tocount <> 1 Then
                            MyClasseur.Sheets("SIC-TERM").Cells(I, 6).Interior.ColorIndex = 3
                            MyClasseur.Sheets("WireGroup").Cells(J, 1).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                        If refdesvalue = MyClasseur.Sheets("WireGroup").Cells(J, 6) And Tocount <> 1 Then
                            MyClasseur.Sheets("SIC-TERM").Cells(I, 6).Interior.ColorIndex = 3
                            MyClasseur.Sheets("WireGroup").Cells(J, 6).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'WIREGROUP FIN
             End If
        Next

    MyClasseur.Sheets("Create").Cells(7, 2).Value = 19 / 66
    
        For I = 2 To numSC

        
            If MyClasseur.Sheets("SIC-CONT").Cells(I, 8).Value <> "" Then

            refdesvalue = MyClasseur.Sheets("SIC-CONT").Cells(I, 8)
            

                'SIC-CONT DEBUT
                    For J = 2 To numSC
                        If refdesvalue = MyClasseur.Sheets("SIC-CONT").Cells(J, 8) And J <> I Then
                            MyClasseur.Sheets("SIC-CONT").Cells(I, 8).Interior.ColorIndex = 3
                            MyClasseur.Sheets("SIC-CONT").Cells(J, 8).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'SIC-CONT FIN
                'IS DEBUT
                    For J = 2 To numIS
                        If refdesvalue = MyClasseur.Sheets("IS").Cells(J, 5) Then
                            MyClasseur.Sheets("SIC-CONT").Cells(I, 8).Interior.ColorIndex = 3
                            MyClasseur.Sheets("IS").Cells(J, 5).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'IS FIN
                'SHELL DEBUT
                    For J = 2 To numS
                        If refdesvalue = MyClasseur.Sheets("Shell").Cells(J, 8) Then
                            MyClasseur.Sheets("SIC-CONT").Cells(I, 6).Interior.ColorIndex = 3
                            MyClasseur.Sheets("Shell").Cells(J, 8).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                        If refdesvalue = MyClasseur.Sheets("Shell").Cells(J, 3) Then
                            MyClasseur.Sheets("SIC-CONT").Cells(I, 6).Interior.ColorIndex = 3
                            MyClasseur.Sheets("Shell").Cells(J, 3).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If

                    Next
                'SHELLFIN
                'EQT DEBUT
                    For J = 2 To numE
                        If refdesvalue = MyClasseur.Sheets("EQT").Cells(J, 3) Then
                            MyClasseur.Sheets("SIC-CONT").Cells(I, 8).Interior.ColorIndex = 3
                            MyClasseur.Sheets("EQT").Cells(J, 3).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                        If refdesvalue = MyClasseur.Sheets("EQT").Cells(J, 6) Then
                            MyClasseur.Sheets("SIC-CONT").Cells(I, 8).Interior.ColorIndex = 3
                            MyClasseur.Sheets("EQT").Cells(J, 6).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'EQT FIN
                 'WIRE DEBUT
                    For J = 2 To numW
                        If refdesvalue = MyClasseur.Sheets("WIRE").Cells(J, 1) Then
                            MyClasseur.Sheets("SIC-CONT").Cells(I, 8).Interior.ColorIndex = 3
                            MyClasseur.Sheets("WIRE").Cells(J, 1).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'WIRE FIN
                'WIRE GROUP DEBUT
                    For J = 2 To numWG
                        If refdesvalue = MyClasseur.Sheets("WireGroup").Cells(J, 1) Then
                            MyClasseur.Sheets("SIC-CONT").Cells(I, 8).Interior.ColorIndex = 3
                            MyClasseur.Sheets("WireGroup").Cells(J, 1).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                        If refdesvalue = MyClasseur.Sheets("WireGroup").Cells(J, 6) Then
                            MyClasseur.Sheets("SIC-CONT").Cells(I, 8).Interior.ColorIndex = 3
                            MyClasseur.Sheets("WireGroup").Cells(J, 6).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'WIREGROUP FIN
             End If
        Next

MyClasseur.Sheets("Create").Cells(7, 2).Value = 28 / 66

        For I = 2 To numIS

        
            If MyClasseur.Sheets("IS").Cells(I, 5).Value <> "" Then

            refdesvalue = MyClasseur.Sheets("IS").Cells(I, 5)
            
                
                'IS DEBUT
                    For J = 2 To numIS
                        If refdesvalue = MyClasseur.Sheets("IS").Cells(J, 5) And J <> I Then
                            MyClasseur.Sheets("IS").Cells(I, 5).Interior.ColorIndex = 3
                            MyClasseur.Sheets("IS").Cells(J, 5).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'IS FIN
                'SHELL DEBUT
                    For J = 2 To numS
                        If refdesvalue = MyClasseur.Sheets("Shell").Cells(J, 8) Then
                            MyClasseur.Sheets("IS").Cells(I, 5).Interior.ColorIndex = 3
                            MyClasseur.Sheets("Shell").Cells(J, 8).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                        If refdesvalue = MyClasseur.Sheets("Shell").Cells(J, 3) Then
                            MyClasseur.Sheets("IS").Cells(I, 5).Interior.ColorIndex = 3
                            MyClasseur.Sheets("Shell").Cells(J, 3).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If

                    Next
                'SHELLFIN
                'EQT DEBUT
                    For J = 2 To numE
                        If refdesvalue = MyClasseur.Sheets("EQT").Cells(J, 3) Then
                            MyClasseur.Sheets("IS").Cells(I, 5).Interior.ColorIndex = 3
                            MyClasseur.Sheets("EQT").Cells(J, 3).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                        If refdesvalue = MyClasseur.Sheets("EQT").Cells(J, 6) Then
                            MyClasseur.Sheets("IS").Cells(I, 5).Interior.ColorIndex = 3
                            MyClasseur.Sheets("EQT").Cells(J, 6).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'EQT FIN
                 'WIRE DEBUT
                    For J = 2 To numW
                        If refdesvalue = MyClasseur.Sheets("WIRE").Cells(J, 1) Then
                            MyClasseur.Sheets("IS").Cells(I, 5).Interior.ColorIndex = 3
                            MyClasseur.Sheets("WIRE").Cells(J, 1).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'WIRE FIN
                'WIRE GROUP DEBUT
                    For J = 2 To numWG
                        If refdesvalue = MyClasseur.Sheets("WireGroup").Cells(J, 1) Then
                            MyClasseur.Sheets("IS").Cells(I, 5).Interior.ColorIndex = 3
                            MyClasseur.Sheets("WireGroup").Cells(J, 1).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                        If refdesvalue = MyClasseur.Sheets("WireGroup").Cells(J, 6) Then
                            MyClasseur.Sheets("IS").Cells(I, 5).Interior.ColorIndex = 3
                            MyClasseur.Sheets("WireGroup").Cells(J, 6).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'WIREGROUP FIN
             End If
        Next

MyClasseur.Sheets("Create").Cells(7, 2).Value = 36 / 66

For I = 2 To numS

        
            If MyClasseur.Sheets("Shell").Cells(I, 8).Value <> "" Then

            refdesvalue = MyClasseur.Sheets("Shell").Cells(I, 8)
            
                
                'SHELL DEBUT
                    For J = 2 To numS
                        If refdesvalue = MyClasseur.Sheets("Shell").Cells(J, 8) And J <> I Then
                            MyClasseur.Sheets("Shell").Cells(I, 8).Interior.ColorIndex = 3
                            MyClasseur.Sheets("Shell").Cells(J, 8).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                        If refdesvalue = MyClasseur.Sheets("Shell").Cells(J, 3) Then
                            MyClasseur.Sheets("Shell").Cells(I, 8).Interior.ColorIndex = 3
                            MyClasseur.Sheets("Shell").Cells(J, 3).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
  
                    Next
                'SHELLFIN
                'EQT DEBUT
                    For J = 2 To numE
                        If refdesvalue = MyClasseur.Sheets("EQT").Cells(J, 3) Then
                            MyClasseur.Sheets("Shell").Cells(I, 8).Interior.ColorIndex = 3
                            MyClasseur.Sheets("EQT").Cells(J, 3).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                        If refdesvalue = MyClasseur.Sheets("EQT").Cells(J, 6) Then
                            MyClasseur.Sheets("Shell").Cells(I, 8).Interior.ColorIndex = 3
                            MyClasseur.Sheets("EQT").Cells(J, 6).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'EQT FIN
                 'WIRE DEBUT
                    For J = 2 To numW
                        If refdesvalue = MyClasseur.Sheets("WIRE").Cells(J, 1) Then
                            MyClasseur.Sheets("Shell").Cells(I, 8).Interior.ColorIndex = 3
                            MyClasseur.Sheets("WIRE").Cells(J, 1).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'WIRE FIN
                'WIRE GROUP DEBUT
                    For J = 2 To numWG
                        If refdesvalue = MyClasseur.Sheets("WireGroup").Cells(J, 1) Then
                            MyClasseur.Sheets("Shell").Cells(I, 8).Interior.ColorIndex = 3
                            MyClasseur.Sheets("WireGroup").Cells(J, 1).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                        If refdesvalue = MyClasseur.Sheets("WireGroup").Cells(J, 6) Then
                            MyClasseur.Sheets("Shell").Cells(I, 8).Interior.ColorIndex = 3
                            MyClasseur.Sheets("WireGroup").Cells(J, 6).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'WIREGROUP FIN
             End If
        Next


MyClasseur.Sheets("Create").Cells(7, 2).Value = 43 / 66

For I = 2 To numS

        
            If MyClasseur.Sheets("Shell").Cells(I, 3).Value <> "" Then
            refdesvalue = MyClasseur.Sheets("Shell").Cells(I, 3)
            
                
                'SHELL DEBUT
                    For J = 2 To numS
                        
                        If refdesvalue = MyClasseur.Sheets("Shell").Cells(J, 3) And J <> I Then
                            MyClasseur.Sheets("Shell").Cells(I, 3).Interior.ColorIndex = 3
                            MyClasseur.Sheets("Shell").Cells(J, 3).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If

                    Next
                'SHELLFIN
                'EQT DEBUT
                    For J = 2 To numE
                        If refdesvalue = MyClasseur.Sheets("EQT").Cells(J, 3) Then
                            MyClasseur.Sheets("Shell").Cells(I, 3).Interior.ColorIndex = 3
                            MyClasseur.Sheets("EQT").Cells(J, 3).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                        If refdesvalue = MyClasseur.Sheets("EQT").Cells(J, 6) Then
                            MyClasseur.Sheets("Shell").Cells(I, 3).Interior.ColorIndex = 3
                            MyClasseur.Sheets("EQT").Cells(J, 6).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'EQT FIN
                 'WIRE DEBUT
                    For J = 2 To numW
                        If refdesvalue = MyClasseur.Sheets("WIRE").Cells(J, 1) Then
                            MyClasseur.Sheets("Shell").Cells(I, 3).Interior.ColorIndex = 3
                            MyClasseur.Sheets("WIRE").Cells(J, 1).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'WIRE FIN
                'WIRE GROUP DEBUT
                    For J = 2 To numWG
                        If refdesvalue = MyClasseur.Sheets("WireGroup").Cells(J, 1) Then
                            MyClasseur.Sheets("Shell").Cells(I, 3).Interior.ColorIndex = 3
                            MyClasseur.Sheets("WireGroup").Cells(J, 1).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                        If refdesvalue = MyClasseur.Sheets("WireGroup").Cells(J, 6) Then
                            MyClasseur.Sheets("Shell").Cells(I, 3).Interior.ColorIndex = 3
                            MyClasseur.Sheets("WireGroup").Cells(J, 6).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'WIREGROUP FIN
             End If
        Next




MyClasseur.Sheets("Create").Cells(7, 2).Value = 49 / 66

For I = 2 To numE

        
            If MyClasseur.Sheets("EQT").Cells(I, 3).Value <> "" Then
            refdesvalue = MyClasseur.Sheets("EQT").Cells(I, 3)
            
                
                'EQT DEBUT
                    For J = 2 To numE
                        If refdesvalue = MyClasseur.Sheets("EQT").Cells(J, 3) And J <> I Then
                            MyClasseur.Sheets("EQT").Cells(I, 3).Interior.ColorIndex = 3
                            MyClasseur.Sheets("EQT").Cells(J, 3).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                        If refdesvalue = MyClasseur.Sheets("EQT").Cells(J, 6) Then
                            MyClasseur.Sheets("EQT").Cells(I, 3).Interior.ColorIndex = 3
                            MyClasseur.Sheets("EQT").Cells(J, 6).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'EQT FIN
                 'WIRE DEBUT
                    For J = 2 To numW
                        If refdesvalue = MyClasseur.Sheets("WIRE").Cells(J, 1) Then
                            MyClasseur.Sheets("EQT").Cells(I, 3).Interior.ColorIndex = 3
                            MyClasseur.Sheets("WIRE").Cells(J, 1).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'WIRE FIN
                'WIRE GROUP DEBUT
                    For J = 2 To numWG
                        If refdesvalue = MyClasseur.Sheets("WireGroup").Cells(J, 1) Then
                            MyClasseur.Sheets("EQT").Cells(I, 3).Interior.ColorIndex = 3
                            MyClasseur.Sheets("WireGroup").Cells(J, 1).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                        If refdesvalue = MyClasseur.Sheets("WireGroup").Cells(J, 6) Then
                            MyClasseur.Sheets("EQT").Cells(I, 3).Interior.ColorIndex = 3
                            MyClasseur.Sheets("WireGroup").Cells(J, 6).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'WIREGROUP FIN
             End If
        Next

MyClasseur.Sheets("Create").Cells(7, 2).Value = 54 / 66

For I = 2 To numE

        
            If MyClasseur.Sheets("EQT").Cells(I, 6).Value <> "" Then

            refdesvalue = MyClasseur.Sheets("EQT").Cells(I, 6)
            
                
                'EQT DEBUT
                    For J = 2 To numE
                        
                        If refdesvalue = MyClasseur.Sheets("EQT").Cells(J, 6) And J <> I Then
                            MyClasseur.Sheets("EQT").Cells(I, 6).Interior.ColorIndex = 3
                            MyClasseur.Sheets("EQT").Cells(J, 6).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'EQT FIN
                 'WIRE DEBUT
                    For J = 2 To numW
                        If refdesvalue = MyClasseur.Sheets("WIRE").Cells(J, 1) Then
                            MyClasseur.Sheets("EQT").Cells(I, 6).Interior.ColorIndex = 3
                            MyClasseur.Sheets("WIRE").Cells(J, 1).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'WIRE FIN
                'WIRE GROUP DEBUT
                    For J = 2 To numWG
                        If refdesvalue = MyClasseur.Sheets("WireGroup").Cells(J, 1) Then
                            MyClasseur.Sheets("EQT").Cells(I, 6).Interior.ColorIndex = 3
                            MyClasseur.Sheets("WireGroup").Cells(J, 1).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                        If refdesvalue = MyClasseur.Sheets("WireGroup").Cells(J, 6) Then
                            MyClasseur.Sheets("EQT").Cells(I, 6).Interior.ColorIndex = 3
                            MyClasseur.Sheets("WireGroup").Cells(J, 6).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'WIREGROUP FIN
             End If
        Next

MyClasseur.Sheets("Create").Cells(7, 2).Value = 58 / 66

For I = 2 To numW

        
            If MyClasseur.Sheets("WIRE").Cells(I, 1).Value <> "" Then
            numW = I
            refdesvalue = MyClasseur.Sheets("wire").Cells(I, 1)
            
               
                 'WIRE DEBUT
                    For J = 2 To numW
                        If refdesvalue = MyClasseur.Sheets("WIRE").Cells(J, 1) And J <> I Then
                            MyClasseur.Sheets("WIRE").Cells(I, 1).Interior.ColorIndex = 3
                            MyClasseur.Sheets("WIRE").Cells(J, 1).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'WIRE FIN
                'WIRE GROUP DEBUT
                    For J = 2 To numWG
                        If refdesvalue = MyClasseur.Sheets("WireGroup").Cells(J, 1) Then
                            MyClasseur.Sheets("WIRE").Cells(I, 1).Interior.ColorIndex = 3
                            MyClasseur.Sheets("WireGroup").Cells(J, 1).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                        If refdesvalue = MyClasseur.Sheets("WireGroup").Cells(J, 6) Then
                            MyClasseur.Sheets("WIRE").Cells(I, 1).Interior.ColorIndex = 3
                            MyClasseur.Sheets("WireGroup").Cells(J, 6).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'WIREGROUP FIN
             End If
        Next

MyClasseur.Sheets("Create").Cells(7, 2).Value = 61 / 66

For I = 2 To numWG

        
            If MyClasseur.Sheets("WireGroup").Cells(I, 1).Value <> "" Then
            
            refdesvalue = MyClasseur.Sheets("WireGroup").Cells(I, 1)
            
                
                'WIRE GROUP DEBUT
                    For J = 2 To numWG
                        If refdesvalue = MyClasseur.Sheets("WireGroup").Cells(J, 1) And J <> I Then
                            MyClasseur.Sheets("WireGroup").Cells(I, 1).Interior.ColorIndex = 3
                            MyClasseur.Sheets("WireGroup").Cells(J, 1).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                        If refdesvalue = MyClasseur.Sheets("WireGroup").Cells(J, 6) Then
                            MyClasseur.Sheets("WireGroup").Cells(I, 1).Interior.ColorIndex = 3
                            MyClasseur.Sheets("WireGroup").Cells(J, 6).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS
                        End If
                    Next
                'WIREGROUP FIN
             End If
        Next

MyClasseur.Sheets("Create").Cells(7, 2).Value = 63 / 66

For I = 2 To numWG
        
            If MyClasseur.Sheets("WireGroup").Cells(I, 6).Value <> "" Then
            refdesvalue = MyClasseur.Sheets("WireGroup").Cells(I, 6)
            
                
                'WIRE GROUP DEBUT
                    For J = 2 To numWG
                        If refdesvalue = MyClasseur.Sheets("WireGroup").Cells(J, 6) And J <> I Then
                            MyClasseur.Sheets("WireGroup").Cells(I, 6).Interior.ColorIndex = 3
                            MyClasseur.Sheets("WireGroup").Cells(J, 6).Interior.ColorIndex = 3
                            NBERRORS = NBERRORS + 1
                        End If
                    Next
                'WIREGROUP FIN
             End If
        Next

MyClasseur.Sheets("Create").Cells(7, 2).Value = 64 / 66

    Dim numForDevices As Integer
    numForWires = EXCEL.WorksheetFunction.Max(numWG, numW)
    numForDevices = EXCEL.WorksheetFunction.Max(numST, numSC, numIS, numS)
  


For I = 2 To numForWires

        Set C_W0 = MyClasseur.Sheets("wire").Cells(I, 1)
        Set C_W1 = MyClasseur.Sheets("wire").Cells(I, 7)
        Set C_W2 = MyClasseur.Sheets("wire").Cells(I, 8)
        Set C_WGM = MyClasseur.Sheets("WireGroup").Cells(I, 1)
        Set C_WGN = MyClasseur.Sheets("WireGroup").Cells(I + 1, 1)
        Set c_wgn0 = MyClasseur.Sheets("WireGroup").Cells(I + 1, 6)
        Set C_WG0 = MyClasseur.Sheets("WireGroup").Cells(I, 6)
        Set C_WG1 = MyClasseur.Sheets("WireGroup").Cells(I, 12)
        Set C_WG2 = MyClasseur.Sheets("WireGroup").Cells(I, 13)
        
        
        W0 = C_W0.Value
        W1 = C_W1.Value
        W2 = C_W2.Value
        WGM = C_WGM.Value
        WGN = C_WGN.Value
        WGN0 = c_wgn0.Value
        WG0 = C_WG0.Value
        WG1 = C_WG1.Value
        WG2 = C_WG2.Value
        
           'SIC-TERM DEBUT
                    For J = 2 To numForDevices
                        RV = MyClasseur.Sheets("SIC-TERM").Cells(J, 6).Value
                            If RV = W1 And RV <> "" And W1 <> "" Then
                                C_W1.Interior.ColorIndex = 4
                            End If
                            If RV = W2 And RV <> "" And W2 <> "" Then
                                C_W2.Interior.ColorIndex = 4
                            End If
                            If RV = WG1 And RV <> "" And WG1 <> "" Then
                                C_WG1.Interior.ColorIndex = 4
                            End If
                            If RV = WG2 And RV <> "" And WG2 <> "" Then
                                C_WG2.Interior.ColorIndex = 4
                            End If

            'SIC-TERM FIN
            'SIC-CONT DEBUT
     
                        RV = MyClasseur.Sheets("SIC-CONT").Cells(J, 8).Value
                            If RV = W1 And RV <> "" And W1 <> "" Then
                                C_W1.Interior.ColorIndex = 4
                            End If
                            If RV = W2 And RV <> "" And W2 <> "" Then
                                C_W2.Interior.ColorIndex = 4
                            End If
                            If RV = WG1 And RV <> "" And WG1 <> "" Then
                                C_WG1.Interior.ColorIndex = 4
                            End If
                            If RV = WG2 And RV <> "" And WG2 <> "" Then
                                C_WG2.Interior.ColorIndex = 4
                            End If
   
                'SIC-CONT FIN
                'IS DEBUT

                        RV = MyClasseur.Sheets("IS").Cells(J, 5).Value
                            If RV = W1 And RV <> "" And W1 <> "" Then
                                C_W1.Interior.ColorIndex = 4
                            End If
                            If RV = W2 And RV <> "" And W2 <> "" Then
                                C_W2.Interior.ColorIndex = 4
                            End If
                            If RV = WG1 And RV <> "" And WG1 <> "" Then
                                C_WG1.Interior.ColorIndex = 4
                            End If
                            If RV = WG2 And RV <> "" And WG2 <> "" Then
                                C_WG2.Interior.ColorIndex = 4
                            End If

                'IS FIN
                'SHELL DEBUT

                        RV = MyClasseur.Sheets("Shell").Cells(J, 8).Value
                            If RV = W1 And RV <> "" And W1 <> "" Then
                                C_W1.Interior.ColorIndex = 4
                            End If
                            If RV = W2 And RV <> "" And W2 <> "" Then
                                C_W2.Interior.ColorIndex = 4
                            End If
                            If RV = WG1 And RV <> "" And WG1 <> "" Then
                                C_WG1.Interior.ColorIndex = 4
                            End If
                            If RV = WG2 And RV <> "" And WG2 <> "" Then
                                C_WG2.Interior.ColorIndex = 4
                            End If

                'SHELLFIN

        Next
        
     If C_W1.Interior.ColorIndex <> 4 And W1 <> "" Then
         NBERRORSW = NBERRORSW + 1
         C_W1.Interior.ColorIndex = 3
         Else
         C_W1.Interior.ColorIndex = 0
     End If
     If C_W2.Interior.ColorIndex <> 4 And W2 <> "" Then
         NBERRORSW = NBERRORSW + 1
         C_W2.Interior.ColorIndex = 3
         Else
         C_W2.Interior.ColorIndex = 0
     End If
      If C_WG1.Interior.ColorIndex <> 4 And WG1 <> "" Then
         NBERRORSW = NBERRORSW + 1
         C_WG1.Interior.ColorIndex = 3
         Else
         C_WG1.Interior.ColorIndex = 0
     End If
      If C_WG2.Interior.ColorIndex <> 4 And WG2 <> "" Then
         NBERRORSW = NBERRORSW + 1
         C_WG2.Interior.ColorIndex = 3
         Else
         C_WG2.Interior.ColorIndex = 0
     End If
     
     
     If W0 <> "" Then
        If W1 = "" Then
            NBERRORSW = NBERRORSW + 1
            C_W1.Interior.ColorIndex = 3
        End If
        If W2 = "" Then
            NBERRORSW = NBERRORSW + 1
            C_W2.Interior.ColorIndex = 3
        End If
     End If
     
     If WG0 <> "" Then
       If WG1 = "" Then
           NBERRORSW = NBERRORSW + 1
           C_WG1.Interior.ColorIndex = 3
       End If
       If WG2 = "" Then
           NBERRORSW = NBERRORSW + 1
           C_WG2.Interior.ColorIndex = 3
       End If
      End If
     
     If WGM <> "" Then
        If WGN = "" Then
            If WGN0 = "" Then
            NBERRORSW = NBERRORSW + 1
            c_wgn0.Interior.ColorIndex = 3
            End If
        End If
     End If
     
     
     
     
     
        
Next
        





MyClasseur.Sheets("Create").Cells(7, 2).Value = NBERRORS + NBERRORSW & " errors detected"
Fin:
Resume Next
End Sub

Sub AddSic(MySeet As Worksheet)
MySeet.Select
For iRow = 2 To MySeet.Range("A1").CurrentRegion.Rows.Count
'MySeet.Select

MySeet.Cells(iRow, 1).Select
    If (MySeet.Cells(iRow, 1) <> 0 And MySeet.Cells(iRow, 2) <> 0 And MySeet.Cells(iRow, 3) <> 0) Then
        jRow = iRow + 1
            Do While (MySeet.Cells(jRow, 5) <> 0)
                jRow = jRow + 1
            Loop
            If (MySeet.Cells(jRow, 5) = 0) Then
                MySeet.Range("A" & jRow, "G" & jRow).Interior.ColorIndex = 16
                
                If MySeet.Name = "SIC-TERM" Then
                MySeet.Cells(jRow + 1, 1).Select
                End If
                
            End If
    Else
    If (((MySeet.Cells(iRow, 1) = 0 And MySeet.Cells(iRow, 2) <> 0 And MySeet.Cells(iRow, 3) <> 0)) Or ((MySeet.Cells(iRow, 1) <> 0 And MySeet.Cells(iRow, 2) = 0 And MySeet.Cells(iRow, 3) <> 0)) Or ((MySeet.Cells(iRow, 1) <> 0 And MySeet.Cells(iRow, 2) <> 0 And MySeet.Cells(iRow, 3) = 0)) Or ((MySeet.Cells(iRow, 1) <> 0 And MySeet.Cells(iRow, 2) = 0 And MySeet.Cells(iRow, 3) = 0)) Or ((MySeet.Cells(iRow, 1) = 0 And MySeet.Cells(iRow, 2) <> 0 And MySeet.Cells(iRow, 3) = 0)) Or ((MySeet.Cells(iRow, 1) = 0 And MySeet.Cells(iRow, 2) = 0 And MySeet.Cells(iRow, 3) <> 0))) Then
               MsgBox "There is an error on this line: " & iRow & vbCrLf & "At least one fill is missing.", vbInformation
        End If
    End If
    DoEvents
Next
End Sub
Sub UpdateSIC(MySeet As Worksheet)

'MySeet.Select
Dim Formule As String

For iRow = 2 To MySeet.Range("A1").CurrentRegion.Rows.Count
MySeet.Select
MySeet.Cells(iRow, 1).Select
    If MySeet.Cells(iRow, 1) <> "" Then
        SicId = MySeet.Cells(iRow, 3)
    End If
    If MySeet.Cells(iRow, 5) <> "" Then
        MySeet.Cells(iRow, 6) = SicId & "." & MySeet.Cells(iRow, 5)
        Formule = "=IF(ISERROR(VLOOKUP(F" & iRow & ",Wire!G:G,1,0)),IF(ISERROR(VLOOKUP(F" & iRow & ",Wire!H:H,1,0)),IF(ISERROR(VLOOKUP(F" & iRow & ",WireGroup!L:L,1,0)),IF(ISERROR(VLOOKUP(F" & iRow & ",WireGroup!M:M,1,0)),""N"",""Y""),""Y""),""Y""),""Y"")"
        MySeet.Cells(iRow, 7).Formula = Formule
        numST = iRow
    End If
    If MySeet.Cells(iRow, 1) <> "" Then
        MySeet.Cells(iRow, 6) = SicId
        Formule = "=IF(ISERROR(VLOOKUP(F" & iRow & ",Wire!G:G,1,0)),IF(ISERROR(VLOOKUP(F" & iRow & ",Wire!H:H,1,0)),IF(ISERROR(VLOOKUP(F" & iRow & ",WireGroup!L:L,1,0)),IF(ISERROR(VLOOKUP(F" & iRow & ",WireGroup!M:M,1,0)),""N"",""Y""),""Y""),""Y""),""Y"")"
        MySeet.Cells(iRow, 7).Formula = Formule
        numST = iRow
    End If

Next

End Sub

Sub AddContacts(MySeet As Worksheet)
MySeet.Select
For iRow = 2 To MySeet.Range("A1").CurrentRegion.Rows.Count
MySeet.Cells(iRow, 1).Select
    If (MySeet.Cells(iRow, 1) <> 0 And MySeet.Cells(iRow, 2) <> 0 And MySeet.Cells(iRow, 3) <> 0) Then
        jRow = iRow + 1
            Do While (MySeet.Cells(jRow, 5) <> 0 And MySeet.Cells(jRow, 6) <> 0)
                jRow = jRow + 1
            Loop
            If (MySeet.Cells(jRow, 5) = 0 And MySeet.Cells(jRow, 6) = 0) Then
                MySeet.Range("A" & jRow, "I" & jRow).Interior.ColorIndex = 16
                
                If MySeet.Name = "SIC-CONT" Then
                MySeet.Cells(jRow + 1, 1).Select
                End If
                
            Else
                If (((MySeet.Cells(jRow, 5) = 0 And MySeet.Cells(jRow, 6) <> 0)) Or ((MySeet.Cells(jRow, 5) <> 0 And MySeet.Cells(jRow, 6) = 0))) Then
                    MsgBox "There is an error on line: " & jRow & vbCrLf & "At least one fill is missing.", vbInformation
            End If
            End If
    Else
    If (((MySeet.Cells(iRow, 1) = 0 And MySeet.Cells(iRow, 2) <> 0 And MySeet.Cells(iRow, 3) <> 0)) Or ((MySeet.Cells(iRow, 1) <> 0 And MySeet.Cells(iRow, 2) = 0 And MySeet.Cells(iRow, 3) <> 0)) Or ((MySeet.Cells(iRow, 1) <> 0 And MySeet.Cells(iRow, 2) <> 0 And MySeet.Cells(iRow, 3) = 0)) Or ((MySeet.Cells(iRow, 1) <> 0 And MySeet.Cells(iRow, 2) = 0 And MySeet.Cells(iRow, 3) = 0)) Or ((MySeet.Cells(iRow, 1) = 0 And MySeet.Cells(iRow, 2) <> 0 And MySeet.Cells(iRow, 3) = 0)) Or ((MySeet.Cells(iRow, 1) = 0 And MySeet.Cells(iRow, 2) = 0 And MySeet.Cells(iRow, 3) <> 0))) Then
        MsgBox "There is an error on this line: " & iRow & vbCrLf & "At least one fill is missing.", vbInformation
    End If
    End If
Next
End Sub
Sub UpdateContacts(MySeet As Worksheet)

MySeet.Select
Dim Formule As String

For iRow = 2 To MySeet.Range("A1").CurrentRegion.Rows.Count
MySeet.Cells(iRow, 1).Select
    If MySeet.Cells(iRow, 1) <> 0 Then
        SicId = MySeet.Cells(iRow, 3)
    End If
    If MySeet.Cells(iRow, 5) <> 0 Then
        MySeet.Cells(iRow, 8) = SicId & "." & MySeet.Cells(iRow, 5)
        Formule = "=IF(ISERROR(VLOOKUP(h" & iRow & ",Wire!G:G,1,0)),IF(ISERROR(VLOOKUP(h" & iRow & ",Wire!H:H,1,0)),IF(ISERROR(VLOOKUP(h" & iRow & ",WireGroup!L:L,1,0)),IF(ISERROR(VLOOKUP(h" & iRow & ",WireGroup!M:M,1,0)),""N"",""Y""),""Y""),""Y""),""Y"")"
        MySeet.Cells(iRow, 9).Formula = Formule
        MySeet.Cells(iRow, 7) = SicId & "." & MySeet.Cells(iRow, 5) & "." & MySeet.Cells(iRow, 6)
        numSC = iRow
        
    End If
    If MySeet.Cells(iRow, 1) <> 0 Then
        MySeet.Cells(iRow, 8) = SicId
        Formule = "=IF(ISERROR(VLOOKUP(h" & iRow & ",Wire!G:G,1,0)),IF(ISERROR(VLOOKUP(h" & iRow & ",Wire!H:H,1,0)),IF(ISERROR(VLOOKUP(h" & iRow & ",WireGroup!L:L,1,0)),IF(ISERROR(VLOOKUP(h" & iRow & ",WireGroup!M:M,1,0)),""N"",""Y""),""Y""),""Y""),""Y"")"
        MySeet.Cells(iRow, 9).Formula = Formule
        numSC = iRow
    End If

Next

End Sub

