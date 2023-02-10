VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOnglet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Onglet As String
Public Liason As New Collection
Public epissure As Boolean
Public Sub RetourneTableau(TXT As String)
Dim I As Long
Dim Tableau As String
Dim Tableau2 As String
Tableau = ""
    For I = 1 To Liason.Count
    
        
        
         
                Tableau = Trim("" & Liason(I).LIAI) & ";"
                Tableau = Tableau & Trim("" & Liason(I).SECT) & ";"
                Tableau = Tableau & Trim("" & Liason(I).TEINT) & ";"
                Tableau = Tableau & Trim("" & Liason(I).POS_OUT) & ";"
                Tableau = Tableau & Trim("" & Liason(I).APP) & ";"
                Tableau = Tableau & Trim("" & Liason(I).VOI) & ";"
                Tableau = Tableau & Trim("" & Liason(I).REF_CONNECTEUR) & ";"
                Tableau = Tableau & Trim("" & Liason(I).POS_OUT2) & ";"
                Tableau = Tableau & Trim("" & Liason(I).App2) & ";"
                Tableau = Tableau & Trim("" & Liason(I).VOI2) & ";"
                Tableau = Tableau & Trim("" & Liason(I).REF_CONNECTEUR2) & ";"
                Tableau = Tableau & Trim("" & Liason(I).Fil)
                Tableau = Tableau & vbCrLf
                
               
              

              
            TXT = TXT & Tableau
                       
               
        
    Next


'Debug.Print RetourneTableau
End Sub
Public Sub AjouterClass(RS As Recordset)
Dim T_L As Liaison

 Set T_L = New Liaison


T_L.Fil = RS("FIL")

T_L.LIAI = RS("LIAI")
T_L.SECT = RS("SECT")
T_L.TEINT = RS("TEINT")
T_L.Fil = RS("FIL")


If UCase("" & RS("Onglet")) = UCase("" & RS("app")) Then
        T_L.POS_OUT = RS("POS-OUT")
        T_L.POS_OUT2 = RS("POS-OUT2")
        T_L.REF_CONNECTEUR = RS("REF CONNECTEUR_FOUR")
        T_L.REF_CONNECTEUR2 = RS("REF CONNECTEUR_FOUR2")
        T_L.APP = RS("APP")
        T_L.App2 = RS("APP2")
        T_L.VOI = RS("VOI")
        T_L.VOI2 = RS("VOI2")

Else
    T_L.POS_OUT = RS("POS-OUT2")
    T_L.POS_OUT2 = RS("POS-OUT")
    T_L.REF_CONNECTEUR = RS("REF CONNECTEUR_FOUR2")
    T_L.REF_CONNECTEUR2 = RS("REF CONNECTEUR_FOUR")
    
    T_L.APP = RS("APP2")
    T_L.App2 = RS("APP")
    T_L.VOI = RS("VOI2")
    T_L.VOI2 = RS("VOI")
End If
    Liason.Add T_L, "F" & RS("FIL")
     Set T_L = Nothing
    
End Sub
Public Sub ReplaceEpissure(TableauOnglet)
On Error Resume Next
Dim I As Long
Dim I2 As Long
Dim Trouve As Boolean

For I = 1 To Liason.Count
    If TableauOnglet(Liason(I).App2).epissure = False Then
        For I2 = 1 To Liason.Count
            Liason(I2).APP = Liason(I).App2
            Liason(I2).VOI = Liason(I).VOI2
            Liason(I2).POS_OUT = Liason(I).POS_OUT2
            Liason(I2).REF_CONNECTEUR = Liason(I).REF_CONNECTEUR2
            
        Next
       Liason.Remove I
       Exit For
    End If
Next
Err.Clear
For I = 1 To Liason.Count
     If TableauOnglet(Liason(I).App2).epissure = True Then
         For I2 = 1 To TableauOnglet(Liason(I).App2).Liason.Count
            If TableauOnglet(TableauOnglet(Liason(I).App2).Liason(I2).App2).epissure = False Then
                Liason(I).App2 = TableauOnglet(Liason(I).App2).Liason(I2).App2
                Liason(I).VOI2 = TableauOnglet(Liason(I).App2).Liason(I2).VOI2
                Liason(I).POS2_OUT = TableauOnglet(Liason(I).App2).Liason(I2).POS_OUT2
                Liason(I).REF_CONNECTEUR2 = TableauOnglet(Liason(I).App2).Liason(I2).REF_CONNECTEUR2
                Exit For
            End If
           
         Next
        For I2 = 1 To TableauOnglet(Liason(I).App2).Liason.Count
            If TableauOnglet(TableauOnglet(Liason(I).App2).Liason(I2).App2).epissure = True Then
               If TableauOnglet(TableauOnglet(Liason(I).App2).Liason(I2).App2).App2 = Liason(I).App2 Then
                   TableauOnglet(TableauOnglet(Liason(I).App2).Liason(I2).App2).Remove I2
                   Exit For
                End If
            End If
            
        Next
       Exit For
     End If
Next
End Sub
