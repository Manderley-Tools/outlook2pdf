' --------------------------------------------------
' Macro Outlook pour enregistrer un ou plusieurs éléments sélectionnés
' dans un répertoire Outlook au format pdf.
' Vous pouvez sélectionner autant d'e-mails que vous le souhaitez et
' chaque e-mail sera sauvegardé dans le répertoire local ou distant de
' votre choix.
'
' Nécessite :
' - Winword (référencé par late-bindings)
' - Excel
' - Microsoft Scripting Runtime
' - Acrobat
' - PDFToolsKit
'
' Macro VBA Inspirée de https://github.com/cavo789/vba_outlook_save_pdf
' --------------------------------------------------

Option Explicit
Private Const cFolder As String = """C:\"""
Private objWord As Object
Declare Function AtEndOfStream Lib "msvbvm60.dll" (ByVal hFile As Long) As Long

' --------------------------------------------------
' Demander à l'utilisateur le dossier dans lequel stocker les courriels
' --------------------------------------------------
Private Function AskForTargetFolder(Optional ByVal sTargetFolder As String) As String

    Dim dlgSaveAs As FileDialog

    sTargetFolder = Trim(sTargetFolder)

    ' Vérifier que sTargetFolder se termine bien par un slash
    If Not (Right(sTargetFolder, 1) = "\") Then
        sTargetFolder = sTargetFolder & "\"
    End If

    ' Si déjà initialisé, récupération de l'objet
    Set dlgSaveAs = objWord.FileDialog(msoFileDialogFolderPicker)

    With dlgSaveAs
        .Title = "Sélectionner le répertoire dans lequel vous souhaitez sauvegarder vos courriels"
        .AllowMultiSelect = False
        .InitialFileName = sTargetFolder
        .Show

        On Error Resume Next

        sTargetFolder = .SelectedItems(1)

        If Err.Number <> 0 Then
            sTargetFolder = cFolder
            Err.Clear
        End If

        On Error GoTo 0

    End With

    ' Vérifier que le chemin du répertoire de destinataion se termine bien par un slash
    If Not (Right(sTargetFolder, 1) = "\") Then
        sTargetFolder = sTargetFolder & "\"
    End If

    AskForTargetFolder = sTargetFolder

End Function

' --------------------------------------------------
' Demander un nom de fichier à l'utilisateur
' --------------------------------------------------
Private Function AskForFileName(ByVal sFileName As String) As String

    Dim dlgSaveAs As FileDialog
    Dim wResponse As VBA.VbMsgBoxResult
    Dim wPos As Integer

    Set dlgSaveAs = objWord.FileDialog(msoFileDialogSaveAs)

    ' Définir l'emplacement initial et le nom du fichier
    ' pour la boîte de dialogue "Enregistrer sous"
    dlgSaveAs.InitialFileName = sFileName

    ' Afficher la boîte de dialogue "Enregistrer sous"
    ' et enregistrer le message au format pdf
    If dlgSaveAs.Show = -1 Then

        sFileName = dlgSaveAs.SelectedItems(1)

        ' Vérifier que le pdf est sélectionné
        If Right(sFileName, 4) <> ".pdf" Then

            wResponse = MsgBox("Désolé, seul l'enregistrement au format PDF " & _
                "est supporté pour le moment." & vbNewLine & vbNewLine & _
                "Sauvegarder au format PDF ?", vbInformation + vbOKCancel)

            If wResponse = vbCancel Then
                sFileName = ""
            ElseIf wResponse = vbOK Then
                wPos = InStrRev(sFileName, ".")
                If wPos > 0 Then
                    sFileName = Left(sFileName, wPos - 1)
                End If
                sFileName = sFileName & ".pdf"
            End If

        End If
    End If

    ' Renvoie le nom du fichier
    AskForFileName = sFileName

End Function

' --------------------------------------------------
' Récupère le domaine d'une adresse e-mail sans son extension
' --------------------------------------------------
Private Function GetDomain(strMail As String) As String

  Dim intAtPos As Integer
  Dim intDotPos As Integer
  Dim intLength As Integer
  
  ' Déterminer la position du @
  intAtPos = InStr(strMail, "@") + 1
  
  ' Déterminer la position du .
  intDotPos = InStrRev(strMail, ".")
  
  ' Compter le nombre de caractère du domaine sans son extension
  intLength = intDotPos - intAtPos
    
  ' Renvoyer le domaine sans son extension
  If intLength > 0 Then
    GetDomain = Mid(strMail, intAtPos, intLength)
  Else
    GetDomain = "MANDERLEY"
  End If
End Function

' --------------------------------------------------
' Fonction pour extraire les mots de plus de 3 lettres de l'objet
' --------------------------------------------------
Private Function CleanSubject(strSubject As String) As String

  Dim i As Long
  Dim strWord As String

  CleanSubject = ""
  ' Diviser la chaîne de caractères en mots en utilisant un espace comme séparateur
   For i = LBound(Split(strSubject, " ")) To UBound(Split(strSubject, " "))
    strWord = Split(strSubject, " ")(i)
    If Len(strWord) > 3 Then
      CleanSubject = Replace(UCase(CleanSubject & strWord & " "), "_", " ")
    End If
  Next

  ' Suppression du dernier "_"
  If Len(CleanSubject) > 0 Then
    CleanSubject = Replace(UCase(Left(CleanSubject, Len(CleanSubject) - 1)), "_", " ")
  End If

End Function

Private Function GetNumberOfPages(pdfFile As String) As Long

    Dim commandLine As String
    Dim process As Object
    Dim pdfData As String
    Dim found As Boolean
    Dim startPos As Long
    Dim endPos As Long

    ' Initialiser les variables
    found = False
    startPos = 0
    endPos = 0

    ' Contruire la ligne de commande
    commandLine = "pdftk.exe """ & pdfFile & """ dump_data /VERYSILENT"

    ' Executer la ligne de commande et récupérer le résultat
    Set process = CreateObject("WScript.Shell").Exec(commandLine)
    Do While Not process.StdOut.AtEndOfStream
        pdfData = pdfData & process.StdOut.ReadLine
    Loop
    
    ' Déterminer la position de l'information recherchée
    startPos = InStr(pdfData, "NumberOfPages:") + 14
    endPos = InStr(startPos, pdfData, "PageMediaBegin")
    
    ' Récupérer le nombre ed pages du documents PDF
    GetNumberOfPages = CLng(Mid(pdfData, startPos, endPos - startPos))
   
End Function

'------------------------------------------------------------------------------------------------
' Transformer une plage en Tableau Structuré.
'------------------------------------------------------------------------------------------------
' TD : La plage concernée (ou la cellule haut gauche) du tableau de données.
' Nom : Le nom à donner au tableau ou vide pour prendre le nom attribué par EXCEL.
' Style : Le style ou * pour le style par défaut ou vide pour aucun style.
' AvecEntete : Indique si la première ligne contient des en-têtes :
'              xlYes : la plage contient des en-têtes;
'              xlNo : la plage ne contient pas d'en-tête et Excel les rajoute;
'              xlGuess : Excel détecte automatiquement si la plage contient ou non des en-têtes.
'------------------------------------------------------------------------------------------------
' Renvoie : La plage Range qui représente la plage du Tableau Structuré créé.
'------------------------------------------------------------------------------------------------
' Exemple :
' Dim TS As Range
' Set TS = ConvertirPlageEnTS(Range("K3"))
'------------------------------------------------------------------------------------------------
Private Function TS_ConvertirPlageEnTS(TD As Range, _
                                      Optional ByRef Nom As String = "", _
                                      Optional ByRef Style As String = "*", _
                                      Optional AvecEntete As XlYesNoGuess = xlYes) As Range

    'On Error GoTo Gest_Err
    Err.Clear
     
    ' Si le TD n'existe pas déjà alors le créer:
    If TD.ListObject Is Nothing Then
        
        ' Si TD ne représente qu'une seule cellule alors étend la plage:
        If TD.Count = 1 Then Set TD = TD.CurrentRegion
         
        ' Création du Tableau Structuré en attribuant le nom passé ou en prenant celui attribué par EXCEL:
        If Nom > "" Then
            TD.Parent.ListObjects.Add(xlSrcRange, TD, , AvecEntete).Name = Nom
        Else
            TD.Parent.ListObjects.Add xlSrcRange, TD, , AvecEntete
        End If
         
        ' Modifie le style s'il ne faut pas prendre celui par défaut, ou pas de style si vide:
        If Style <> "*" Then TD.Parent.ListObjects(TD.ListObject.Name).TableStyle = Style
    
    End If
      
    ' Renseigne le nom du Tableau Structuré et son style:
    Nom = TD.ListObject.DisplayName
    Style = TD.Parent.ListObjects(TD.ListObject.Name).TableStyle
    TD.ListObject.ShowTotals = True
    
    ' Renvoie la plage du Tableau Structuré:
    Set TS_ConvertirPlageEnTS = TD
     
    ' Fin du traitement:
' Gest_Err:
'    TS_Err_Number = Err.Number
'    TS_Err_Description = Err.Description
'    If Err.Number <> 0 Then
'        If TS_Méthode_Err = TS_Générer_Erreur Then Err.Raise TS_Err_Number, "TS_ConvertirPlageEnTS", TS_Err_Description
'        If TS_Méthode_Err = TS_MsgBox_Erreur Then MsgBox TS_Err_Number & " : " & TS_Err_Description, vbInformation, "TS_ConvertirPlageEnTS"
'    End If
'    Err.Clear

End Function
'------------------------------------------------------------------------------------------------


' --------------------------------------------------
' Fait le taf, traite tous les mails sélectionnés en
' les exportant au format PDF.
' Si l'utilisateur demande à ce que les messages soient
' supprimés une fois exportés, alors ils le sont.
' --------------------------------------------------
Sub SaveAsPDF()

    Const wdExportFormatPDF = 17
    Const wdExportOptimizeForPrint = 0
    Const wdExportAllDocument = 0
    Const wdExportDocumentContent = 0
    Const wdExportCreateNoBookmarks = 0
    
    Dim stRep As String
    Dim WshShell As Object
    Set WshShell = CreateObject("WScript.Shell")
    stRep = WshShell.SpecialFolders("MyDocuments")
    Set WshShell = Nothing
    
    Dim oSelection As Outlook.Selection
    Dim oMail As Outlook.MailItem
    Dim objAtt As Outlook.Attachment
    Dim objFSO As FileSystemObject
    Dim objPdf As Acrobat.AcroPDDoc

    ' Utiliser des liaisons tardives
    Dim objDoc As Object
    Dim oRegEx As Object
    
    ' Tableau de bord
    Dim objExcel As New Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objData As New Dictionary

    Dim dlgSaveAs As FileDialog
    Dim objFDFS As FileDialogFilters
    Dim fdf As FileDialogFilter
    Dim i As Integer, y As Integer, x As Integer, N As Integer, P As Integer, wSelectedeMails As Integer
    Dim sFileName As String, nb As String
    Dim sTempFolder As String, sTempFileName As String, strFilePath As String
    Dim sTargetFolder As String, strCurrentFile As String
    Dim sExt As String
    
    Dim strSender As String
    Dim strReceiver As String
    
    Dim bContinue As Boolean
    Dim bAskForFileName As Boolean
    Dim bRemoveMailAfterExport As Boolean
    
    Dim Fichier As String
    Dim Description As String
    Dim ID As Integer
    Dim oDate As String
    Dim sType As String
    Dim Emetteur As String
    Dim Destinataire As String
    Dim Objet As String
    Dim Reference As String
    Dim Montant As Currency
    Dim Synthese As String
    Dim Observation As String
    Dim Questions As String
    Dim Demandes As String
    Dim nbPages As Integer
    
    ' Obtenir toute la sélection
    Set oSelection = Application.ActiveExplorer.Selection

    ' Obtenir le nombre de courriels sélectionnés
    wSelectedeMails = oSelection.Count

    ' Vérifier qu'au moins un élément est sélectionné
    If wSelectedeMails < 1 Then
        Call MsgBox("Veuillez sélectionner au moins un courriel", vbExclamation, "Enregistrer en PDF")
        Exit Sub
    End If

    ' --------------------------------------------------
    '
    bContinue = MsgBox("Vous êtes sur le point de convertir " & wSelectedeMails & " " & _
        "courriels en autant de fichiers PDF qui pourront être seront automatiquement renommés. " & vbCrLf & vbCrLf & _
        "Voulez-vous continuer ? " & vbCrLf & vbCrLf & _
        "Si oui, vous devrez sélectionner le répertoire dans lequel vous souhaitez les enregistrer.", _
        vbQuestion + vbYesNo + vbDefaultButton1) = vbYes

    If Not bContinue Then
        Exit Sub
    End If

    ' --------------------------------------------------
    ' Démarrer Word et initialiser le fichier
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = False

    ' --------------------------------------------------
    ' Définir le répertoire cible, où enregistrer les courriels
    'sTargetFolder = AskForTargetFolder(cFolder)
    sTargetFolder = AskForTargetFolder(Environ("USERPROFILE"))
    
    If sTargetFolder = "" Then
        objWord.Quit
        Set objWord = Nothing
        Exit Sub
    End If

    ' --------------------------------------------------
    ' Une fois que le courrier a été enregistré au format PDF,
    ' faut-il le supprimer ?
    bRemoveMailAfterExport = MsgBox("Une fois que le courrier électronique aura été enregistré " & _
    "sur votre disque au format PDF, souhaitez-vous le conserver ou le déplacer dans les éléments supprimés ? " & vbCrLf & vbCrLf & _
    "Cliquez sur Oui pour le conserver " & vbCrLf & _
    "Cliquez sur Non pour le supprimer après l'exportation", _
    vbQuestion + vbYesNo + vbDefaultButton1) = vbNo

    ' --------------------------------------------------
    ' Lorsque plusieurs courriels ont été sélectionnés,
    ' demander à l'utilisateur s'il souhaite modifier les noms
    ' de fichiers à chaque fois (ce qui peut être fastidieux).
    bAskForFileName = True

    If (wSelectedeMails > 1) Then
        bAskForFileName = MsgBox("Vous êtes sur le point de convertir " & wSelectedeMails & " " & _
            "courriels en autant de fichiers PDF. " & vbCrLf & vbCrLf & _
            "Voulez-vous voir chacun des " & wSelectedeMails & " " & _
            "messages pour renommer le fichier ou utiliser le mode automatisé ? " & vbCrLf & vbCrLf & _
            "Cliquez sur Oui pour voir les invites d'enregistrement, ou sur non pour laisser faire le mode automatique.", _
            vbQuestion + vbYesNo + vbDefaultButton2) = vbYes

        MsgBox "ATTENTION : Vous ne verrez pas de progression à l'écran (malheureusement, " & _
            "Outlook ne le permet pas)." & vbCrLf & vbCrLf & _
            "Si vous exportez beaucoup d'e-mails, le processus peut prendre un certain temps. " & vbCrLf & vbCrLf & _
            "Pour vérifier que les choses fonctionnent ouvrez une fenêtre de l'explorateur pour " & _
            "voir les fichiers s'ajouter au dossier. " & vbCrLf & vbCrLf & _
            "Une fois l'opération terminée, un message de confirmation s'affiche.", _
            vbInformation + vbOKOnly
    End If

    ' --------------------------------------------------
    ' Définir la boîte de dialogue "Enregistrer sous"
    If bAskForFileName Then

        Set dlgSaveAs = objWord.FileDialog(msoFileDialogSaveAs)

        ' --------------------------------------------------
        ' Déterminer l'indice de filtre pour l'enregistrement d'un fichier pdf
        ' Obtenir tous les filtres et s'assurer que nous avons "pdf".
        Set objFDFS = dlgSaveAs.Filters

        i = 0
        For Each fdf In objFDFS
            i = i + 1
            If InStr(1, fdf.Extensions, "pdf", vbTextCompare) > 0 Then
                Exit For
            End If
        Next fdf

        Set objFDFS = Nothing

        ' Définir le FilterIndex à pdf-files
        dlgSaveAs.FilterIndex = i

    End If

    ' ----------------------------------------------------
    ' Obtenir le dossier temporaire de l'utilisateur dans lequel
    ' l'élément doit être stocké
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    sTempFolder = objFSO.GetSpecialFolder(2)
    

    ' ----------------------------------------------------
    ' Démarrage du traitement des courriels
    ' Traiter chaque courriel unitairement
    On Error Resume Next
    
    strFilePath = sTargetFolder & "ANALYSE.xlsx"
    Debug.Print objFSO.FileExists(strFilePath)
    
    ' Vérifier si le fichier existe déjà
    If Dir(strFilePath) <> "" Then
    Debug.Print Dir(strFilePath)
        ' Ouverture du classeur existant
        Set objWorkbook = objExcel.Workbooks.Open(strFilePath)
        
        ' Vérification si la feuille de calcul existe
        If Not objWorkbook.Sheets("ANALYSE") <> "" Then
            ' Créer une nouvelle feuille de calcul
            Set objWorksheet = objWorkbook.Sheets.Add(After:=objWorkbook.Sheets(objWorkbook.Sheets.Count))
            ' Définir le nom de la feuille de calcul
            objWorksheet.Name = "ANALYSE"
        End If
    
        ' Activation de l'onglet ANALYSE
        Set objWorksheet = objWorkbook.Sheets("ANALYSE")
        
        ' Incrémentation du nombre de lignes
        y = objWorksheet.Cells.Find(What:="*").Row + 1
       
    Else
    
        ' Création du classeur Excel
        Set objWorkbook = objExcel.Workbooks.Add
        objWorkbook.SaveAs FileName:=strFilePath, FileFormat:=xlOpenXMLWorkbook
        
        ' Créer une nouvelle feuille de calcul
        Set objWorksheet = objWorkbook.Sheets.Add(After:=objWorkbook.Sheets(objWorkbook.Sheets.Count))
        
        ' Définir le nom de la feuille de calcul
        objWorksheet.Name = "ANALYSE"
        
        ' Activation de l'onglet ANALYSE
        Set objWorksheet = objWorkbook.Sheets("ANALYSE")
        objWorkbook.Sheets("Feuil1").Delete
    
        ' En-tête de la feuille
        objWorksheet.Cells(1, 1).Value = "Fichier"
        objWorksheet.Cells(1, 2).Value = "Description"
        objWorksheet.Cells(1, 3).Value = "#"
        objWorksheet.Cells(1, 4).Value = "Date"
        objWorksheet.Cells(1, 5).Value = "Heure"
        objWorksheet.Cells(1, 6).Value = "Type"
        objWorksheet.Cells(1, 7).Value = "Emetteur"
        objWorksheet.Cells(1, 8).Value = "Destinataire"
        objWorksheet.Cells(1, 9).Value = "Objet"
        objWorksheet.Cells(1, 10).Value = "Référence"
        objWorksheet.Cells(1, 11).Value = "Montant"
        objWorksheet.Cells(1, 12).Value = "Synthèse"
        objWorksheet.Cells(1, 13).Value = "Observations"
        objWorksheet.Cells(1, 14).Value = "Questions"
        objWorksheet.Cells(1, 15).Value = "Demandes"
        objWorksheet.Cells(1, 16).Value = "Pages"
        
        y = 2
        
    End If

    For i = wSelectedeMails To 1 Step -1
    
        ' Récupérer le courriel sélectionné
        Set oMail = oSelection.Item(i)
        
        ' Récupérer le domaine de l'expéditeur ou son nom
        strSender = UCase(GetDomain(oMail.Sender.Address))
        
        If strSender = "" Then
            strSender = UCase(oMail.Sender.Name)
        End If
        
        ' A revoir
        If InStr(strSender, "TALARICO") _
        + InStr(strSender, "KALFAT") _
        + InStr(strSender, "BUTTI") _
        + InStr(strSender, "GELLER") _
        + InStr(strSender, "DEMAIMAY") _
        + InStr(strSender, "BOURRIER") _
        > 0 Then
            strSender = "MANDERLEY"
        End If
        
        ' Récupérer le domaine du premier destinataire ou son nom
        strReceiver = UCase(GetDomain(oMail.Recipients(1).Name))
        
        If strReceiver = "" Then
            strReceiver = UCase(oMail.Recipients(1).Name)
        End If
      
        ' A revoir
        If InStr(strReceiver, "TALARICO") _
        + InStr(strReceiver, "KALFAT") _
        + InStr(strReceiver, "BUTTI") _
        + InStr(strReceiver, "GELLER") _
        + InStr(strReceiver, "DEMAIMAY") _
        + InStr(strReceiver, "BOURRIER") _
        > 0 Then
            strReceiver = "MANDERLEY"
        End If
        
        ' Construire le nom de fichier pour le fichier mht temporaire
        sTempFileName = sTempFolder & "\outlook.mht"

        Debug.Print sTempFileName
        ' Ecraser le fichier précédent s'il est déjà présent
        If Dir(sTempFileName) Then Kill (sTempFileName)
        
        ' Enregistrer le fichier mht
        oMail.SaveAs sTempFileName, olMHTML

        ' Ouvrir le fichier mht dans Word sans que Word soit visible
        Set objDoc = objWord.Documents.Open(FileName:=sTempFileName, Visible:=False, ReadOnly:=True)
        
        ' Assainissement du nom de fichier, suppression des caractères indésirables
        Set oRegEx = CreateObject("vbscript.regexp")
        oRegEx.Global = True
        oRegEx.Pattern = "[\\/:*?""<>|]"
        
        ' Affecter les variables
        Objet = UCase(Trim(oRegEx.Replace(CleanSubject(oMail.Subject), "")))
        Fichier = Format(oMail.ReceivedTime, "yyyymmdd_HhNnSs") & "_MEL_" & strSender & "_" & strReceiver & "_" & Objet & ".pdf"

        ' Construire un nom de fichier sûr à partir de l'objet du message
        sFileName = sTargetFolder & Fichier
            
        If bAskForFileName Then
            sFileName = AskForFileName(sFileName)
        End If

        If Not (Trim(sFileName) = "") Then

            ' S'il existe déjà, supprimer d'abord le fichier
            If Dir(sFileName) <> "" Then
                Kill (sFileName)
            End If

            ' Enregistrer en PDF
            objDoc.ExportAsFixedFormat OutputFileName:=sFileName, _
                ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
                wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=0, To:=0, _
                Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
                CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
                BitmapMissingFonts:=True, UseISO19005_1:=False

            ' Fermer une fois sauvegardé sur le disque
            objDoc.Close
            
            ' Affecter les variables
            nbPages = GetNumberOfPages(sFileName)
            
            ' Renseigner le tableau de bord Excel avec le courriel
            objWorksheet.Cells(y, 1).Value = Fichier
            objWorksheet.Cells(y, 2).Value = Objet
            objWorksheet.Cells(y, 3).Value = y - 1
            objWorksheet.Cells(y, 4).Value = Format(oMail.ReceivedTime, "dd/mm/yyyy")
            objWorksheet.Cells(y, 5).Value = Format(oMail.ReceivedTime, "Hh:Nn:Ss")
            objWorksheet.Cells(y, 6).Value = "Courriel"
            objWorksheet.Cells(y, 7).Value = strSender
            objWorksheet.Cells(y, 8).Value = strReceiver
            objWorksheet.Hyperlinks.Add _
                Anchor:=objWorksheet.Cells(y, 9), _
                Address:="" & Fichier & "", _
                TextToDisplay:="" & Objet & ""
            'objWorksheet.Cells(y, 9).Value = "=LIEN_HYPERTEXTE(A" & y & ", B" & y & ")"
            objWorksheet.Cells(y, 10).Value = ""
            objWorksheet.Cells(y, 11).Value = 0
            objWorksheet.Cells(y, 12).Value = ""
            objWorksheet.Cells(y, 13).Value = ""
            objWorksheet.Cells(y, 14).Value = ""
            objWorksheet.Cells(y, 15).Value = ""
            objWorksheet.Cells(y, 16).Value = nbPages
            
            ' Enregistrement du classeur Excel
            objWorkbook.Save
            
            ' Incrémentation de la position de la ligne dans le tableau Excel
            y = y + 1
            
            ' Réinitialisation du nombre de page
            nbPages = 0
                        
            ' Supprimer le courriel ?
            If bRemoveMailAfterExport Then
                ' Ok, seulement si le courrier a bien été exporté.
                If Dir(sFileName) <> "" Then
                    oMail.Delete
                End If
            End If

        End If
        
        ' Ajout des pièces jointes
        P = 0
        For N = 1 To oMail.Attachments.Count
          Set objAtt = oMail.Attachments.Item(N)
          
          If Not Left(objAtt.FileName, 6) = "image0" Then
            
            'Encodage du n° de la pièce jointe
            P = P + 1
            
            If P < 10 Then
                nb = "0" & P
            Else
                nb = P
            End If
            
            ' Récupération de l'extension de la pièce jointe
            Set objFSO = CreateObject("Scripting.FileSystemObject")
            sExt = objFSO.GetExtensionName(oMail.Attachments.Item(N).FileName)
            
            Debug.Print "Date de dernière modification de la pièce jointe : " & objFSO.GetFile(oMail.Attachments.Item(N).FileName).DateLastModified
            
            Set objFSO = Nothing
            
            ' Traitement conditionnel en fonction du type de pièce jointe
            Select Case sExt
                
                ' Pour les pièces jointes directement convertibles en PDF
                Case "doc", "docx", "xls", "xlsx", "ppt", "pptx"
                    'Debug.Print "L'extension de la pièce jointe est : " & UCase(sExt) & ".", vbInformation
                    ' TODO : Coder la conversion au format PDF
                
                ' Pour les pièces jointes susceptibles de comporter des pièces jointes
                Case "msg", "eml"
                    'Debug.Print "ATTENTION => La pièce jointe est un mail au format : " & UCase(sExt) & ".", vbCritical
                    ' TODO : Coder la conversion au format PDF
                    ' TODO : Vérifier s'ils comportent des pièces jointes ou non et les traiter de la même manière le cas échéant
                    ' TODO : Factoriser le tout pour alléger le code
                
                ' Pour les pièces jointes déjà au format PDF
                Case "pdf"
                    'Debug.Print "La pièce jointe est déjà au format " & UCase(sExt) & ".", vbInformation
                    nbPages = GetNumberOfPages(oMail.Attachments.Item(N).FileName)
                
                ' Pour tous les autres types de pièces jointes
                Case Else
                    'Debug.Print "ATTENTION => L'extension de la pièce jointe est : " & UCase(sExt) & ".", vbCritical
                    ' TODO : Cf. Traitement ci-dessous
            End Select
            
            Objet = UCase(CleanSubject(objAtt.FileName))
            Fichier = Format(oMail.ReceivedTime, "yyyymmdd_HhNnSs") & "_P" & nb & "_" & strSender & "_" & strReceiver & "_" & Objet
            Objet = UCase(Replace(objAtt.FileName, "." & sExt, ""))
            
            objAtt.SaveAsFile (sTargetFolder & Fichier)
            oDate = FileDateTime(sTargetFolder & Fichier)
            
            'Debug.Print "Date de dernière modification de la pièce jointe : " & oDate
                          
            ' Renseigner le tableau de bord Excel avec la pièce jointe
            objWorksheet.Cells(y, 1).Value = Fichier
            objWorksheet.Cells(y, 2).Value = Objet
            objWorksheet.Cells(y, 3).Value = y - 1
            objWorksheet.Cells(y, 4).Value = Format(oDate, "dd/mm/yyyy")
            objWorksheet.Cells(y, 5).Value = Format(oDate, "Hh:Nn:Ss")
            objWorksheet.Cells(y, 6).Value = "Pièce jointe"
            objWorksheet.Cells(y, 7).Value = strSender
            objWorksheet.Cells(y, 8).Value = strReceiver
            objWorksheet.Hyperlinks.Add _
                Anchor:=objWorksheet.Cells(y, 9), _
                Address:="" & Fichier & "", _
                TextToDisplay:="" & Objet & ""
            objWorksheet.Cells(y, 10).Value = UCase(sExt)
            objWorksheet.Cells(y, 11).Value = 0
            objWorksheet.Cells(y, 12).Value = "Dernière modification le " & Format(oDate, "dd/mm/yyyy à Hh:Nn:Ss")
            objWorksheet.Cells(y, 13).Value = "Pièce jointe n°" & nb
            objWorksheet.Cells(y, 14).Value = ""
            objWorksheet.Cells(y, 15).Value = ""
            objWorksheet.Cells(y, 16).Value = nbPages
            
            ' Enregistrement du classeur Excel
            objWorkbook.Save
            
            'Incrémentation de la position de la ligne dans le fichier Excel
            y = y + 1
            
            ' Réinitialisation du nombre de page
            nbPages = 0
                    
          End If
        
        Next N
        
        strSender = ""
        strReceiver = ""
                
        'Debug.Print _
        '"Date : " & Format(oMail.ReceivedTime, "yyyymmdd_HhNn") & vbCrLf & _
        '"Type : MEL" & vbCrLf & _
        '"Expéditeur : " & UCase(oMail.Sender.Name) & " (" & oMail.Sender.Address & ")" & vbCrLf & _
        '"Destinataire : " & UCase(oMail.Recipients(1).Name) & " (" & oMail.Recipients(1).Address & ")" & vbCrLf & _
        '"Objet : " & UCase(Trim(oRegEx.Replace(CleanSubject(oMail.Subject), "")))
        
    Next i
       
    ' Formatter le tableau Excel
    TS_ConvertirPlageEnTS objWorksheet.Range("A1"), "PIECES", "*", xlYes
    objWorksheet.Range("A1:O" & y & "").NumberFormat = "jj/mm/aaaa"
    objWorksheet.Range("A1:O" & y & "").NumberFormat = "Comptabilité"
    objWorksheet.Range("A1:O" & y & "").NumberFormat = "Nombre"
    objWorksheet.Cells(y, 3).Formula2 = "=SOUS.TOTAL(104,['#])"
    objWorksheet.Cells(y, 11).Formula2 = "=SOUS.TOTAL(109,[Montant])"
    objWorksheet.Cells(y, 16).Formula2 = "=SOUS.TOTAL(109,[Pages])"
       
    ' Création et fermeture du groupe de colonnes
    ' A supprimer plus tard
    objWorksheet.Columns("A:C").Group
    objWorksheet.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
    
    ' Enregistrement du classeur Excel
    objWorkbook.Save
    
    ' Enregistrement du classeur dans le dossier spécifié
    objWorkbook.Close savechanges:=True
   
    Set dlgSaveAs = Nothing

    On Error GoTo 0

    ' Fermez le document et Word
    On Error Resume Next
    objWord.Quit
    objWorkbook.Quit
    objExcel.Quit

    On Error GoTo 0

    ' Nettoyage
    Set oSelection = Nothing
    Set oMail = Nothing
    Set objDoc = Nothing
    Set objExcel = Nothing
    Set objWorkbook = Nothing
    Set objWorksheet = Nothing
    Set objWord = Nothing
    Set oRegEx = Nothing

    MsgBox "Vos fichiers sont prêts ! " & vbCrLf & vbCrLf & _
    "Les e-mails sélectionnés ont été exportés vers " & sTargetFolder & vbCrLf & vbCrLf & _
    "Manderley-AI espère vous avoir pu vous être utile !" & vbCrLf & vbCrLf & _
    "Si oui, n'hésitez pas à adresser vos dons à :" & vbCrLf & _
    "l.talarico@ciblexperts.com ;-)", vbSystemModal, "Manderley-AI vous remercie !"
    
    ' TODO : Optimiser le code
    ' TODO : Ajouter un peu d'intelligence à ce petit automate ;-)

End Sub


