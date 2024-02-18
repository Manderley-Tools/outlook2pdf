' --------------------------------------------------
' Macro Outlook pour enregistrer un ou plusieurs
' éléments sélectionnés en tant que fichiers pdf
' sur votre disque dur. Vous pouvez sélectionner 
' autant de mails ' que vous voulez et chaque mail 
' sera sauvegardé sur votre disque.
' Nécessite : 
' - Winword (référencé par late-bindings)
' @voir https://github.com/Manderley-tools/outlook2pdf
' --------------------------------------------------
Option Explicit
Private Const cFolder As String = "C:\"
Private objWord As Object
' --------------------------------------------------
' Demander à l'utilisateur le dossier dans lequel stocker les courriels
' --------------------------------------------------
Private Function AskForTargetFolder(ByVal sTargetFolder As String) As String
    Dim dlgSaveAs As FileDialog
    sTargetFolder = Trim(sTargetFolder)
    If Not (Right(sTargetFolder, 1) = "\") Then                         ' Vérifier que sTargetFolder se termine par une barre oblique.
        sTargetFolder = sTargetFolder & "\"
    End If
    Set dlgSaveAs = objWord.FileDialog(msoFileDialogFolderPicker)       ' Récupérer l'objet
    With dlgSaveAs
        .Title = "Sélectionnez le dossier dans lequel enregistrer ces courriels"
        .AllowMultiSelect = False
        .InitialFileName = sTargetFolder
        .Show
        On Error Resume Next
        sTargetFolder = .SelectedItems(1)
        If Err.Number <> 0 Then
            sTargetFolder = ""
            Err.Clear
        End If
        On Error GoTo 0
    End With
    If Not (Right(sTargetFolder, 1) = "\") Then                         ' Vérifier que sTargetFolder se termine par une barre oblique.
        sTargetFolder = sTargetFolder & "\"
    End If
    AskForTargetFolder = sTargetFolder
End Function
' --------------------------------------------------
' Demander à l'utilisateur un nom de fichier
' --------------------------------------------------
Private Function AskForFileName(ByVal sFileName As String) As String
    Dim dlgSaveAs As FileDialog
    Dim wResponse As VBA.VbMsgBoxResult
    Dim wPos As Integer
    Set dlgSaveAs = objWord.FileDialog(msoFileDialogSaveAs)
    dlgSaveAs.InitialFileName = sFileName                               ' Définir l'emplacement initial pour la boîte de dialogue Enregistrer sous
    If dlgSaveAs.Show = -1 Then                                         ' Afficher la boîte de dialogue Enregistrer sous et enregistrer le message au format PDF
        sFileName = dlgSaveAs.SelectedItems(1)                          ' Définir le nom du fichier pour la boîte de dialogue Enregistrer sous
        If Right(sFileName, 4) <> ".pdf" Then                           ' Vérifier si le format PDF est bien sélectionné
            wResponse = MsgBox("Désolé, seul l'enregistrement " & _
                "au format pdf est pris en charge." & _
                vbNewLine & vbNewLine & _
                "Enregistrer au format pdf à la place ?", _
                vbInformation + vbOKCancel)
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
    AskForFileName = sFileName                                          ' Renvoyer le nom du fichier
End Function
' --------------------------------------------------
' Récupère le domaine d'une adresse e-mail sans son extension
' --------------------------------------------------
Private Function GetDomain(strMail As String) As String
    Dim intAtPos As Integer
    Dim intDotPos As Integer
    Dim intLength As Integer
    intAtPos = InStr(strMail, "@") + 1                                    ' Déterminer la position du @
    intDotPos = InStrRev(strMail, ".")                                    ' Déterminer la position du .
    intLength = intDotPos - intAtPos                                      ' Compter le nombre de caractère du domaine sans son extension
    If intLength > 0 Then                                                 ' Renvoyer le domaine sans son extension
        GetDomain = Mid(strMail, intAtPos, intLength)
    Else
        GetDomain = "MANDERLEY"
    End If
End Function
' --------------------------------------------------
' Extraire les mots de plus de 3 lettres de l'objet 
' --------------------------------------------------
Private Function CleanSubject(strSubject As String) As String
  Dim i As Long
  Dim strWord As String
  CleanSubject = ""
   For i = LBound(Split(strSubject, " ")) To UBound(Split(strSubject, " "))' Diviser la chaîne de caractères en mots en utilisant un espace comme séparateur
    strWord = Split(strSubject, " ")(i)
    If Len(strWord) > 3 Then
      CleanSubject = Replace(UCase(CleanSubject & strWord & " "), "_", " ")
    End If
  Next
  If Len(CleanSubject) > 0 Then                                             ' Suppression du dernier "_"
    CleanSubject = Replace(UCase(Left(CleanSubject, Len(CleanSubject) - 1)), "_", " ")
  End If
End Function
' --------------------------------------------------
' Extraire le nombre de pages d'un fichier PDF
' --------------------------------------------------
Private Function GetNumberOfPages(pdfFile As String) As Long
    Dim commandLine As String                                               ' Initialiser les variables
    Dim process As Object
    Dim pdfData As String
    Dim found As Boolean
    Dim startPos As Long
    Dim endPos As Long
    found = False
    startPos = 0
    endPos = 0
    commandLine = "pdftk.exe """ & pdfFile & """ dump_data /VERYSILENT"     ' Contruire la ligne de commande
    Set process = CreateObject("WScript.Shell").Exec(commandLine)           ' Executer la ligne de commande et récupérer le résultat
    Do While Not process.StdOut.AtEndOfStream
        pdfData = pdfData & process.StdOut.ReadLine
    Loop
    startPos = InStr(pdfData, "NumberOfPages:") + 14                        ' Déterminer la position de l'information recherchée
    endPos = InStr(startPos, pdfData, "PageMediaBegin")
    GetNumberOfPages = CLng(Mid(pdfData, startPos, endPos - startPos))      ' Récupérer le nombre ed pages du documents PDF
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
    On Error GoTo Gest_Err
    Err.Clear
    If TD.ListObject Is Nothing Then                                        ' Si le TD n'existe pas déjà alors le créer:
        If TD.Count = 1 Then Set TD = TD.CurrentRegion                      ' Si TD ne représente qu'une seule cellule alors étend la plage:
        If Nom > "" Then                                                    ' Création du Tableau Structuré en attribuant le nom passé ou en prenant celui attribué par EXCEL:
            TD.Parent.ListObjects.Add(xlSrcRange, TD, , AvecEntete).Name = Nom
        Else
            TD.Parent.ListObjects.Add xlSrcRange, TD, , AvecEntete
        End If
        If Style <> "*" Then                                                ' Modifie le style s'il ne faut pas prendre celui par défaut, ou pas de style si vide:
            TD.Parent.ListObjects(TD.ListObject.Name).TableStyle = Style
        End If                                                          
    End If
    Nom = TD.ListObject.DisplayName                                         ' Renseigne le nom du Tableau Structuré et son style:
    Style = TD.Parent.ListObjects(TD.ListObject.Name).TableStyle
    TD.ListObject.ShowTotals = True
    Set TS_ConvertirPlageEnTS = TD                                          ' Renvoie la plage du Tableau Structuré:
    
    Gest_Err:                                                               ' Fin du traitement:
    TS_Err_Number = Err.Number
    TS_Err_Description = Err.Description
    If Err.Number <> 0 Then
        If TS_Méthode_Err = TS_Générer_Erreur Then Err.Raise TS_Err_Number, "TS_ConvertirPlageEnTS", TS_Err_Description
        If TS_Méthode_Err = TS_MsgBox_Erreur Then MsgBox TS_Err_Number & " : " & TS_Err_Description, vbInformation, "TS_ConvertirPlageEnTS"
    End If
    Err.Clear
End Function
' --------------------------------------------------
' Faire le travail, traiter les courriels sélectionnés et les exporter au format PDF
' Déplace les courriels dans éléments supprimés si deamndé.
' --------------------------------------------------
Sub SaveAsPDFfile()
    Const wdExportFormatPDF = 17                                        ' Initialisation des constantes
    Const wdExportOptimizeForPrint = 0
    Const wdExportAllDocument = 0
    Const wdExportDocumentContent = 0
    Const wdExportCreateNoBookmarks = 0
    Dim oSelection As Outlook.Selection                                 ' Initialisation des variables
    Dim oMail As Outlook.MailItem                                       
    Dim objFSO As FileSystemObject                      
    Dim objDoc As Object                                                ' Utilise late-bindings
    Dim oRegEx As Object
    Dim dlgSaveAs As FileDialog                                         ' Boites de dialogues
    Dim objFDFS As FileDialogFilters
    Dim fdf As FileDialogFilter
    Dim I As Integer, wSelectedeMails As Integer                        ' Itérateurs
    Dim sFileName As String                                             ' Fichiers
    Dim sTempFolder As String, sTempFileName As String                  ' Répertoire temporaire
    Dim sTargetFolder As String, strCurrentFile As String               ' Répertoire cible
    Dim bContinue As Boolean                                            ' Initialisation des variables booléennes
    Dim bAskForFileName As Boolean
    Dim bRemoveMailAfterExport As Boolean
    Set oSelection = Application.ActiveExplorer.Selection               ' Obtenir tous les courriels sélectionnés
    wSelectedeMails = oSelection.Count                                  ' Obtenir le nombre de courriels sélectionnés
    If wSelectedeMails < 1 Then                                         ' Assurez-vous qu'au moins un élément est sélectionné
        Call MsgBox("Veuillez sélectionner au moins un email", _
            vbExclamation, "Enregistrer en PDF")
        Exit Sub
    End If
    bContinue = MsgBox("Vous êtes sur le point d'exporter " & wSelectedeMails & " " & _
        "emails en tant que fichiers PDF, voulez-vous continuer ? If you Yes, you'll " & _
        "devrez d'abord spécifier le nom du dossier dans lequel les fichiers seront stockés", _
        vbQuestion + vbYesNo + vbDefaultButton1) = vbYes
    If Not bContinue Then Exit Sub
    Set objWord = CreateObject("Word.Application")                      ' Démarrer Word et initialiser l'objet
        objWord.Visible = False                                         ' Ne pazs afficher Word
    sTargetFolder = AskForTargetFolder(cFolder)                         ' Définir le dossier cible, où enregistrer les courriels
    If sTargetFolder = "" Then
        objWord.Quit
        Set objWord = Nothing
        Exit Sub
    End If
    ' Une fois enregistré en PDF, supprimer le courriel ?
    bRemoveMailAfterExport = False 
    bRemoveMailAfterExport = MsgBox("Une fois que l'e-mail a été " & _
        "exporté et enregistré sur votre disque, souhaitez-vous " & _
        "le conserver dans votre boîte aux lettres ou le supprimer ?" & _ 
        vbCrLf & vbCrLf & _
        "Cliquez sur Oui pour le conserver. " & vbCrLf & _ 
        "Cliquez sur Non pour le supprimer.", _
        vbQuestion + vbYesNo + vbDefaultButton1) = vbNo
    bAskForFileName = True                                              
    If (wSelectedeMails > 1) Then                                       ' Si plusieurs courriels, choisir s'il faut voir les noms de fichiers.
        bAskForFileName = MsgBox("Vous êtes sur le point de sauvegarder " & _ 
            wSelectedeMails & " " & "les courriers électroniques sous " & _  
            "forme de fichiers PDF. Voulez-vous voir " & wSelectedeMails & _ 
            " des invites pour que vous puissiez mettre à jour le nom " & _ 
            "du fichier ou utiliser le fichier automatique automatisé " & _ 
            "(donc pas d'invite)." & vbCrLf & vbCrLf & _
            "Cliquez sur Oui pour voir les invites."  & vbCrLf & _
            "Cliquez sur Non pour laisserr faire l'automate.", _
            vbQuestion + vbYesNo + vbDefaultButton2) = vbYes
        MsgBox "ATTENTION : Vous ne verrez pas de progression à l'écran " & _ 
            "(malheureusement, Outlook ne le permet pas)." & _ 
            vbCrLf & vbCrLf & _
            "Si vous exportez beaucoup d'e-mails, le processus peut " & _
            "prendre un certain temps. La meilleure façon de voir que " & _ 
            "les choses fonctionnent consiste à ouvrir une fenêtre " & _ 
            "d'explorateur et de voir comment les fichiers sont ajoutés " & _ 
            "au dossier sélectionné." & _ 
            vbCrLf & vbCrLf & _
            "Une fois l'opération terminée, vous verrez un message de " & _ 
            "retour d'information.", _
            vbInformation + vbOKOnly
    End If
    If bAskForFileName Then                                             
        Set dlgSaveAs = objWord.FileDialog(msoFileDialogSaveAs)         ' Ouvrir la boîte de dialogue Enregistrer sous
        Set objFDFS = dlgSaveAs.Filters                                 ' Déterminer l'indice de filtre pour l'enregistrement d'un fichier PDF
        I = 0                                                           ' Obtenir tous les filtres et vérifier l'existence de "PDF".
        For Each fdf In objFDFS
            I = I + 1
            If InStr(1, fdf.Extensions, "pdf", vbTextCompare) > 0 Then
                Exit For
            End If
        Next fdf
        Set objFDFS = Nothing
        dlgSaveAs.FilterIndex = I                                       ' Définir l'indice de filtre à pdf-files
    End If
    Set objFSO = CreateObject("Scripting.FileSystemObject")             ' Obtenir le dossier temporaire de l'utilisateur où l'élément doit être stocké
    sTempFolder = objFSO.GetSpecialFolder(2)
    Set objFSO = Nothing
    On Error Resume Next                                                ' Commencer le traitement unitaire des courriels sélectionnés.
    For I = 1 To wSelectedeMails
        Set oMail = oSelection.Item(I)                                  ' Récupérer le courriel sélectionné
        sTempFileName = sTempFolder & "\outlook.mht"                    ' Construire le nom de fichier pour le fichier mht temporaire
        If Dir(sTempFileName) Then Kill (sTempFileName)                 ' Tuer le fichier précédent s'il est déjà présent
        oMail.SaveAs sTempFileName, olMHTML                             ' Enregistrez le fichier mht et l'ouvrir dans Word sans l'afficher.
        Set objDoc = objWord.Documents.Open (FileName:=sTempFileName, Visible:=False, ReadOnly:=True)
        sFileName = oMail.Subject                                       ' Construire le nom de fichier à partir de l'objet du message
        Set oRegEx = CreateObject("vbscript.regexp")                    ' Assainir le nom de fichier, supprimer les caractères indésirables
        oRegEx.Global = True
        oRegEx.Pattern = "[\\/:*?""<>|]"
        ' Ajouter la date du courriel reçu comme préfixe
        sFileName = sTargetFolder & Format(oMail.ReceivedTime, "yyyy-mm-dd_Hh-Nn") & _
            "_" & Trim(oRegEx.Replace(sFileName, "")) & ".pdf"
        If bAskForFileName Then
            sFileName = AskForFileName(sFileName)
        End If
        If Not (Trim(sFileName) = "") Then
            Debug.Print "Save " & sFileName
            If Dir(sFileName) <> "" Then                                ' S'il existe déjà, supprimer d'abord le fichier
                Kill (sFileName)
            End If
            ' Enregister au format PDF
            objDoc.ExportAsFixedFormat OutputFileName:=sFileName, _
                ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
                wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=0, To:=0, _
                Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
                CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
                BitmapMissingFonts:=True, UseISO19005_1:=False
            objDoc.Close (False)                                        ' Fermer une fois sauvegardé sur le disque
            If bRemoveMailAfterExport Then                              ' Déplacer le courriel dans les éléments supprimés ?
                If Dir(sFileName) <> "" Then                            ' Seulement si le courriel a bien été exporté.
                    oMail.Delete
                End If
            End If
        End If
    Next I
    Set dlgSaveAs = Nothing
    On Error GoTo 0
    On Error Resume Next
    objWord.Quit                                                        ' Fermer le document et Word
    On Error GoTo 0
    Set oSelection = Nothing                                            ' Nettoyage des objets
    Set oMail = Nothing
    Set objDoc = Nothing
    Set objWord = Nothing
    Set oRegEx = Nothing
    MsgBox "Vos fichiers sont prêts ! " & vbCrLf & vbCrLf & _
    "Les e-mails sélectionnés ont été exportés vers " & sTargetFolder & vbCrLf & vbCrLf & _
    "Manderley-AI espère vous avoir pu vous être utile !" & vbCrLf & vbCrLf & _
    "Dans l'affirmative, n'hésitez pas à adresser vos dons à :" & vbCrLf & _
    "l.talarico@ciblexperts.com ;-)", vbSystemModal, "Manderley-AI vous remercie !"
    ' TODO : Optimiser le code
    ' TODO : Ajouter un peu d'intelligence à ce petit automate ;-)
End Sub