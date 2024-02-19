' --------------------------------------------------
' Macro Outlook pour enregistrer un ou plusieurs
' éléments sélectionnés en tant que fichiers pdf
' sur votre disque dur. Vous pouvez sélectionner
' autant de mails ' que vous voulez et chaque mail
' sera sauvegardé sur votre disque.
'
' Nécessite :
' - Winword (référencé par late-bindings)
' - Excel
' - Microsoft Scripting Runtime
' - Acrobat
' - PDFToolsKit
'
' @voir Inspirée de https://github.com/cavo789/vba_outlook_save_pdf
' @voir https://github.com/Manderley-tools/outlook2pdf
' --------------------------------------------------
Option Explicit
Private Const cFolder As String = """C:\"""
Private objWord As Object
'Declare Function AtEndOfStream Lib "msvbvm60.dll" (ByVal hFile As Long) As Long

' --------------------------------------------------
' Demander à l'utilisateur le dossier dans lequel stocker les courriels
' --------------------------------------------------
Private Function AskForTargetFolder(ByVal sTgtFolder As String) As String
    Dim dlgSaveAs As FileDialog
    sTgtFolder = Trim(sTgtFolder)
    If Not (Right(sTgtFolder, 1) = "\") Then                            ' Vérifier que sTgtFolder se termine par une barre oblique.
        sTgtFolder = sTgtFolder & "\"
    End If
    Set dlgSaveAs = objWord.FileDialog(msoFileDialogFolderPicker)       ' Récupérer l'objet
    With dlgSaveAs
        .Title = "Sélectionnez le dossier dans lequel enregistrer ces courriels"
        .AllowMultiSelect = False
        .InitialFileName = sTgtFolder
        .Show
        On Error Resume Next
        sTgtFolder = .SelectedItems(1)
        If Err.Number <> 0 Then
            sTgtFolder = ""
            Err.Clear
        End If
        On Error GoTo 0
    End With
    If Not (Right(sTgtFolder, 1) = "\") Then                            ' Vérifier que sTgtFolder se termine par une barre oblique.
        sTgtFolder = sTgtFolder & "\"
    End If
    AskForTargetFolder = sTgtFolder
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
   For i = LBound(Split(strSubject, " ")) To UBound(Split(strSubject, " ")) ' Diviser la chaîne de caractères en mots en utilisant un espace comme séparateur
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
    Dim acrApp As New Acrobat.AcroApp
    Dim objPDF As New Acrobat.AcroPDDoc
    Set acrApp = CreateObject("AcroExch.App")
    Set objPDF = CreateObject("AcroExch.PDDoc")
    If objPDF.Open(pdfFile) Then
        GetNumberOfPages = objPDF.GetNumPages
    Else
        GetNumberOfPages = 0
    End If
End Function

Private Function GetPDFDate(pdfFile As String, Optional dDate As String) As String
    Dim acrApp As New Acrobat.AcroApp
    Dim objPDF As New Acrobat.AcroPDDoc
    Set acrApp = CreateObject("AcroExch.App")
    Set objPDF = CreateObject("AcroExch.PDDoc")
    If objPDF.Open(pdfFile) Then
        If objPDF.GetInfo("Modified") <> "" Then
            GetPDFDate = objPDF.GetInfo("Modified")
            Debug.Print "Date de modification du PDF : " & objPDF.GetInfo("Modified")
        Else
            GetPDFDate = objPDF.GetInfo("Created")
            Debug.Print "Date de modification du PDF : " & objPDF.GetInfo("Created")
        End If
    End If
End Function

Function GetLastDate(strFilePath As String, Optional dDate As String) As String
 
    Dim oFSO As Scripting.FileSystemObject
    Dim oFl As Scripting.File
     
    Set oFSO = New Scripting.FileSystemObject
    Set oFl = oFSO.GetFile(strFilePath)
    
    GetLastDate = oFl.DateLastModified
    MsgBox oFl.DateLastModified
 
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
    
    Dim TS_Err_Number As Integer
    Dim TS_Err_Description As String
    Dim TS_Methode_Err As String
    Dim TS_Generer_Erreur As String
    Dim TS_MsgBox_Erreur As String
    
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
    
Gest_Err:                                                                   ' Fin du traitement:
    TS_Err_Number = Err.Number
    TS_Err_Description = Err.Description
    If Err.Number <> 0 Then
        If TS_Methode_Err = TS_Generer_Erreur Then
            Err.Raise TS_Err_Number, _
            "TS_ConvertirPlageEnTS", _
            TS_Err_Description
        End If
        If TS_Methode_Err = TS_MsgBox_Erreur Then
            MsgBox TS_Err_Number & " : " & TS_Err_Description, vbInformation, "TS_ConvertirPlageEnTS"
        End If
    End If
    Err.Clear
End Function
' --------------------------------------------------
' Faire le travail, traiter les courriels sélectionnés et les exporter au format PDF
' Déplace les courriels dans éléments supprimés si deamndé.
' --------------------------------------------------
Sub SaveAsPDFfiles()
' Définition des constantes et des variables
    Const wdExportFormatPDF = 17                                        ' Initialisation des constantes
    Const wdExportOptimizeForPrint = 0
    Const wdExportAllDocument = 0
    Const wdExportDocumentContent = 0
    Const wdExportCreateNoBookmarks = 0
    Dim oSel As Outlook.Selection                                       ' Initialisation des variables
    Dim oMail As Outlook.MailItem
    Dim objAtt As Outlook.Attachment
    Dim objFSO As FileSystemObject
    Dim objPDF As Acrobat.AcroPDDoc
    Dim objDoc As Object                                                ' Utilise late-bindings
    Dim oRegEx As Object
    Dim objExcel As New Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objData As New Dictionary
    Dim dlgSaveAs As FileDialog                                         ' Boites de dialogues
    Dim objFDFS As FileDialogFilters
    Dim fdf As FileDialogFilter
    Dim i As Integer                                                    ' Itérateurs
    Dim x As String
    Dim y As Integer
    Dim N As Integer
    Dim P As Integer
    Dim oSelCount As Integer
    Dim sFileName As String                                             ' Fichiers
    Dim sTmpFolder As String                                            ' Répertoire temporaire
    Dim sTmpFileName As String                                          ' Nom du fichier temporaire
    Dim sTmpFilePath As String                                         ' Chemin complet du fichier temporaire
    Dim sTgtFolder As String                                            ' Répertoire cible
    Dim sTgtFileName As String                                          ' Nom du fichier cible
    Dim sTgtFilePath As String                                          ' Chemin complet du fichier cible
    Dim sCurFile As String '                                             ' Nom du fichier courant
    Dim sExt As String                                                  ' Extension du fichier
    Dim strSender As String                                             ' Emetteur du courriel
    Dim strReceiver As String                                           ' Premier destinataire du courriel
    Dim bContinue As Boolean                                            ' Initialisation des variables booléennes
    Dim bAskForFileName As Boolean
    Dim bRemoveMailAfterExport As Boolean
    Dim isFile As Boolean
    Dim myFile As String                                                ' Nom du fichier retraité
    Dim myDesc As String                                                ' Description
    Dim myID As Integer                                                 ' Identifiant numérique
    Dim myDate As String                                                  ' Date d'émission du document
    Dim myTime As String                                                  ' Heure d'émission du document
    Dim myType As String                                                ' Type de document
    Dim mySender As String                                              ' Emetteur du document
    Dim myReceiver As String                                            ' Destinataire du document
    Dim myObject As String                                              ' Objet du document
    Dim myRef As String                                                 ' Référence du document
    Dim myAmount As Currency                                            ' Montant
    Dim mySynt As String                                                ' Synthèse
    Dim myObs As String                                                 ' Observations
    Dim myQuest As String                                               ' Questions
    Dim myAsk As String                                                 ' Demande de pièces
    Dim nbPages As Integer                                              ' Nombre de pages
'
    Set oSel = Application.ActiveExplorer.Selection                     ' Obtenir tous les courriels sélectionnés
    oSelCount = oSel.Count                                              ' Obtenir le nombre de courriels sélectionnés
    If oSelCount < 1 Then                                               ' Assurez-vous qu'au moins un élément est sélectionné
        Call MsgBox("Veuillez sélectionner au moins un email", _
            vbExclamation, "Enregistrer en PDF")
        Exit Sub
    End If
    bContinue = MsgBox("Vous êtes sur le point d'exporter " & oSelCount & " " & _
        "emails en tant que fichiers PDF, voulez-vous continuer ? If you Yes, you'll " & _
        "devrez d'abord spécifier le nom du dossier dans lequel les fichiers seront stockés", _
        vbQuestion + vbYesNo + vbDefaultButton1) = vbYes
    If Not bContinue Then Exit Sub
    Set objWord = CreateObject("Word.Application")                      ' Démarrer Word et initialiser l'objet
        objWord.Visible = False                                         ' Ne pas afficher la fenetre Word
    sTgtFolder = AskForTargetFolder(Environ("USERPROFILE"))                            ' Définir le dossier cible, où enregistrer les courriels
    If sTgtFolder = "" Then
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
    If (oSelCount > 1) Then                                             ' Si plusieurs courriels, choisir s'il faut voir les noms de fichiers.
        bAskForFileName = MsgBox("Vous êtes sur le point de sauvegarder " & _
            oSelCount & " " & "les courriers électroniques sous " & _
            "forme de fichiers PDF. Voulez-vous voir " & oSelCount & _
            " des invites pour que vous puissiez mettre à jour le nom " & _
            "du fichier ou utiliser le fichier automatique automatisé " & _
            "(donc pas d'invite)." & vbCrLf & vbCrLf & _
            "Cliquez sur Oui pour voir les invites." & vbCrLf & _
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
        i = 0                                                           ' Obtenir tous les filtres et vérifier l'existence de "PDF".
        For Each fdf In objFDFS
            i = i + 1
            If InStr(1, fdf.Extensions, "pdf", vbTextCompare) > 0 Then
                Exit For
            End If
        Next fdf
        Set objFDFS = Nothing
        dlgSaveAs.FilterIndex = i                                       ' Définir l'indice de filtre à pdf-files
    End If
    Set objFSO = CreateObject("Scripting.FileSystemObject")             ' Obtenir le dossier temporaire de l'utilisateur où l'élément doit être stocké
    sTmpFolder = objFSO.GetSpecialFolder(2)
    Set objFSO = Nothing
    On Error Resume Next                                                ' Commencer le traitement unitaire des courriels sélectionnés.
    sTgtFilePath = sTgtFolder & "ANALYSE.xlsx"
    Debug.Print "Le fichier " & sTgtFilePath & " existe : " & isFile
    If Dir(sTgtFilePath) <> "" Then                                     ' Vérifier si le fichier Excel existe déjà
        Debug.Print Dir(sTgtFilePath)
        Set objWorkbook = objExcel.Workbooks.Open(sTgtFilePath)         ' Ouverture du fichier Excel existant
        If Not objWorkbook.Sheets("ANALYSE DE PIECES") <> "" Then       ' Vérification si la feuille de calcul existe
            Set objWorksheet = objWorkbook.Sheets.Add(After:=objWorkbook.Sheets(objWorkbook.Sheets.Count))   ' Créer une nouvelle feuille de calcul
            objWorksheet.Name = "ANALYSE DE PIECES"                     ' Définir le nom de la feuille de calcul
        End If
        Set objWorksheet = objWorkbook.Sheets("ANALYSE DE PIECES")      ' Activation de l'onglet ANALYSE
        y = objWorksheet.Cells.Find(What:="*").Row + 1                  ' Incrémentation du nombre de lignes
    Else
        Set objWorkbook = objExcel.Workbooks.Add                        ' Création du classeur Excel
        objWorkbook.SaveAs FileName:=sTgtFilePath, _
            FileFormat:=xlOpenXMLWorkbook                               ' Créer une nouvelle feuille de calcul
        Set objWorksheet = objWorkbook.Sheets.Add _
            (After:=objWorkbook.Sheets(objWorkbook.Sheets.Count))
        objWorksheet.Name = "ANALYSE DE PIECES"                         ' Définir le nom de la feuille de calcul
        Set objWorksheet = objWorkbook.Sheets("ANALYSE DE PIECES")      ' Activation de l'onglet ANALYSE
        objWorkbook.Sheets("Feuil1").Delete
        objWorksheet.Cells(1, 1).Value = "Fichier"                      ' En-tête de la feuille
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
    For i = oSelCount To 1 Step -1
        Set oMail = oSel.Item(i)                                        ' Récupérer le courriel sélectionné
        strSender = UCase(GetDomain(oMail.Sender.Address))              ' Récupérer l'adresse de l'expéditeur
        If strSender = "" Then                                          ' SI l'adresse n'est pas trouvée
            strSender = UCase(oMail.Sender.Name)                        ' Récupérer le nom de l'expéditeur
        End If
        Debug.Print "Emetteur : " & strSender
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
        strReceiver = UCase(GetDomain(oMail.Recipients(1).Address))     ' Récupérer le domaine du premier destinataire ou son nom
        If strReceiver = "" Then
            strReceiver = UCase(oMail.Recipients(1).Name)
        End If
        Debug.Print "Destinataire : " & strReceiver
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
        sTmpFileName = sTmpFolder & "\outlook.mht"                      ' Construire le nom de fichier pour le fichier mht temporaire
        Debug.Print "Fichier temporaire : " & sTmpFileName
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        isFile = objFSO.FileExists(sTmpFileName)
        Set objFSO = Nothing
        Debug.Print "Le fichier temporaire existe : " & isFile
        If isFile Then Kill sTmpFileName                                ' Effacer le fichier précédent s'il est déjà présent
        oMail.SaveAs sTmpFileName, olMHTML                              ' Enregistrez le fichier mht et l'ouvrir dans Word sans l'afficher.
        Set objDoc = objWord.Documents.Open _
            (FileName:=sTmpFileName, Visible:=False, ReadOnly:=True)
        sFileName = oMail.Subject                                       ' Construire le nom de fichier à partir de l'objet du message
        Set oRegEx = CreateObject("VBScript.RegExp")                    ' Assainir le nom de fichier, supprimer les caractères indésirables
        oRegEx.Global = True
        oRegEx.Pattern = "[\\/:*?""<>|]"
        myObject = UCase(Trim(oRegEx.Replace(CleanSubject(sFileName), "")))
        myDate = Format(oMail.ReceivedTime, "yyyymmdd_HhNnSs")
        myTime = Format(oMail.ReceivedTime, "Hh:Nn:Ss")
        myType = "_MEL_"
        myFile = UCase(myDate & myType & strSender & "_" & strReceiver & "_" & myObject) & ".pdf"
        myDate = Format(oMail.ReceivedTime, "dd/mm/yyyy")
        Debug.Print "Nom du fichier cible : " & sFileName
        sFileName = sTgtFolder & myFile                                 ' Ajouter la date du courriel reçu comme préfixe
        Debug.Print "Chemin du fichier cible : " & sFileName
        If bAskForFileName Then sFileName = AskForFileName(sFileName)
        If Not (Trim(sFileName) = "") Then
            Debug.Print "Chemin du fichier cible : " & sFileName
            If Dir(sFileName) <> "" Then Kill (sFileName)               ' S'il existe déjà, supprimer d'abord le fichier
            objDoc.ExportAsFixedFormat OutputFileName:=sFileName, _
                ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
                wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=0, To:=0, _
                Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
                CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
                BitmapMissingFonts:=True, UseISO19005_1:=False
            objDoc.Close (False)                                        ' Fermer une fois sauvegardé sur le disque
            nbPages = GetNumberOfPages(sFileName)
            objWorksheet.Cells(y, 1).Value = myFile                     ' Renseigner le tableau de bord Excel avec le courriel
            objWorksheet.Cells(y, 2).Value = myObject
            objWorksheet.Cells(y, 3).Value = y - 1
            objWorksheet.Cells(y, 4).Value = myDate
            objWorksheet.Cells(y, 5).Value = myTime
            objWorksheet.Cells(y, 6).Value = "Courriel"
            objWorksheet.Cells(y, 7).Value = strSender
            objWorksheet.Cells(y, 8).Value = strReceiver
            objWorksheet.Hyperlinks.Add _
                Anchor:=objWorksheet.Cells(y, 9), _
                Address:="" & myFile & "", _
                TextToDisplay:="" & myObject & ""
            objWorksheet.Cells(y, 10).Value = ""
            objWorksheet.Cells(y, 11).Value = 0
            objWorksheet.Cells(y, 12).Value = ""
            objWorksheet.Cells(y, 13).Value = ""
            objWorksheet.Cells(y, 14).Value = ""
            objWorksheet.Cells(y, 15).Value = ""
            objWorksheet.Cells(y, 16).Value = nbPages
            objWorkbook.Save                                            ' Enregistrement du classeur Excel
            y = y + 1                                                   ' Incrémentation de la position de la ligne dans le tableau Excel
            nbPages = 0                                                 ' Réinitialisation du nombre de page
            If bRemoveMailAfterExport Then                              ' Déplacer le courriel dans les éléments supprimés ?
                If Dir(sFileName) <> "" Then oMail.Delete               ' Seulement si le courriel a bien été exporté.
            End If
        End If
        P = 0                                                           ' Ajout des pièces jointes
        For N = 1 To oMail.Attachments.Count
            Set objAtt = oMail.Attachments.Item(N)
            If Not Left(objAtt.FileName, 6) = "image0" Then
                P = P + 1                                               'Encodage du n° de la pièce jointe
                x = P
                If P < 10 Then x = "0" & P
                Set objFSO = CreateObject("Scripting.FileSystemObject") ' Récupération de l'extension de la pièce jointe
                sExt = objFSO.GetExtensionName(oMail.Attachments.Item(N).FileName)
                Select Case sExt                                        ' Traitement conditionnel en fonction de l'extension
                Case "doc", "docx", "xls", "xlsx", "ppt", "pptx"        ' Pour les pièces jointes directement convertibles en PDF
                    Debug.Print "L'extension de la pièce jointe est : " & UCase(sExt)
                    ' TODO : Coder la conversion au format PDF
                    ' TODO : Récupérer la date du dernier enregistrement (cf. BuiltinDocumentProperties("Last Save Time") ?)
                    ' TODO : Récuperer l'auteur  (cf. BuiltinDocumentProperties("Author") ?)
                Case "msg", "eml"                                       ' Pour les pièces jointes susceptibles d'en contenir d'autres
                    Debug.Print "ATTENTION => La pièce jointe est un mail au format : " & UCase(sExt)
                    ' TODO : Coder la conversion au format PDF
                    ' TODO : Vérifier s'ils comportent des pièces jointes ou non et les traiter de la même manière le cas échéant
                    ' TODO : Factoriser le tout pour alléger le code
                Case "pdf"                                              ' Pour les pièces jointes déjà au format PDF
                    Debug.Print "La pièce jointe est déjà au format " & UCase(sExt)
                    objAtt.SaveAsFile (sTgtFolder & oMail.Attachments.Item(N).FileName)
                    nbPages = GetNumberOfPages(sTgtFolder & oMail.Attachments.Item(N).FileName)
                    Debug.Print "Date de modification de la pièce : " & GetPDFDate(sTgtFolder & oMail.Attachments.Item(N).FileName)
                    Kill (sTgtFolder & oMail.Attachments.Item(N).FileName)
                Case Else                                               ' Pour tous les autres types de pièces jointes
                    Debug.Print "ATTENTION => L'extension de la pièce jointe est : " & UCase(sExt)
                    ' TODO : Cf. Traitement ci-dessous
            End Select
            myObject = UCase(Trim(oRegEx.Replace(CleanSubject(objAtt.FileName), "")))
            myDate = Format(oMail.ReceivedTime, "yyyymmdd_HhNnSs")
            myType = "_P" & x & "_"
            myFile = myDate & myType & strSender & "_" & strReceiver & "_" & myObject
            myObject = UCase(Replace(objAtt.FileName, "." & sExt, ""))
            objAtt.SaveAsFile (sTgtFolder & myFile)
            myDate = FileDateTime(sTgtFolder & myFile)
            Debug.Print "Date de dernière modification de la pièce jointe : " & myDate
            objWorksheet.Cells(y, 1).Value = myFile                     ' Renseigner le tableau de bord Excel avec la pièce jointe
            objWorksheet.Cells(y, 2).Value = myObject
            objWorksheet.Cells(y, 3).Value = y - 1
            objWorksheet.Cells(y, 4).Value = Format(myDate, "dd/mm/yyyy")
            objWorksheet.Cells(y, 5).Value = Format(myDate, "Hh:Nn:Ss")
            objWorksheet.Cells(y, 6).Value = "Pièce jointe"
            objWorksheet.Cells(y, 7).Value = strSender
            objWorksheet.Cells(y, 8).Value = strReceiver
            objWorksheet.Hyperlinks.Add _
                Anchor:=objWorksheet.Cells(y, 9), _
                Address:="" & myFile & "", _
                TextToDisplay:="" & myObject & ""
            objWorksheet.Cells(y, 10).Value = UCase(sExt)
            objWorksheet.Cells(y, 11).Value = Format(0, "# ##0,00 €")
            objWorksheet.Cells(y, 12).Value = "Dernière modification le " & Format(myDate, "dd/mm/yyyy à Hh:Nn:Ss")
            objWorksheet.Cells(y, 13).Value = "Pièce jointe n°" & x
            objWorksheet.Cells(y, 14).Value = ""
            objWorksheet.Cells(y, 15).Value = ""
            objWorksheet.Cells(y, 16).Value = nbPages
            objWorkbook.Save                                            ' Enregistrement du classeur Excel
            y = y + 1                                                   ' Incrémentation de la position de la ligne dans le fichier Excel
            nbPages = 0                                                 ' Réinitialisation du nombre de page
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
    
    TS_ConvertirPlageEnTS objWorksheet.Range("A1"), "PIECES", "*", xlYes ' Formatter le tableau Excel
    Dim str As String
    ' A reprendre pour tenir compte de la variable y
    objWorksheet.Range("C2:C16").NumberFormatLocal = "# ##0"
    objWorksheet.Range("D2:D16").NumberFormatLocal = "jj/mm/aaaa"
    objWorksheet.Range("E2:E16").NumberFormatLocal = "hh:mm:ss"
    objWorksheet.Range("K2:K16").NumberFormatLocal = "# ##0,00 €;[Rouge]- # ##0 €"
    objWorksheet.Range("P2:P16").NumberFormatLocal = "# ##0"
    ' -----------------------------------------
    objWorksheet.Cells(y, 3).Formula2 = "=SOUS.TOTAL(104,['#])"
    objWorksheet.Cells(y, 11).Formula2 = "=SOUS.TOTAL(109,[Montant])"
    objWorksheet.Cells(y, 16).Formula2 = "=SOUS.TOTAL(109,[Pages])"
    objWorksheet.Columns("A:C").Group                                   ' Création et fermeture du groupe de colonnes (à supprimer plus tard)
    objWorksheet.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
    objWorksheet.Columns("D:P").AutoFit
    objWorksheet.Columns("D:E").HorizontalAlignment = xlCenter
    objWorksheet.Columns("J").HorizontalAlignment = xlLeft
    objWorkbook.Save                                                    ' Enregistrement des modifications apportées au Classeur
    objWorkbook.Close savechanges:=True                                 ' Enregistrement du classeur dans le dossier spécifié
    Set dlgSaveAs = Nothing
    On Error GoTo 0
    On Error Resume Next
    objWord.Quit                                                        ' Fermer le document et Word
    On Error GoTo 0
    Set oSel = Nothing                                                  ' Nettoyage des objets
    Set oMail = Nothing
    Set objDoc = Nothing
    Set objExcel = Nothing
    Set objWorkbook = Nothing
    Set objWorksheet = Nothing
    Set objWord = Nothing
    Set oRegEx = Nothing
    MsgBox "Vos fichiers sont prêts ! " & vbCrLf & vbCrLf & _
    "Les e-mails sélectionnés ont été exportés vers " & sTgtFolder & vbCrLf & vbCrLf & _
    "Manderley-AI espère vous avoir pu vous être utile !" & vbCrLf & vbCrLf & _
    "Dans l'affirmative, n'hésitez pas à adresser vos dons à :" & vbCrLf & _
    "l.talarico@ciblexperts.com ;-)", vbSystemModal, "Manderley-AI vous remercie !"
    ' TODO : Optimiser le code
    ' TODO : Ajouter un peu d'intelligence à ce petit automate ;-)
End Sub

