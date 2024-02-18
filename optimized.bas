Option Explicit
Private Const xlsHeaders As Variant = ("")
Private Const strFolder As String = """C:\"""
Declare Function AtEndOfStream Lib "msvbvm60.dll" (ByVal hFile As Long) As Long
Private Function GetSelectedEmails() As Outlook.Selection
    Dim olApp As Outlook.Application                                        ' Initialisation des variables
    Dim olNs As Outlook.NameSpace
    Dim olFolder As Outlook.Folder
    Dim olItems As Outlook.Items
    Set olApp = Outlook.Application                                         ' Récupération de l'application Outlook
    Set olNs = olApp.GetNamespace("MAPI")                                   ' Récupération du Namespace par défaut
    Set olFolder = olNs.GetDefaultFolder(olFolderInbox)                     ' Accès au dossier Boîte de réception
    Set olItems = olFolder.Selection                                        ' Récupération de tous les e-mails sélectionnés
    If olItems.Count = 0 Then                                               ' Gestion d'erreur
        MsgBox "Aucun e-mail n'est sélectionné.", vbCritical, _
            "Sélectionnez un ou plusieurs e-mails."
        Exit Function
    End If
    GetSelectedEmails = olItems                                             ' Renvoi de la sélection
    Set olItems = Nothing                                                   ' Libération de la mémoire
    Set olFolder = Nothing
    Set olNs = Nothing
    Set olApp = Nothing
End Function
Private Function GetDestFolder(Optional strPath As String = "Environ('USERPROFILE')") As String
  Dim fd As FileDialog                                                      ' Déclaration de l'objet FileDialog
  fd.Title = "Sélectionnez le répertoire de destination"                    ' Initialisation du titre
  fd.AllowMultiSelect = False                                               '
  fd.MsoFileDialogType = msoFileDialogFolderPicker                          ' Définition du type de sélection (dossier)
  If fd.Show = True Then                                                    ' Affichage de la boîte de dialogue
    GetDestFolder = fd.SelectedItems(1)                                     ' Récuperation du chemin du dossier sélectionné
  Else                                                                      '
    GetDestFolder = strFolder                                               ' Annulation de la sélection
  End If                                                                    '
End Function
Private Function GetDomainName(ByVal strEmail As String) As String
  Dim intAtPos As Integer                                                   ' Initialisation des variables
  Dim intDotPos As Integer
  Dim intLength As Integer
  intAtPos = InStr(strEmail, "@") + 1                                       ' Déterminer la position du @
  intDotPos = InStrRev(strEmail, ".")                                       ' Déterminer la position du .
  intLength = intDotPos - intAtPos                                          ' Compter le nombre de caractère du domaine sans son extension
  GetDomainName = Mid(strEmail, intAtPos, intLength)                        ' Renvoyer le domaine sans son extension
End Function
Private Function GetCleanSubject(strSubject As String) As String
    Dim i As Long                                                           ' Initialisation des variables
    Dim strWord As String
    For i = LBound(Split(strSubject, " ")) _
        To UBound(Split(strSubject, " "))
    strWord = Split(strSubject, " ")(i)
    If Len(strWord) > 3 Then
        CleanSubject = Replace(UCase(CleanSubject & _
            strWord & " "), "_", " ")
    End If
  Next
  If Len(CleanSubject) > 0 Then                                             ' Suppression du dernier "_"
    GetCleanSubject = Replace(UCase(Left(CleanSubject, _
        Len(CleanSubject) - 1)), "_", " ")
  End If
End Function
Private Function SetFileName(ByVal strFileName As String) As String
    Dim dlgSaveAs As FileDialog
    Dim wResponse As VBA.VbMsgBoxResult
    Dim wPos As Integer
    Set dlgSaveAs = objWord.FileDialog(msoFileDialogSaveAs)
    dlgSaveAs.InitialFileName = strFileName                                 ' Définir le nom du fichier
    If dlgSaveAs.Show = -1 Then                                             ' Afficher la boîte de dialogue "Enregistrer sous"
        strFileName = dlgSaveAs.SelectedItems(1)
        If Right(strFileName, 4) <> ".pdf" Then                             ' Vérifier que le pdf est sélectionné
            wResponse = MsgBox( _
                "Désolé, seul l'enregistrement au format PDF " & _
                "est supporté pour le moment." & vbNewLine & vbNewLine & _
                "Sauvegarder au format PDF ?", vbInformation + vbOKCancel)
            If wResponse = vbCancel Then
                sFileName = ""
            ElseIf wResponse = vbOK Then
                wPos = InStrRev(strFileName, ".")
                If wPos > 0 Then
                    strFileName = Left(strFileName, wPos - 1)
                End If
                strFileName = strFileName & ".pdf"
            End If
        End If
    End If
    SetFileName = strFileName                                               ' Renvoie le nom du fichier
End Function
Private Function GetPagesCount(objPdf As Acrobat.AcroPDDoc)
    GetPagesCount = objPdf.GetNumPages                                      ' Récupération du nombre de pages du document
End Function
Private Function SetTbFormat(Tb As Range, _
            Optional ByRef strTbName As String = "", _
            Optional ByRef strTbStyle As String = "*", _
            Optional wHeaders As XlYesNoGuess = xlYes, _
            Optional wTotals As XlYesNoGuess = xlYes) As Range
    If Tb.ListObject Is Nothing Then                                        ' Si le Tb n'existe pas déjà alors le créer:
        If Tb.Count = 1 Then Set Tb = Tb.CurrentRegion                      ' Si Tb ne représente qu'une seule cellule alors étend la plage:
        If strTbName > "" Then                                              ' Création du Tableau Structuré en attribuant le nom passé ou en prenant celui attribué par EXCEL:
            Tb.Parent.ListObjects.Add(xlSrcRange, Tb, , wHeaders).Name = strTbName
        Else
            Tb.Parent.ListObjects.Add xlSrcRange, Tb, , wHeaders
        End If
        If strTbStyle <> "*" Then                                           ' Modifie le style s'il ne faut pas prendre celui par défaut, ou pas de style si vide:
            Tb.Parent.ListObjects(Tb.ListObject.Name).TableStyle = strTbStyle
        End If
    End If
    strTbName = Tb.ListObject.DisplayName                                   ' Renseigne le nom du Tableau Structuré et son style:
    strTbStyle = Tb.Parent.ListObjects(Tb.ListObject.Name).TableStyle
    Tb.ListObject.ShowTotals = True
    Set SetTbFormat = Tb                                                    ' Renvoie la plage du Tableau Structuré:
End Function
Private Function GetExcelSheet(strFileName As String, _
            Optional strTbName As String = "ANALYSE DE PIECES", _
            Optional arrHeaders As Variant = xlsHeaders) As Worksheet

    Dim fso As New FileSystemObject                                         ' Déclaration de variables
    Dim xlApp As Excel.Application
    Dim xlWb As Excel.Workbook
    Dim ws As Worksheet
    Dim i As Integer
    
    If fso.FileExists(strFileName) Is Nothing Then                          ' Vérifier l'existence du fichier
        Set xlWb = xlApp.Workbooks.Add                                      ' Création du nouveau fichier
        xlWb.SaveAs FileName:=strFileName                                   ' Enregistrement du nouveau fichier
        Set ws = xlWb.Sheets.Add(After:=xlWb.Sheets(xlWb.Sheets.Count))     ' Créer la feuille de travail
        ws.Name = strTbName                                                 ' Renommer la feuille active
    Else                                                                    ' Si le fichier existe déjà
        Set xlWb = xlApp.Workbooks.Open(strFileName)                        ' Ouvrir le fichier existant
        Set ws = xlWb.Sheets(strTbName)                                     ' Activer la feuille de travail
        If ws Is Nothing Then                                               ' Vérifier l'existence de la feuille
            Set ws = xlWb.Sheets.Add(After:=xlWb.Sheets(xlWb.Sheets.Count)) ' Créer la feuille de travail
            ws.Name = strTbName                                             ' Renommer la feuille active
        End If
    End If

    For Each strHeader In xlsHeaders                                        ' Mettre en forme les en-têtes
        i = i + 1                                                           ' Incrémenter l'index de colonne
        ws.Cells(1, i).Value = strHeader                                    ' Affecter les valeurs d'en-tête
    Next strHeader

    GetExcelSheet = ws                                                      ' Renvoyer la feuille
    
    Set ws = Nothing                                                        ' Libérer la mémoire
    Set xlWb = Nothing
    Set xlApp = Nothing
    Set fso = Nothing

End Function
Private Function SetExcelData(ws As Excel.Worksheet, arrData As Variant, Optional intPos As Integer = 2)
    SetExcelData = ws                                                      ' Renvoyer la feuille
End Function
Private Function GetPdfFile(FilePath As String) As String
    Dim fso As New FileSystemObject                                         ' Déclaration de variables
    Dim objAcro As Acrobat.CAcroApp
    Dim objDoc As Acrobat.CAcroPDDoc
    Dim NewFilePath As String
    If fso.FileExists(FilePath) Then                                        ' Vérification de l'existence du fichier
        Set objAcro = New Acrobat.CAcroApp                                  ' Création de l'objet Acrobat
        Set objDoc = objAcro.OpenDoc(FilePath)                              ' Conversion du fichier
        NewFilePath = Replace(FilePath, _
            fso.GetExtensionName(FilePath), ".pdf")
        objDoc.ExportAsFixedFormat _
            OutputFileName:=NewFilePath, _
            ExportFormat:=17, _
            OpenAfterExport:=False
        objDoc.Close                                                        ' Fermeture du document et de l'application
        objAcro.Quit
        Set objDoc = Nothing                                                ' Libération de la mémoire
        Set objAcro = Nothing
    Else
        MsgBox "Le fichier " & strFilePath & _
            " n'existe pas.", _
            vbCritical, "Fichier introuvable !"
        Exit Function
    End If
    GetPdfFile = NewFilePath                                                ' Renvoi du chemin du fichier PDF
End Function
Private Function SetMailAsDone(oMail As Outlook.MailItem) As Boolean
    SetMailAsDone = True
End Function
Private Function DeleteMail(oMail As Outlook.MailItem) As Boolean
    oMail.Delete
    DeleteMail = True
End Function
Sub DoMyJob()

    Dim oSel As Outlook.Selection                                                   ' Initialisation des variables
    Dim oMail As Outlook.MailItem
    Dim objAtt As Outlook.Attachment
    Dim objFSO As FileSystemObject
    
    Dim objPdf As Acrobat.AcroPDDoc
    
    Dim bSetFileName As Boolean, _
        bContinue As Boolean, _
        bDelMail As Boolean
    
    Dim strTargetFolder As String, _
        strCurrentFile As String, _
        strTargetFile As String, _
        strExt As String
    
    Dim sTmpFolder As String, _
        sTmpFileName As String, _
        strTmpFilePath As String
        
    Dim dlgSaveAs As FileDialog
    Dim objFDFS As FileDialogFilters
    Dim fdf As FileDialogFilter

    Set oSel = Application.ActiveExplorer.Selection                         ' Obtenir toute la sélection
    If oSel.Count < 1 Then                                                  ' Vérifier qu'au moins un élément est sélectionné
        Call MsgBox("Veuillez sélectionner au moins un courriel", _
            vbExclamation, _
            "Enregistrer au format PDF")
        Exit Sub
    Else
        bSetFileName = True
        bSetFileName = _
            MsgBox("Vous êtes sur le point de convertir " & oSel.Count & " " & _
                "courriels en autant de fichiers PDF. " & vbCrLf & vbCrLf & _
                "Voulez-vous voir chacun des " & oSel.Count & " " & _
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
        bContinue = MsgBox( _
            "Vous êtes sur le point de convertir " & oSel.Count & _
            "courriels en autant de fichiers PDF." & vbCrLf & _
            "Ils pourront être renommés automatiquement." & vbCrLf & _
            "Soutaitez-vous continuer ? " & vbCrLf & vbCrLf & _
            "Si oui, vous devrez sélectionner le répertoire dans lequel" & _
            " vous souhaitez les enregistrer.", _
            vbQuestion + vbYesNo + vbDefaultButton1) = vbYes
        If Not bContinue Then Exit Sub
    End If

    strTargetFolder = GetDestFolder(Environ("USERPROFILE"))
    If strTargetFolder = "" Then Exit Sub

    If bSetFileName Then                                                    ' Ouvrir la boîte de dialogue "Enregistrer sous"
        Set dlgSaveAs = objWord.FileDialog(msoFileDialogSaveAs)
        Set objFDFS = dlgSaveAs.Filters                                     ' Déterminer l'indice de filtre pour l'enregistrement d'un fichier pdf
        i = 0
        For Each fdf In objFDFS                                             ' Obtenir tous les filtres pour vérifier l'existence du "pdf".
            i = i + 1
            If InStr(1, fdf.Extensions, "pdf", vbTextCompare) > 0 Then
                Exit For
            End If
        Next fdf
        Set objFDFS = Nothing
        dlgSaveAs.FilterIndex = i                                           ' Définir le FilterIndex à pdf-files
    End If

End Sub