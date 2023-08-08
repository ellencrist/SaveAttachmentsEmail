Attribute VB_Name = "Módulo1"

'strSaveFolder = "C:\Users\SeuNome\Downloads\PastaTeste"
    
    Sub SalvarAnexos()
    Dim objOutlook As Object
    Dim objNamespace As Object
    Dim objFolder As Object
    Dim objMail As Object
    Dim objAttachment As Object
    Dim strSaveFolder As String
    Dim attachmentCount As Integer ' Variável para contar o número de anexos salvos
    
    ' Pasta de destino para salvar os anexos
    strSaveFolder = "C:\Users\ellencrist\Downloads\PastaTeste\"
    
    ' Endereço de e-mail da conta que enviou o anexo para filtrar os e-mails
    Dim targetEmail As String
    targetEmail = "emaildoremetente@gmail/outlook.com"
    
    ' Inicializar o objeto Outlook
    Set objOutlook = CreateObject("Outlook.Application")
    ' Obter o namespace do Outlook
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    
    ' Obter a pasta da conta específica
    Dim targetFolder As Object
    Set targetFolder = objNamespace.Folders("meuemail_@outlook.com").Folders("Caixa de Entrada") ' Substitua "seu_email@example.com" pelo endereço da conta desejada
    
    ' Inicializar a contagem de anexos
    attachmentCount = 0
    
    ' Percorrer os e-mails na pasta da conta específica
    For Each objMail In targetFolder.Items
        ' Verificar se o e-mail contém anexos
        If objMail.Attachments.Count > 0 Then
            ' Verificar o endereço de e-mail do remetente
            If objMail.SenderEmailAddress = targetEmail Then
                ' Percorrer os anexos do e-mail
                For Each objAttachment In objMail.Attachments
                    ' Salvar o anexo na pasta especificada
                    objAttachment.SaveAsFile strSaveFolder & objAttachment.FileName
                    ' Incrementar a contagem de anexos
                    attachmentCount = attachmentCount + 1
                Next objAttachment
            End If
        End If
    Next objMail
    
    ' Exibir o aviso com base na contagem de anexos
    If attachmentCount > 0 Then
        MsgBox attachmentCount & " anexo(s) salvo(s) na pasta especificada.", vbInformation, "Anexos Salvos"
    Else
        MsgBox "Nenhum anexo encontrado na caixa de entrada da conta específica.", vbInformation, "Nenhum Anexo"
    End If
    
    ' Limpar a memória
    Set objAttachment = Nothing
    Set objMail = Nothing
    Set targetFolder = Nothing
    Set objNamespace = Nothing
    Set objOutlook = Nothing
End Sub

