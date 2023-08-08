# Automatically download attachments

This VBA script is designed to extract attachments from emails in a specific folder in your Outlook account. It cycles through the emails in the selected account's incoming folder, checks if they contain attachments, and if the sender's email address matches the specified address, saves the attachments to a predefined destination folder.

## Presets
Enable the developer tab: File>Options>Customize Ribbon. In the box on the right, select the 'Developer' checkbox.

<a href="https://uploaddeimagens.com.br/images/004/570/473/full/Captura_de_tela_2023-08-08_180151.png?1691528548"><img src="https://uploaddeimagens.com.br/images/004/570/473/full/Captura_de_tela_2023-08-08_180151.png?1691528548" height="300px"></a>

## Usage
1. Open Microsoft Outlook.
2. Press `ALT + F11` to open the Visual Basic for Applications Editor.
3. Create a new module (if you don't already have one) and paste the provided code into it.
4. Modify the `strSaveFolder` and `targetEmail` variables according to your needs:
    - `strSaveFolder`: Specifies the folder where attachments will be saved. By default, it is set to `"C:\Users\ellencrist\Downloads\TestFolder\"`. Replace with the desired path.
    - `targetEmail`: Enter the email address of the sender whose emails you want to filter and extract attachments.
5. Modify the `Set targetFolder` line to point to the correct folder in your Outlook account.
6. Close the VBA Editor and return to Outlook.
7. Run the script by pressing `ALT + F8`, select "Save Attachments" and click "Run".

Note: Make sure Microsoft Outlook referral is enabled.
<a href="https://uploaddeimagens.com.br/images/004/570/415/full/Captura_de_tela_2023-08-08_173356.png?1691527142"><img src="https://uploaddeimagens.com.br/images/004/570/415/full/Captura_de_tela_2023-08-08_173356.png?1691527142" alt="references vba" border="0"> </a>


## Operation Description

1. The script starts by creating an instance of Outlook and getting the Outlook namespace.
2. It specifies the destination folder where the attachments will be saved and the email address of the sender whose emails will be filtered.
3. The script iterates over the emails in the specific account's incoming folder.
4. Checks whether the email contains attachments and whether the sender's email address matches the specified address.
5. If both conditions are met, the script cycles through the email attachments and saves them in the destination folder.
6. After completing the iteration, displays a message stating how many attachments were saved or if no attachments were found.

## Demonstration
<img src="https://s11.gifyu.com/images/ScZqu.gif">

## Warnings

- Make sure you replace the `targetFolder` variable value with the correct folder path in your Outlook account.
- This script is designed to run in Microsoft Outlook and requires permissions to access your email folders.
- Remember that Outlook automation can vary between different Outlook versions and security settings.
