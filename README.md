# smtpro
A simple email automation utility

For first time users, run the SMTPro_Setup.py, it will generate the required folders and template files for use.

If you need to modify the config you could do so by either rerunning the setup script, or by modifying the .yaml file in the config folder.

For now this utility is mainly for sending invoices out via Email without needing to use a paid software like Mailchimp or anything more advanced than that.

The setup script will create an excel file called 'OutboxTemplateFile.xlsx' which you could rename or copy as much as you'd like.
The only limits to using this template is that all of the headings in there by default must remain unchanged. You can add as many other columns as you'd like for organizational purposes, but the default columns must remain for the program to work properly.

The template does require some work to setup, but if you have a standard message you can make it work.
If you have multiple recipients or copied recipients, just separate those emails with '; '.
If you have attachments, just put the attachment name (including filetype) in the attachment column. As of yet, this program only supports one attachment per email, but I plan on developing that functionality in the future.

Also keep in mind, the program will export a .txt file with the results of every email sent (success, or failure).

I'll add more to this README overtime as I make changes and think of better ways of explaining how to use the utility.
