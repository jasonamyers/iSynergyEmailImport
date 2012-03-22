This program uses IMAP to read a mailbox, saves all the attachments to the specified directory, and deletes the messages.

To use this program set the following:
- client.Host = "host";  change the word host to your IMAP host name
- client.Username = "User";  Change User to the IMAP user, and remember if you include a domain use <domain>\\<user>
- client.Password = "password"; Change password to the IMAP user's password
- client.AttachmentDirectory = "PathForAttachments";  Change Path for Attachments remember to use \\ instead of \ for example: c:\\directory\\subdirectory\file.txt


It makes use of the libraries from http://www.componentace.com/.NET_Email_component_pop3_smtp_imap_csharp.htm.  They are awesome.
