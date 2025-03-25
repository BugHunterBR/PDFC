import win32com.client as client

outlook = client.Dispatch('Outlook.Application')
sender = 'fixed-term.Pedro.Kruger@boschrexroth.com.br '

message = outlook.CreateItem(0)
message.Display()
message.To = sender
message.Subject = 'Illegible certificate'
message.Body = 'The attached file is illegible. \n\nPlease return a new file.'
message.Attachments.Add('attachment_path')
message.Save()
message.Send()