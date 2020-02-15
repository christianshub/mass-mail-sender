# MailSender
Send personalized mass emails through gmail (Windows)

### 1. About

a WPF application written for a client.

<img src="https://github.com/christianshub/MailSender/blob/master/Billeder/App.jpg" width="500">


'Mail sender' sends multiple e-mails using an excel-file made by the
user. The excel-file should include one column with names (column 1),
and one column with corrosponding e-mails (column 2). 
This way, the user will be able to send multiple e-mails with the same
content, but with different names in the beginning of the content.

- The excel file:

![](https://github.com/christianshub/MailSender/blob/master/Billeder/Excel.jpg)

- Mail 1:

![](https://github.com/christianshub/MailSender/blob/master/Billeder/Mail1.jpg)

- Mail 2:

![](https://github.com/christianshub/MailSender/blob/master/Billeder/Mail2.jpg)


### 2. REQUIREMENTS:   

- 2.1. Create an excel file
  + 2.1.1. Fill column 1 (top-down) with names (e.g.: Christian)
  + 2.1.2. Fill column 2 (top-down) with the corrosponding mails (e.g.: Christian@yahoo.com)

- 2.2. Allow less secure apps:
  + 2.2.1. Sign in to your Google Admin console. ...
  + 2.2.2. Click Security > Basic settings. ...
  + 2.2.3. Under Less secure apps, select Go to settings for less secure apps.
  + 2.2.4. In the subwindow, select the Allow users to manage their access to less secure apps radio button.

### 3. INSTRUCTIONS:

- 3.1. Open the program
- 3.2. Login
- 3.3. Input subject
- 3.4. Input your text (content)
- 3.5. Attach the created excel file by pressing "Open..."

### 4. TODO:

- 4.1. Add login-screen
- 4.2. Allow other mail clients
- 4.3. Choose between danish and english (currently only in danish)
    
