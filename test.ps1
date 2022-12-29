$outlook = new-object -comobject outlook.application

 

$email = $outlook.CreateItem(0)

$email.To = biostar1020@gmail.com

$email.Subject = "New email test"

$email.Body = "This is a testing email"

$email.Attachments.add("C:\Users\talharthi\Desktop\0.txt")

 

$email.Send()

$outlook.Quit()
