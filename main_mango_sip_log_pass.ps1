$login = "XXXXXXXX" #Логин для входа в кабинет
$password = "AHDlS03327" #Пароль для входа в кабинет
$IURL = "https://lk.mango-office.ru/400183801/400254558/members/index" # Путь на страницу где список сотрудников


$ie = New-Object -com "InternetExplorer.Application"
$ie.visible = $true
$ie.silent = $true
$ie.Navigate($IURL)
Wait-Event -Timeout 3
$IE.Document.Forms | %{$_.item("login")} | % {$_.value = $login}
$IE.Document.Forms | %{$_.item("password")} | % {$_.value = $password }
Wait-Event -Timeout 3
$IE.Document.Forms | %{$_.getElementsByClassName("prime")} | % {$_.Click()}
Wait-Event -Timeout 3
$a1 = (($ie.Document.IHTMLDocument3_getElementByID("b-members-data")).text | ConvertFrom-Json)
$a1
$IE.quit()
