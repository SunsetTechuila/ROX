$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

$body = @{
	"LoginForm[username]" = "login"
	"LoginForm[password]" = "password"
}
$Parameters = @{
	Uri = "https://placeholder.salesdoc.io/site/login"
	UseBasicParsing = $true
	Method = "POST"
	Body = $body
	WebSession = $session
}
Invoke-WebRequest @Parameters | Out-Null

$body = @{
	"CLIENT_ID"                         = "n2_193"
	"orderDate"                         = ""
	"AGENT_ID"                          = "d0_4"
	"HAND_EDIT"                         = "0"
	"PRICE_TYPE"                        = "e6_15"
	"store"                             = "d0_1"
	"OrderDetail[167][PRICE]"           = "4120"
	"OrderDetail[167][COUNT]"           = "1"
	"OrderDetail[167][VOLUME]"          = "400"
	"OrderDetail[167][SUMMA]"           = "4120"
	"OrderDetail[167][PRODUCT_ID]"      = "d0_156"
	"OrderDetail[167][PRODUCT_TYPE_ID]" = "d0_27"
	"OrderDetail[167][UNIT]"            = "d0_4"
	"OrderDetail[167][UNIT_SYMBOL]"     = "гр."
	"COMMENT"                           = ""
	"COMMENT_2"                         = ""
	"skidkaManual"                      = "0"
	"bonus_type"                        = "-1"
	"save_and_add"                      = ""
	"startCell"                         = ""
}
$Parameters.Uri = "https://placeholder.salesdoc.io/orders/createOrder/createAjax"

for ($i = 0; $i -lt 500; $i++) {
	$body.orderDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
	$Parameters.Body = $body
	Invoke-WebRequest @Parameters | Out-Null
	Start-Sleep 2
}
