$defaultLaptop = "HP 255 G8 AMD Ryzen 5 3500U 8GB 256GB SSD 15.6IN FHD Win 10 Pro"
$defaultDesktop = "Dell Vostro 3681 SFF (Core i5, 8GB RAM, VGA & HDMI output)"
$defaultPhone = "Samsung Galaxy A12 (New)"

$quoteString = ""
$totalPrice = 0


while ($true) {
    # First, prompt for what we are quoting for
    $modelPC = ""
    Write-Output "Ender product description. Quick access:"
    Write-Output "1. $defaultLaptop"
    Write-Output "2. $defaultDesktop"
    Write-Output "3. $defaultPhone"
    $modelPC = Read-Host "or any other text. (blank to finish)"

    if (-not $modelPC) {
        break
    }

    # Yes this could be made generic, but it's Powershell, not ASM
    if ($modelPC -eq "1") {
        Write-Host "Selecting $defaultLaptop"
        $modelPC = $defaultLaptop
    }

    elseif ($modelPC -eq "2") {
        Write-Host "Selecting $defaultDesktop"
        $modelPC = $defaultDesktop
    }

    elseif ($modelPC -eq "3") {
        Write-Host "Selecting $defaultPhone"
        $modelPC = $defaultPhone
    }

    # Read the quantity, check its a round number
    $NumberPC = 0
    do {
        try {
            $numOk = $true
            [int]$NumberPC = Read-host "> Quantity"
            } # end try
        catch {$numOK = $false}
        } # end do 
    until ($numOK)

    $pricePC = 0
    do {
        try {
            $numOk = $true
            Write-Output "Price as integer (or float!), excluding VAT & currency symbols"
            [float]$pricePC = Read-Host "> Price"
            } # end try
        catch {$numOK = $false}
        } # end do 
    until ($numOK)
    
    # Add the price of all to the total
    $totalPrice += ($pricePC * $NumberPC)

    # Format the price value
    if ($NumberPC -gt 1) {
        $strPricePC = "£{0:N2} + vat ea" -f $pricePC
    } else {
        $strPricePC = "£{0:N2} + vat" -f $pricePC
    }

    # Add a line to the quote
    $quoteString += "$NumberPC x $modelPC @ <b>$strPricePC</b><br>"
}

# Check if shipping price is confirmed
$doShipping = ""
while($doShipping -ne "y")
{
    if ($doShipping -eq 'n') {
        break
    }
    $doShipping = Read-Host "Is shipping price TBC? [y/n]"
}

# If we know how much shipping is, get the price
if ($doShipping -eq 'n') {
    do {
        try {
            $numOk = $true
            Write-Output "Shipping price as float, excluding VAT & currency symbols"
            [float]$shippingPrice = Read-Host "> Price"
            } # end try
        catch {$numOK = $false}
        } # end do 
    until ($numOK)

    if ($shippingPrice -gt 0) {
        # Add a line for shipping
        $quoteString += "Carriage incurred and to forward on to end user (inc. increased liability) @ <b>£{0:N2}</b><br>" -f $shippingPrice
    }
} else {
    # Otherwise add the generic TBC
    $quoteString += "Carriage incurred and to forward on to end user (inc. increased liability) @ <b>TBC</b> (<Reason here>)<br>"
}

# Add the total to the quote
$quoteString += "<b>Total: £{0:N2} + vat</b><br><i>Please note all quotes are subject to availability of stock at time of confirmation</i>" -f $totalPrice

# Print the total amount & turn the string into RTF in the clipboard
Write-Output ("Total price is {0:N2}. Copied full quote to clipboard" -f $totalPrice)
Set-Clipboard -Value $quoteString -AsHtml