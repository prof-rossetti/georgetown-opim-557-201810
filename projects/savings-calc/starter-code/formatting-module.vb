'
' SAVINGS CALCULATOR PROJECT > STARTER CODE > FORMATTING MODULE
'
' Author: Prof. Rossetti <prof.mj.rossetti@gmail.com>
' Source: https://github.com/prof-rossetti/georgetown-opim-557-201810/blob/master/projects/savings-calc/starter-code/formatting-module.vb
' License: Students, feel free but not obligated to use this module in your project as long as you retain this attribution section. If you wrote something like this on your own, no need to attribute. If this code inspired you to write your own code, please still consider providing an attribution link to this file's GitHub URL.
'

' FormatUSD returns a string formatted as US Dollar currency.
Public Function FormatUSD(ByVal Price) ' not declaring a datatype here because price can be integer or double.
    FormatUSD = Format(Price, "Currency") ' or ... Format(Price, "$##,##0.00")
End Function

' FormatPct returns a string formatted as a percentage.
Public Function FormatPct(ByVal Percentage As Double)
    FormatPct = Format(Percentage, "Percent") ' or ... Format(Percentage, "###0.00%")
End Function
