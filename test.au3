#include <Array.au3>
#include <MsgBoxConstants.au3>
Global $aTest[0][12]
ReDim $aTest[1][12]
for $i =0 to 11
   $aTest[0][$i]=""
next
ReDim $aTest[2][12]

$aTest[1][0]=1
$aTest[1][1]=2

_ArrayDisplay($aTest)

MsgBox($MB_SYSTEMMODAL, "Delimited string to add", ubound($aTest))