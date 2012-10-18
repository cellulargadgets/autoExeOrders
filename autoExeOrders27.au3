#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author: John L.
 Version: 27.0
 Script Function: Process Orders


; ## TODO LIST ## ;


; Stamp.com does not recognize entire state name. needs to appreviate
; snake skin with clip would header snake skin only.
; ambigous address


; when buy.com doesn't have any order the window doesn't close
; completed order Log

;buy.com 3T = 10oz  when it is 7.2+0.4oz

; buy.com multiple item would proceed even when a ref_id isn't found
; paypal doens't show when packaging string does not exist
; paypal 3 doesn't stop when in_stock = 0 
; no new order does not close broswer
; don't log ebay purse organizer entries

; done - if paypal has "Coupon Discount" then skip item  line 1400
; debug - Print international label
; done - change address for amazon buy.com and paypal123

#ce ----------------------------------------------------------------------------



#include <Array.au3>
#include <IE.au3>
#include <file.au3>
#include <Date.au3>
#include <GUIListView.au3>
#include "mysql.au3"
AutoItSetOption("WinWaitDelay", 1000)
AutoItSetOption("SendKeyDelay", 100)
AutoItSetOption("SendKeyDownDelay", 3)

Global $timestamp = @MON & "/" & @MDAY & "/" & @YEAR & "  " & @HOUR & ":" & @MIN & ":" & @SEC
Global $timestamp_filename = @YEAR&"_"&@MON&"_"&@MDAY&@HOUR&@MIN
Global $body=""
Global $global_log=""
Global $gmailShipstreamPasswd="", $emailTo=""
Global $amazon_user="", $amazon_pw="", $buycom_user="", $buycom_pw="", $paypal1_user="", $paypal1_pw="", $paypal2_user="", $paypal2_pw="", $paypal3_user="", $paypal3_pw=""
;Global $todayDate = StringReplace(_NowDate(),"/","_")

getSecret()

Start_Programs()

Sleep(3000)

;ieA()
;ieB()
ieP1()
;ieP2()
;ieP3()


Func ieA ()
$ieA = _IECreate()
_IENavigate($ieA, "https://sellercentral.amazon.com/gp/homepage.html")
_IEWaitForLoad($ieA, "Seller Central - Windows Internet Explorer", "https://sellercentral.amazon.com/gp/homepage.html")
Login($ieA, $amazon_user, $amazon_pw, "signin", "email", "password", "sign-in-button" )
sleep(2000)
_IEWaitForLoad($ieA, "Seller Central - Windows Internet Explorer", "https://sellercentral.amazon.com/gp/homepage.html")
processeAmazonOrders($ieA)
_IEQuit($ieA)
EndFunc

Func ieB ()
$ieB = _IECreate()
_IENavigate($ieB, "https://sellertools.marketplace.buy.com/login.aspx")
_IEWaitForLoad($ieB, "Buy.com - Marketplace Seller Tools - Windows Internet Explorer", "https://sellertools.marketplace.buy.com/login.aspx")
Login($ieB, $buycom_user, $buycom_pw, "aspnetForm", "ctl00$cphMiddle$txtEmail", "ctl00$cphMiddle$txtPassword", "ctl00$cphMiddle$btnSubmit" )
sleep(2000)
_IEWaitForLoad($ieB, "Buy.com - Marketplace Seller Tools - Windows Internet Explorer", "https://sellertools.marketplace.buy.com/login.aspx")
Buy_com_Marketplace($ieB)
Sleep(2000)
_IEQuit($ieB)
EndFunc


Func ieP1 ()
$ieP1 = _IECreate()
_IENavigate($ieP1, "http://paypal.com")
_IELoadWait($ieP1)
Login($ieP1, $paypal1_user, $paypal1_pw, "login_form", "login_email", "login_password", "submit.x" )
sleep(3000)
paypalOrders($ieP1, 1)
_IEQuit($ieP1)
EndFunc


Func ieP2 ()
$ieP2 = _IECreate()
_IENavigate($ieP2, "http://paypal.com")
_IELoadWait($ieP2)
Login($ieP2, $paypal2_user, $paypal2_pw, "login_form", "login_email", "login_password", "submit.x" )
sleep(3000)
paypalOrders($ieP2, 2)
_IEQuit($ieP2)
EndFunc


Func ieP3()
$ieP3 = _IECreate()
_IENavigate($ieP3, "http://paypal.com")
_IELoadWait($ieP3)
Login($ieP3, $paypal3_user, $paypal3_pw, "login_form", "login_email", "login_password", "submit.x" )
sleep(3000)
paypalOrders($ieP3, 3)
_IEQuit($ieP3)
EndFunc



WinClose("Stamps.com Pro")

	If $body <> "" Then
		_sendmail("Process Error Error(s)", $body)
	Endif
				
;	If $global_log <> "" Then
;		_Write_Log($global_log)
;	Endif		
		
Func Buy_com_Marketplace($oIE)
	;;; three different type item has problems
	
	Local $new_orders = True
	
	$new_orders = _Download_Open_Orders($oIE)
	
	If $new_orders = True Then
		
		Local $import_file = "OpenOrderExport_"&$timestamp_filename&".txt"
		;Local $import_file = "OpenOrderExport_2012_06_132050.txt"
		Local $i_t = TimerInit()
		Local $_array = _DelimFile_To_Array2D($import_file, @TAB, 34)
		Local $i_d = TimerDiff($i_t)
					
		$_array = BuyComCheckForDuplicate($_array) ;; Check for Order ID Duplicates
	
		processeBuycomOrders($oIE, $_array)
	
	EndIf

	;_IEQuit($oIE)

EndFunc



Func processeAmazonOrders($oAIE)
	
	Local $oOrderTable, $oOrderTableData, $oTable, $oTableData, $oProductTable, $oProductTableData, $splitString="", $address="", $splitAddress=""
	Local $i=0, $j=0, $w=0, $itemNum=1, $item=0, $productSplit="", $productString, $found = True, $inStock = True, $itemString="", $count=0
	Local $oShipTable, $oShipTableData, $shipType="", $emptyFound=False, $packagingType = "", $itemWeight = 0, $itemSKU = "", $packagingString=""
	Local $packaging="", $packagingWeight = 0, $link=""	, $noMatch = False, $addressName="", $itemType="", $arraySize=0, $addressChange= False, $skipTracking=False
	
	_IENavigate($oAIE, "https://sellercentral.amazon.com/gp/orders-v2/list/ref=id_myo_dnav_xx_?ie=UTF8&useSavedSearch=default")
	_IELoadWait($oAIE)
   Sleep(2000)  	
	;For $i=1 To 20 
	$oOrderTable = _IETableGetCollection ($oAIE, 4)
	$oOrderTableData = _IETableWriteToArray ($oOrderTable, True)
	
	;MsgBox(0, "UBound($oOrderTableData)", UBound($oOrderTableData))
	;Next
	If UBound($oOrderTableData) = 1 Then
		$oOrderTable = _IETableGetCollection ($oAIE, 6)
		$oOrderTableData = _IETableWriteToArray ($oOrderTable, True)
	EndIf

	;_ArrayDisplay($oOrderTableData)
	;MsgBox(0, $i , UBound($oOrderTableData)-1)

	Local $open_id_name_duplicate[UBound($oOrderTableData)-4][3]
	
	For $i=4 To UBound($oOrderTableData)-1
	   
		If StringInStr($oOrderTableData[$i][8], "Unshipped") Then
			
			$open_id_name_duplicate[$j][1] = $oOrderTableData[$i][5]
			$splitString = StringSplit($oOrderTableData[$i][4], " ", 2)
			
			$open_id_name_duplicate[$j][0] = $splitString[0]
			
			$j+=1
		
		Else
			_ArrayDelete($open_id_name_duplicate, $i)
			
		EndIf
		
    Next

	;_ArrayDisplay($open_id_name_duplicate)
	
	$open_id_name_duplicate = amazonNameDupCheck($open_id_name_duplicate)
	
	;_ArrayDisplay($open_id_name_duplicate)
	
	;MsgBox(0, "Ubound($open_id_name_duplicate)-1", Ubound($open_id_name_duplicate)-1)
    
	$w=0
	;MsgBox(0,"Ubound($open_id_name_duplicate)", Ubound($open_id_name_duplicate))
	
	
	
	While $w < Ubound($open_id_name_duplicate)
		
		If $open_id_name_duplicate[$w][2] = 1 Then ;;if duplicate name exist
			;_Log(@CR&$timestamp & @TAB & "Duplicated Name Exist " & " Order ID: " &$open_id_name_duplicate[$w][0]&@CR)
			$body &= "Amazon Order ID: "&$open_id_name_duplicate[$w][0]& " Multiple Orders with Same Person And Address" & @CR			
		
		ElseIf define_skip($open_id_name_duplicate[$w][0]) = True Then
		
			$body &= "Amazon Order ID: "&$open_id_name_duplicate[$w][0]& " Skipped" & @CR
		
		Else
				
				;_ArrayDisplay($open_id_name_duplicate)
				
				$oLinks = _IELinkGetCollection($oAIE)
				  For $oLink In $oLinks
					 If StringInStr($oLink.innerText, $open_id_name_duplicate[$w][0]) Then
						 ;MsgBox(0, "$oLink.innerText - $open_id_name_duplicate[$w][0]", $oLink.innerText&" - "&$open_id_name_duplicate[$w][0])
						_IEAction($oLink, "click")
						ExitLoop
					 EndIf
				  Next
			   
			   _IELoadWait($oAIE)
			   Sleep(2000)
			
			
			$getHTML = ""
			$getHTML = _IEBodyReadHTML($oAIE)
			$bLateShipment = False
			If StringInStr($getHTML, "Late Shipment") <> 0 Then
				$bLateShipment = True
			EndIf	
				
			;;Check for Late Shipment
			If	$bLateShipment = True Then
				$oProductTable = _IETableGetCollection ($oAIE, 10)
			   $oProductTableData = _IETableWriteToArray ($oProductTable, True)
			Else
			
			;;Check for Product ID Existence
			   ;For $b=4 to 20
			   $oProductTable = _IETableGetCollection ($oAIE, 9)
			   $oProductTableData = _IETableWriteToArray ($oProductTable, True)

			EndIf

			   ;_arraydisplay($oProductTableData)
			   ;Next
			   $item = 0
			   For $k=1 To UBound($oProductTableData)-1 Step 1
					If StringInStr($oProductTableData[$k][1], "Unshipped") <> 0 Then
						$item +=1
					EndIf
			   Next
			
			;_arraydisplay($oProductTableData)
		
			   ;MsgBox(0,"$item", $item)
			   
			   Local $itemName_SKU_Qty_Weight_Ptype_ItemType[$item][6]
			   
			   $itemNum=0

				For $k=1 To UBound($oProductTableData)-1 Step 1
				  
				  If StringInStr($oProductTableData[$k][1], "Unshipped") <> 0 Then
					
					$productSplit = StringSplit($oProductTableData[$k][0], @CR)

					$itemName_SKU_Qty_Weight_Ptype_ItemType[$itemNum][0] = $productSplit[1]
					 
				
					 $oTable = _IETableGetCollection ($oAIE, 8+$k*2)
					 $oTableData = _IETableWriteToArray ($oTable, True)
					 
					 ;_arraydisplay($oTableData)
					 
					 
					 
					If $bLateShipment = True Then
						$itemName_SKU_Qty_Weight_Ptype_ItemType[$itemNum][2] = $oTableData[1][2]
						$oTable = _IETableGetCollection ($oAIE, 9+$k*2)
						$oTableData = _IETableWriteToArray ($oTable, True)
						;_arraydisplay($oProductTableData)
						
						$itemName_SKU_Qty_Weight_Ptype_ItemType[$itemNum][1] = $oTableData[1][1]
					Else
						$itemName_SKU_Qty_Weight_Ptype_ItemType[$itemNum][2] = $oTableData[0][1]
						$itemName_SKU_Qty_Weight_Ptype_ItemType[$itemNum][1] = $oTableData[1][1]
					EndIf
					 $itemNum+=1
				  EndIf
				  
				  ;MsgBox(0,"$k",$k)

			   Next
					
				;; delete blanks in the array	
				
			  ;_arraydisplay($itemName_SKU_Qty_Weight_Ptype_ItemType) ;; check to see if there is any empty cells

				;For $k=0 to UBound($itemName_SKU_Qty_Weight_Ptype)-1 Step 1
				;   If $itemName_SKU_Qty_Weight_Ptype_ItemType[$k][0] = "" Or $itemName_SKU_Qty_Weight_Ptype_ItemType[$k][0] = " "Then
				;	 _ArrayDelete($itemName_SKU_Qty_Weight_Ptype_ItemType, $k)
				;   EndIf
			    ;Next
					
				;_arraydisplay($itemName_SKU_Qty)
				
				$found=True
				$inStock = True	
				
				;_arraydisplay($itemName_SKU_Qty_Weight_Ptype_ItemType)
				
				For $k=0 To UBound($itemName_SKU_Qty_Weight_Ptype_ItemType)-1 Step 1
				  ;If $itemName_SKU_Qty_Weight_Ptype_ItemType[$k][0] <> "" Or $itemName_SKU_Qty_Weight_Ptype_ItemType[$k][0] <> " "Then 
					If db_item_title_existence($itemName_SKU_Qty_Weight_Ptype_ItemType[$k][0], "amazon") = False Then						
						add_to_DB($itemName_SKU_Qty_Weight_Ptype_ItemType[$k][0], "name", "amazon", $itemName_SKU_Qty_Weight_Ptype_ItemType[$k][1], "amazon_id")
						$found = False
					EndIf
					
					If itemInStock($itemName_SKU_Qty_Weight_Ptype_ItemType[$k][0], "amazon") = 0 Then
						
						$inStock = False
					EndIf	
					
					
					If $found = False Then
						$body &= "Amazon Item Name: "&$itemName_SKU_Qty_Weight_Ptype_ItemType[$k][0]& " not found" & @CR
					ElseIf $inStock = False Then
						$body &= "Amazon Item Name: "&$itemName_SKU_Qty_Weight_Ptype_ItemType[$k][0]& " not in stock" & @CR
					EndIf
						
				Next
			
			;; End of Finding Proudct existence
			
			
			If $found = True And $inStock = True Then
				
				;;Get Customer Address
			   $address = ""
			   $addressName = ""

			   
			If addressChangeExist($open_id_name_duplicate[$w][0]) = True Then
				$address = addressChange($open_id_name_duplicate[$w][0])
				$addressName = addressChangeName($open_id_name_duplicate[$w][0])

			Else
				
				If $bLateShipment = True Then
				
				   $oTable = _IETableGetCollection ($oAIE, 7)
				   $oTableData = _IETableWriteToArray ($oTable, True)
				
				Else
				   $oTable = _IETableGetCollection ($oAIE, 6)
				   $oTableData = _IETableWriteToArray ($oTable, True)				
				
				EndIf
			   
			   $splitAddress = StringSplit($oTableData[0][0], @LF, 2)
			   
			   
			   ;_ArrayDisplay($oTableData)
			   ;_ArrayDisplay($splitAddress)
			   ;MsgBox(0,"UBound($splitAddress)", UBound($splitAddress)) ;; want to see if the UBound($splitAddress) need -1 or not
			   
				For $k=1 To UBound($splitAddress)-1
				   If StringinStr($splitAddress[$k], "Phone:") = 0 Then
						If $splitAddress[$k] <> "" Then
							If $k = 1 Then
								$address &= StringReplace ($splitAddress[$k], @CR, "")&@CR
								$addressName = StringReplace ($splitAddress[$k], @CR, "")
							Else
								$address &= StringReplace ($splitAddress[$k], @CR, "")&@CR
							EndIf
	
						EndIf
				   EndIf
				Next

				$addressNameSplit = StringSplit($addressName, " ",2)

					If UBound($addressNameSplit) > 2 then
						$addressName=""
						;MsgBox(0,"test",UBound($addressNameSplit)-1)
						For $k=0 To UBound($addressNameSplit)-1
							
							If $addressNameSplit[$k] <> "" Then

								If $addressName <> "" Then
									$addressName&=" "&$addressNameSplit[$k]
								Else
									$addressName&=$addressNameSplit[$k]
								EndIf
	
							EndIf
						Next
						
					EndIf


			   ;MsgBox(0,"address", $address)
		    EndIf ;; end of Address change
		   
		   	   $shipType = ""
		   
			   ;;Get Shipping Type
			   
			   If $bLateShipment = True Then			   
					$oShipTable = _IETableGetCollection ($oAIE, 8)
					$oShipTableData = _IETableWriteToArray ($oShipTable, True)
			   
			   Else
					$oShipTable = _IETableGetCollection ($oAIE, 7)
					$oShipTableData = _IETableWriteToArray ($oShipTable, True)			   
			   EndIf
			   
			   $shipType = $oShipTableData[3][1] ;;standard or expedite
			   
			   ;MsgBox(0,"$shipType", $shipType)
			   
			   ;;Print Shipping Label
				  ;;Fetch Packaging Type and weight
				  $emptyFound=False
				  $packagingType = ""
				  $itemWeight = 0
				  $itemSKU = ""
				  $itemType = ""
				  
					 For $k=0 To UBound($itemName_SKU_Qty_Weight_Ptype_ItemType)-1
						
						$packagingType = retrievePackagingType($itemName_SKU_Qty_Weight_Ptype_ItemType[$k][0], "amazon")  
						If $packagingType = "" Then
						   $emptyFound = True 
						Else
						   $itemName_SKU_Qty_Weight_Ptype_ItemType[$k][4] = $packagingType
						EndIf
						
						$itemWeight = retrieveWeight($itemName_SKU_Qty_Weight_Ptype_ItemType[$k][0], "amazon")
						If $itemWeight = 0 Then
						   $emptyFound = True 
						Else
						   $itemName_SKU_Qty_Weight_Ptype_ItemType[$k][3] = $itemWeight			   
						EndIf
						
						$itemType = retrieveItemType($itemName_SKU_Qty_Weight_Ptype_ItemType[$k][0], "amazon")
						If $itemType = "" Then
						   $emptyFound = True 
						Else					   
						   $itemName_SKU_Qty_Weight_Ptype_ItemType[$k][5] = $itemType
						EndIf   

						$itemSKU = retrieveSKU($itemName_SKU_Qty_Weight_Ptype_ItemType[$k][0], "amazon")
						If $itemSKU = "" Then
						   $emptyFound = True 
						Else
						   $itemName_SKU_Qty_Weight_Ptype_ItemType[$k][1] = $itemSKU
						EndIf   
						
						;MsgBox(0,"$k",$k)
					 Next
				  
				  ;MsgBox(0, "$emptyFound", $emptyFound)
				  ;_arraydisplay($itemName_SKU_Qty_Weight_Ptype_ItemType)
				  
				  $itemString=""
				  $arraySize=0
				  
				  ;;package String
					 If UBound($itemName_SKU_Qty_Weight_Ptype_ItemType) > 1 Then
						$arraySize = UBound($itemName_SKU_Qty_Weight_Ptype_ItemType)
						Local $itemArray[$arraySize][2]
						$count=0
						
						
						;MsgBox(0,"",Ubound($itemArray))
						
						For $m=0 To UBound($itemName_SKU_Qty_Weight_Ptype_ItemType)-1
						   
						   If $itemArray[0][0] = "" Then
							  $itemArray[0][0] = $itemName_SKU_Qty_Weight_Ptype_ItemType[$m][4]
							  $itemArray[0][1] = $itemName_SKU_Qty_Weight_Ptype_ItemType[$m][2]
							  $count+=1
							  
						   Else
							  $found = False
							  For $n=0 To UBound($itemArray)-1
								 
								 If $itemArray[$n][0] = $itemName_SKU_Qty_Weight_Ptype_ItemType[$m][4] Then
									$itemArray[$n][1]+=1
									$found = True
								 EndIf
							  
							  Next
							  
							  If $found <> True Then
								 $itemArray[$count][0] = $itemName_SKU_Qty_Weight_Ptype_ItemType[$m][4]	
								 $itemArray[$count][1] = $itemName_SKU_Qty_Weight_Ptype_ItemType[$m][2]
								 $count+=1
							  EndIf
						   
						   EndIf
						Next
						
						;_ArrayDisplay($itemArray)
						_ArraySort($itemArray)
						
						For $k=0 To Ubound($itemArray)-1
						   If $itemArray[$k][0] <> "" Then
							  $itemString &= $itemArray[$k][0]
							  $itemString &= $itemArray[$k][1]
						   EndIf
						Next
					 Else
						$itemString = $itemName_SKU_Qty_Weight_Ptype_ItemType[0][4]
						$itemString &= $itemName_SKU_Qty_Weight_Ptype_ItemType[0][2]
					 EndIf
					 
				 ; MsgBox(0,"$itemString",$itemString)
				  
				  ;_ArrayDisplay($itemString)
				  ;;package string existence
				  $packaging = ""
				  $packaging = define_packaging($itemString) 

				  ;;package string weight
				  $packagingWeight = 0
				  $packagingWeight = define_packaging_weight($itemString)
				  
				  If $packaging="" or $packagingWeight=0 Then
					 $emptyFound = True
					 
				  EndIf
				  
				  ;; total weight item weight * qty + packaging weight
				  $weightTotal=0
			   
				  ;MsgBox(0, "$emptyfound", $emptyFound)
				  
				  If $emptyFound = False Then

					 If UBound($itemName_SKU_Qty_Weight_Ptype_ItemType) > 1 Then
						
						For $k=0 To UBound($itemName_SKU_Qty_Weight_Ptype_ItemType)-1
						   $weightTotal+= $itemName_SKU_Qty_Weight_Ptype_ItemType[$k][3] * $itemName_SKU_Qty_Weight_Ptype_ItemType[$k][2]
						Next
						
					 Else
						
						$weightTotal = $itemName_SKU_Qty_Weight_Ptype_ItemType[0][3] * $itemName_SKU_Qty_Weight_Ptype_ItemType[0][2]
						
					 EndIf
					 
					 $weightTotal+= $packagingWeight
					 
					 ;MsgBox(0, "$weightTotal", $weightTotal)
					 
					 $noMatch = False
					 
					 ;;Open First-Class Layout or Priority Mail Layout and Print Shipping label
					 
					 
					 If $weightTotal < 13 Or $weightTotal = 13 Then
						Select
						   Case $shipType = "Standard"
							;MsgBox(0, "$shipType = Standard", $shipType)
							printFirstClass($address, $weightTotal, $open_id_name_duplicate[$w][0], "CellularGadgets", "BUBBLE_MAILER")
							  
						   Case $shipType = "Expedited"
							  If $itemString = "P112" Then
								;MsgBox(0, "$shipType = Expedite and itemstring = p112", $shipType)		
								printSmallFlatRateBox($address, $weightTotal, $open_id_name_duplicate[$w][0], "CellularGadgets")
								
							  Else
								;MsgBox(0, "$shipType = Expedite", $shipType)
								 printFirstClass($address, $weightTotal, $open_id_name_duplicate[$w][0], "CellularGadgets", "BUBBLE_MAILER")
								
							  EndIf
						   
						   Case Else
							 $noMatch = True

						EndSelect						
						
					 ElseIf $weightTotal > 13 Then
						;MsgBox(0, "$weightTotal > 13", $weightTotal)
						$weightTotal = $weightTotal / 16
						printPriority($address, $weightTotal, $open_id_name_duplicate[$w][0], "CellularGadgets")
						
					 Else
						$noMatch = True
					 
					 EndIf ;; end of $weight_total < 13 Or $weight_total = 13
					 
					
					;MsgBox(0, "$noMatch", $noMatch)
					
					 If $noMatch <> True Then
					
					ControlClick("Stamps.com Pro", "", 32551)
					Sleep(2000)
					
					$bodyTxt = ""
					$skipTracking = False
					 ;;Add Tracking Number
						WinWaitActive("Stamps.com Pro")
						Sleep(2000)
						$trackingIE = _IEAttach("Stamps.com Pro", "Embedded")
						_IELoadWait($trackingIE)
						;MsgBox(0,"$addressName", $addressName)
					For $k = 0 to 15
						 $bodyTxt = _IEBodyReadText($trackingIE)
						
						;MsgBox(0,"StringInStr ($bodyTxt, $addressName)", StringInStr ($bodyTxt, $addressName))
						If StringInStr ($bodyTxt, $addressName) <> 0 Then
							$trackingTable = _IETableGetCollection($trackingIE, 1)
							$trackingData = _IETableWriteToArray($trackingTable, True)
							;MsgBox(0,StringInStr ($trackingData[1][4], $addressName), $trackingData[1][4]&"addressName="&$addressName)
							If StringInStr($trackingData[1][4], $addressName) <> 0 Then
								$temp_tracking = $trackingData[1][6]
								ExitLoop
							EndIf
						
							ControlClick("Stamps.com Pro", "", "[CLASS:ToolbarWindow32; INSTANCE:1]", "Primary", 1, 124,15 ) 
							_IELoadWait($trackingIE)
							Sleep(20000)
						EndIf
						
						If $k > 7 Then
							ControlClick("Stamps.com Pro", "", "[CLASS:ToolbarWindow32; INSTANCE:1]", "Primary", 1, 124,15 ) 
							_IELoadWait($trackingIE)							
							Sleep(60000)
						EndIf
					Next
					
					If $k > 14 Then
						$body &= "Unable to locate " & $addressName & "on Stamp.com Shipping Table" & "with $k value of " & $k & "on Amazon" & @CR
						$skipTracking = True
					EndIf
					
					Sleep(2000)
					ControlClick("Stamps.com Pro", "", 32513)
					Sleep(2000)					
					
					
					If $skipTracking <> True Then
					   $oConfirmShipment = _IEGetObjByName($oAIE, "Confirm shipment")
					   _IEAction($oConfirmShipment, "click")
					   
					   _IELoadWait($oAIE)
					   Sleep(1000)
					   
					   Local $oInput = _IEGetObjById($oAIE, "carrierNameDropDown_UNSHIPPEDITEMS")
						   _IEFormElementOptionselect($oInput, "USPS" , 1, "byValue")

						   $oInput = _IEGetObjByName($oAIE, "shippingMethod_UNSHIPPEDITEMS")
						   _IEFormElementSetValue($oInput, "First-Class")

						   $oInput = _IEGetObjByName($oAIE, "trackingID_UNSHIPPEDITEMS")
						   _IEFormElementSetValue($oInput, $temp_tracking)
						   
						   $oInput = _IEGetObjByName($oAIE, "shipping_co_id")
						   _IEFormElementOptionselect($oInput, "1" , 1, "byValue")
						   
						   
						   $oInput = _IEGetObjByName($oAIE, "Confirm Shipment")
						   _IEAction($oInput, "click")				  
				  
					EndIf
			  
				  ;;Print Packing Slip
					 ;;Fetch Universal SKU
					 ;;Fetch Packaging Type Name
					
					$link = "https://sellercentral.amazon.com/gp/orders-v2/packing-slip?ie=UTF8&orderID="&$open_id_name_duplicate[$w][0]
					$iePrint = _IECreate()
					_IENavigate($iePrint, $link)
					
					Sleep(1000)
					
					WinWaitActive("Print")
					;WinClose("Print")
					
					Sleep(2000)
					
					Send("!{F4}")
					
					Sleep(10000)
					
					printAmazonPackingSlip($itemName_SKU_Qty_Weight_Ptype_ItemType, $packaging)
					
					_IEQuit($iePrint)
					
					 ;;Update Inventory
						
					inventoryUpdate($itemName_SKU_Qty_Weight_Ptype_ItemType, 1, 2)
						
					EndIf ;;  end of if $noMatch <> True
				  EndIf ;;end of $empty Found
			   
			   
			EndIf	;; end of $found true and $instock true

			;;Clean Array 
			 If $open_id_name_duplicate[$w][2] <> 1 Then 
				For $k=0 To UBound($itemName_SKU_Qty_Weight_Ptype_ItemType) Step 1
					_ArrayDelete($itemName_SKU_Qty_Weight_Ptype_ItemType, $k)
				Next
			 EndIf

		EndIf ;;end of dup name exist or skip exist
		
		


		_IENavigate($oAIE, "https://sellercentral.amazon.com/gp/orders-v2/list/ref=id_myo_dnav_xx_?ie=UTF8&useSavedSearch=default")
		_IELoadWait($oAIE)
	   Sleep(2000)  
		
		$w+=1
	
	WEnd
	
	_IELinkClickByText($oAIE, "LOGOUT")
	_IELoadWait($oAIE)
	Sleep(3000)
	
	_IEQuit($oAIE)
	
	;MsgBox(0,"$w",$w)
	
	
	
	;;Log off	
	;;Update Inventory									

	

EndFunc



Func processeBuycomOrders($BoIE, $OpenOrdersArray)


	Local $incomplete = 0, $completed=0
	Local $order_timestamp = @MON & "/" & @MDAY & "/" & @YEAR & "  " & @HOUR & ":" & @MIN & ":" & @SEC
	Local $singleItemWeight=0, $totalItemWeight=0, $totalWeight=0, $singleItemString="", $packageWeight=0, $packageType="", $itemString="", $totalQty=0

	Local $orderid="", $singleItemRefid="", $singleItemQty=0, $ShipToName="", $ShipToCompany="", $ShipToStreet1="", $refidString=""
	Local $ShipToStreet2="", $ShipToCity="", $ShipToState="", $ShipToZipCode="", $ShippingMethodId=0, $found=False
	Local $ItemStringCount=0, $address="", $inStock=True, $noQtyRefid="", $singleItemQtyCheck=True, $addressName="", $bodyTxt=""
	Local $CurrentOrderArray[15][34]
	Local $itemStringArray[15][2]
	Local $upperCaseId="", $skuType="", $itemCount=0, $skipAddTracking=False
				
	;_ArrayDisplay($OpenOrdersArray)
	;MsgBox(0, "", UBound($OpenOrdersArray, 2))


	For $i=2 To UBound($OpenOrdersArray)-1
		

		
		;$orderid = $OpenOrdersArray[$i][1]
		;$refid = $OpenOrdersArray[$i][6]
		;$qty = $OpenOrdersArray[$i][7]
		;$ShipToName = $OpenOrdersArray[$i][25]
		;$ShipToCompany = $OpenOrdersArray[$i][26]
		;$ShipToStreet1 = $OpenOrdersArray[$i][27]
		;$ShipToStreet2 = $OpenOrdersArray[$i][28]
		;$ShipToCity = $OpenOrdersArray[$i][29]
		;$ShipToState = $OpenOrdersArray[$i][30]	
		;$ShipToZipCode = $OpenOrdersArray[$i][31]		
		;$ShippingMethodId = $OpenOrdersArray[$i][32]
		

		If $OpenOrdersArray[$i][0] <> "" And $OpenOrdersArray[$i][0] <> "done" And $OpenOrdersArray[$i][0] <> " " Then
			
			If $OpenOrdersArray[$i][0]  = "Same_Person_Diff_Order" Then
				
				;_Log($timestamp & @TAB & "incomplete - Order ID: " & $a_[$row][1] &" Multiple Order with Same Person And Address" &@CR)
				$body &= "Buy.com Order ID: "&$OpenOrdersArray[$i][1]&" Multiple Orders with Same Person And Address" & @CR
			
			ElseIf define_skip($OpenOrdersArray[$i][1]) = True Then
		
				$body &= "Buy.com Order ID: "&$OpenOrdersArray[$i][1]& " Skipped" & @CR
					
			Else
				$itemCount=0

				For $a=0 To UBound($OpenOrdersArray,2)-1
					$CurrentOrderArray[$itemCount][$a] = $OpenOrdersArray[$i][$a]
				Next
				
				$itemCount+=1
					
				For $j=$i+1 To UBound($OpenOrdersArray)-1
					If $OpenOrdersArray[$i][1] = $OpenOrdersArray[$j][1] Then
							
						$OpenOrdersArray[$j][0] = "done"
							
						For $a=0 To UBound($OpenOrdersArray,2)-1
							$CurrentOrderArray[$itemCount][$a] = $OpenOrdersArray[$j][$a]
						Next
						$itemCount+=1
					EndIf
				Next
					
					;_ArrayDisplay($CurrentOrderArray)
					;_ArrayDisplay($OpenOrdersArray)
					

				$totalWeight=0
				$totalQty=0
				$refidString=""	
				$ItemStringCount=0
				$totalItemWeight=0
				$address=""
				$inStock=True
				$noQtyRefid=""
				$addressName=""
				$bodyTxt=""
				$upperCaseId=""
				$skuType=""
				
				;_ArrayDisplay($CurrentOrderArray)
				;MsgBox(0, "Ubound($CurrentOrderArray)", Ubound($CurrentOrderArray))
				
				
				For $m=0 To Ubound($CurrentOrderArray)-1
					
					If $CurrentOrderArray[$m][0] <> "" Then
					
						$singleItemWeight=0
						$singleItemString=""
						$singleItemRefid=""
						$singleItemQty=0
						$singleItemQtyCheck=True
						
						;_ArrayDisplay($CurrentOrderArray)
						
						If $CurrentOrderArray[$m][30] = "CA" Then  ;;; If is within CA, First Class Shipment will be used for Expedited Shipping
							$CurrentOrderArray[$m][32] = 1
							$OpenOrdersArray[$i][32] = 1
						EndIf
						
						;MsgBox(0,"$CurrentOrderArray[$m][30]",$CurrentOrderArray[$m][30])
						
						;_ArrayDisplay($CurrentOrderArray)
						
						;_ArrayDisplay($CurrentOrderArray)
						;MsgBox(0,"$CurrentOrderArray[$m][6]",$CurrentOrderArray[$m][6])
						
						$singleItemRefid = define_single_item_ref_id($CurrentOrderArray[$m][6])
						
						;MsgBox(0,"$singleItemRefid",$singleItemRefid)
						
						If $singleItemRefid = "" Then
							;_Log(@CR&$timestamp & @TAB & "incomplete - Order ID: " & $a_[$row][1] &" Qty: " &$a_[$row][7]&" Sku: "&$a_[$row][6]& " No Matching SKU" & @CR)
							;$body &= "Order ID: "&$a_[$row][1]& " Qty: " &$a_[$row][7]&" Sku: "&$a_[$row][6]& " No Matching SKU" & @CR
							$body &= "Buy.com Order ID: "& $CurrentOrderArray[$m][1] & " SKU: "&$CurrentOrderArray[$m][6]&" doesn't exist in the Inventory Table" & @CR								
							;MsgBox(0, "$singleItemRefid Error", '$singleItemRefid is empty')							
							
						Else
							
							$CurrentOrderArray[$m][6] = $singleItemRefid

							If $refidString <> "" Then
								$refidString &=", "
							EndIf
							
								$upperCaseId = StringUpper ($singleItemRefid)
								$skuType = StringLeft($upperCaseId, 1)
								
								Select
									Case $skuType = "T"
										$refidString &= "HOME-"
								
									Case $skuType = "P"
										$refidString &= "CAR-"

									Case $skuType = "H"
										$refidString &= "POUCH-"
								
									Case $skuType = "S"
										$refidString &= "V3CASE-"

									Case $skuType = "D"
										$refidString &= "DATACABLE-"
										
									Case Else
										$refidString &= ""

								EndSelect
							
							$refidString &= $singleItemRefid		
							
							$singleItemWeight = define_ref_id_weight($singleItemRefid)
							$singleItemString = buy_com_ref_id_packaging_type($singleItemRefid)
							
						
							If $singleItemWeight = 0 Or $singleItemString = "" Then
								;_Log(@CR&$timestamp & @TAB & "incomplete - Order ID: " & $a_[$row][1] &" Qty: " &$a_[$row][7]&" Sku: "&$a_[$row][6]& " No Matching SKU" & @CR)
								If $singleItemWeight = 0 Then 
									$body &= "Reference ID: "&$singleItemRefid& " No Matching Weight in Inventory Table" & @CR
								EndIf
								
								If $singleItemString = "" Then
									$body &= "Reference ID: "&$singleItemRefid& " No Matching Package Type in Inventory Table " & @CR						
								EndIf
								
								;MsgBox(0, "$singleItemWeight = 0 Or $item_string = "" Error", '$singleItemWeight = 0 Or $item_string = ""')
								
							Else
								$singleItemQty = $CurrentOrderArray[$m][7]
								$totalQty+=$CurrentOrderArray[$m][7]
								$totalItemWeight+=$singleItemWeight*$singleItemQty		
								$singleItemQtyCheck = inventoryQtyCheck($singleItemRefid, $singleItemQty)
								
								If $singleItemQtyCheck = False Then
									$inStock = False
									$noQtyRefid = $singleItemRefid
								EndIf	
								
								
								;MsgBox(0, "$singleItemString and $singleItemQty", $singleItemString&" "&$singleItemQty)
								
								;_ArrayDisplay($itemStringArray)
								
								
								If $itemStringArray[0][0] = "" Then
									$itemStringArray[0][0] = $singleItemString
									$itemStringArray[0][1] = $singleItemQty
									$ItemStringCount+=1
								Else
									$found = False
									
									For $n=0 To UBound($itemStringArray)-1
										If $itemStringArray[$n][0] <> "" Then
											If $itemStringArray[$n][0] = $singleItemString Then
												$itemStringArray[$n][1] += $singleItemQty
												$found = True
												ExitLoop
											EndIf
										EndIf
									Next
									
									If $found = False Then
										$itemStringArray[$ItemStringCount][0] = $singleItemString
										$itemStringArray[$ItemStringCount][1] = $singleItemQty
										$ItemStringCount+=1
									EndIf	
									
								
								EndIf	;; End of $itemStringArray[0][0] = ""

							EndIf ;; End of $weightTotal = 0 Or $item_string = "" 
					
						EndIf ;; End of if refid is empty
				
					EndIf ;; End of if the cell is empty
			
				Next
								
				;_ArrayDisplay($itemStringArray)
				
					$itemString=""
					;;package String
						_ArraySort($itemStringArray)
						For $m=0 To UBound($itemStringArray)-1
							If $itemStringArray[$m][0] <> "" And $itemStringArray[$m][0] <> " " And $itemStringArray[$m][0] <> 1 Then
								$itemString &= $itemStringArray[$m][0] & $itemStringArray[$m][1]							
							EndIf						
						Next
						
						;MsgBox(0, "$itemString", $itemString)
						   
						   
				$packageType=""
				$packageWeight=0
				
				;; define packaging and define packaging weight
				$packageType = define_packaging($itemString) 
				$packageWeight = define_packaging_weight($itemString)		

					;MsgBox(0, "$packageType $packagingWeight", $packageType&" "&)
					
					
					If $packageWeight = 0 Or $packageWeight = "" Or $packageType="" Or $inStock=False Then
						
						If $packageWeight=0 Or $packageType="" Then
							$body &= "Item String: "&$itemString& " is missing weight or type in the Packaging Table"& @CR	
						EndIf
						
						If $inStock = False Then
							$body &= "Item(s): "& $noQtyRefid & " do not have enough qty to fulfill the order(s)"& @CR
						EndIf
						
						;$body &= "Order ID: "&$a_[$row][1]& " Qty: " &$a_[$row][7]&" Sku: "&$a_[$row][6]& " No Matching SKU" & @CR						
						;MsgBox(0, "$packageWeight = 0 Or $packageWeight = "" Or $packageType="" Error", '$packageWeight = 0 Or $packageWeight = "" Or $packageType=""')
					
					Else
						
						$weightTotal = $totalItemWeight + $packageWeight
					
					
					
					;MsgBox(0, "$weightTotal", $weightTotal)
					
					
					;;address

					If addressChangeExist($OpenOrdersArray[$i][1]) = True Then
						$address = addressChange($OpenOrdersArray[$i][1])
						$addressName = addressChangeName($OpenOrdersArray[$i][1])
					
					Else
						$addressName = StringReplace ($OpenOrdersArray[$i][25], @CR, "")
						$addressName = StringReplace ($addressName, @LF, "")
						
						$addressNameSplit = StringSplit($addressName, " ",2)
						
						If UBound($addressNameSplit) > 2 then
							$addressName=""
							For $k=0 To UBound($addressNameSplit)-1
								If $addressNameSplit[$k] <> "" Then
									If $addressName <> "" Then
										$addressName&=" "&$addressNameSplit[$k]
									Else
										$addressName&=$addressNameSplit[$k]
									EndIf
		
								EndIf
							Next
							
						EndIf
						
						$address = $OpenOrdersArray[$i][25]&@CR
						If $OpenOrdersArray[$i][26] <> "" Then
							$address &= $OpenOrdersArray[$i][26]&@CR
						EndIf
						$address &= $OpenOrdersArray[$i][27]&@CR
						If $OpenOrdersArray[$i][28] <> "" Then
							$address &= $OpenOrdersArray[$i][28]&@CR
						EndIf
						$address &= $OpenOrdersArray[$i][29]&", "
						$address &= $OpenOrdersArray[$i][30]&" "
						$address &= $OpenOrdersArray[$i][31]
						;$ShipToName = $OpenOrdersArray[$i][25]
						;$ShipToCompany = $OpenOrdersArray[$i][26]
						;$ShipToStreet1 = $OpenOrdersArray[$i][27]
						;$ShipToStreet2 = $OpenOrdersArray[$i][28]
						;$ShipToCity = $OpenOrdersArray[$i][29]
						;$ShipToState = $OpenOrdersArray[$i][30]	
						;$ShipToZipCode = $OpenOrdersArray[$i][31]		
						;$ShippingMethodId = $OpenOrdersArray[$i][32]		
									
						;MsgBox(0, "$address", $address)
						;MsgBox(0, "$weightTotal", $weightTotal)
					EndIf
					
						If $weightTotal < 13 Or $weightTotal = 13 Then
									
							Select
								Case $packageType = "BUBBLE_MAILER"
		
									If $OpenOrdersArray[$i][32] = "1" Then
										printFirstClass($address, $weightTotal, $OpenOrdersArray[$i][1], "EMD Store", "BUBBLE_MAILER")
												
									ElseIf $OpenOrdersArray[$i][32] = "2" Then
										printSmallFlatRateBox($address, $weightTotal, $OpenOrdersArray[$i][1], "EMD Store")
								
									Else 
										MsgBox(0, "$packageType Error Less Than <13", "No Matching Package Found") 
								
									EndIf
											
								
								Case $packageType = "SMALL_FLAT_RATE_BOX"
										printSmallFlatRateBox($address, $weightTotal, $OpenOrdersArray[$i][1], "EMD Store")
											
								Case $packageType = "MEDIUM_FLAT_RATE_BOX"
										printMediumFlatRateBox($address, $weightTotal, $OpenOrdersArray[$i][1], "EMD Store")

								Case $packageType = "LARGE_FLAT_RATE_BOX"
											_Print_Large_Flat_Rate_Box($a_, $row)
										
								Case $packageType = "BOX444" 
									If $OpenOrdersArray[$i][32] = "1" Then
										printFirstClass($address, $weightTotal, $OpenOrdersArray[$i][1], "EMD Store", "BUBBLE_MAILER")
												
									Else
										printSmallFlatRateBox($address, $weightTotal, $OpenOrdersArray[$i][1], "EMD Store")
				
									EndIf									
										
								Case $packageType = "BOX644" 
									If $OpenOrdersArray[$i][32] = "1" Then
										printFirstClass($address, $weightTotal, $OpenOrdersArray[$i][1], "EMD Store", "BUBBLE_MAILER")
												
									Else
										MsgBox(0, "$packageType = BOX644", "No Printing Class Found") 
				
									EndIf
										
								Case $packageType = "BOX664"
									If $OpenOrdersArray[$i][32] = "1" Then
										printFirstClass($address, $weightTotal, $OpenOrdersArray[$i][1], "EMD Store", "BUBBLE_MAILER")
												
									Else
										MsgBox(0, "$packageType = BOX644", "No Printing Class Found") 
				
									EndIf


								EndSelect
										
							Else
									
								Select
									Case $packageType = "BUBBLE_MAILER" 
										printPriority($address, $weightTotal, $OpenOrdersArray[$i][1], "EMD Store", "BUBBLE_MAILER")
											
									Case $packageType = "BOX444" 
										printPriority($address, $weightTotal, $OpenOrdersArray[$i][1], "EMD Store", "BOX444")
										
									Case $packageType = "BOX644"
										printPriority($address, $weightTotal, $OpenOrdersArray[$i][1], "EMD Store", "BOX644")
										
									Case $packageType = "BOX664"		
										printPriority($address, $weightTotal, $OpenOrdersArray[$i][1], "EMD Store", "BOX664")
										
									Case $packageType = "SMALL_FLAT_RATE_BOX"
										printSmallFlatRateBox($address, $weightTotal, $OpenOrdersArray[$i][1], "EMD Store")
											
									Case $packageType = "MEDIUM_FLAT_RATE_BOX"
										printMediumFlatRateBox($address, $weightTotal, $OpenOrdersArray[$i][1], "EMD Store")

									Case $packageType = "LARGE_FLAT_RATE_BOX"
										_Print_Large_Flat_Rate_Box($a_, $row)
		
								EndSelect
			
							EndIf
		
					;; Add Tracking Number
							
							ControlClick("Stamps.com Pro", "", 32551)
							Sleep(2000)
							
							$bodyTxt = ""
							$skipAddTracking = False
	
								WinWaitActive("Stamps.com Pro")
								Sleep(2000)
								
								$trackingIE = _IEAttach("Stamps.com Pro", "Embedded")
								_IELoadWait($trackingIE)
								
								;MsgBox(0,"$addressName", $addressName)
							
							For $k = 0 to 15 
								 $bodyTxt = _IEBodyReadText($trackingIE)
								
								;MsgBox(0,"StringInStr ($bodyTxt, $addressName)", StringInStr ($bodyTxt, $addressName))
								If StringInStr ($bodyTxt, $addressName) <> 0 Then
									$trackingTable = _IETableGetCollection($trackingIE, 1)
									$trackingData = _IETableWriteToArray($trackingTable, True)
									;MsgBox(0,StringInStr ($trackingData[1][4], $addressName), $trackingData[1][4]&"addressName="&$addressName)
									If StringInStr($trackingData[1][4], $addressName) <> 0 Then
										$temp_tracking = $trackingData[1][6]
										ExitLoop
									EndIf
									
									ControlClick("Stamps.com Pro", "", "[CLASS:ToolbarWindow32; INSTANCE:1]", "Primary", 1, 124,15 ) 
									_IELoadWait($trackingIE)
									Sleep(10000)
								EndIf
								
								If $k > 7 Then
									ControlClick("Stamps.com Pro", "", "[CLASS:ToolbarWindow32; INSTANCE:1]", "Primary", 1, 124,15 ) 
									_IELoadWait($trackingIE)									
									Sleep(60000)
									
								EndIf

							Next

							If $k > 14 Then
								;MsgBox(0, "Error", "Cannot find address name: "&$addressName)
								$body &= "Cannot find address name "& $addressName & "Buy.com"
								$skipAddTracking = True
							EndIf


							Sleep(2000)
							ControlClick("Stamps.com Pro", "", 32513)
							Sleep(2000)							
							
							If $skipAddTracking <> True Then
								addBuycomTracking($CurrentOrderArray, $temp_tracking)
								Sleep(500)
							EndIf 
							
							printBuycomPackingSlip($CurrentOrderArray, $packageType, $refidString)
							Sleep(500)							
						
							;;Update Inventory
							inventoryUpdate($CurrentOrderArray, 6, 7)

							For $m=0 To Ubound($addressNameSplit)-1
								$addressNameSplit[$m]=""

							Next							
							
					EndIf ;; end of $packageWeight = 0 Or $packageWeight = "" Or $packageType="" Or $inStock=False
				
			EndIf ;; end of if not same person with different order ID

		EndIf  ;; end of if array not empty or done

			;_ArrayDisplay($itemStringArray)
			;_ArrayDisplay($CurrentOrderArray)
							

			;;Clean Arrays
			For $m=0 To UBound($itemStringArray)-1
				$itemStringArray[$m][0]=""
				$itemStringArray[$m][1]=""
			Next
							
							
			For $m=0 To $itemCount Step 1
				For $n=0 To UBound($CurrentOrderArray, 2)-1 
					$CurrentOrderArray[$m][$n]=""
					;$CurrentOrderArray[$m][0]=""
				Next
				
			Next
			
			

			
			
							
			;_ArrayDisplay($itemStringArray)
			;_ArrayDisplay($CurrentOrderArray)



	Next

EndFunc

Func retrieveTrackingID($name)
	
	Local $bodyTxt="", $trackingID=""
	Local $addressSplit="", $address="", $domestic=False, $company="", $transactionid="", $transSplit=""
	
	
	WinActivate("Stamps.com Pro")
	ControlClick("Stamps.com Pro", "", 32551)
		
	Sleep(3000)
							
	$bodyTxt = ""

	WinWaitActive("Stamps.com Pro")
	Sleep(2000)
	$trackingIE = _IEAttach("Stamps.com Pro", "Embedded")
	_IELoadWait($trackingIE)
							
				
	For $k = 0 to 20
		 $bodyTxt = _IEBodyReadText($trackingIE)
						
		If StringInStr ($bodyTxt, $name) <> 0 Then
			$trackingTable = _IETableGetCollection($trackingIE, 1)
			$trackingData = _IETableWriteToArray($trackingTable, True)
				If StringInStr($trackingData[1][4], $name) <> 0 Then
					$trackingID = $trackingData[1][6]
					ExitLoop
				EndIf
									
			ControlClick("Stamps.com Pro", "", "[CLASS:ToolbarWindow32; INSTANCE:1]", "Primary", 1, 124,15 ) 
			_IELoadWait($trackingIE)
			Sleep(15000)
		EndIf
								
		If $k > 10 Then
			ControlClick("Stamps.com Pro", "", "[CLASS:ToolbarWindow32; INSTANCE:1]", "Primary", 1, 124,15 ) 
			_IELoadWait($trackingIE)					
			Sleep(100000)
		EndIf

	Next
						
	If $k > 19 Then
		$body &= "Address Name " & $addressName & "Cannot be found on Stamp.com table; no tracking info entered for Paypal "& $acct & @CR
	EndIf
						
	Sleep(2000)
	ControlClick("Stamps.com Pro", "", 32513)
	Sleep(2000)
	
	Return $trackingID
	
EndFunc

Func paypalMultipleOrdersAddTrackingPrintPacking($oIE, $name, $trackingID, $multipleOrders, $itemArray, $packaging, $domestic=True)
	

	
	WinActivate("PayPal Website Payment Details")
	WinWaitActive("PayPal Website Payment Details")

	If $domestic = False Then
		$trackingID &= "None-tracking, custom declaration# "
	EndIf

		_IELinkClickByText($oIE, "Add Tracking Info")								
		_IELoadWait($oIE)

		Local $oInput = _IEGetObjByName($oIE, "shipping_status")
		_IEFormElementOptionselect($oInput, "S" , 1, "byValue")
									
		$oInput = _IEGetObjByName($oIE, "track_num")
		_IEFormElementSetValue($oInput, $trackingID)
									
		$oInput = _IEGetObjByName($oIE, "shipping_co_id")
		_IEFormElementOptionselect($oInput, "1" , 1, "byValue")
									
		MsgBox(0, "okay", "okay?")
									
		$oInput = _IEGetObjByName($oIE, "Save")
		_IEAction($oInput, "click")
									
		_IELoadWait($oIE)

		_IELinkClickByText($oIE, "Print Packing Slip")
						
						
		WinWaitActive("Packing Slip - PayPal - Windows Internet Explorer")
		sleep(2000)
		
		printPaypalPackingSlip($itemArray ,$packaging)
							

		$oInput = _IEGetObjByName($oIE, "cancel.x")
		_IEAction($oInput, "click")							
		_IELoadWait($oIE)
		
		
		_IELinkClickByText($oIE, "History")		
		_IELoadWait($oIE)


	For $numClick=2 To $multipleOrders 
		
		$oLinks = _IELinkGetCollection($oIE)
		$clickCount=1
		For $oLink In $oLinks	
						
			If StringInStr($oLink.innerText, $name) <> 0 And StringInStr($oLink.innerText, "Details") <> 0  Then
				;MsgBox(0,$clickCount&" - "&$numClick,$oLink.innerText)
				If $numClick = $clickCount Then
					;MsgBox(0,"$clickCount",$clickCount)
					_IEAction($oLink, "click")
					ExitLoop
				EndIf
				
				$clickCount+=1
				
			EndIf
		Next	
			

			_IELinkClickByText($oIE, "Add Tracking Info")								
			_IELoadWait($oIE)

			Local $oInput = _IEGetObjByName($oIE, "shipping_status")
			_IEFormElementOptionselect($oInput, "S" , 1, "byValue")
										
			$oInput = _IEGetObjByName($oIE, "track_num")
			_IEFormElementSetValue($oInput, $trackingID)
										
			$oInput = _IEGetObjByName($oIE, "shipping_co_id")
			_IEFormElementOptionselect($oInput, "1" , 1, "byValue")
										
			MsgBox(0, "okay", "okay?")
										
			$oInput = _IEGetObjByName($oIE, "Save")
			_IEAction($oInput, "click")
										
			_IELoadWait($oIE)


		_IELinkClickByText($oIE, "History")		
		_IELoadWait($oIE)		
		
		
	Next


EndFunc

Func paypayOrdersProceed($oIE, $name, $orderNumber, $acct, $flagnote="")

	Local $array, $itemArray[10][8] ;; itemNum_itemName_qty_price
	Local $itemCount=0, $trackingID=""

	For $numClick=1 To $orderNumber 
		;MsgBox(0,$orderNumber,$numClick)
		
		$oLinks = _IELinkGetCollection($oIE)
		$clickCount=1
		For $oLink In $oLinks	
			
			
			If StringInStr($oLink.innerText, $name) <> 0 And StringInStr($oLink.innerText, "Details") <> 0  Then
				;MsgBox(0,$clickCount&" - "&$numClick,$oLink.innerText)
				If $numClick = $clickCount Then
					;MsgBox(0,"$clickCount",$clickCount)
					_IEAction($oLink, "click")
					ExitLoop
				EndIf
				
				$clickCount+=1
				
			EndIf
		Next
	
		_IELoadWait($oIE)

		$transTable = _IETableGetCollection ($oIE, 3)
		$transTableData = _IETableWriteToArray ($transTable, True)
							
		;_ArrayDisplay($transTableData)
		$transSplit = StringSplit($transTableData[1][0], "#", 2)
		;_ArrayDisplay($transSplit)
		$transactionid = StringStripWS($transSplit[1], 8)
		$transactionid = StringTrimRight($transactionid, 1)
	

		If define_skip($transactionid) = True Then
			$body &= "Paypal "&$acct&" transaction id: "&$transactionid&" skipped"&@CR

		Else

		$htmlbody = _IEBodyReadText($oIE)
		
			If StringInStr($htmlbody, "iOffer") Then
				$body &= "iOffer Transaction Exist" & @CR
				;_Log($timestamp &@TAB& "PayPal " & $acct & " - iOffer Transaction Exist"& @CR)
				$back = _IEGetObjByName($oIE, "cancel.x")
				_IEAction($back, "click")
				;_IEWaitForTitle($oIE, "Account overview - PayPal - Windows Internet Explorer")
				_IELinkClickByText($oIE, "History")		
				;_IEWaitForTitle($oIE, "History - PayPal - Windows Internet Explorer")
					
			
			ElseIf StringInStr($htmlbody, "Atomic Mall Order") Then
				$body &= "Atomic Mall Order Exist" & @CR
				;_Log($timestamp &@TAB& "PayPal " & $acct & " - Atomic Mall Order Exist"& @CR)
				$back = _IEGetObjByName($oIE, "cancel.x")
				_IEAction($back, "click")
				;_IEWaitForTitle($oIE, "Account overview - PayPal - Windows Internet Explorer")
				_IELinkClickByText($oIE, "History")		
				;_IEWaitForTitle($oIE, "History - PayPal - Windows Internet Explorer")			

			Else
				

				$itemTable = _IETableGetCollection ($oIE, 5)
				$itemTableData = _IETableWriteToArray ($itemTable, True)
				
				;_ArrayDisplay($itemTableData)
				
				
				$itemNameSplit=""
				;$isNumber=0
					For $i=1 to UBound($itemTableData)-1
								
						;MsgBox(0,$i, $itemTableData[$i][1])
						
						If StringInStr($itemTableData[$i][0], "Amount") = 0 And StringInStr($itemTableData[$i][1], "Coupon Discount") = 0 And StringInStr($itemTableData[$i][1], "Coupon") = 0 Then
							$itemArray[$itemCount][2] = $itemTableData[$i][0] ;; qty
							$itemArray[$itemCount][5] = $itemTableData[$i][3] ;; price
							$itemNameSplit = StringSplit($itemTableData[$i][1], "#", 2)
							;MsgBox(0,"StringInStr($itemTableData[$i][1])", StringInStr($itemTableData[$i][1], "Coupon Discount"))
							;_ArrayDisplay($itemNameSplit)
							
							$itemArray[$itemCount][0] = StringStripWS($itemNameSplit[1], 8) ;; item number
							$itemArray[$itemCount][1] = StringTrimRight($itemNameSplit[0], 5) ;; item name
							$itemCount+=1
						EndIf
						;_ArrayDisplay($itemArray)
					Next
		
			;_ArrayDisplay($itemArray)
		
			_IELinkClickByText($oIE, "History")		
			_IELoadWait($oIE)			
			
			EndIf ;; End of iOffer if and Atomic Mall if
		
		EndIf ;; End of define Skip if
		
	
	Next		
;;;;;;;;;;;; end building item table ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;



			$found = True
				
			;; define existence 
					
			For $i=0 To UBound($itemArray)-1
						
				If $itemArray[$i][0] <> "" Then
							
					If db_item_title_exist($itemArray[$i][1]) = True Then
						If db_item_check_stock($itemArray[$i][1]) <> True Then
							$found=False
						EndIf
					Else
						$found=False
						add_to_DB($itemArray[$i][1], "name", "ebay", $itemArray[$i][0], "item_number")
					
					EndIf
				EndIf
			Next
					
					
			;;define per item weight
			$proceed=True
			$packagingType = ""
			$itemWeight = 0
			$itemSKU = ""
			$itemType = ""
			$arraySize=0
			$emptyFound_packagingType=False
			$emptyFound_itemWeight=False
			$emptyFound_itemType=False
			$emptyFound_itemSKU=False

			If $flagnote <> "" Then
				$approval = flagged($flagnote, $transactionid, $name)
						;MsgBox(0,"approval",$approval)
				
				If $approval <> True Then
					$proceed = False
				EndIf
			EndIf
				  
				  ;; $itemArray[0] - item number
				  ;; $itemArray[1] - item name
				  ;; $itemArray[2] - qty
				  ;; $itemArray[3] - weight
				  ;; $itemArray[4] - package type
				  ;; $itemArray[5] - price
				  ;; $itemArray[6] - item type
				  ;; $itemArray[7] - item sku				  
				  
			 For $k=0 To UBound($itemArray)-1
				If $itemArray[$k][0] <> "" Then
					$packagingType = retrievePackagingType($itemArray[$k][1], "ebay")  
					;MsgBox(0,"$packagingType",$packagingType)
					If $packagingType = "" Then
					   $proceed = False
					   $emptyFound_packagingType=True
					Else
					   $itemArray[$k][4] = $packagingType
					EndIf
						
					$itemWeight = retrieveWeight($itemArray[$k][1], "ebay")
					If $itemWeight = 0 Then
					   $proceed = False
					   $emptyFound_itemWeight= True
							  
					Else
					   $itemArray[$k][3] = $itemWeight			   
					EndIf
							
					$itemType = retrieveItemType($itemArray[$k][1], "ebay")
					If $itemType = "" Then
					   $proceed = False
					   $emptyFound_itemType=True
							  
					Else					   
					   $itemArray[$k][6] = $itemType
					EndIf   

					$itemSKU = retrieveSKU($itemArray[$k][1], "ebay")
					If $itemSKU = "" Then
					   $proceed = False
					   $emptyFound_itemSKU=True
					Else					   
					   $itemArray[$k][7] = $itemSKU

					EndIf
													
				If StringInStr($itemArray[$k][1], "Purse") = 0 Or StringInStr ($itemArray[$k][1], "purse") = 0 Then 
					If  $emptyFound_packagingType=True Then							
						$body &= $itemArray[$k][1] & "itemWeight is empty"
					EndIf
					If $emptyFound_itemWeight= True Then
						$body &= $itemArray[$k][1] & "packagingType is empty"
					EndIf
					If $emptyFound_itemType=True Then
						$body &= $itemArray[$k][1] & "itemSKU is empty"
					EndIf
					If $emptyFound_itemSKU=True Then
						$body &= $itemArray[$k][1] & "itemType is empty"
						EndIf
					EndIf
						$arraySize += 1
							

				EndIf
						;MsgBox(0,"$k",$k)
			 Next

				  
				  $itemString=""
				  
				  
					 If $arraySize > 1 Then
						 ;MsgBox(0, "$arraySize > 1","true")
						Local $packageArray[$arraySize][2]
						$count=0

						For $m=0 To UBound($itemArray)-1
							If $itemArray[$m][0] <> "" Then
							   If $packageArray[0][0] = "" Then
								  $packageArray[0][0] = $itemArray[$m][4]
								  $packageArray[0][1] = $itemArray[$m][2]
								  $count+=1
								  
							   Else
								  $found = False
								  For $n=0 To UBound($packageArray)-1
									 
									 If $packageArray[$n][0] = $itemArray[$m][4] Then
										$packageArray[$n][1]+=1
										$found = True
									 EndIf
								  
								  Next
								  
								  If $found <> True Then
									 $packageArray[$count][0] = $itemArray[$m][4]	
									 $packageArray[$count][1] = $itemArray[$m][2]
									 $count+=1
								  EndIf
							   
								EndIf
							EndIf
						Next
						
						;_ArrayDisplay($itemArray)
						_ArraySort($packageArray)
						;_ArrayDisplay($packageArray)
						
						For $k=0 To Ubound($packageArray)-1
						   If $packageArray[$k][0] <> "" Then
							  $itemString &= $packageArray[$k][0]
							  $itemString &= $packageArray[$k][1]
						   EndIf
						Next
					
						;_ArrayDisplay($packageArray)
						
						For $i=0 To UBound($packageArray)-1
							For $j=0 To UBound($packageArray,2)-1
								$packageArray[$i][$j] = ""
							Next
						Next
					
					Else
					;MsgBox(0, "$arraySize > 1","else")
						If $itemArray[0][0] <> "" Then
							;_ArrayDisplay($itemArray)
							$itemString = $itemArray[0][4]
							$itemString &= $itemArray[0][2]
						EndIf	
					EndIf

					
				
				  ;MsgBox(0,"$itemString",$itemString)
				  
				  ;_ArrayDisplay($packageArray)
				  ;;package string existence
				  $packaging = ""
				  $packaging = define_packaging($itemString) 

				  ;;package string weight
				  $packagingWeight = 0
				  $packagingWeight = define_packaging_weight($itemString)
				  
				  If $packaging="" or $packagingWeight=0 Then
					 $proceed = False
					 $body &= "Item String: "&$itemString& " Packaging Missing" & @CR
				  EndIf
				  
				  ;; total weight item weight * qty + packaging weight
					$weightTotal=0

				  
				  
				  If $proceed = True Then

					 If $arraySize > 1 Then
						For $k=0 To UBound($itemArray)-1
							If $itemArray[$k][0] <> "" Then
								$weightTotal+= $itemArray[$k][3] * $itemArray[$k][2]
							EndIf
						Next
						
					 Else
						If $itemArray[0][0] <> "" Then
							$weightTotal = $itemArray[0][3] * $itemArray[0][2]
						EndIf
					 EndIf
					 
					 $weightTotal+= $packagingWeight

				EndIf
		
		
					$start=0
					
				If $proceed = True And $weightTotal <> 0 Then

				$oLinks = _IELinkGetCollection($oIE)
				For $oLink In $oLinks	

					If StringInStr($oLink.innerText, $name) <> 0 And StringInStr($oLink.innerText, "Details") <> 0  Then
						_IEAction($oLink, "click")
						ExitLoop
					EndIf
				Next

				;_ArrayDisplay($itemArray)

				_IELoadWait($oIE)
				Sleep(1500)

					$addressName=""
					$address=""
					
					If addressChangeExist($transactionid) = True Then
						$address = addressChange($transactionid)
						$addressName = addressChangeName($transactionid)
					
					Else
					
						$addressTable = _IETableGetCollection ($oIE, 2)
						$addressTableData = _IETableWriteToArray ($addressTable, True)

						;_ArrayDisplay($addressTableData)

						$addressSplit = $addressTableData[1][1]
									
						;MsgBox(0,"test",$data)
						$addressSplit = StringSplit ($addressSplit, @LF)
						;_ArrayDisplay($addressSplit)
								
									
						For $i = 1 To UBound($addressSplit)-1
						;MsgBox(0,"$data[$i] = $open_order_names[$w]", $data[$i] &" - "& $open_order_names[$w])
							If StringRegExp($addressSplit[$i], "Protection address") = 1 OR StringRegExp($addressSplit[$i], "Ship to address") = 1 OR StringRegExp($addressSplit[$i], "Tips to sell securely") = 1 Then
							;MsgBox(0, "test", StringRegExp($data[$i], "Tips to sell securely"))
										
								$start = $i+1
									
							EndIf
						Next
						
						;_ArrayDisplay($addressSplit)
						$domestic = False
						
						For $i=$start To UBound($addressSplit)-1
							
							If StringInStr($addressSplit[$i], "United States") <> 0 Then
								$domestic = True
							ElseIf StringInStr($addressSplit[$i], "Confirmed") = 0 Then
								$address &= StringReplace ($addressSplit[$i], @CR, "")&@CR		
							EndIf

							If $i=$start Then
								$addressName = StringReplace ($addressSplit[$i], @CR, "")
							EndIf


						Next

						;MsgBox(0,"address",$address)

						$addressNameSplit = StringSplit($addressName, " ",2)

							If UBound($addressNameSplit) > 2 then
								$addressName=""
								;MsgBox(0,"test",UBound($addressNameSplit)-1)
								For $k=0 To UBound($addressNameSplit)-1
									
									If $addressNameSplit[$k] <> "" Then

										If $addressName <> "" Then
											$addressName&=" "&$addressNameSplit[$k]
										Else
											$addressName&=$addressNameSplit[$k]
										EndIf
			
									EndIf
								Next
								
							EndIf
					EndIf
						;MsgBox(0,$addressName,$address)
						;MsgBox(0,"$weightTotal",$weightTotal)
						;_ArrayDisplay($itemArray)

							If $acct = 1 Then
								$company = "EMD Store"
							ElseIf $acct = 2 Then
								$company = "EMD Merchant"
							ElseIf $acct = 3 Then
								$company = "EMDCELL"
							Else
								$body &= "unknown paypal account number $acct= " & $acct&@CR
								$noMatch = True
							EndIf

						
						If $domestic = True Then
							

							
							If $company <> "" Then
								If $weightTotal < 13 Or $weightTotal = 13 Then									
									printFirstClass($address, $weightTotal, $transactionid, $company, $packaging)
								ElseIf $weightTotal > 13
									
									printPriority($address, $weightTotal, $transactionid, $company, $packaging)
								Else
									$body &= "Paypal print shipping label total weight unknown"
									$noMatch = True
								EndIf
							
							
							
							EndIf
							;paypalPrintDomesticLabel()
						
						ElseIf StringInStr($address, "Puerto Rico") <> 0 Then
						
							;_Log($timestamp & @TAB & "Puerto Rico Order Exist" & @CR)	
							$body &= "Puerto Rico Order Exist"
							$noMatch = True
						Else
							;;skip
							;$noMatch = True
							
							;; international shipping
							If $company <> "" Then
								
								printInternationalFirstClassPayPal($address, $weightTotal, $transactionid, $company, $packaging, $itemArray)

							EndIf
						
						EndIf
							
						Sleep(1700)
						$trackingID=retrieveTrackingID($addressName)
						Sleep(1700)
						paypalMultipleOrdersAddTrackingPrintPacking($oIE, $addressName, $trackingID, $orderNumber, $itemArray, $packagingType, $domestic)
						Sleep(1700)
							
							
						inventoryUpdate($itemArray, 7, 2)
							
				
						

						
				EndIf ;; end of $proceed = True And $weightTotal <> 0			


EndFunc



Func paypalOrders($oIE, $acct)
   
   dim $array, $itemArray[10][8] ;; itemNum_itemName_qty_price
   Local $addressSplit="", $address="", $domestic=False, $company="", $transactionid="", $transSplit=""
   Local $emptyFound=False, $packagingType = "", $itemWeight = 0, $itemSKU = "", $itemType = ""
   Local $itemString="", $count=0, $arraySize=0, $packaging = "", $packagingWeight = 0, $weightTotal=0, $noMatch = False, $addressName=""
   Local $proceed = True, $approval=False, $finish = False, $totalQty=0, $skipTracking=False
   Local $sameName="", $flagnote=""
   _IENavigate($oIE, "https://www.paypal.com/us/cgi-bin/webscr?cmd=_account&nav=0")
	  
   ;;Get fast access to your PayPal cash
$finish = False

For $reload=0 To 11 
   $winTitle = winGetTitle("[active]")
	
	If $finish <> False Then
		ExitLoop

	ElseIf StringInStr( $wintitle, "Get fast access to your PayPal cash") <> 0 Then
		_IELinkClickByText($oIE, "Go to My Account")
	  
	Elseif StringinStr($wintitle, "Logging in - PayPal") <> 0 then	
	   Sleep(5000)
	

	ElseIf StringInStr( $wintitle, "My Account - PayPal") <> 0 Or StringinStr($wintitle, "Account overview - PayPal") <> 0 Or StringInStr($wintitle, "History - PayPal") <> 0 Then
	  _IELinkClickByText($oIE, "History")
	  _IELoadwait($oIE)
	  sleep(1000)

	  ;$rawOrderTable = _TableArray($oIE, 0)
	  	Local $rawTable = _IETableGetCollection ($oIE, 0)
		Local $rawOrderTable = _IETableWriteToArray ($rawTable, True)
		
		; _ArrayDisplay($rawOrderTable)

	  $itemCount=0
	For $i=0 To Ubound($rawOrderTable)-1
		;MsgBox(0,$i,$rawOrderTable[$i][8])
		If StringinStr($rawOrderTable[$itemCount][6], "Cleared") <> 0 Or $rawOrderTable[$itemCount][6] = "Completed" Then
			;If StringinStr($rawOrderTable[$itemCount][8], "Print shipping label") <> 0 And $rawOrderTable[$itemCount][6] = "Cleared" Then
			$itemCount+=1
		
		Else
			_ArrayDelete($rawOrderTable, $itemCount)
			
		EndIf
	Next
	
	;_ArrayDisplay($rawOrderTable)
	
	$itemCount=0

	For $i=0 To Ubound($rawOrderTable)-1
		;MsgBox(0,$i,$rawOrderTable[$i][8])
		If StringinStr($rawOrderTable[$itemCount][8], "Print shipping label") <> 0 Then
			;If StringinStr($rawOrderTable[$itemCount][8], "Print shipping label") <> 0 And $rawOrderTable[$itemCount][6] = "Cleared" Then
			$itemCount+=1
		
		Else
			_ArrayDelete($rawOrderTable, $itemCount)
			
		EndIf
	Next


	;_ArrayDisplay($rawOrderTable)
	
	$rawOrderTable =checkPaypalDup($rawOrderTable)
	
	;_ArrayDisplay($rawOrderTable)

;col 0 - select record 1
;col 1 - 0
;col 2 - Date
;col 3 - comment
;col 4 - payment from
;col 5 - name
;col 6 - completed
;col 7 - details payment from
;col 8 - choose action
;col 9 - price
;col 10 - fee
;col 11 - total	


	;if skip transcation id exist 
	
	;if comment exist  then check for approval
	
	;if dup exist
	
	;else execute the order
	
	
	
	For $w=0 to UBound($rawOrderTable)-1
		
		$sameOrderQty=0
		
		If ($rawOrderTable[$w][0] <> "") Or  ($rawOrderTable[$w][0] <> "done") Then
			If StringInStr ($rawOrderTable[$w][1], "samename") <> 0 Then
				
				$sameNameSplit=StringSplit($rawOrderTable[$w][1], "_")
				
				$sameName = $sameNameSplit[2]
				
				;MsgBox(0, "samename", $sameName)

				For $z=0 To UBound($rawOrderTable)-1 
					If StringInStr($rawOrderTable[$z][5], $sameName) <> 0   Then
						$sameOrderQty+=1
						$rawOrderTable[$z][0]="done"
					EndIf
					
					If $rawOrderTable[$z][3] <> "" And $rawOrderTable[$z][3] <> " " Then
						$flagnote &= $rawOrderTable[$z][3] & " | "
					EndIf
				Next	

				;MsgBox(0, "$sameOrderQty", $sameOrderQty)
				paypayOrdersProceed($oIE, $sameName, $sameOrderQty, $acct, $flagnote)
				
				

			Else
				$proceed = True
				;;;; Use the names in open order names array and find the detail button and click to the detail order page
				$oTR = _IETagNameGetCollection($rawTable, "TR", $w) ; reference to TR tag
				$oTD = _IETagNameGetCollection($oTR, "TD", 7) ; reference to TD tag
				$detail =  _IEPropertyGet($oTD, "innertext")

				$oLinks = _IELinkGetCollection($oIE)
				For $oLink In $oLinks			
					If StringInStr($oLink.innerText, $rawOrderTable[$w][5]) <> 0 and StringInStr($oLink.innerText, "Details") <> 0 Then
						;msgbox(0,"FOUND","innerText = " & $oLink.innerText)
						;MsgBox(0,"$open_order_names[$i]",$open_order_names[$w])
						_IEAction($oLink, "click")
						ExitLoop
					EndIf
				Next
			
			_IELoadWait($oIE)

				$transTable = _IETableGetCollection ($oIE, 3)
				$transTableData = _IETableWriteToArray ($transTable, True)
							
				;_ArrayDisplay($transTableData)
				$transSplit = StringSplit($transTableData[1][0], "#", 2)
				;_ArrayDisplay($transSplit)
				$transactionid = StringStripWS($transSplit[1], 8)
				$transactionid = StringTrimRight($transactionid, 1)

				If define_skip($transactionid) = True Then
					$body &= "Paypal "&$acct&" transaction id: "&$transactionid&" skipped"&@CR
					$proceed = False
					

				EndIf
		
		;MsgBox(0,"$proceed",$proceed)
		;MsgBox(0,"$transactionid",$transactionid)
		;MsgBox(0,"$rawOrderTable[$w][1]",$rawOrderTable[$w][1])
		
		If $proceed <> False Then

		$htmlbody = _IEBodyReadText($oIE)
		
			If StringInStr($htmlbody, "iOffer") Then
				$body &= "iOffer Transaction Exist" & @CR
				;_Log($timestamp &@TAB& "PayPal " & $acct & " - iOffer Transaction Exist"& @CR)
				$back = _IEGetObjByName($oIE, "cancel.x")
				_IEAction($back, "click")
				;_IEWaitForTitle($oIE, "Account overview - PayPal - Windows Internet Explorer")
				_IELinkClickByText($oIE, "History")		
				;_IEWaitForTitle($oIE, "History - PayPal - Windows Internet Explorer")
					
			
			ElseIf StringInStr($htmlbody, "Atomic Mall Order") Then
				$body &= "Atomic Mall Order Exist" & @CR
				;_Log($timestamp &@TAB& "PayPal " & $acct & " - Atomic Mall Order Exist"& @CR)
				$back = _IEGetObjByName($oIE, "cancel.x")
				_IEAction($back, "click")
				;_IEWaitForTitle($oIE, "Account overview - PayPal - Windows Internet Explorer")
				_IELinkClickByText($oIE, "History")		
				;_IEWaitForTitle($oIE, "History - PayPal - Windows Internet Explorer")			

			Else
				

				$itemTable = _IETableGetCollection ($oIE, 5)
				$itemTableData = _IETableWriteToArray ($itemTable, True)
				
				;_ArrayDisplay($itemTableData)
				
				$itemCount=0
				$itemNameSplit=""
				;$isNumber=0
					For $i=1 to UBound($itemTableData)-1
								
						;MsgBox(0,$i, $itemTableData[$i][1])
						
						If StringInStr($itemTableData[$i][0], "Amount") = 0 And StringInStr($itemTableData[$i][1], "Coupon Discount") = 0 And StringInStr($itemTableData[$i][1], "Coupon") = 0 Then
							$itemArray[$itemCount][2] = $itemTableData[$i][0] ;; qty
							$itemArray[$itemCount][5] = $itemTableData[$i][3] ;; price
							$itemNameSplit = StringSplit($itemTableData[$i][1], "#", 2)
							;MsgBox(0,"StringInStr($itemTableData[$i][1])", StringInStr($itemTableData[$i][1], "Coupon Discount"))
							;_ArrayDisplay($itemNameSplit)
							
							$itemArray[$itemCount][0] = StringStripWS($itemNameSplit[1], 8) ;; item number
							$itemArray[$itemCount][1] = StringTrimRight($itemNameSplit[0], 5) ;; item name
							$itemCount+=1
						EndIf
						;_ArrayDisplay($itemArray)
					Next
					
					$found = True
					
					;; define existence 
					
					;MsgBox(0, "$found", $found)
					
					For $i=0 To UBound($itemArray)-1
						
						If $itemArray[$i][0] <> "" Then
							
							If db_item_title_exist($itemArray[$i][1]) = True Then
								If db_item_check_stock($itemArray[$i][1]) <> True Then
									$found=False
								EndIf
							Else
								$found=False
								add_to_DB($itemArray[$i][1], "name", "ebay", $itemArray[$i][0], "item_number")
					
							EndIf
						EndIf
					Next
					
					;MsgBox(0, "$found", $found)
					
					;;define per item weight
				  $emptyFound=False
				  $packagingType = ""
				  $itemWeight = 0
				  $itemSKU = ""
				  $itemType = ""
				  $arraySize=0
				  $emptyFound_packagingType=False
				  $emptyFound_itemWeight=False
				  $emptyFound_itemType=False
				  $emptyFound_itemSKU=False

					If $rawOrderTable[$w][3] <> "" and $rawOrderTable[$w][3] <> " " Then
						;MsgBox(0,"$rawOrderTable[$w][3]",$rawOrderTable[$w][3])
						;MsgBox(0,"$transcation id", $transactionid)
						$approval = flagged($rawOrderTable[$w][3], $transactionid, $rawOrderTable[$w][5])
						;MsgBox(0,"approval",$approval)
				
						If $approval <> True Then
							$emptyFound = True
						EndIf
					EndIf
				  
				  ;; $itemArray[0] - item number
				  ;; $itemArray[1] - item name
				  ;; $itemArray[2] - qty
				  ;; $itemArray[3] - weight
				  ;; $itemArray[4] - package type
				  ;; $itemArray[5] - price
				  ;; $itemArray[6] - item type
				  ;; $itemArray[7] - item sku				  
				  
					 For $k=0 To UBound($itemArray)-1
						If $itemArray[$k][0] <> "" Then
							$packagingType = retrievePackagingType($itemArray[$k][1], "ebay")  
							;MsgBox(0,"$packagingType",$packagingType)
							If $packagingType = "" Then
							   $emptyFound = True 
							   $emptyFound_packagingType=True
							Else
							   $itemArray[$k][4] = $packagingType
							EndIf
							
							$itemWeight = retrieveWeight($itemArray[$k][1], "ebay")
							If $itemWeight = 0 Then
							   $emptyFound = True 
							   $emptyFound_itemWeight= True
							  
							Else
							   $itemArray[$k][3] = $itemWeight			   
							EndIf
							
							$itemType = retrieveItemType($itemArray[$k][1], "ebay")
							If $itemType = "" Then
							   $emptyFound = True 
							   $emptyFound_itemType=True
							  
							Else					   
							   $itemArray[$k][6] = $itemType
							EndIf   

							$itemSKU = retrieveSKU($itemArray[$k][1], "ebay")
							If $itemSKU = "" Then
							   $emptyFound = True 
							   $emptyFound_itemSKU=True
							Else					   
							   $itemArray[$k][7] = $itemSKU

							EndIf
													
					If StringInStr($itemArray[$k][1], "Purse") = 0 Or StringInStr ($itemArray[$k][1], "purse") = 0 Then 
						If  $emptyFound_packagingType=True Then							
							$body &= $itemArray[$k][1] & "itemWeight is empty"
						EndIf
						If $emptyFound_itemWeight= True Then
							$body &= $itemArray[$k][1] & "packagingType is empty"
						EndIf
						If $emptyFound_itemType=True Then
							$body &= $itemArray[$k][1] & "itemSKU is empty"
						EndIf
						If $emptyFound_itemSKU=True Then
							$body &= $itemArray[$k][1] & "itemType is empty"
						EndIf
					EndIf
							$arraySize += 1
						

						EndIf
						;MsgBox(0,"$k",$k)
					 Next

				  
				  

				  ;MsgBox(0, "$emptyFound", $emptyFound)
				  ;MsgBox(0, "$packagingType", $packagingType)
				  ;MsgBox(0, "$itemType", $itemType)
				  ;MsgBox(0, "$itemSKU", $itemSKU)
				  ;MsgBox(0, "$itemWeight", $itemWeight)
					;_arraydisplay($itemArray)
				  
				  $itemString=""
				  
				  
				  ;;package String
				  
				  ;MsgBox(0, "$arraySize", $arraySize)
				  
					 If $arraySize > 1 Then
						 ;MsgBox(0, "$arraySize > 1","true")
						Local $packageArray[$arraySize][2]
						$count=0

						For $m=0 To UBound($itemArray)-1
							If $itemArray[$m][0] <> "" Then
							   If $packageArray[0][0] = "" Then
								  $packageArray[0][0] = $itemArray[$m][4]
								  $packageArray[0][1] = $itemArray[$m][2]
								  $count+=1
								  
							   Else
								  $found = False
								  For $n=0 To UBound($packageArray)-1
									 
									 If $packageArray[$n][0] = $itemArray[$m][4] Then
										$packageArray[$n][1]+=1
										$found = True
									 EndIf
								  
								  Next
								  
								  If $found <> True Then
									 $packageArray[$count][0] = $itemArray[$m][4]	
									 $packageArray[$count][1] = $itemArray[$m][2]
									 $count+=1
								  EndIf
							   
								EndIf
							EndIf
						Next
						
						;_ArrayDisplay($itemArray)
						_ArraySort($packageArray)
						;_ArrayDisplay($packageArray)
						
						For $k=0 To Ubound($packageArray)-1
						   If $packageArray[$k][0] <> "" Then
							  $itemString &= $packageArray[$k][0]
							  $itemString &= $packageArray[$k][1]
						   EndIf
						Next
					
						;_ArrayDisplay($packageArray)
						
						For $i=0 To UBound($packageArray)-1
							For $j=0 To UBound($packageArray,2)-1
								$packageArray[$i][$j] = ""
							Next
						Next
					
					Else
					;MsgBox(0, "$arraySize > 1","else")
						If $itemArray[0][0] <> "" Then
							;_ArrayDisplay($itemArray)
							$itemString = $itemArray[0][4]
							$itemString &= $itemArray[0][2]
						EndIf	
					EndIf

					
				
				  ;MsgBox(0,"$itemString",$itemString)
				  
				  ;_ArrayDisplay($packageArray)
				  ;;package string existence
				  $packaging = ""
				  $packaging = define_packaging($itemString) 

				  ;;package string weight
				  $packagingWeight = 0
				  $packagingWeight = define_packaging_weight($itemString)
				  
				  If $packaging="" or $packagingWeight=0 Then
					 $emptyFound = True
					 $body &= "Item String: "&$itemString& " Packaging Missing" & @CR
				  EndIf
				  
				  ;; total weight item weight * qty + packaging weight
					$weightTotal=0
				  
				  ;MsgBox(0, "$emptyFound", $emptyFound)
				 ; MsgBox(0, "$packaging - $packagingWeight", $packaging & " - " & $packagingWeight)
				  
				  
				  If $emptyFound = False Then

					 If $arraySize > 1 Then
						For $k=0 To UBound($itemArray)-1
							If $itemArray[$k][0] <> "" Then
								$weightTotal+= $itemArray[$k][3] * $itemArray[$k][2]
							EndIf
						Next
						
					 Else
						If $itemArray[0][0] <> "" Then
							$weightTotal = $itemArray[0][3] * $itemArray[0][2]
						EndIf
					 EndIf
					 
					 $weightTotal+= $packagingWeight

				EndIf


;;multiple orders end function paypalOrderProceeds End

					$start=0
					
					If $emptyFound = False And $weightTotal <> 0 Then
						
						;$transTable = _IETableGetCollection ($oIE, 3)
						;$transTableData = _IETableWriteToArray ($transTable, True)
						
						;_ArrayDisplay($transTableData)
						;$transSplit = StringSplit($transTableData[1][0], "#", 2)
						;_ArrayDisplay($transSplit)
						;$transactionid = StringStripWS($transSplit[1], 8)
						;$transactionid = StringTrimRight($transactionid, 1)
						
						
						;MsgBox(0, "$transactionid", $transactionid)
					$addressName=""
					$address=""
					
					If addressChangeExist($transactionid) = True Then
						$address = addressChange($transactionid)
						$addressName = addressChangeName($transactionid)
					
					Else
						$addressTable = _IETableGetCollection ($oIE, 2)
						$addressTableData = _IETableWriteToArray ($addressTable, True)

						$addressSplit = $addressTableData[1][1]
									
						;MsgBox(0,"test",$data)
						$addressSplit = StringSplit ($addressSplit, @LF)
						;_ArrayDisplay($addressSplit)
								
									
						For $i = 1 To UBound($addressSplit)-1
						;MsgBox(0,"$data[$i] = $open_order_names[$w]", $data[$i] &" - "& $open_order_names[$w])
							If StringRegExp($addressSplit[$i], "Protection address") = 1 OR StringRegExp($addressSplit[$i], "Ship to address") = 1 OR StringRegExp($addressSplit[$i], "Tips to sell securely") = 1 Then
							;MsgBox(0, "test", StringRegExp($data[$i], "Tips to sell securely"))
										
								$start = $i+1
									
							EndIf
						Next
						
						;_ArrayDisplay($addressSplit)
						$domestic = False
						
						For $i=$start To UBound($addressSplit)-1
							
							If StringInStr($addressSplit[$i], "United States") <> 0 Then
								$domestic = True
							ElseIf StringInStr($addressSplit[$i], "Confirmed") = 0 Then
								$address &= StringReplace ($addressSplit[$i], @CR, "")&@CR		
							EndIf

							If $i=$start Then
								$addressName = StringReplace ($addressSplit[$i], @CR, "")
							EndIf


						Next

						;MsgBox(0,"address",$address)

						$addressNameSplit = StringSplit($addressName, " ",2)

							If UBound($addressNameSplit) > 2 then
								$addressName=""
								;MsgBox(0,"test",UBound($addressNameSplit)-1)
								For $k=0 To UBound($addressNameSplit)-1
									
									If $addressNameSplit[$k] <> "" Then

										If $addressName <> "" Then
											$addressName&=" "&$addressNameSplit[$k]
										Else
											$addressName&=$addressNameSplit[$k]
										EndIf
			
									EndIf
								Next
								
							EndIf
					EndIf
						;MsgBox(0,$addressName,$address)

							If $acct = 1 Then
								$company = "EMD Store"
							ElseIf $acct = 2 Then
								$company = "EMD Merchant"
							ElseIf $acct = 3 Then
								$company = "EMDCELL"
							Else
								$body &= "unknown paypal account number $acct= " & $acct&@CR
								$noMatch = True
							EndIf


						;If isAlphabet($address) = False Then
							;MsgBox(0,"unknown character exist", $address)
							;_Log($timestamp & @TAB & "unknown character exist in address " & $address&@CR)	
						;	$body &= "unknown character exist in address " & $address&@CR	
						;	$noMatch = True
						
						If $domestic = True Then
							

							
							If $company <> "" Then
								If $weightTotal < 13 Or $weightTotal = 13 Then									
									printFirstClass($address, $weightTotal, $transactionid, $company, $packaging)
								ElseIf $weightTotal > 13
									
									printPriority($address, $weightTotal, $transactionid, $company, $packaging)
								Else
									$body &= "Paypal print shipping label total weight unknown"
									$noMatch = True
								EndIf
							
							
							
							EndIf
							;paypalPrintDomesticLabel()
						
						ElseIf StringInStr($address, "Puerto Rico") <> 0 Then
						
							;_Log($timestamp & @TAB & "Puerto Rico Order Exist" & @CR)	
							$body &= "Puerto Rico Order Exist"
							$noMatch = True
						Else
							;;skip
							;$noMatch = True
							
							;; international shipping
							If $company <> "" Then
									;$totalQty = 0
								   ;For $k=0 To UBound($itemArray)-1
								;	   $totalQty += $itemArray[$k][2]
								 ;  Next
								
								printInternationalFirstClassPayPal($address, $weightTotal, $transactionid, $company, $packaging, $itemArray)
								
							
							EndIf
						
							
						
						
							
						EndIf 	;; end of is alphabet
						
						
						If $noMatch <> True Then
					
							ControlClick("Stamps.com Pro", "", 32551)
							Sleep(3000)
							
							$bodyTxt = ""
							 ;;Add Tracking Number
								WinWaitActive("Stamps.com Pro")
								Sleep(2000)
								$trackingIE = _IEAttach("Stamps.com Pro", "Embedded")
								_IELoadWait($trackingIE)
								;MsgBox(0,"$addressName", $addressName)
							
							$skipTracking = False
							
							For $k = 0 to 15
								 $bodyTxt = _IEBodyReadText($trackingIE)
								
								;MsgBox(0,"StringInStr ($bodyTxt, $addressName)", StringInStr ($bodyTxt, $addressName))
								If StringInStr ($bodyTxt, $addressName) <> 0 Then
									$trackingTable = _IETableGetCollection($trackingIE, 1)
									$trackingData = _IETableWriteToArray($trackingTable, True)
									;MsgBox(0,StringInStr ($trackingData[1][4], $addressName), $trackingData[1][4]&"addressName="&$addressName)
									If StringInStr($trackingData[1][4], $addressName) <> 0 Then
										$temp_tracking = $trackingData[1][6]
										ExitLoop
									EndIf
									
									ControlClick("Stamps.com Pro", "", "[CLASS:ToolbarWindow32; INSTANCE:1]", "Primary", 1, 124,15 ) 
									_IELoadWait($trackingIE)
									Sleep(10000)
								EndIf
								
								If $k > 7 Then
									ControlClick("Stamps.com Pro", "", "[CLASS:ToolbarWindow32; INSTANCE:1]", "Primary", 1, 124,15 ) 
									_IELoadWait($trackingIE)					
									Sleep(60000)
								EndIf

							Next
						
							If $k > 14 Then
								$body &= "Address Name " & $addressName & "Cannot be found on Stamp.com table; no tracking info entered for Paypal "& $acct & @CR
								$skipTracking = True
							EndIf
						
							Sleep(2000)
							ControlClick("Stamps.com Pro", "", 32513)
							Sleep(2000)


								WinActivate("PayPal Website Payment Details")
								WinWaitActive("PayPal Website Payment Details")

							If $domestic = False Then
								$temp_tracking = "None-tracking, custom declaration# " & $temp_tracking
							EndIf

							If ($skipTracking <> true) Then
								
								_IELinkClickByText($oIE, "Add Tracking Info")
								
								;_IELoadWait($oIE)
								_IEWaitForTitle($oIE, "Add Tracking Info and Order Status")
								;_IEWaitFor("Add Tracking Info and Order Status")
								
								Local $oInput = _IEGetObjByName($oIE, "shipping_status")
								_IEFormElementOptionselect($oInput, "S" , 1, "byValue")
								
								$oInput = _IEGetObjByName($oIE, "track_num")
								_IEFormElementSetValue($oInput, $temp_tracking)
								
								$oInput = _IEGetObjByName($oIE, "shipping_co_id")
								_IEFormElementOptionselect($oInput, "1" , 1, "byValue")
								
								;MsgBox(0, "okay", "okay?")
								
								$oInput = _IEGetObjByName($oIE, "Save")
								_IEAction($oInput, "click")
								
								;_IELoadWait($oIE)
								_IEWaitForTitle($oIE, "PayPal Website Payment Details")

							EndIf
							
								Sleep(700)
				
		
								_IELinkClickByText($oIE, "Print Packing Slip")
						
						
							WinWaitActive("Packing Slip - PayPal - Windows Internet Explorer")
							sleep(2000)
							printPaypalPackingSlip($itemArray ,$packaging)
							
								;$oInput = _IEFrameGetObjByName($oIE, "cancel.x")
								$oInput = _IEGetObjByName($oIE, "cancel.x")
								_IEAction($oInput, "click")
							
							_IELoadWait($oIE)
							
							
							
							inventoryUpdate($itemArray, 7, 2)
							
						EndIf ;; end of noMatch 
						

						
					EndIf ;; end of found is True
					
					
				
			
			EndIf ;; if order is from iOffer or Atomic Mall
		
			
	
			For $i=0 To UBound($itemArray)-1
				For $j=0 To UBound($itemArray,2)-1
					$itemArray[$i][$j] = ""
				Next
			Next	
			
			;_ArrayDisplay($itemArray)
	
		EndIf ;; end of if proceed is not false
	
		
			EndIf
		EndIf ;; $rawOrderTable is not empty

	  _IELinkClickByText($oIE, "History")
	  _IELoadwait($oIE)
	  Sleep(2000)

   Next
	  
	$finish = True
	  
	ElseIf $reload > 10 Then
		$body &= "Paypal "&$acct&" with $wintitle "&$wintitle&@CR
	  
	EndIf ;; end of $winTitle check



Next

;_IELinkClickByText($oIE, "Log Out")
_IELoadWait($oIE)
sleep(2000)

   
EndFunc



Func printPaypalPackingSlip($itemArray, $packagingType)

   Local $ref_id="", $qtyString="", $totalQty=0, $itemSKU=""
   
   ;_ArrayDisplay($itemName_SKU_Qty_Weight_Ptype_ItemType)
   ;MsgBox(0, "UBound($itemName_SKU_Qty_Weight_Ptype_ItemType)", UBound($itemName_SKU_Qty_Weight_Ptype))

   For $i=0 To UBound($itemArray)-1
	  
	  If $itemArray[$i][0] <> "" Then
		  If $ref_id <> "" And $qtyString <> "" Then			  
			 $ref_id &= ", "
			 $qtyString &= ", "			 
		  EndIf
		  
		  $itemSKU = define_single_item_ref_id($itemArray[$i][7])
		  $ref_id &= $itemArray[$i][6]
		  $ref_id &= "-"&$itemSKU
		  $qtyString &= $itemArray[$i][2]
	   
		  $totalQty += $itemArray[$i][2]
	  
	  EndIf
   Next
   
   
   
    Send("!f")
	Sleep(700)
	Send("u")

	;Send("v")
	;Sleep(700)
	;WinWaitActive("Print Preview")
	Sleep(1000)

	;Send("!u")
	;Sleep(700)

	WinWaitActive("Page Setup")
	Sleep(700)

	Send("!h")
	Sleep(700)

	Send("c")
	Sleep(700)
	

    Send($ref_id)
	Send("{ENTER}")
					
	Send("{TAB 2}")
	Send("c")

	Send("Qty: " & $qtyString)
	Send("{ENTER}")
					
	Send("{TAB}")
	Send("c")
	Send($packagingType)
	Send("{ENTER}")
					
	Send("{TAB 2}")
	Send("c")					
	Send("Total Item: " & $totalQty)
	Send("{ENTER}")					
					
	Send("!n")
	Send("Verdana")
	Sleep(700)

	Send("{TAB}")
	Sleep(700)


	If $totalQty > 1 Then
		Send("Bold")

	Else
		Send("Regular")

	EndIf

	;Sleep(1000)

	Send("{TAB}")
	Sleep(500)
	
	If $totalQty < 2 Then
		Send("11")
	ElseIf $totalQty > 1 And $totalQty < 3 Then
		Send("10")
	
	ElseIf $totalQty > 2 And $totalQty < 4 Then
		Send("9")	
	
	Else
		Send("8")

	EndIf
					
					
	Sleep(500)
	Send("{ENTER}")
	Sleep(1000)

	Send("{TAB 2}")
	Sleep(500)
	Send("{ENTER}")
	Sleep(3000)

	;WinWaitActive("Print Preview")
	;Sleep(700)
	
	Send("^p")
	Sleep(3000)
	WinWaitActive ("Print")
	;MsgBox(0,"okay?","okay?")
	Sleep(2000)
	Send("!p")  ;print
					
	;Sleep(700)
	;WinWaitActive("Buy.com")
	;Sleep(700)

EndFunc

Func printAmazonPackingSlip($itemName_SKU_Qty_Weight_Ptype_ItemType, $packagingType)
   
   Local $ref_id="", $qtyString="", $totalQty=0, $itemSKU=""
   
   ;_ArrayDisplay($itemName_SKU_Qty_Weight_Ptype_ItemType)
   ;MsgBox(0, "UBound($itemName_SKU_Qty_Weight_Ptype_ItemType)", UBound($itemName_SKU_Qty_Weight_Ptype))

   For $i=0 To UBound($itemName_SKU_Qty_Weight_Ptype_ItemType)-1 Step 1
	  
	  If $ref_id <> "" And $qtyString <> "" Then
		  
		 $ref_id &= ", "
		 $qtyString &= ", "
		 
	  EndIf
	  
	  $itemSKU = define_single_item_ref_id($itemName_SKU_Qty_Weight_Ptype_ItemType[$i][1])
	  $ref_id &= $itemName_SKU_Qty_Weight_Ptype_ItemType[$i][5]
	  $ref_id &= "-"&$itemName_SKU_Qty_Weight_Ptype_ItemType[$i][1]
	  $qtyString &= $itemName_SKU_Qty_Weight_Ptype_ItemType[$i][2]
   
	  $totalQty += $itemName_SKU_Qty_Weight_Ptype_ItemType[$i][2]
	  
	  
   Next
   
   
   
    Send("!f")
	Sleep(700)
	Send("u")

	;Send("v")
	;Sleep(700)
	;WinWaitActive("Print Preview")
	Sleep(1000)

	;Send("!u")
	;Sleep(700)

	WinWaitActive("Page Setup")
	Sleep(700)

	Send("!h")
	Sleep(700)

	Send("c")
	Sleep(700)
	

    Send($ref_id)
	Send("{ENTER}")
					
	Send("{TAB 2}")
	Send("c")

	Send("Qty: " & $qtyString)
	Send("{ENTER}")
					
	Send("{TAB}")
	Send("c")
	Send($packagingType)
	Send("{ENTER}")
					
	Send("{TAB 2}")
	Send("c")					
	Send("Total Item: " & $totalQty)
	Send("{ENTER}")					
					
	Send("!n")
	Send("Verdana")
	Sleep(700)

	Send("{TAB}")
	Sleep(700)


	If $totalQty > 1 Then
		Send("Bold")

	Else
		Send("Regular")

	EndIf

	;Sleep(1000)

	Send("{TAB}")
	Sleep(500)
	
	If $totalQty < 2 Then
		Send("11")
	ElseIf $totalQty > 1 And $totalQty < 3 Then
		Send("10")
	
	ElseIf $totalQty > 2 And $totalQty < 4 Then
		Send("9")	
	
	Else
		Send("8")

	EndIf
					
					
	Sleep(500)
	Send("{ENTER}")
	Sleep(1000)

	Send("{TAB 2}")
	Sleep(500)
	Send("{ENTER}")
	Sleep(3000)

	;WinWaitActive("Print Preview")
	;Sleep(700)
	
	Send("^p")
	Sleep(1000)
	WinWaitActive ("Print")
	Sleep(2000)
	Send("!p")  ;print
					
	;Sleep(700)
	;WinWaitActive("Buy.com")
	;Sleep(700)
   
   
EndFunc

;;;; Print USPS International First Class ;;;;

Func printInternationalFirstClassPayPal($address, $weightTotal, $transactionid, $company, $packageType, $itemArray)
	
	
	Local $lb=0, $oz=0, $qty=0, $desc="", $singleItemWeight=0, $price=0
	Local $add2="", $city="", $returnAddress="", $chkadd=""
	
	WinActivate("Stamps.com Pro")
	WinWaitActive("Stamps.com Pro")	
	ControlClick("Stamps.com Pro","", 32514)
	Sleep(1000)	
	
	
	$returnAddress = ControlGetText("Stamps.com Pro","",1711)

	If StringInStr($returnAddress, $company) = 0 Then
	
		ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit1]", "{UP}{UP}{UP}{HOME}{SHIFTDOWN}{END}{SHIFTUP}{DEL}")
		Sleep(500)
		ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit1]", $company)
		Sleep(500)
	
		Send("{tab}")
		Sleep(500)
	
		WinWaitActive("Confirm Return Address")
		ControlSetText("Confirm Return Address","","[CLASSNN:Edit1]", "")
		ControlSetText("Confirm Return Address","","[CLASSNN:Edit2]", $company)
		Sleep(500)
		ControlClick("Confirm Return Address","", 1)
		WinWaitActive("Stamps.com Pro")
		Sleep(500)
	
	EndIf
	
	If StringInStr($address, "United Kingdom") Then
		$address = StringReplace($address, "United Kingdom", "Great Britain")
		
	EndIf
	
	Run("notepad.exe")
	WinWaitActive("Untitled - Notepad")
	ControlSend("Untitled - Notepad", "", "[CLASSNN:Edit1]", $address)
	ControlSend("Untitled - Notepad", "", "[CLASSNN:Edit1]", "^{a}")
	ControlSend("Untitled - Notepad", "", "[CLASSNN:Edit1]", "^{c}")
	$address = Clipget()
	
	
	WinClose("Untitled - Notepad")
	ControlClick("Notepad","","[CLASSNN:Button2]")
	
	
	;MsgBox(0,"$address",$address)


	;WinActivate("Untitled - Notepad")
	
	;WinActivate("Stamps.com Pro")
	;ControlSetText("Stamps.com Pro", "", "[CLASSNN:Edit3]", "")
	;ControlFocus("Stamps.com Pro", "", 20002)	
	;ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit3]", "^{p}")
	
	;ControlSend("Untitled - Notepad", $address)
	;ControlSetText("Stamps.com Pro", "", "[CLASSNN:Edit3]", $address)
	
	
	;$chkadd = ControlGetText("Confirm International Address","",20002)
	
	
		
	WinActivate("Stamps.com Pro")
	WinWaitActive("Stamps.com Pro")
	ControlFocus("Stamps.com Pro", "", 20002)
	Send("^v")
	;ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit3]", $address)
	;Sleep(5000)	
	
	Send("{tab}")
	Sleep(500)

	
	;; confirming internaional address
	$add2 = ControlGetText("Confirm International Address","",1021)
	$city = ControlGetText("Confirm International Address","",1018)
			
	If ( $add2 <> "" And $city = "") Then
	
		ControlSetText("Confirm International Address", "", "[CLASSNN:Edit5]", $add2) 
		ControlSetText("Confirm International Address", "", "[CLASSNN:Edit4]", "" )

	EndIf
	
	
	sleep(500)
	ControlClick("Confirm International Address", "", "[CLASSNN:Button2]")
	
	WinWaitActive("Stamps.com Pro")

	;MsgBox(0,"okay","okay?")

	ControlSend("Stamps.com Pro", "", "", "!{l}") ;; Email Recipent checkbox
	Sleep(500)
	ControlSetText("Stamps.com Pro", "", "[CLASSNN:Edit4]", "emdcell.shipstream@gmail.com")
	Sleep(500)

	ControlSetText("Stamps.com Pro", "", "[CLASSNN:Edit5]", "9095819572")
	Sleep(500)


	;MsgBox(0,"weight",$weightTotal)
	
	;; entering weight
	If $weightTotal < 13 Then
		$weightTotal = Ceiling($weightTotal)
		ControlSetText("Stamps.com Pro", "", "[CLASSNN:Edit7]", $weightTotal)
	
	ElseIf $weightTotal > 13  And $weightTotal < 16 Then
		$lb = 1
		$oz = 0 
		
		ControlSetText("Stamps.com Pro", "", "[CLASSNN:Edit6]", $lb)
	
	Else
		$lb = $weightTotal/16
		$lb = Floor($lb)
		$oz = Mod($weightTotal, 16)
		$oz = Ceiling($oz)
	
		ControlSetText("Stamps.com Pro", "", "[CLASSNN:Edit6]", $lb)
		ControlSetText("Stamps.com Pro", "", "[CLASSNN:Edit7]", $oz)
	EndIf
	
	

ControlFocus("Stamps.com Pro", "", 1964)
			Send ("{Down 13}")	
			Sleep(1000)
	
		;MsgBox(0,"$packageType", $packageType)
	
	
	Select
   
		Case $packageType = "BUBBLE_MAILER"
			
			ControlFocus("Stamps.com Pro", "", 1964)
			Send ("{Down 13}")	
			Sleep(1000)
		
		Case $packageType = "BOX444"
			ControlCommand("Stamps.com Pro", "", "[CLASS:ComboBox; INSTANCE:1]", "SelectString", "Packages")
			Sleep(1000)

			ControlSetText("Stamps.com Pro", "", "[CLASSNN:Edit6]", "4")
			Sleep(300)
			ControlSetText("Stamps.com Pro", "", "[CLASSNN:Edit7]", "4")
			Sleep(300)
			ControlSetText("Stamps.com Pro", "", "[CLASSNN:Edit8]", "4")
			Sleep(300)


	EndSelect
	
	
	
	
	;; enter custom information
	   ControlClick("Stamps.com Pro", "", "[CLASSNN:Button7]")
	   Sleep(1000)
		WinWaitActive("Customs Information")
		
		ControlCommand("Customs Information", "", "[CLASSNN:ComboBox1]", "SelectString", "Merchandise")
	
	;_ArrayDisplay($itemArray)
	
	For $i=0 To UBound($itemArray)-1
		$price=0
		If $itemArray[$i][0] <> "" Then
			$desc = define_international_desc($itemArray[$i][1])
			$singleItemWeight = $itemArray[$i][3]*$itemArray[$i][2]
			
			;MsgBox(0,"$desc",$desc)
			;MsgBox(0,"$singleItemWeight",$singleItemWeight)
			;MsgBox(0,"$itemArray[$i][2]",$itemArray[$i][2])
			;MsgBox(0,"$itemArray[$i][5]",$itemArray[$i][5])
			
			$priceSplit = StringSplit($itemArray[$i][5], " ", 2)
			$price = $priceSplit[0]
			$price = StringTrimLeft($price, 1)
			If $price < 0.05 Then
				$price = 1
			EndIf
			
			ControlSetText("Customs Information", "", "[CLASSNN:Edit6]", $itemArray[$i][2]) ;;qty
			ControlSetText("Customs Information", "", "[CLASSNN:Edit7]", $desc) ;;item description
			;ControlSetText("Customs Information", "", "[CLASSNN:Edit8]", "") ;;lb
			ControlSetText("Customs Information", "", "[CLASSNN:Edit9]", $singleItemWeight) ;;oz
			ControlSetText("Customs Information", "", "[CLASSNN:Edit11]", $price) ;; price	
		   ;; complete custom form
			
			ControlClick("Customs Information", "", "[CLASSNN:Button4]")
		    Sleep(1000)
				
			;MsgBox(0,"$i","okay?")		
			
		EndIf
   Next
   
	ControlClick("Customs Information", "", "[CLASSNN:Button7]")
	  Sleep(1000)

	ControlClick("Customs Information", "", "[CLASSNN:Button8]")
	  Sleep(1000)

	WinWaitActive("Stamps.com Pro")
	
	ControlCommand("Stamps.com Pro", "", "[CLASS:ComboBox; INSTANCE:4]", "SelectString", "Zebra/Eltron Type - Standard 4x6 label - roll")
	Sleep(500)

   	MsgBox(0,"confirm","okay?")

	ControlClick("Stamps.com Pro", "", "[CLASSNN:Button9]")
   Sleep(5000)


	
   printStampcomLabel($address)
			
	WinWaitActive("Stamps.com Pro")		
				
	

EndFunc


;;;; Print USPS First Class ;;;;

Func printFirstClass($address, $weightTotal, $description, $company, $packageType)
	
	Local $sTitle = "", $complete=True, $getText="", $state=""
	Local $rejectedAddressSplitName=""
	
   $weightTotal = Ceiling($weightTotal)
   
   WinActivate("Stamps.com Pro")
   WinWaitActive("Stamps.com Pro")
   
   ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit1]", "{UP}{UP}{UP}{HOME}{SHIFTDOWN}{END}{SHIFTUP}{DEL}")
   Sleep(1000)
   ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit1]", $company)
	Sleep(1000)
   
	Select
   
		Case $packageType = "BUBBLE_MAILER"
			ControlCommand("Stamps.com Pro", "", "[CLASS:ComboBox; INSTANCE:1]", "SelectString", "Thick Envelopes (over 3/4 inches)")
			Sleep(1000)
		
		Case $packageType = "BOX444"
			ControlCommand("Stamps.com Pro", "", "[CLASS:ComboBox; INSTANCE:1]", "SelectString", "Packages")
			Sleep(1000)

			ControlSetText("Stamps.com Pro", "", "[CLASSNN:Edit6]", "4")
			Sleep(300)
			ControlSetText("Stamps.com Pro", "", "[CLASSNN:Edit7]", "4")
			Sleep(300)
			ControlSetText("Stamps.com Pro", "", "[CLASSNN:Edit8]", "4")
			Sleep(300)


	EndSelect
   
   ControlSend("Stamps.com Pro", "", "[CLASS:Edit; INSTANCE:5]", "{HOME}{SHIFTDOWN}{END}{SHIFTUP}{DEL}") 
   Sleep(1000)
   ControlSend("Stamps.com Pro", "", "[CLASS:Edit; INSTANCE:5]", $weightTotal) ;; select weight
   Sleep(1000)
   
	$hList = ControlGetHandle("Stamps.com Pro", "", "SysListView322")
	_GUICtrlListView_SetItemSelected($hList, 0) ;; select first-class mail
   Sleep(1000)
   
   
   ControlCommand("Stamps.com Pro", "", "[CLASS:ComboBox; INSTANCE:2]", "SelectString", "Delivery Confirmation")
   Sleep(1000)
	

	ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit2]", "{UP}{UP}{UP}{UP}{UP}{UP}{HOME}{SHIFTDOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{END}{SHIFTUP}{DEL}")
	Sleep(400)
	ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit2]", $address)


	
	If ControlCommand ("Stamps.com Pro", "", "[CLASSNN:Button7]", "IsChecked", "") <> 1 Then
		ControlClick("Stamps.com Pro", "", "[CLASSNN:Button7]") ;; Email Recipent checkbox
		Sleep(1000)
		ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit3]", "emdcell.shipstream@gmail.com") ;; Email Recipent
		Sleep(1000)	
	
	EndIf

	ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit11]", "holder")
	ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit11]", "{HOME}{SHIFTDOWN}{END}{SHIFTUP}{DEL}")
	ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit11]", "Order ID: "&$description) ;; Order ID
	;ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit11]", $description)
    Sleep(1000)
   	
	;MsgBox(0, "confirm", "okay?")
	
	ControlClick("Stamps.com Pro", "", "[CLASSNN:Button21]")
   Sleep(5000)
   	
   printStampcomLabel($address)
			
	WinWaitActive("Stamps.com Pro")		
			
	
EndFunc

Func rejectAddress($address)
	
	;;Blank out everything first
	
	ControlSetText("Rejected Address", "", "[CLASSNN:Edit1]", "")
	ControlSetText("Rejected Address", "", "[CLASSNN:Edit2]", "")
	ControlSetText("Rejected Address", "", "[CLASSNN:Edit3]", "")
	ControlSetText("Rejected Address", "", "[CLASSNN:Edit4]", "")
	ControlSetText("Rejected Address", "", "[CLASSNN:Edit5]", "")
	ControlSetText("Rejected Address", "", "[CLASSNN:Edit7]", "")
	ControlSetText("Rejected Address", "", "[CLASSNN:Edit8]", "")
	ControlSetText("Rejected Address", "", "[CLASSNN:Edit9]", "")
	ControlSetText("Rejected Address", "", "[CLASSNN:Edit10]", "")
	ControlSetText("Rejected Address", "", "[CLASSNN:Edit11]", "")
	ControlSetText("Rejected Address", "", "[CLASSNN:Edit12]", "")
	ControlSetText("Rejected Address", "", "[CLASSNN:Edit13]", "")
	ControlSetText("Rejected Address", "", "[CLASSNN:Edit14]", "")
	ControlSetText("Rejected Address", "", "[CLASSNN:Edit15]", "")
	ControlSetText("Rejected Address", "", "[CLASSNN:Edit16]", "")
	ControlSetText("Rejected Address", "", "[CLASSNN:Edit17]", "")
	ControlSetText("Rejected Address", "", "[CLASSNN:Edit18]", "")
	
	
			$rejectedAddressSplit = StringSplit($address, @CR, 2)
			$rejectedAddressSplitName = StringSplit($rejectedAddressSplit[0], " ", 2)
			
			
			
			If UBound($rejectedAddressSplitName) > 1 Then	
				
				ControlSetText("Rejected Address", "", "[CLASSNN:Edit3]", $rejectedAddressSplitName[0]) ;; first name
				ControlSetText("Rejected Address", "", "[CLASSNN:Edit5]", $rejectedAddressSplitName[1]) ;; last name
			Else
				ControlSetText("Rejected Address", "", "[CLASSNN:Edit3]", $rejectedAddressSplitName[0]) ;; first name
			EndIf
			
			If UBound($rejectedAddressSplit) > 3 Then
				ControlSetText("Rejected Address", "", "[CLASSNN:Edit9]", $rejectedAddressSplit[1]) ;; Company
				ControlSetText("Rejected Address", "", "[CLASSNN:Edit11]",$rejectedAddressSplit[2]) ;; Address
				$rejectedAddressSplitCityState = StringSplit($rejectedAddressSplit[3], ",", 2)
			
			Else
				ControlSetText("Rejected Address", "", "[CLASSNN:Edit11]",$rejectedAddressSplit[1]) ;; Address
				$rejectedAddressSplitCityState = StringSplit($rejectedAddressSplit[2], ",", 2)
				
			EndIf
			
			;ControlSend("Rejected Address", "", "[CLASSNN:Edit7]", $address) ;; Title
			;ControlSend("Rejected Address", "", "[CLASSNN:Edit8]", $address) ;; Department
			
			ControlSetText("Rejected Address", "", "[CLASSNN:Edit14]", $rejectedAddressSplitCityState[0]) ;; City
			
			$rejectedAddressSplitStateZip = StringSplit($rejectedAddressSplitCityState[1], " ", 2)
			
			;;C:\Documents and Settings\Cellular\Desktop\Autoit\auto_emdcell_24.au3 (1407) : ==> Array variable has incorrect number of subscripts or subscript dimension range exceeded.:
			;;$rejectedAddressSplitStateZip = StringSplit($rejectedAddressSplitCityState[1], " ", 2)
			;;$rejectedAddressSplitStateZip = StringSplit(^ ERROR
			
			$state = defineState($rejectedAddressSplitStateZip[0])
			
			ControlSetText("Rejected Address", "", "[CLASSNN:Edit17]", $state) ;; State
			ControlSetText("Rejected Address", "", "[CLASSNN:Edit18]", $rejectedAddressSplitStateZip[1]) ;; zip code
			
			MsgBox(0, "confirm", "okay?")
			
EndFunc

Func buyPostage()	
	
	ControlClick("Insufficient Postage", "", "[CLASSNN:Button1]")
	Sleep(500)
	WinWaitActive("Purchase Postage")
			
	ControlClick("Purchase Postage", "", "[CLASSNN:Button6]") ;; select $100
	Sleep(1000)
	ControlClick("Purchase Postage", "", "[CLASSNN:Button1]") ;; click ok
	Sleep(1000)
	
	WinWaitActive("Stamps.com")
	;;Would you like to submit your request to purchase $100.00 of postage?  Once you have submitted a request, it cannot be undone.
	ControlClick("Stamps.com", "", "[CLASSNN:Button1]") ;; click yes
	
	;; wait for approval
	
	WinWaitActive("Stamps.com")
	Sleep(1000)
	;;Your postage purchase request for $100.00 has been approved.
	ControlClick("Stamps.com", "", "[CLASSNN:Button1]") ;; ok
	Sleep(2000)
			
EndFunc

Func printStampcomLabel($address)
	
	For $k=0 to 19
		$winTitle = WinGetTitle("")
		
		If $winTitle = "Insufficient Postage" Then
			buyPostage()
				
		ElseIf $winTitle = "Modified Address" Then
		
			ControlClick("Modified Address", "", "[CLASSNN:Button2]")
			Sleep(1000)
			
		ElseIf $winTitle = "USPS Restrictions" Then

			ControlClick("USPS Restrictions", "", "[CLASSNN:Button1]")
			Sleep(1000)
		
		;ElseIf $winTitle = "Ambiguous Address" Then

		ElseIf $winTitle = "Address Match Details" Then

			ControlClick("Address Match Details", "", "[CLASSNN:Button1]")
			Sleep(1000)

		ElseIf $winTitle = "Stamps.com" Then
			$getText = WinGetText("Stamps.com")
				If StringInStr($getText, "You are printing postage with today's date and it is") <> 0 Then
					ControlClick("Stamps.com", "", "[CLASSNN:Button1]")
				
				ElseIf StringInStr($getText, "Invalid address.") <> 0 Then
					MsgBox(0,"Address Error","Check the address again")
				;;More Stamps.com Messages Can be added here.
					
				EndIf
		
		ElseIf $winTitle = "Rejected Address" Then
			WinWaitActive("Rejected Address")
			
			rejectAddress($address)
		
		ElseIf $winTitle = "Print Label"Then
		
			ExitLoop
		
		ElseIf $winTitle = "Stamps.com Pro" Or $winTitle = "" Then
		
			Sleep(3000)
		
		ElseIf $k > 20 Then
		
			MsgBox(0, "error", "$k exceeded 20 " & $k)
		
		Else	
			MsgBox(0, "error - no matching window", $winTitle)
			
		EndIf	
		
		
	Next
	
	
	WinWaitActive("Print Label")
	Sleep(1000)

	ControlCommand("Print Label", "", "[CLASSNN:ComboBox1]", "SelectString", "\\http://192.168.1.15\ZebraLP")
	Sleep(2000)
   
	ControlClick("Print Label", "", "[CLASSNN:Button3]")
	Sleep(3000)	
	
EndFunc


Func printPriority($address, $weightTotal, $description, $company, $packageType)

	Local $sTitle = "", $complete=True, $getText="", $lb=0, $oz=0
	
	;MsgBox(0,"$address, $weightTotal, $description, $company, $packageType", $address&" "&$weightTotal&" "&$description&" "&$company&" "&$packageType)
	
	
	If $weightTotal < 13 Then
		MsgBox(0, "Error", "wrong shipping label - weight less than 13 oz")
	
	Else
	
		If $weightTotal > 13 And $weightTotal < 16 Then
			$lb = 1
			$oz = 0
		
		ElseIf $weightTotal > 16 Then
			;MsgBox(0, "$weightTotal", $weightTotal)
			$lb = $weightTotal/16
			$lb = Floor($lb)
			;$oz = $weightTotal - $lb * 16
			;MsgBox(0, "$oz", $oz)
			$oz = Mod($weightTotal, 16)
			;MsgBox(0, "Mod($weightTotal, 16)", Mod($weightTotal, 16))
			$oz = Ceiling($oz)
			;MsgBox(0, "$oz", $oz)
		
		EndIf
		
	
   WinActivate("Stamps.com Pro")
   WinWaitActive("Stamps.com Pro")
   
   ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit1]", "{UP}{UP}{UP}{HOME}{SHIFTDOWN}{END}{SHIFTUP}{DEL}")
   Sleep(1000)
   ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit1]", $company)
	Sleep(1000)
   
	Select
   
		Case $packageType = "BUBBLE_MAILER"
			ControlCommand("Stamps.com Pro", "", "[CLASS:ComboBox; INSTANCE:1]", "SelectString", "Thick Envelopes (over 3/4 inches)")
			Sleep(1000)
		
		Case $packageType = "BOX444"
			ControlCommand("Stamps.com Pro", "", "[CLASS:ComboBox; INSTANCE:1]", "SelectString", "Packages")
			ControlSetText("Stamps.com Pro", "", "[CLASSNN:Edit6]", "4")
			Sleep(300)
			ControlSetText("Stamps.com Pro", "", "[CLASSNN:Edit7]", "4")
			Sleep(300)
			ControlSetText("Stamps.com Pro", "", "[CLASSNN:Edit8]", "4")
			Sleep(300)

		Case $packageType = "BOX644"
			ControlCommand("Stamps.com Pro", "", "[CLASS:ComboBox; INSTANCE:1]", "SelectString", "Packages")
			ControlSetText("Stamps.com Pro", "", "[CLASSNN:Edit6]", "6")
			Sleep(300)
			ControlSetText("Stamps.com Pro", "", "[CLASSNN:Edit7]", "4")
			Sleep(300)
			ControlSetText("Stamps.com Pro", "", "[CLASSNN:Edit8]", "4")
			Sleep(300)

		Case $packageType = "BOX664"
			ControlCommand("Stamps.com Pro", "", "[CLASS:ComboBox; INSTANCE:1]", "SelectString", "Packages")
			ControlSetText("Stamps.com Pro", "", "[CLASSNN:Edit6]", "6")
			Sleep(300)
			ControlSetText("Stamps.com Pro", "", "[CLASSNN:Edit7]", "6")
			Sleep(300)
			ControlSetText("Stamps.com Pro", "", "[CLASSNN:Edit8]", "4")
			Sleep(300)

	EndSelect
 
   ;ControlSetText("Stamps.com Pro", "", "[CLASS:Edit; INSTANCE:4]", "{HOME}{SHIFTDOWN}{END}{SHIFTUP}{DEL}") 
   ;Sleep(1000)
   ControlSetText("Stamps.com Pro", "", "[CLASS:Edit; INSTANCE:4]", $lb) ;; select weight
   Sleep(1000)

   ;ControlSetText("Stamps.com Pro", "", "[CLASS:Edit; INSTANCE:5]", "{HOME}{SHIFTDOWN}{END}{SHIFTUP}{DEL}") 
   ;Sleep(1000)
   ControlSetText("Stamps.com Pro", "", "[CLASS:Edit; INSTANCE:5]", $oz) ;; select weight
   Sleep(1000)
   
	$hList = ControlGetHandle("Stamps.com Pro", "", "SysListView322")
	_GUICtrlListView_SetItemSelected($hList, 1) ;; select first-class mail
   Sleep(1000)
   
   
   ControlCommand("Stamps.com Pro", "", "[CLASS:ComboBox; INSTANCE:2]", "SelectString", "Delivery Confirmation")
   Sleep(1000)
	

	ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit2]", "{UP}{UP}{UP}{UP}{UP}{UP}{HOME}{SHIFTDOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{END}{SHIFTUP}{DEL}")
	Sleep(400)
	ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit2]", $address)


	
	If ControlCommand ("Stamps.com Pro", "", "[CLASSNN:Button7]", "IsChecked", "") <> 1 Then
		ControlClick("Stamps.com Pro", "", "[CLASSNN:Button7]") ;; Email Recipent checkbox
		Sleep(1000)
		ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit3]", "emdcell.shipstream@gmail.com") ;; Email Recipent
		Sleep(1000)	
	
	EndIf

	ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit11]", "holder")
	ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit11]", "{HOME}{SHIFTDOWN}{END}{SHIFTUP}{DEL}")
	ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit11]", "Order ID: "&$description) ;; Order ID
	;ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit11]", $description)
    Sleep(1000)
   	
	MsgBox(0, "confirm", "okay?")
	
	ControlClick("Stamps.com Pro", "", "[CLASSNN:Button21]")
   Sleep(5000)

	printStampcomLabel($address)
	
			
	WinWaitActive("Stamps.com Pro")		

	ControlSetText("Stamps.com Pro", "", "[CLASSNN:Edit4]", "0")
   
	EndIf ;; weight leses than 13 oz 
   
EndFunc



Func printSmallFlatRateBox($address, $weightTotal, $description, $company)

	
	Local $sTitle = "", $complete=True, $getText=""
	
   $weightTotal = Ceiling($weightTotal)
   
   WinActivate("Stamps.com Pro")
   WinWaitActive("Stamps.com Pro")
   
   ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit1]", "{UP}{UP}{UP}{HOME}{SHIFTDOWN}{END}{SHIFTUP}{DEL}")
   Sleep(1000)
   ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit1]", $company)
	Sleep(1000)

	ControlCommand("Stamps.com Pro", "", "[CLASS:ComboBox; INSTANCE:1]", "SelectString", "USPS Small Flat Rate Priority Mail Box")
	Sleep(1000)
   
   ControlSend("Stamps.com Pro", "", "[CLASS:Edit; INSTANCE:5]", "{HOME}{SHIFTDOWN}{END}{SHIFTUP}{DEL}") 
   Sleep(1000)
   ControlSend("Stamps.com Pro", "", "[CLASS:Edit; INSTANCE:5]", $weightTotal) ;; select weight
   Sleep(1000)
   
	$hList = ControlGetHandle("Stamps.com Pro", "", "SysListView322")
	_GUICtrlListView_SetItemSelected($hList, 1) ;; select first-class mail
   Sleep(1000)
   
   
   ControlCommand("Stamps.com Pro", "", "[CLASS:ComboBox; INSTANCE:2]", "SelectString", "Delivery Confirmation")
   Sleep(1000)
	

	ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit2]", "{UP}{UP}{UP}{UP}{UP}{UP}{HOME}{SHIFTDOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{END}{SHIFTUP}{DEL}")
	Sleep(400)
	ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit2]", $address)


	
	If ControlCommand ("Stamps.com Pro", "", "[CLASSNN:Button7]", "IsChecked", "") <> 1 Then
		ControlClick("Stamps.com Pro", "", "[CLASSNN:Button7]") ;; Email Recipent checkbox
		Sleep(1000)
		ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit3]", "emdcell.shipstream@gmail.com") ;; Email Recipent
		Sleep(1000)	
	
	EndIf

	ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit11]", "holder")
	ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit11]", "{HOME}{SHIFTDOWN}{END}{SHIFTUP}{DEL}")
	ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit11]", "Order ID: "&$description) ;; Order ID
	;ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit11]", $description)
    Sleep(1000)
   	
	;MsgBox(0, "confirm", "okay?")
	
	ControlClick("Stamps.com Pro", "", "[CLASSNN:Button21]")
   Sleep(5000)

	printStampcomLabel($address)
			
	WinWaitActive("Stamps.com Pro")		
			

EndFunc


Func printMediumFlatRateBox($address, $weightTotal, $description, $company)

	
	Local $sTitle = "", $complete=True, $getText=""
	
   $weightTotal = Ceiling($weightTotal)
   
   WinActivate("Stamps.com Pro")
   WinWaitActive("Stamps.com Pro")
   
   ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit1]", "{UP}{UP}{UP}{HOME}{SHIFTDOWN}{END}{SHIFTUP}{DEL}")
   Sleep(1000)
   ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit1]", $company)
	Sleep(1000)

	ControlCommand("Stamps.com Pro", "", "[CLASS:ComboBox; INSTANCE:1]", "SelectString", "USPS Medium Flat Rate Box")
	Sleep(1000)
   
   ControlSend("Stamps.com Pro", "", "[CLASS:Edit; INSTANCE:5]", "{HOME}{SHIFTDOWN}{END}{SHIFTUP}{DEL}") 
   Sleep(1000)
   ControlSend("Stamps.com Pro", "", "[CLASS:Edit; INSTANCE:5]", $weightTotal) ;; select weight
   Sleep(1000)
   
	$hList = ControlGetHandle("Stamps.com Pro", "", "SysListView322")
	_GUICtrlListView_SetItemSelected($hList, 1) ;; select first-class mail
   Sleep(1000)
   
   
   ControlCommand("Stamps.com Pro", "", "[CLASS:ComboBox; INSTANCE:2]", "SelectString", "Delivery Confirmation")
   Sleep(1000)
	

	ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit2]", "{UP}{UP}{UP}{UP}{UP}{UP}{HOME}{SHIFTDOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{END}{SHIFTUP}{DEL}")
	Sleep(400)
	ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit2]", $address)


	
	If ControlCommand ("Stamps.com Pro", "", "[CLASSNN:Button7]", "IsChecked", "") <> 1 Then
		ControlClick("Stamps.com Pro", "", "[CLASSNN:Button7]") ;; Email Recipent checkbox
		Sleep(1000)
		ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit3]", "emdcell.shipstream@gmail.com") ;; Email Recipent
		Sleep(1000)	
	
	EndIf

	ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit11]", "holder")
	ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit11]", "{HOME}{SHIFTDOWN}{END}{SHIFTUP}{DEL}")
	ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit11]", "Order ID: "&$description) ;; Order ID
	;ControlSend("Stamps.com Pro", "", "[CLASSNN:Edit11]", $description)
    Sleep(1000)
   	
	MsgBox(0, "confirm", "okay?")
	
	ControlClick("Stamps.com Pro", "", "[CLASSNN:Button21]")
   Sleep(5000)

	printStampcomLabel($address)
			
	WinWaitActive("Stamps.com Pro")		
			

EndFunc


Func retrievePackagingType($title, $tablename)
   
	Local $PackagingType=""

	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	
	Local $SQLCode_getRecord = ""
	Local $GetContent = ""

		$SQLCode_getRecord = 'SELECT * FROM '& $tablename &  ' WHERE name = "' & $title & '"'
		$GetContent = _Query($SQLInstance, $SQLCode_getRecord)
				
			With $GetContent
				While Not .EOF
					$PackagingType = .Fields("package" ).value
					.MoveNext
				WEnd
			EndWith
			
	Return $PackagingType
   
EndFunc

Func retrieveItemType($title, $tablename)
   
	Local $PackagingType=""


	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	
	Local $SQLCode_getRecord = ""
	Local $GetContent = ""

		$SQLCode_getRecord = 'SELECT * FROM '& $tablename &  ' WHERE name = "' & $title & '"'
		$GetContent = _Query($SQLInstance, $SQLCode_getRecord)
				
			With $GetContent
				While Not .EOF
					$PackagingType = .Fields("type" ).value
					.MoveNext
				WEnd
			EndWith
			
	Return $PackagingType
   
EndFunc

Func retrieveWeight($title, $tablename)
	
   	Local $weight=0


	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	
	Local $SQLCode_getRecord = ""
	Local $GetContent = ""

		$SQLCode_getRecord = 'SELECT * FROM '& $tablename &  ' WHERE name = "' & $title & '"'
		$GetContent = _Query($SQLInstance, $SQLCode_getRecord)
				
			With $GetContent
				While Not .EOF
					$weight = .Fields("weight" ).value
					.MoveNext
				WEnd
			EndWith
			
	Return $weight
	
EndFunc



Func retrieveSKU($title, $tablename)
   
   	Local $sku=""

	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	
	Local $SQLCode_getRecord = ""
	Local $GetContent = ""

		$SQLCode_getRecord = 'SELECT * FROM '& $tablename &  ' WHERE name = "' & $title & '"'
		$GetContent = _Query($SQLInstance, $SQLCode_getRecord)
				
			With $GetContent
				While Not .EOF
					$sku = .Fields("ref_id" ).value
					.MoveNext
				WEnd
			EndWith
			
	Return $sku
   
   
EndFunc



Func _Print_Medium_Flat_Rate_Box($array, $row)
	
		WinActivate("Shipstream™ Manager")
		
		$sTitle =WinGetTitle("")
					
		If $sTitle <> "Shipstream™ Manager - Zebra_label_medium_flat_rate_emdstore" Then
			Send("!f")
			Sleep(500)
			Send("o")
			Sleep(500)
			if WinExists("Shipstream™ Manager") Then
				Send("!n")
			EndIf
			WinWaitActive("Open Layout File","")
			WinActivate("Open Layout File")
			Sleep(500)
			Send("Zebra_label_medium_flat_rate_emdstore.LYT")
			Sleep(500)
											
			Send("{ENTER}")
		EndIf
				
		Sleep(1000)
				
		_Address_Input($array[$row][25], $array[$row][26], $array[$row][27], $array[$row][28], $array[$row][29], $array[$row][30], $array[$row][31])
					
		WinWaitActive("Shipstream™ Manager")
		Send("^p")
						

		_Check_Time()		
					
		Sleep(1000)
		
		_Print_label("Order ID: "&$array[$row][1], "mediumflat")
	

EndFunc


Func _Print_Large_Flat_Rate_Box($array, $row)
		
		WinActivate("Shipstream™ Manager")
		
		$sTitle =WinGetTitle("")
		
					
		If $sTitle <> "Shipstream™ Manager - Zebra_label_large_flat_rate_emdstore" Then
			Send("!f")
			Sleep(500)
			Send("o")
			Sleep(500)
			if WinExists("Shipstream™ Manager") Then
				Send("!n")
			EndIf
			WinWaitActive("Open Layout File","")
			WinActivate("Open Layout File")
			Sleep(500)
			Send("Zebra_label_large_flat_rate_emdstore.LYT")
			Sleep(500)
											
			Send("{ENTER}")
		EndIf
				
		Sleep(1000)
				
		_Address_Input($array[$row][25], $array[$row][26], $array[$row][27], $array[$row][28], $array[$row][29], $array[$row][30], $array[$row][31])
					
		WinWaitActive("Shipstream™ Manager")
		Send("^p")
						

		_Check_Time()		
					
		Sleep(1000)
		
		
		_Print_label("Order ID: "&$array[$row][1], "smallflat")


EndFunc



;;;;;;;;;;;;; PRINT PACKING SLIP ;;;;;;;;;;;;;;;
Func printBuycomPackingSlip($orderArray, $packagingType, $headerRefid)
	
	Local $upperCaseId="", $type="", $headerQty="", $totalqty=0
	
	For $i=0 To UBound($orderArray)-1 Step 1
		If $orderArray[$i][0] <> "" Then
			;$upperCaseId = StringUpper ($orderArray[$i][6])
			;$type = StringLeft($upperCaseId, 1)
			
			If  $headerQty <> "" Then
			;	$headerRefid &=", "
				$headerQty &=", "
			EndIf
			
			;Select
			;	Case $type = "T"
			;		$type = "HOME"
			
			;	Case $type = "P"
			;		$type = "CAR"

			;	Case $type = "H"
			;		$type = "POUCH"
			
			;	Case $type = "S"
			;		$type = "V3CASE"

			;	Case $type = "D"
			;		$type = "DATACABLE"
					
			;	Case Else
			;		$type = ""

			;EndSelect
		
			$totalQty += $orderArray[$i][7]
			$headerQty &= $orderArray[$i][7]
			;$headerRefid &= $type&" - "&$orderArray[$i][6]
		EndIf
	Next


	Sleep(500)

	$oIE = _IECreate()
	_IENavigate($oIE, "https://sellertools.marketplace.buy.com/OrderPackingSlip.aspx?o="&$orderArray[0][1])
	
	
	Sleep(700)

	Send("!f")
	Sleep(200)

	Send("v")
	Sleep(700)
	WinWaitActive("Print Preview")
	Sleep(1000)

	Send("!u")
	Sleep(700)

	WinWaitActive("Page Setup")
	Sleep(700)

	Send("!h")
	Sleep(700)

	Send("c")
	Sleep(700)
	
	
	Send( $headerRefid)
	Send("{ENTER}")
					
	Send("{TAB 2}")
	Send("c")

	Send("Qty: " & $headerQty)
	Send("{ENTER}")
					
	Send("{TAB}")
	Send("c")
	Send($packagingType)
	Send("{ENTER}")
					
	Send("{TAB 2}")
	Send("c")					
	Send("Total Item: " & $totalQty)
	Send("{ENTER}")					
					
	Send("!n")
	Send("Verdana")
	Sleep(700)

	Send("{TAB}")
	Sleep(700)




	If $totalQty > 1 Then
		Send("Bold")

	Else
		Send("Regular")

	EndIf

	;Sleep(1000)

	Send("{TAB}")
	Sleep(500)
	

	If $totalQty = 2  Then
		Send("9")	
	ElseIf $totalQty > 2 Then
		Send("7")						
	Else
		Send("10")	
	EndIf

	Sleep(500)
	Send("{ENTER}")
	Sleep(1000)

	Send("{TAB 2}")
	Sleep(500)
	Send("{ENTER}")
	Sleep(700)

	WinWaitActive("Print Preview")
	Sleep(700)

	Send("!p")
	Sleep(700)
	WinWaitActive ("Print")
	Sleep(4000)
	Send("!p")  ;print
					
	Sleep(700)
	WinWaitActive("Buy.com")
	Sleep(700)

	_IEQuit($oIE)

EndFunc

;;;;;;;;;;;;; ENTER TRACKING NUMBER ;;;;;;;;;;;;;;;
Func addBuycomTracking($orderArray, $trackingNum)
	
	;$orderid = $orderArray[$i][1]
	;$orderqty = $orderArray[$i][7]
	
	Local $getHTML=""
	
	$oIE = _IECreate()
	_IENavigate($oIE, "https://sellertools.marketplace.buy.com/OrderE.aspx?o="&$orderArray[0][1]&"&pg=1")
	_IELoadWait($oIE)
	
	$oTrack = _IEFormGetObjByName ($oIE, "aspnetForm")
	
	For $i=0 To UBound($orderArray)-1 Step 1

		$oQty = _IEFormElementGetObjByName ($oTrack, "ctl00$cphMiddle$drOrderItems$ctl0"&$i&"$txtShipQty")
		$oTrackingNum = _IEFormElementGetObjByName ($oTrack, "ctl00$cphMiddle$drOrderItems$ctl0"&$i&"$txtShipTrackingNum")
		$oShipCarrier = _IEFormElementGetObjByName ($oTrack, "ctl00$cphMiddle$drOrderItems$ctl0"&$i&"$ddlShipCarrier")

		; Set field values and submit the form
		_IEFormElementSetValue ($oQty, $orderArray[$i][7])
		_IEFormElementSetValue ($oTrackingNum, $trackingNum)
		_IEFormElementSetValue ($oShipCarrier, "USPS")		
		
	Next
	
	;MsgBox(0, "confirm", "okay?")
	
	$o_submit = _IEGetObjByName ($oIE, "ctl00$cphMiddle$btnSubmit")
	_IEAction ($o_submit, "click")
		
		
	_IELoadWait($oIE)
	Sleep(3000)
	

	
	For $j=0 To 20 Step 1
		
		$getHTML = ""
		$getHTML = _IEBodyReadHTML($oIE)
		
		If StringInStr($getHTML, "ctl00$cphMiddle$btnSubmit") = 0 Then
			ExitLoop
		
		ElseIf $j>20 Then
			$body &= "Order ID: "&$orderArray[$i][1]&" tracking ID need to be double checked"
			ExitLoop
		
		Else
			Sleep(5000)
			_IENavigate($oIE, "https://sellertools.marketplace.buy.com/OrderE.aspx?o="&$orderArray[$i][1]"&pg=1")
			_IELoadWait($oIE)
		
		EndIf
	
	Next
	

	_IEQuit($oIE)

			
EndFunc 




Func define_ref_id_weight($ref_id)

	Local $weight=0
	
	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	
	Local $SQLCode_getRecord = "SELECT * FROM inventory WHERE ref_id = '" & $ref_id & "'"
	Local $GetContent = _Query($SQLInstance, $SQLCode_getRecord)

		With $GetContent
			While Not .EOF
				$weight = .Fields("weight" ).value
				
				.MoveNext
			WEnd
		EndWith
		
;MsgBox(0,"$weight",$weight)

	Return $weight
	
EndFunc


Func buy_com_ref_id_packaging_type($ref_id)

	Local $packing_id=""
	
	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	
	Local $SQLCode_getRecord = "SELECT * FROM inventory WHERE ref_id = '" & $ref_id & "'"
	Local $GetContent = _Query($SQLInstance, $SQLCode_getRecord)

		With $GetContent
			While Not .EOF
				$packing_id = .Fields("packing_id" ).value
				
				.MoveNext
			WEnd
		EndWith
		
;MsgBox(0,"$packing_id",$packing_id)

	Return $packing_id
	
EndFunc

Func defineDesc($sku, $table)
	
	Local $weight=0, $return_value="", $field_value1="", $field_value2="", $field_value3="", $field_value4="", $desc=""

	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	
	Local $SQLCode_getRecord = ""
	Local $GetContent = ""
	
		$SQLCode_getRecord = "SELECT * FROM "& $table 
		$GetContent = _Query($SQLInstance, $SQLCode_getRecord)

			With $GetContent
				While Not .EOF
					$field_value1 = .Fields($field1).value
					$field_value1 = StringUpper($field_value1)
					$field_value2 = .Fields($field2).value
					$field_value2 = StringUpper($field_value2)
					$field_value3 = .Fields($field3).value
					$field_value3 = StringUpper($field_value3)
					$field_value4 = .Fields($field4).value
					$field_value4 = StringUpper($field_value4)
					$desc = .Fields($desc).value

					
;MsgBox(0, "$field_value1234", $field_value1 & ", " & $field_value2 & ", " & $field_value3 & ", " & $field_value4)
					
					If StringInStr($sku, $field_value1) <> 0 OR StringInStr($sku, $field_value2) <> 0 OR StringInStr($sku, $field_value3) <> 0 OR StringInStr($sku, $field_value4) <> 0 Then
						$return_value = $desc
						ExitLoop
					EndIf
					.MoveNext
				WEnd
			EndWith
		

	Return $return_value
	
EndFunc

Func define_single_item_ref_id($item_name)
	
	Local $weight=0, $return_value="", $field_value="", $field_value1="", $field_value2="", $field_value3="", $field_value4="", $field_value5="", $field_value6=""
	Local $field="ref_id", $field1="ref_id1", $field2="ref_id2", $field3="ref_id3", $field4="ref_id4", $field5="ref_id5", $field6="ref_id6"

	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	
	Local $SQLCode_getRecord = ""
	Local $GetContent = ""
	
		$SQLCode_getRecord = "SELECT * FROM inventory" 
		$GetContent = _Query($SQLInstance, $SQLCode_getRecord)

			With $GetContent
				While Not .EOF
					$field_value = .Fields($field).value
					$field_value = StringUpper($field_value)					
					$field_value1 = .Fields($field1).value
					$field_value1 = StringUpper($field_value1)
					$field_value2 = .Fields($field2).value
					$field_value2 = StringUpper($field_value2)
					$field_value3 = .Fields($field3).value
					$field_value3 = StringUpper($field_value3)
					$field_value4 = .Fields($field4).value
					$field_value4 = StringUpper($field_value4)
					$field_value5 = .Fields($field5).value
					$field_value5 = StringUpper($field_value5)
					$field_value6 = .Fields($field6).value
					$field_value6 = StringUpper($field_value6)
					
					
;MsgBox(0, "$field_value1234", $field_value1 & ", " & $field_value2 & ", " & $field_value3 & ", " & $field_value4)
					
					If StringInStr($item_name, $field_value) <> 0 OR StringInStr($item_name, $field_value1) <> 0 OR StringInStr($item_name, $field_value2) <> 0 OR StringInStr($item_name, $field_value3) <> 0 OR StringInStr($item_name, $field_value4) <> 0 OR StringInStr($item_name, $field_value5) <> 0 OR StringInStr($item_name, $field_value6) <> 0 Then
							$return_value = $field_value
						ExitLoop
					EndIf
					.MoveNext
				WEnd
			EndWith
		

	Return $return_value
	
EndFunc






Func _Add_Multiple_Tracking_Number($id, $qty_array, $temp_tracking)
	
	;_ArrayDisplay($qty_array)
	
	Local $ShipCarrier = "USPS"
	Local $trackingNum = $temp_tracking
	Local $text = "", $stop=False, $w = 0

While ($stop = False And $w < 20)
	
	If $w <> 0 Then
		Sleep(9000)
	EndIf	
	
	Local $oIE = _IECreate()
	_IENavigate($oIE, "https://sellertools.marketplace.buy.com/OrderE.aspx?o="&$id&"&pg=1")

	Sleep(1000)

	$oTrack = _IEFormGetObjByName ($oIE, "aspnetForm")

	For $i=0 To UBound($qty_array)-1
		
		If $qty_array <> "" Then
			$oQty = _IEFormElementGetObjByName ($oTrack, "ctl00$cphMiddle$drOrderItems$ctl0"&$i&"$txtShipQty")
			$oTrackingNum = _IEFormElementGetObjByName ($oTrack, "ctl00$cphMiddle$drOrderItems$ctl0"&$i&"$txtShipTrackingNum")
			$oShipCarrier = _IEFormElementGetObjByName ($oTrack, "ctl00$cphMiddle$drOrderItems$ctl0"&$i&"$ddlShipCarrier")

			; Set field values and submit the form
			_IEFormElementSetValue ($oQty, $qty_array[$i][1])
			_IEFormElementSetValue ($oTrackingNum, $trackingNum)
			_IEFormElementSetValue ($oShipCarrier, $ShipCarrier)
		EndIf
	Next
	
	Sleep(1000)
	
	Local $o_submit = _IEGetObjByName ($oIE, "ctl00$cphMiddle$btnSubmit")
	_IEAction ($o_submit, "click")
	_IELoadWait($oIE)
	
	Sleep(1500)
	
	_IEQuit($oIE)
	
	Sleep(5000)
	
	$testIE = _IECreate()
	_IENavigate($testIE, "https://sellertools.marketplace.buy.com/OrderE.aspx?o="&$id&"&pg=1")
	_IELoadWait($testIE)
	
	$text = _IEBodyReadHTML($testIE)
	
	If StringInStr($text, "ctl00$cphMiddle$btnSubmit") = 0 Then
		$stop = True

	EndIf
	
	
	_IEQuit($testIE)
	
	$w+=1
WEnd	
	
EndFunc

Func Start_Programs()
	
	Local $winGetTitle = ""
	
		;Run ("C:\Program Files\Microsoft Office\Office12\OUTLOOK.EXE")
		;Sleep(5000)
			
		;WinSetState("Inbox - Microsoft Outlook","",@SW_MINIMIZE)	
			
		Run ("C:\Program Files\Stamps.com Internet Postage\ipostage.exe")

		WinWaitActive("Stamps.com Login")
				
		ControlClick("Stamps.com Login", "", "[CLASSNN:Button2]")
		Sleep(2000)
		
		For $i=0 To 20 Step 1
			$winGetTitle = WinGetTitle("")
			
			If $winGetTitle = "Stamps.com Pro" Then
				
				ExitLoop
			
			ElseIf $winGetTitle = "Stamps.com Login" Then
				ControlClick("Stamps.com Login", "", "[CLASSNN:Button2]")
				Sleep(1000)

			ElseIf $i > 20 Then
				MsgBox(0, "Start Program Error", "Couldn't Login to Stamps.com")
				Exit
			EndIf
			
			sleep(2000)			
			
		Next
		
		;While ($i < 50)
		;	
		;	$oIE_Loading = _IEAttach("Stamps.com Pro", "Embedded")
		;	_IELoadWait($oIE_Loading)
			
		;	Sleep(3000)
			
		;	$wingettxt = WinGetText("Stamps.com Pro")
			;$wingettxt = _IEBodyReadText($oIE_Loading)
			
		;	If StringRegExp ($wingettxt, "SdcBrowserView") = 1 Then
			;If StringInStr ($wingettxt, "Learning Center") <> 0 Then	
		;		ExitLoop				
		;	EndIf
			
			;Sleep(1500)

		;	$i += 1
		;WEnd
	
	WinWaitActive("Stamps.com Pro")
	Sleep(5000)
	
	ControlClick("Stamps.com Pro","", 32513)
	Sleep(3000)

	
EndFunc

;;;;;;;;;;;;; DOWNLOAD OPEN ORDER TXT FILE ;;;;;;;;;;;;;;;
Func _Download_Open_Orders($DoIE)
	
	Local $hWnd_local, $body="", $result=True

	;$DoIE = _IECreate()
	WinActivate("Buy.com - Marketplace Seller Tools")
	
	_IENavigate($DoIE, "https://sellertools.marketplace.buy.com/OrderDownload.aspx")
		
	_IEWaitForLoad($DoIE, "Buy.com - Marketplace Seller Tools", "https://sellertools.marketplace.buy.com/OrderDownload.aspx")

	Send("{TAB 13}")
	Sleep(1000)
	Send("{ENTER}")
	WinWaitActive("Message from webpage")
	;ControlClick("Message from webpage", "", "[CLASSNN: Button1]")
	Send("{ENTER}")
	
	Sleep(4000)
	
	$body = _IEBodyReadText($DoIE)
	
	;MsgBox(0,"$body", $body)
	
	If StringInStr($body, "no open orders") Then
		$body &= "Buy.com Marketplace have no new orders" & @CR
		_Log($timestamp &@TAB& "Buy.com Marketplace have no new orders" & @CR)
		$result = False
		
	Else

			
		WinWaitActive("File Download")
		;ControlClick("File Download", "", "[CLASSNN: Button2]")		
				
		Send("!s")
		Sleep(1000)
					
		WinWaitActive("Save As") 
		Send("OpenOrderExport_"&$timestamp_filename)
		
		Sleep(1000)
		Send("!s")
				

		WinWaitActive("Download complete")	
		Sleep(500)
		WinClose("Download complete")
		Sleep(500)
		
		_IENavigate($DoIE, "https://sellertools.marketplace.buy.com")
		_IEWaitForLoad($DoIE, "Buy.com - Marketplace Seller Tools", "https://sellertools.marketplace.buy.com")
		

	EndIf
	
	;_IEQuit($DoIE)
	;_IENavigate($DoIE, "https://sellertools.marketplace.buy.com/")
	
	
	Return $result

EndFunc


Func _Check_For_Duplicates_Buy_com($array)
	
	Local $i = 2
	Local $end = UBound($array)
	;Local $same_id = ""
	Local $same_addy = "Same_Person_But_Diff_Order"
	
	
	While $i < $end
		Local $j = $i+1
			While $j < $end
				if $array[$i][1] = $array[$j][1] Then
					$array[$i][0] = $array[$i][1]
					$array[$j][0] = $array[$i][1]
				
				ElseIf	$array[$i][27] = $array[$j][27] Then
					If $array[$i][25] = $array[$j][25] Then
						$array[$i][0] = $same_addy
						$array[$j][0] = $same_addy
					EndIf
				EndIf	
				
				$j += 1
			WEnd	
		
		$i += 1
	WEnd
	
	Return $array
	
	
EndFunc

Func _IEWaitForLoad($oIE, $title, $url)
	
	Local $current_title = "", $reload = True, $i=0
	
	_IELoadWait($oIE)
	
	While ($reload = True OR $i < 50)
		
		$current_title = WinGetTitle("")
		
		
		If StringRegExp($current_title , $title) = 1 Then
			$reload = False
		
		ElseIf StringRegExp($current_title , "Internet Explorer cannot display the webpage - Windows Internet Explorer") = 1 Then
			$reload	= True
		
		EndIf
			
		If $reload = True Then
			_IENavigate($oIE, $url)
			_IELoadWait($oIE)
			Sleep(1000)
		EndIf
			
		If ($i > 50) Then
			Exit
		EndIf
		
		$i += 1
	WEnd
		
EndFunc









Func print_packing_slip($item_ref_id, $qty_string, $packaging_type, $qty_total, $item_type=0)
		Send("!f")
		Sleep(1000)

		Send("v")
		Sleep(1000)
		WinWaitActive("Print Preview")
		Sleep(1000)

		Send("!u")
		Sleep(1000)

		WinWaitActive("Page Setup")
		;Sleep(1000)

		Send("!h")
		Send("c")
					
					Send( $item_ref_id)
					Send("{ENTER}")
					
					Send("{TAB 2}")
					Send("c")
					

					
					Send("Qty: " & $qty_string)
					Send("{ENTER}")
					
					Send("{TAB}")
					Send("c")
					Send($packaging_type)
					Send("{ENTER}")
					
					Send("{TAB 2}")
					Send("c")					
					Send("Total Item: " & $qty_total)
					Send("{ENTER}")					
					
					Send("!n")
					Send("Verdana")
					Sleep(1000)

					Send("{TAB}")
					Sleep(1000)
					
					
					If $qty_total > 1 Then
						Send("Bold")

					Else
						Send("Regular")

					EndIf

					;Sleep(1000)

					Send("{TAB}")
					Sleep(500)
					
					Select
						Case $item_type < 3
						Send("11")	
					
						Case $item_type = 3
						Send("10")			

						Case $item_type = 4
						Send("9")	
						
						Case $item_type = 5
						Send("8")

						Case $item_type = 6
						Send("7")	

						Case $item_type > 6
						Send("6")	
					EndSelect
					
					
					Sleep(500)
					Send("{ENTER}")
					Sleep(1000)

					Send("{TAB 2}")
					Sleep(500)
					Send("{ENTER}")
					Sleep(1000)

					WinWaitActive("Print Preview")
					Sleep(1000)

					Send("!p")
					Sleep(1000)
					WinWaitActive ("Print")
					Sleep(2000)
					Send("!p")  ;print
EndFunc


Func define_international_desc($item_name)

	Local $desc=""
	
	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	
	Local $SQLCode_getRecord = "SELECT * FROM ebay WHERE name = '" & $item_name & "'"
	Local $GetContent = _Query($SQLInstance, $SQLCode_getRecord)

		With $GetContent
			While Not .EOF
				$desc = .Fields("international_desc" ).value
				
				.MoveNext
			WEnd
		EndWith
;MsgBox(0,"$desc",$desc)
	Return $desc
EndFunc


Func addressChangeName($orderid)
	
	Local $firstname = "", $lastname = ""
	Local $addressName = ""
	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	Local $SQLCode_getRecord = "SELECT * FROM address_change WHERE orderid = '" & $orderid & "'"
	Local $GetContent = _Query($SQLInstance, $SQLCode_getRecord)


		With $GetContent
			While Not .EOF
				$firstname = .Fields("firstname" ).value
				$lastname = .Fields("lastname" ).value
				.MoveNext
			WEnd
		EndWith
	
	If $firstname <> "" And $firstname <> " "  Then
		$addressName &= $firstname & " "
	EndIf
	
	If  $lastname <> "" And $lastname <> " "  Then
		$addressName &= $lastname
	EndIf
	
	Return $addressName

EndFunc


Func addressChange($orderid)
	
	Local $firstname = "", $lastname = "", $address1="", $address2="", $zip="", $city="", $state="", $company, $address3
	Local $ofirstname = "", $olastname = "", $oaddress1="", $oaddress2="", $ozip="", $ocity="", $ostate=""
	Local $changed_address = ""
	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	Local $SQLCode_getRecord = "SELECT * FROM address_change WHERE orderid = '" & $orderid & "'"
	Local $GetContent = _Query($SQLInstance, $SQLCode_getRecord)


		With $GetContent
			While Not .EOF
				$firstname = .Fields("firstname" ).value
				$lastname = .Fields("lastname" ).value
				$company = .Fields("company" ).value
				$address1 = .Fields("address1" ).value
				$address2 = .Fields("address2" ).value
				$address3 = .Fields("address3" ).value
				$zip = .Fields("zip" ).value
				$city = .Fields("city" ).value
				$state = .Fields("state" ).value
				.MoveNext
			WEnd
		EndWith

	$changed_address = $firstname & " " & $lastname & @CR

	If $company <> "" And $company <> " " Then
	
		 $changed_address &= $company & @CR 
	
	EndIf

	$changed_address &= $address1 & @CR

	If $address2 <> "" And $address2 <> " " Then
	
		$changed_address &= $address2 & @CR
	
	EndIf

	If $address3 <> "" And $address3 <> " " Then
	
		$changed_address &= $address3 & @CR
	
	EndIf
	
	$changed_address &= $city & ", " &  $state & " " & $zip
	
	
	Return $changed_address

EndFunc



Func paypal_change_address($ie_obj, $transaction_id)

	Local $firstname = "", $lastname = "", $address1="", $address2="", $zip="", $city="", $state=""
	Local $ofirstname = "", $olastname = "", $oaddress1="", $oaddress2="", $ozip="", $ocity="", $ostate=""
	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	Local $SQLCode_getRecord = "SELECT * FROM address_change WHERE transaction_id = '" & $transaction_id & "'"
	Local $GetContent = _Query($SQLInstance, $SQLCode_getRecord)


		With $GetContent
			While Not .EOF
				$firstname = .Fields("firstname" ).value
				$lastname = .Fields("lastname" ).value
				$address1 = .Fields("address1" ).value
				$address2 = .Fields("address2" ).value
				$zip = .Fields("zip" ).value
				$city = .Fields("city" ).value
				$state = .Fields("state" ).value
				.MoveNext
			WEnd
		EndWith
	
	If $firstname <> "" Then
		$ofirstname = _IEGetObjByName($ie_obj, "first_name")
		_IEFormElementSetValue($ofirstname, $firstname )
	EndIf
	
	If $lastname <> "" Then
		$olastname = _IEGetObjByName($ie_obj, "last_name")
		_IEFormElementSetValue($olastname, $lastname )
	EndIf
	
	If $address1 <> "" Then
		$oaddress1 = _IEGetObjByName($ie_obj, "address1")
		_IEFormElementSetValue($oaddress1, $address1 )
	EndIf

	If $address2 <> "" Then
		$oaddress2 = _IEGetObjByName($ie_obj, "address2")
		_IEFormElementSetValue($oaddress2, "" )
	Else
		$oaddress2 = _IEGetObjByName($ie_obj, "address2")
		_IEFormElementSetValue($oaddress2, $address2 )
	EndIf
	
	If 	$zip <> "" Then
		$ozip = _IEGetObjByName($ie_obj, "zip")
		_IEFormElementSetValue($ozip, $zip )
	EndIf

	If $city <> "" Then
		$ocity = _IEGetObjByName($ie_obj, "city")
		_IEFormElementSetValue($ocity, $city )
	EndIf

	If $state <> "" Then
		$ostate = _IEGetObjByName($ie_obj, "state")
		_IEFormElementSetValue($ostate, $state )	
	EndIf	
	
	
EndFunc

Func addressChangeExist($orderid)
	
	Local $found = False
	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	Local $SQLCode_getRecord = "SELECT * FROM address_change"
	Local $GetContent = _Query($SQLInstance, $SQLCode_getRecord)
	
		If $orderid <> "" Then
			With $GetContent
				While Not .EOF
					If ($orderid = .Fields("orderid" ).value) Then
						$found = True
					EndIf
					.MoveNext
				WEnd
			EndWith
		EndIf

	;MsgBox(0,"test",$found)
	Return $found

EndFunc

Func define_single_item_weight($item_name, $table, $field1, $field2, $field3, $field4)
	
	Local $weight=0, $found=False, $field_value="", $field_value1="", $field_value2="", $field_value3="", $field_value4=""

	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	
	Local $SQLCode_getRecord = ""
	Local $GetContent = ""
	
		$SQLCode_getRecord = "SELECT * FROM "& $table 
		$GetContent = _Query($SQLInstance, $SQLCode_getRecord)

			With $GetContent
				While Not .EOF
					$field_value1 = .Fields($field1).value
					$field_value1 = StringUpper($field_value1)
					$field_value2 = .Fields($field2).value
					$field_value2 = StringUpper($field_value2)
					$field_value3 = .Fields($field3).value
					$field_value3 = StringUpper($field_value3)
					$field_value4 = .Fields($field4).value
					$field_value4 = StringUpper($field_value4)					
;MsgBox(0, "$field_value1234", $field_value1 & ", " & $field_value2 & ", " & $field_value3 & ", " & $field_value4)
					
					If StringInStr($item_name, $field_value1) <> 0 OR StringInStr($item_name, $field_value2) <> 0 OR StringInStr($item_name, $field_value3) <> 0 OR StringInStr($item_name, $field_value4) <> 0 Then
						$found = True
						$field_value = .Fields($field1).value
						ExitLoop
					EndIf
					.MoveNext
				WEnd
			EndWith
		
		;MsgBox(0, "$field_value", $field_value)
		
		If $found = True Then
			
			$SQLCode_getRecord = "SELECT * FROM "&$table&" WHERE "& $field1 & " = '" & $field_value & "'"
			
			With $GetContent
				While Not .EOF
					$weight = .Fields("weight" ).value
					.MoveNext
				WEnd
			EndWith
		
		EndIf

			;MsgBox(0, "$weight", $weight)
	

	Return $weight
	
EndFunc

Func define_item_weight($item_name, $item_qty)
	
	Local $weight_per=0, $weight_total=0, $i=1, $type="";, $item_multiple=0

	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	
	Local $SQLCode_getRecord = ""
	Local $GetContent = ""

		$SQLCode_getRecord = 'SELECT * FROM ebay WHERE name = "' & $item_name & '"'
		$GetContent = _Query($SQLInstance, $SQLCode_getRecord)
				
			With $GetContent
				While Not .EOF
					$weight_per = .Fields("weight" ).value
					$type = .Fields("type" ).value
					.MoveNext
				WEnd
			EndWith
			
			;MsgBox(0, "define_item_total_weight", "Qty" & $item_array[$i][2] & "$weight_per= "& $weight_per)
			
			If $weight_per = "" Then
				$weight_total = 0
				
			Else
				;$item_multiple = StringLeft($type, 1)
				;If StringIsDigit ($item_multiple) = 1 Then
				;		MsgBox(0,"$item_multiple",$item_multiple)
				;		$weight_total += $weight_per * $item_qty * $item_multiple
				;	Else
						$weight_total += $weight_per * $item_qty
				;EndIf
				
			EndIf


	Return $weight_total
	
EndFunc


Func define_item_total_weight($item_array)
	
	Local $weight_per=0, $weight_total=0, $i=1
	Local $array[UBound($item_array)]

	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	
	Local $SQLCode_getRecord = ""
	Local $GetContent = ""

	While $i < UBound($item_array)
		
		$SQLCode_getRecord = 'SELECT * FROM ebay WHERE name = "' & $item_array[$i][1] & '"'
		$GetContent = _Query($SQLInstance, $SQLCode_getRecord)
				
			With $GetContent
				While Not .EOF
					$weight_per = .Fields("weight" ).value
				
					.MoveNext
				WEnd
			EndWith
			
			;MsgBox(0, "define_item_total_weight", "Qty" & $item_array[$i][2] & "$weight_per= "& $weight_per)
			
			If $weight_per = "" Then
				$weight_total = 0
				ExitLoop
				
			Else
					$weight_total += $weight_per * $item_array[$i][2]
			EndIf
			
		$i+=1
	WEnd

	Return $weight_total
	
EndFunc

Func _qty_copy($item_array)
	
	If $item_array[0][0] = "Qty" Then
		
		For $i=1 to UBound($item_array)-1
			$item_array[$i][2] = $item_array[$i][0]
		Next
	
	EndIf

	Return $item_array
	
EndFunc


Func flagged($msg, $transaction, $name)
	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	Local $SQLCode_getRecord = "SELECT * FROM flagged"
	Local $GetContent = _Query($SQLInstance, $SQLCode_getRecord)
	Local $found = False, $result = False, $approval = 0

	$msg = StringReplace ($msg, "'", " ")

		If $transaction <> "" Then
			With $GetContent
				While Not .EOF
					If ($transaction = .Fields("transaction_id" ).value) Then
						$found = True
					EndIf
					.MoveNext
				WEnd
			EndWith
		EndIf
		
		;MsgBox(0,"$found", $found)
		;MsgBox(0,"$flag, $transaction, $name", $msg&", "&$transaction&", " &$name)
		
		If $found = True Then
			
			$SQLCode_getRecord = "SELECT * FROM flagged WHERE transaction_id = '" & $transaction & "'"
			$GetContent = _Query($SQLInstance, $SQLCode_getRecord)
				With $GetContent
					While Not .EOF
						$approval = .Fields("approval").value		
						.MoveNext
					WEnd
				EndWith
			
			If $approval = 1 Then
				$result = True
			EndIf
		
			;MsgBox(0,"$approval", $approval)
			;MsgBox(0,"$result", $result)
		Else
			
			
			$SQLCode_insertRecord = "INSERT INTO flagged(transaction_id, name, note) VALUES('" & $transaction & "', '" & $name & "', '" & $msg & "')"
			;msgBox(0,"$SQLCode_insertRecord",$SQLCode_insertRecord)
			_Query($SQLInstance, $SQLCode_insertRecord)
			_Log(@CR & $timestamp &@TAB& "added "&$transaction&" "&$name&" to approval table" & @CR)
			$body &= "Transaction ID: "&@TAB& "added "&$transaction&" "&$name&" to approval table" & @CR
		EndIf

	Return $result

EndFunc

Func define_item_sku($item_name)
	
	Local $sku_result = ""
	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	Local $SQLCode_getRecord = 'SELECT * FROM ebay WHERE name = "' & $item_name & '"'
	Local $GetContent = _Query($SQLInstance, $SQLCode_getRecord)


		With $GetContent
			While Not .EOF
				$sku_result = .Fields("sku" ).value		
				.MoveNext
			WEnd
		EndWith

	Return $sku_result

EndFunc

Func define_item_ref_id($item_array)
	Local $ref_id_string="", $i=1, $j=0, $ref_id=""
	Local $container[UBound($item_array)]
	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	
	Local $SQLCode_getRecord = ""
	Local $GetContent = ""
	
	For $i=1 To UBound($item_array)-1
		$SQLCode_getRecord = 'SELECT * FROM ebay WHERE name = "' & $item_array[$i][1] & '"'
		$GetContent = _Query($SQLInstance, $SQLCode_getRecord)
		
		With $GetContent
			While Not .EOF
				$ref_id = .Fields("ref_id" ).value		
				.MoveNext
			WEnd
		EndWith
			
		$container[$i] = $ref_id
		
	Next
	
	
		For $j = 1 to UBound($container)-1
			If $container[$j] <> "" Then
				$ref_id_string &=  $container[$j]
			EndIf
			
			If UBound($container)-1 > 1 Then
				$ref_id_string &= ", "
			EndIf
		Next
		
	Return 	$ref_id_string

EndFunc


Func define_item_string($item_array)
	
	Local $packaging_string="", $i=1, $j=0, $k=0, $m=0, $n=1, $found = False, $no_value=False, $type=""
	Local $array[UBound($item_array)][2]
	Local $type_container[UBound($item_array)]
	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	
	Local $SQLCode_getRecord = ""
	Local $GetContent = ""

	While $i < UBound($item_array)
		
		$SQLCode_getRecord = 'SELECT * FROM ebay WHERE name = "' & $item_array[$i][1] & '"'
		$GetContent = _Query($SQLInstance, $SQLCode_getRecord)
				
			With $GetContent
				While Not .EOF
					$type = .Fields("type" ).value		
					.MoveNext
				WEnd
			EndWith
			

		
		If $type = "" Then
			$no_value = True
		Else
			$type_container[$i] = $type
		EndIf
			
				;MsgBox(0,"$type_container[$i]",$type_container[$i])
		$i+=1
	WEnd
	
		;_ArrayDisplay($type_container)	
	
	If $no_value <> True Then
		While $n < UBound($type_container)
			;msgbox(0,"test","$i="&$i)
			;$SQLCode_getRecord = "SELECT * FROM ebay WHERE name = '" & $item_array[$i][1] & "'"
			;$GetContent = _Query($SQLInstance, $SQLCode_getRecord)
				$k=0
				;$found = False
			
				While $k < UBound($array)-1
					If $array[$k][0] = $type_container[$n] Then					
						$array[$k][1]+= $item_array[$n][2] 
						$found = True										
					EndIf
					$k+=1
				WEnd
				
				If $found <> true Then
					$array[$j][0] = $type_container[$n]
					$array[$j][1] = $item_array[$n][2]
					$j+=1	

				EndIf
			;_ArrayDisplay($array)	
			
			$n+=1
		WEnd
		
		_ArraySort($array)
		;_ArrayDisplay($array)
		
		
		For $m = 0 to UBound($array)-1
			If $array[$m][0] <> "" Then
				$packaging_string &=  $array[$m][0] & $array[$m][1]
			EndIf
		Next
	EndIf
	
	;MsgBox(0,"$packaging_string",$packaging_string)
	;_ArrayDisplay($array)
	
	Return $packaging_string
	
EndFunc



Func define_packaging_weight($string)
	Local $packaging_weight = 0, $packaging_type="", $found = False

	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	Local $SQLCode_getRecord = "SELECT * FROM packaging"
	Local $GetContent = _Query($SQLInstance, $SQLCode_getRecord)


		If $string <> "" Then
			With $GetContent
				While Not .EOF
					if ($string = .Fields("type" ).value) Then
						$found = True
					EndIf
					.MoveNext
				WEnd
			EndWith
		EndIf
	
		;MsgBox(0,"test",$found)
	
		If $found = True Then
			$SQLCode_getRecord = "SELECT * FROM packaging WHERE type = '" & $string & "'"
			$GetContent = _Query($SQLInstance, $SQLCode_getRecord)

				With $GetContent
					While Not .EOF
						$packaging_weight = .Fields("box_weight" ).value		
						.MoveNext
					WEnd
				EndWith
				
			;Select
				
			;Case $packaging_type = "BUBBLE_MAILER"
			;	$packaging_weight = 0.4

			;Case $packaging_type = "BOX444"
			;	$packaging_weight = 1.9
				
			;Case $packaging_type = "BOX644"
			;	$packaging_weight = 2.6
				
			;Case $packaging_type = "BOX664"
			;	$packaging_weight = 3.5
				
			;EndSelect

		EndIf
		
		;MsgBox(0, "check weight", $packaging_weight)

	Return $packaging_weight
	
	
EndFunc

Func define_packaging($string)
	Local $packaging_type = "", $found = False

	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	Local $SQLCode_getRecord = "SELECT * FROM packaging"
	Local $GetContent = _Query($SQLInstance, $SQLCode_getRecord)


		If $string <> "" Then
			With $GetContent
				While Not .EOF
					if ($string = .Fields("type" ).value) Then
						$found = True
					EndIf
					.MoveNext
				WEnd
			EndWith
		EndIf
	
	
	
		If $found = True Then
				$SQLCode_getRecord = "SELECT * FROM packaging WHERE type = '" & $string & "'"
				$GetContent = _Query($SQLInstance, $SQLCode_getRecord)

				With $GetContent
					While Not .EOF
						$packaging_type = .Fields("box" ).value		
						.MoveNext
					WEnd
				EndWith
		
		Else
			If $packaging_type <> "" And $packaging_type <> " " Then
				add_to_DB($string, "type", "packaging")
			EndIf
		EndIf
		
		

	Return $packaging_type
	
	
EndFunc

Func db_item_title_exist($title)
	
	Local $found = false

	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	Local $SQLCode_getRecord = "SELECT * FROM ebay"
	Local $GetContent = _Query($SQLInstance, $SQLCode_getRecord)


		If $title <> "" Then
			With $GetContent
				While Not .EOF
					If ($title = .Fields("name" ).value) Then
						$found = True
					EndIf
					
					.MoveNext
				WEnd
			EndWith
		EndIf
		
	;MsgBox(0,"in db title item exist - before found = false", $found)
	
		;If $found = False Then
		;	add_to_DB($title, "ebay", "name")
			
		;Else
			
	;	If db_item_check_stock($title) = 0 Then
	;		$found = False
	;		
	;	EndIf
	
		;If db_item_check_stock($title) = False Then
		;	$found = False
		;EndIf
	;MsgBox(0,"in db title item exist", $found)
	
	Return $found

EndFunc

Func db_item_title_existence($title, $tablename, $field="name")
	
	Local $found = false

	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	Local $SQLCode_getRecord = "SELECT * FROM "&$tablename
	Local $GetContent = _Query($SQLInstance, $SQLCode_getRecord)


		If $title <> "" Then
			With $GetContent
				While Not .EOF
					If ($title = .Fields($field ).value) Then
						$found = True
					EndIf
					
					.MoveNext
				WEnd
			EndWith
		EndIf
		
	;MsgBox(0,"in db title item exist - before found = false", $found)
	
		;If $found = False Then
		;	add_to_DB($title, "ebay", "name")
			
		;Else
			
	;	If db_item_check_stock($title) = 0 Then
	;		$found = False
	;		
	;	EndIf
	
		;If db_item_check_stock($title) = False Then
		;	$found = False
		;EndIf
	;MsgBox(0,"in db title item exist", $found)
	
	Return $found

EndFunc

Func inventoryQtyCheck($ref_id, $orderQty)

	Local $inventoryQty=0
	
	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	
	Local $SQLCode_getRecord = "SELECT * FROM inventory WHERE ref_id = '" & $ref_id & "'"
	Local $GetContent = _Query($SQLInstance, $SQLCode_getRecord)

		With $GetContent
			While Not .EOF
				$inventoryQty = .Fields("qty" ).value
				
				.MoveNext
			WEnd
		EndWith
		
	If $orderQty > $inventoryQty Then
		Return False
	Else
		Return True
	EndIf


EndFunc

Func db_item_check_stock($title)
	
	Local $inStock = false
	Local $stock = 0

	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	Local $SQLCode_getRecord = 'SELECT * FROM ebay WHERE name = "' & $title & '"'
	Local $GetContent = _Query($SQLInstance, $SQLCode_getRecord)


		With $GetContent
			While Not .EOF
				$stock = .Fields("in_stock" ).value
				.MoveNext
			WEnd
		EndWith

	;MsgBox(0,"in db item check stock", $stock)

	Return $stock

EndFunc

Func itemInStock($title, $tablename, $field="name", $inStockField="in_stock")
	
	;Local $inStock = false
	Local $stock = 0


	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	Local $SQLCode_getRecord = 'SELECT * FROM '&$tablename&' WHERE '&$field&' = "' & $title & '"'
	Local $GetContent = _Query($SQLInstance, $SQLCode_getRecord)


		With $GetContent
			While Not .EOF
				$stock = .Fields($inStockField).value
				.MoveNext
			WEnd
		EndWith

	;MsgBox(0,"in db item check stock", $stock)

	Return $stock

EndFunc

Func define_skip($id)

	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	Local $SQLCode_getRecord = "SELECT * FROM skip"
	Local $GetContent = _Query($SQLInstance, $SQLCode_getRecord)
	Local $found = False

		If $id <> "" Then
			With $GetContent
				While Not .EOF
					If ($id = .Fields("id" ).value) Then
						$found = True
					EndIf
					.MoveNext
				WEnd
			EndWith
		EndIf	
	
	Return $found
	
EndFunc

Func add_to_DB($add, $field, $db, $add2="", $field2="")
	;add_to_DB($item_title_array[$m][1], "name", "ebay", $item_number_array[$m], "item_number")


	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	Local $SQLCode_insertRecord = ""
	
	;MsgBox(0,"test", "begin add_to_db")
	
	If $add2 <> "" And $field2 <> "" Then
		$SQLCode_insertRecord = 'INSERT INTO '&$db&'('&$field&', '&$field2&')'&' VALUES("' & $add &'", "'&$add2& '")'
		;MsgBox(0,"test", $SQLCode_insertRecord)
		$log = @CR & $timestamp &@TAB& "added "&$add&" and "&$add2&" in fields "&$field&" and "&$field2&" to the "&$db&" database."&@CR
	Else
		$SQLCode_insertRecord = 'INSERT INTO '&$db&'('&$field&')'&' VALUES("' & $add & '")'
		$log = @CR & $timestamp &@TAB& "added "&$add&" to the "&$db&" database."&@CR
	EndIf
	
;MsgBox(0,"test", "end add_to_db")
				;	Local $SQLCode_insertRecord = "INSERT INTO ebay(name) VALUES('" & $title & "')"
				;$GetContent = _Query($SQLInstance, $SQLCode_insertRecord)
	;MsgBox(0,"$SQLCode_insertRecord", $SQLCode_insertRecord)
	
	_Query($SQLInstance, $SQLCode_insertRecord)
	_Log($log)
	;MsgBox(0,"test", "done")
EndFunc

Func Login($oIE, $usern="buy.com@emdcell.com", $passw="cell13579", $form="aspnetForm", $f_usern="ctl00$cphMiddle$txtEmail", $f_passw="ctl00$cphMiddle$txtPassword", $f_submit="ctl00$cphMiddle$btnSubmit")
	
	Local $oForm = _IEFormGetObjByName($oIE, $form)
	Local $oUsername = _IEFormElementGetObjByName ($oForm, $f_usern)
	Local $oPassword = _IEFormElementGetObjByName ($oForm, $f_passw)
	Local $oSubmit = _IEGetObjByName ($oIE, $f_submit)

	_IEFormElementSetValue ($oUsername, $usern)
	_IEFormElementSetValue ($oPassword, $passw)

	_IEAction($oSubmit, "click")
	_IELoadWait($oIE)

EndFunc


					
Func _WaitForWinTitle($wintitle, $wincontent="")
	
	local $getTitle="", $stop = false, $s=0, $result = True, $getContent=""
	Local $max_wait = 50
	
	While ($stop = false AND $s < 50)
		
		If $s < $max_wait  Then	
			
			$getTitle = WinGetTitle("")
		
			If StringRegExp ($getTitle, $wintitle) Then				
				$stop = True
			EndIf
		
		Else
			MsgBox(0, "ERROR - " & $getTitle, "Current WIN Content: "&$getContent&@CR&@CR& "Waiting for WIN Content: " &@CR&@CR& $wincontent) 
			Exit
		
		EndIf
			
		Sleep(1000)
		$s += 1
	WEnd	
		

	
	If $wincontent <> "" AND $s < 50 Then
		$stop = False
		$s = 0
		
		While($stop = false AND $s < 50)
			
			If $s < $max_wait  Then	
				
				$getContent = WinGetText($wintitle)
					If StringInStr ($getContent, $wincontent) Then
						$stop = True
					EndIf
			
			Else 
				MsgBox(0, "ERROR - " & $getTitle, "Current WIN Content: "&$getContent&@CR&@CR& "Waiting for WIN Content: " &@CR&@CR& $wincontent) 
				Exit
				
			EndIf
			
			Sleep(1000)
			$s += 1		
		WEnd
	EndIf

EndFunc

Func isAlphabet ($string)
	$ABC = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890-,.'`/\~!@#$%^&*()_+|;:<>?"
	Local $found = True, $char=""
	
	$string = StringReplace ($string, @CR, "")
	$string = StringStripWS ($string, 8)
	
	If $string <> "" Then
		
		For $i=1  To StringLen($string)
			
			$char = StringMid($string, $i, 1)		
			
			For $j=1 To StringLen($ABC)
				If $char = StringMid($ABC, $j, 1) Then
					$found = True
					ExitLoop
				Else
					$found = False
				EndIf
			Next

			If $found = False Then
				ExitLoop
			EndIf
			
		Next	
	EndIf		
			

	Return $found

EndFunc




Func _IEWaitForTitle($oIE, $ietitle, $ietitle2="")
	
	_IELoadWait($oIE)
	Sleep(700)
	
	Local $refresh = True
	Local $getTitle = ""
	Local $i = 1, $max_wait=10
	
	While $refresh = True And $i > $max_wait

		$getTitle = WinGetTitle("")
		
		If $i < $max_wait  Then		
		
			If StringRegExp ($getTitle, $ietitle) OR StringRegExp ($getTitle, $ietitle2) Then
				$refresh = False
			
			ElseIf $getTitle = "Check out with Bill Me Later, a PayPal service - PayPal - Windows Internet Explorer" Then
				_IENavigate($oIE, "https://www.paypal.com/us/cgi-bin/webscr?cmd=_account")
				_IELoadWait($oIE)
				Sleep(1000)
			
			ElseIf StringRegExp ($getTitle, "Internet Explorer cannot display the webpage - Windows Internet Explorer") Then
				_IEAction($oIE, "refresh")
				_IELoadWait($oIE)
				Sleep(1000)
			
			ElseIf Mod($i, 10) = 0 Then
				_IEAction($oIE, "refresh")
				_IELoadWait($oIE)
				Sleep(1000)
				
			EndIf
		
		Else
			MsgBox(0, "ERROR - " & $getTitle, "Current IE Title: "&$getTitle&@CR&@CR& "Waiting for IE Title: " &@CR&@CR& $ietitle&@CR&@CR&"OR"&@CR&@CR&$ietitle2) 
			Exit
		
		EndIf
		
		;MsgBox(0,"test",$i)
		Sleep(2000)
		$i += 1
	WEnd
	
EndFunc

Func _TableArray($oIE)

	Local $oTable = _IETableGetCollection ($oIE, 0)
	Local $aTableData = _IETableWriteToArray ($oTable, True)
	;_ArrayDisplay($aTableData)

	Return $aTableData

EndFunc


Func BuyComCheckForDuplicate($array)
	
	For $i=2 To UBound($array)-1 Step 1  ;;Start from 2 because row 0 is size of the array row 1 is header
		For $j=$i+1 To UBound($array)-1 Step 1
			If $array[$i][0] <> "" And $array[$j][0] <> "" Then
				If $array[$i][27] = $array[$j][27] And $array[$i][1] <> $array[$j][1] Then
					$array[$i][0] = "Same_Person_Diff_Order"
					$array[$j][0] = "Same_Person_Diff_Order"
				EndIf
			EndIf
		Next
		
	Next
	
	Return $array
	
EndFunc

Func amazonNameDupCheck($array)
	
	Local $same_name = 1
	
	;_ArrayDisplay($array)
	
	For $i=0 To UBound($array)-1 Step 1
		If $i < UBound($array)-1 Then
			For $j = $i+1 To UBound($array)-1 Step 1
				;MsgBox(0, "$i - $j", $i & " - "& $j)
				If $array[$i][1] = $array[$j][1] Then
					$array[$i][1] = $same_name
					$array[$j][1] = $same_name
				EndIf	
			Next
		EndIf
	Next
	
	Return $array
EndFunc

Func checkPaypalDup($array)
	
	For $i=0 To UBound($array)-1
		If UBound($array) > 1 And $array[$i][1] = 0 Then
			For $j=$i+1 To UBound($array)-1
				If $array[$i][5] = $array[$j][5] Then
					If $array[$i][4] <> "Update From" And $array[$j][4] <> "Update From" And StringInStr($array[$j][8], "Print shipping label") <> 0  And StringInStr($array[$i][8], "Print shipping label") <> 0 Then
						$array[$i][1] = "samename_"&$array[$i][5]
						$array[$j][1] = "samename_"&$array[$i][5]
						$j+=1
					EndIf
				EndIf
			Next
		EndIf
	Next
	


	Return $array


EndFunc

Func inventoryUpdateDB($itemSKU, $itemQty)
	

	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	Local $SQLCode_getRecord = ""
	Local $GetContent = ""
	Local $SQLCode_updateRecord = ""	
	Local $log = "", $originalQty=0, $id=0

	Local $todayDate = _NowDate()
	$dateSplit = StringSplit($todayDate, "/", 2)

	If StringLen($dateSplit[0]) < 2 Then
		$dateSplit[0] = "0"&$dateSplit[0]
	EndIf

	If StringLen($dateSplit[1]) < 2 Then
		$dateSplit[1] = "0"&$dateSplit[1]
	EndIf
	
	$todayDate = $dateSplit[2]&$dateSplit[0]&$dateSplit[1]
	
	;MsgBox(0,"$itemSKU", $itemSKU) 

	For $j = 0 To 6
		If $j = 0 Then
			$SQLCode_getRecord = "SELECT * FROM inventory WHERE ref_id = '" & $itemSKU & "'"
			$GetContent = _Query($SQLInstance, $SQLCode_getRecord)
				
			With $GetContent
				While Not .EOF
					$originalQty = .Fields("qty").value
					.MoveNext
				WEnd
			EndWith
						
						
		Else
			$SQLCode_getRecord = "SELECT * FROM inventory WHERE ref_id"&$j&" = '" & $itemSKU & "'"
			$GetContent = _Query($SQLInstance, $SQLCode_getRecord)
				
			With $GetContent
				While Not .EOF
					$originalQty = .Fields("qty").value
					.MoveNext
				WEnd
			EndWith		

			;MsgBox(0,$j,$originalQty)
		EndIf
				
		If $originalQty <> 0 Then
			$id = $j
			ExitLoop
		EndIf
				
			
	Next
	
	;MsgBox(0,"$originalQty", $originalQty) 

		$log &= $itemSKU
				
		$log &= " - Original Qty: " & $originalQty
				
		$newQty = $originalQty-$itemQty
				
		$log &= " - Updated Qty: " & $newQty & @CRLF
				
		If $id = 0 Then
			$SQLCode_updateRecord = "UPDATE inventory SET qty = " & $newQty & " WHERE ref_id = '" & $itemSKU & "'"
			_Query($SQLInstance, $SQLCode_updateRecord)
				
			$SQLCode_updateRecord = "UPDATE inventory SET last_updated = " & $todayDate & " WHERE ref_id = '" & $itemSKU & "'"
			_Query($SQLInstance, $SQLCode_updateRecord)	
				
		ElseIf $originalQty <> 0 Then
			$SQLCode_updateRecord = "UPDATE inventory SET qty = " & $newQty & " WHERE ref_id"&$id&" = '" & $itemSKU & "'"
			_Query($SQLInstance, $SQLCode_updateRecord)
				
			$SQLCode_updateRecord = "UPDATE inventory SET last_updated = " & $todayDate & " WHERE ref_id"&$id&" = '" & $itemSKU & "'"
			_Query($SQLInstance, $SQLCode_updateRecord)						
		
		Else
			$body &= $itemSKU & " not found in inventory table with Qty: " & $itemQty

		EndIf
				
		
	
	ConsoleWrite("Return=" & FileWriteLine ("Inventory_Log" & StringReplace(_NowCalcDate(), "/", "_") & ".txt", $log ))



EndFunc

Func inventoryUpdate($inventoryArray, $refidCol, $qtyCol)

	Local $originalQty=0, $newQty=0
	Local $arrayItem_Qty[UBound($inventoryArray)*2][2]
	Local $itemCount = 0, $id=0
	
	

	
	;_ArrayDisplay($arrayItem_Qty)
			
	For $i=0 to UBound($inventoryArray)-1
		If $inventoryArray[$i][$refidCol] <> "" Then
			
			Select 
				
				Case $inventoryArray[$i][$refidCol] = "T_USBMOV9" Or $inventoryArray[$i][$refidCol] = "T_2IN1V9"
					$arrayItem_Qty[$itemCount][0] = "D_MICROUSB"
					$arrayItem_Qty[$itemCount][1] = $inventoryArray[$i][$qtyCol]
					$itemCount +=1
					
					$arrayItem_Qty[$itemCount][0] = "T_USBBK"
					$arrayItem_Qty[$itemCount][1] = $inventoryArray[$i][$qtyCol]
					$itemCount +=1
					
				Case $inventoryArray[$i][$refidCol] = "P_USBMOV9" Or $inventoryArray[$i][$refidCol] = "P_2IN1V9"
					$arrayItem_Qty[$itemCount][0] = "D_MICROUSB"
					$arrayItem_Qty[$itemCount][1] = $inventoryArray[$i][$qtyCol]
					$itemCount +=1
					
					$arrayItem_Qty[$itemCount][0] = "P_USBBK"
					$arrayItem_Qty[$itemCount][1] = $inventoryArray[$i][$qtyCol]
					$itemCount +=1
					
				Case $inventoryArray[$i][$refidCol] = "T_USBSMM300" Or $inventoryArray[$i][$refidCol] = "T_2IN1M300"
					$arrayItem_Qty[$itemCount][0] = "D_SAM300"
					$arrayItem_Qty[$itemCount][1] = $inventoryArray[$i][$qtyCol]
					$itemCount +=1
					
					$arrayItem_Qty[$itemCount][0] = "T_USBBK"
					$arrayItem_Qty[$itemCount][1] = $inventoryArray[$i][$qtyCol]
					$itemCount +=1
				
				Case $inventoryArray[$i][$refidCol] = "P_USBSMM300" Or $inventoryArray[$i][$refidCol] = "P_2IN1M300"
					$arrayItem_Qty[$itemCount][0] = "D_SAM300"
					$arrayItem_Qty[$itemCount][1] = $inventoryArray[$i][$qtyCol]
					$itemCount +=1
					
					$arrayItem_Qty[$itemCount][0] = "P_USBBK"
					$arrayItem_Qty[$itemCount][1] = $inventoryArray[$i][$qtyCol]
					$itemCount +=1

				Case $inventoryArray[$i][$refidCol] = "T_USBAPPLE"
					$arrayItem_Qty[$itemCount][0] = "D_APPLE"
					$arrayItem_Qty[$itemCount][1] = $inventoryArray[$i][$qtyCol]
					$itemCount +=1
					
					$arrayItem_Qty[$itemCount][0] = "T_USBWH"
					$arrayItem_Qty[$itemCount][1] = $inventoryArray[$i][$qtyCol]
					$itemCount +=1
				
				Case $inventoryArray[$i][$refidCol] = "P_USBAPPLE"
					$arrayItem_Qty[$itemCount][0] = "D_APPLE"
					$arrayItem_Qty[$itemCount][1] = $inventoryArray[$i][$qtyCol]
					$itemCount +=1
					
					$arrayItem_Qty[$itemCount][0] = "P_USBWH"
					$arrayItem_Qty[$itemCount][1] = $inventoryArray[$i][$qtyCol]
					$itemCount +=1

				Case $inventoryArray[$i][$refidCol] = "T_21V9N"
					$arrayItem_Qty[$itemCount][0] = "D_V9"
					$arrayItem_Qty[$itemCount][1] = $inventoryArray[$i][$qtyCol]
					$itemCount +=1
					
					$arrayItem_Qty[$itemCount][0] = "T_USBBK"
					$arrayItem_Qty[$itemCount][1] = $inventoryArray[$i][$qtyCol]
					$itemCount +=1

				Case $inventoryArray[$i][$refidCol] = "P_21V9N"
					$arrayItem_Qty[$itemCount][0] = "D_V9"
					$arrayItem_Qty[$itemCount][1] = $inventoryArray[$i][$qtyCol]
					$itemCount +=1
					
					$arrayItem_Qty[$itemCount][0] = "P_USBBK"
					$arrayItem_Qty[$itemCount][1] = $inventoryArray[$i][$qtyCol]
					$itemCount +=1

				Case Else
					$arrayItem_Qty[$itemCount][0] = $inventoryArray[$i][$refidCol]
					$arrayItem_Qty[$itemCount][1] = $inventoryArray[$i][$qtyCol]
					$itemCount += 1
					
			EndSelect
		EndIf
	Next	
	
	For $i=0 To UBound($arrayItem_Qty)-1
		If $arrayItem_Qty[$i][0] <> "" Then
			inventoryUpdateDB($arrayItem_Qty[$i][0], $arrayItem_Qty[$i][1])
		EndIf
	Next
	
EndFunc

Func defineState($state)
	
	$state = StringStripWS($state, 1)
	$state = StringStripWS($state, 2)
	
	If StringInStr("Alabama", $state) <> 0 Or $state = "AL" Then
		$state = "AL"
	ElseIf StringInStr("Alaska", $state) <> 0 Or $state = "AK" Then
		$state = "AK"
	ElseIf StringInStr("Arizona", $state) <> 0 Or $state = "AZ" Then
		$state = "AZ"
	ElseIf StringInStr("Arkansas", $state) <> 0 Or $state = "AR" Then
		$state = "AR"
	ElseIf StringInStr("California", $state) <> 0 Or $state = "CA" Then
		$state = "CA"		
	ElseIf StringInStr("Arkansas", $state) <> 0 Or $state = "AR" Then
		$state = "AR"	
	ElseIf StringInStr("Colorado", $state) <> 0 Or $state = "CO" Then
		$state = "CO"
	ElseIf StringInStr("Connecticut", $state) <> 0 Or $state = "CT" Then
		$state = "CT"
	ElseIf StringInStr("Delaware", $state) <> 0 Or $state = "DE" Then
		$state = "DE"
	ElseIf StringInStr("Florida", $state) <> 0 Or $state = "FL" Then
		$state = "FL"
	ElseIf StringInStr("Georgia", $state) <> 0 Or $state = "GA" Then
		$state = "GA"
	ElseIf StringInStr("Hawaii", $state) <> 0 Or $state = "HI" Then
		$state = "HI"
	ElseIf StringInStr("Idaho", $state) <> 0 Or $state = "ID" Then
		$state = "ID"
	ElseIf StringInStr("Illinois", $state) <> 0 Or $state = "IL" Then
		$state = "IL"
	ElseIf StringInStr("Indiana", $state) <> 0 Or $state = "IN" Then
		$state = "AR"
	ElseIf StringInStr("Iowa", $state) <> 0 Or $state = "IA" Then
		$state = "IA"	
	ElseIf StringInStr("Kansas", $state) <> 0 Or $state = "KS" Then
		$state = "KS"
	ElseIf StringInStr("Kentucky", $state) <> 0 Or $state = "KY" Then
		$state = "KY"
	ElseIf StringInStr("Louisiana", $state) <> 0 Or $state = "LA" Then
		$state = "LA"		
	ElseIf StringInStr("Maine", $state) <> 0 Or $state = "ME" Then
		$state = "ME"		
	ElseIf StringInStr("Maryland", $state) <> 0 Or $state = "MD" Then
		$state = "MD"		
	ElseIf StringInStr("Massachusetts", $state) <> 0 Or $state = "MA" Then
		$state = "MA"		
	ElseIf StringInStr("Michigan", $state) <> 0 Or $state = "MI" Then
		$state = "MI"		
	ElseIf StringInStr("Minnesota", $state) <> 0 Or $state = "MN" Then
		$state = "MN"		
	ElseIf StringInStr("Mississippi", $state) <> 0 Or $state = "MS" Then
		$state = "MS"		
	ElseIf StringInStr("Missouri", $state) <> 0 Or $state = "MO" Then
		$state = "MO"		
	ElseIf StringInStr("Montana", $state) <> 0 Or $state = "MT" Then
		$state = "MT"		
	ElseIf StringInStr("Nebraska", $state) <> 0 Or $state = "NE" Then
		$state = "NE"		
	ElseIf StringInStr("Nevada", $state) <> 0 Or $state = "NV" Then
		$state = "NV"	
	ElseIf StringInStr("New Hampshire", $state) <> 0 Or $state = "NH" Then
		$state = "NH"
	ElseIf StringInStr("New Jersey", $state) <> 0 Or $state = "NJ" Then
		$state = "NJ"
	ElseIf StringInStr("New Mexico", $state) <> 0 Or $state = "NM" Then
		$state = "NM"
	ElseIf StringInStr("New York", $state) <> 0 Or $state = "NY" Then
		$state = "NY"
	ElseIf StringInStr("North Carolina", $state) <> 0 Or $state = "NC" Then
		$state = "NC"
	ElseIf StringInStr("North Dakota", $state) <> 0 Or $state = "ND" Then
		$state = "ND"
	ElseIf StringInStr("Ohio", $state) <> 0 Or $state = "OH" Then
		$state = "OH"
	ElseIf StringInStr("Oklahoma", $state) <> 0 Or $state = "OK" Then
		$state = "OK"
	ElseIf StringInStr("Oregon", $state) <> 0 Or $state = "OR" Then
		$state = "OR"
	ElseIf StringInStr("Pennsylvania", $state) <> 0 Or $state = "PA" Then
		$state = "PA"
	ElseIf StringInStr("Rhode Island", $state) <> 0 Or $state = "RI" Then
		$state = "RI"
	ElseIf StringInStr("South Carolina", $state) <> 0 Or $state = "SC" Then
		$state = "SC"
	ElseIf StringInStr("South Dakota", $state) <> 0 Or $state = "SD" Then
		$state = "SD"
	ElseIf StringInStr("Tennessee", $state) <> 0 Or $state = "TN" Then
		$state = "TN"
	ElseIf StringInStr("Texas", $state) <> 0 Or $state = "TX" Then
		$state = "TX"
	ElseIf StringInStr("Utah", $state) <> 0 Or $state = "UT" Then
		$state = "UT"
	ElseIf StringInStr("Vermont", $state) <> 0 Or $state = "VT" Then
		$state = "VT"
	ElseIf StringInStr("Virginia", $state) <> 0 Or $state = "VA" Then
		$state = "VA"
	ElseIf StringInStr("Washington", $state) <> 0 Or $state = "WA" Then
		$state = "WA"
	ElseIf StringInStr("West Virginia", $state) <> 0 Or $state = "WV" Then
		$state = "WV"
	ElseIf StringInStr("Wisconsin", $state) <> 0 Or $state = "WI" Then
		$state = "WI"
	ElseIf StringInStr("Wyoming", $state) <> 0 Or $state = "WY" Then
		$state = "WY"
	Else
		MsgBox(0, "Error", "No State Found")
	EndIf
	
	Return $state

EndFunc

Func _DelimFile_To_Array2D($s_file, $s_delim = @TAB, $i_max_2d = 0)
    
    Local $s_str = $s_file
    If FileExists($s_str) Then $s_str = FileRead($s_file)
    
    
    Local $i_enum_max = False
    If Int($i_max_2d) < 1 Then
        $i_enum_max = True
        $i_max_2d = 1
    EndIf
    
    Local $a_split = StringSplit(StringStripCR($s_str), @LF)
    Local $a_ret[$a_split[0] + 1][$i_max_2d] = [[$a_split[0]]], $a_delim
    
    For $i = 1 To $a_split[0]
        $a_delim = StringSplit($a_split[$i], $s_delim, 1)
        If $i_enum_max And $i_max_2d < $a_delim[0] Then
            ReDim $a_ret[$a_split[0] + 1][$a_delim[0]]
            $i_max_2d = $a_delim[0]
        EndIf
        For $j = 1 To $a_delim[0]
            $a_ret[$i][$j - 1] = $a_delim[$j]
        Next
    Next
    
    Return $a_ret
EndFunc

Func getSecret()
	
	Global $UserName = "cellular_cg"
	Global $Password = "sql54321"
	Global $Database = "cellular_warehouse"
	Global $MySQLServerName = "192.168.1.15"
	Local $SQLInstance = _MySQLConnect($UserName, $Password, $Database, $MySQLServerName)
	Local $SQLCode_getRecord = ""
	Local $GetContent = ""
	Local $SQLCode_updateRecord = ""		
	
	Local $SQLCode_getRecord = "SELECT * FROM secret WHERE name = 'gmail_shipstream'"
	Local $GetContent = _Query($SQLInstance, $SQLCode_getRecord)

	;Global $emailTo = ""

	With $GetContent
		While Not .EOF
			$gmailShipstreamPasswd = .Fields("value").value
			.MoveNext
		WEnd
	EndWith
						
	;MsgBox(0, "$gmailShipstreamPasswd", $gmailShipstreamPasswd)	

	$SQLCode_getRecord = "SELECT * FROM secret WHERE name = 'email_to'"
	$GetContent = _Query($SQLInstance, $SQLCode_getRecord)
				
	With $GetContent
		While Not .EOF
			$emailTo = .Fields("value").value
			.MoveNext
		WEnd
	EndWith

	;MsgBox(0, "$emailTo", $emailTo)	

	$SQLCode_getRecord = "SELECT * FROM secret WHERE name = 'amazon_user'"
	$GetContent = _Query($SQLInstance, $SQLCode_getRecord)

	With $GetContent
		While Not .EOF
			$amazon_user = .Fields("value").value
			.MoveNext
		WEnd
	EndWith

	;MsgBox(0, "$amazon_user", $amazon_user)

	$SQLCode_getRecord = "SELECT * FROM secret WHERE name = 'amazon_pw'"
	$GetContent = _Query($SQLInstance, $SQLCode_getRecord)

	With $GetContent
		While Not .EOF
			$amazon_pw = .Fields("value").value
			.MoveNext
		WEnd
	EndWith

	;MsgBox(0, "$amazon_pw", $amazon_pw)
	
	$SQLCode_getRecord = "SELECT * FROM secret WHERE name = 'buycom_user'"
	$GetContent = _Query($SQLInstance, $SQLCode_getRecord)

	With $GetContent
		While Not .EOF
			$buycom_user = .Fields("value").value
			.MoveNext
		WEnd
	EndWith

	;MsgBox(0, "$buycom_user", $buycom_user)	
	
	$SQLCode_getRecord = "SELECT * FROM secret WHERE name = 'buycom_pw'"
	$GetContent = _Query($SQLInstance, $SQLCode_getRecord)

	With $GetContent
		While Not .EOF
			$buycom_pw = .Fields("value").value
			.MoveNext
		WEnd
	EndWith

	;MsgBox(0, "$buycom_pw", $buycom_pw)	

	$SQLCode_getRecord = "SELECT * FROM secret WHERE name = 'paypal1_user'"
	$GetContent = _Query($SQLInstance, $SQLCode_getRecord)

	With $GetContent
		While Not .EOF
			$paypal1_user = .Fields("value").value
			.MoveNext
		WEnd
	EndWith

	;MsgBox(0, "$paypal1_user", $paypal1_user)	

	$SQLCode_getRecord = "SELECT * FROM secret WHERE name = 'paypal1_pw'"
	$GetContent = _Query($SQLInstance, $SQLCode_getRecord)

	With $GetContent
		While Not .EOF
			$paypal1_pw = .Fields("value").value
			.MoveNext
		WEnd
	EndWith

	;MsgBox(0, "$paypa1_pw", $paypal1_pw)	
	
	$SQLCode_getRecord = "SELECT * FROM secret WHERE name = 'paypal2_user'"
	$GetContent = _Query($SQLInstance, $SQLCode_getRecord)

	With $GetContent
		While Not .EOF
			$paypal2_user = .Fields("value").value
			.MoveNext
		WEnd
	EndWith

	;MsgBox(0, "$paypal2_user", $paypal2_user)		
	

	$SQLCode_getRecord = "SELECT * FROM secret WHERE name = 'paypal2_pw'"
	$GetContent = _Query($SQLInstance, $SQLCode_getRecord)

	With $GetContent
		While Not .EOF
			$paypal2_pw = .Fields("value").value
			.MoveNext
		WEnd
	EndWith

	;MsgBox(0, "$paypal2_pw", $paypal2_pw)	
	
	$SQLCode_getRecord = "SELECT * FROM secret WHERE name = 'paypal3_user'"
	$GetContent = _Query($SQLInstance, $SQLCode_getRecord)

	With $GetContent
		While Not .EOF
			$paypal3_user = .Fields("value").value
			.MoveNext
		WEnd
	EndWith

	;MsgBox(0, "$paypal3_user", $paypal3_user)		
	

	$SQLCode_getRecord = "SELECT * FROM secret WHERE name = 'paypal3_pw'"
	$GetContent = _Query($SQLInstance, $SQLCode_getRecord)

	With $GetContent
		While Not .EOF
			$paypal3_pw = .Fields("value").value
			.MoveNext
		WEnd
	EndWith

	;MsgBox(0, "$paypal3_pw", $paypal3_pw)		

EndFunc

Func _Log($msg)
	
	$global_log &= $msg
	
	
	;ConsoleWrite("Return=" & FileWriteLine ("Log_" & StringReplace(_NowCalcDate(), "/", "_") & ".txt", $msg ) & @CR)

EndFunc

Func _Write_Log($msg)
	
	ConsoleWrite("Return=" & FileWriteLine ("Log_" & StringReplace(_NowCalcDate(), "/", "_") & ".txt", $msg ) & @CRLF)

EndFunc

Func Logout_PayPal($oIE)

	_IELinkClickByText($oIE, "Log Out")
	_IELoadWait($oIE)
	Sleep(3000)
	;_IEQuit($oIE)

EndFunc

Func Logout_Marketplace($oIE)

	_IELinkClickByText($oIE, "LogOut")
	_IELoadWait($oIE)
	Sleep(3000)
	;_IEQuit($oIE)

EndFunc
;;;;;;;;;;;;; SENDING EMAIL ;;;;;;;;;;;;;;;

Func _sendmail($Subject ,$Body)
;##################################
; Variables
;##################################
$SmtpServer = "smtp.gmail.com"              ; address for the smtp-server to use - REQUIRED
$FromName = "EMD Store"                      ; name from who the email was sent
$FromAddress = "emdcell.shipstream@gmail.com" ; address from where the mail should come
$ToAddress = $emailTo	   ; destination address of the email - REQUIRED
;$Subject = ""                   ; subject from the email - can be anything you want it to be
$AttachFiles = ""                       ; the file(s) you want to attach seperated with a ; (Semicolon) - leave blank if not needed
$CcAddress = ""       ; address for cc - leave blank if not needed
$BccAddress = ""     ; address for bcc - leave blank if not needed
$Importance = "Normal"                  ; Send message priority: "High", "Normal", "Low"
$Username = "emdcell.shipstream"                    ; username for the account used from where the mail gets sent - REQUIRED
$Password = $gmailShipstreamPasswd                 ; password for the account used from where the mail gets sent - REQUIRED
;$IPPort = 25                            ; port used for sending the mail
;$ssl =                             ; enables/disables secure socket layer sending - put to 1 if using httpS
$IPPort=465                          ; GMAIL port used for sending the mail
$ssl=1                               ; GMAILenables/disables secure socket layer sending - put to 1 if using httpS

;##################################
; Script
;##################################
Global $oMyRet[2]
Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
$rc = _INetSmtpMailCom($SmtpServer, $FromName, $FromAddress, $ToAddress, $Subject, $Body, $AttachFiles, $CcAddress, $BccAddress, $Importance, $Username, $Password, $IPPort, $ssl)
If @error Then
    MsgBox(0, "Error sending message", "Error code:" & @error & "  Description:" & $rc)
EndIf
;
; The UDF

EndFunc

Func _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject = "", $as_Body = "", $s_AttachFiles = "", $s_CcAddress = "", $s_BccAddress = "", $s_Importance="Normal", $s_Username = "", $s_Password = "", $IPPort = 25, $ssl = 0)
    Local $objEmail = ObjCreate("CDO.Message")
    $objEmail.From = '"' & $s_FromName & '" <' & $s_FromAddress & '>'
    $objEmail.To = $s_ToAddress
    Local $i_Error = 0
    Local $i_Error_desciption = ""
    If $s_CcAddress <> "" Then $objEmail.Cc = $s_CcAddress
    If $s_BccAddress <> "" Then $objEmail.Bcc = $s_BccAddress
    $objEmail.Subject = $s_Subject
    If StringInStr($as_Body, "<") And StringInStr($as_Body, ">") Then
        $objEmail.HTMLBody = $as_Body
    Else
        $objEmail.Textbody = $as_Body & @CRLF
    EndIf
    If $s_AttachFiles <> "" Then
        Local $S_Files2Attach = StringSplit($s_AttachFiles, ";")
        For $x = 1 To $S_Files2Attach[0]
            $S_Files2Attach[$x] = _PathFull($S_Files2Attach[$x])
;~          ConsoleWrite('@@ Debug : $S_Files2Attach[$x] = ' & $S_Files2Attach[$x] & @LF & '>Error code: ' & @error & @LF) ;### Debug Console
            If FileExists($S_Files2Attach[$x]) Then
                ConsoleWrite('+> File attachment added: ' & $S_Files2Attach[$x] & @LF)
                $objEmail.AddAttachment($S_Files2Attach[$x])
            Else
                ConsoleWrite('!> File not found to attach: ' & $S_Files2Attach[$x] & @LF)
                SetError(1)
                Return 0
            EndIf
        Next
    EndIf
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer
    If Number($IPPort) = 0 then $IPPort = 25
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $IPPort
    ;Authenticated SMTP
    If $s_Username <> "" Then
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
    EndIf
    If $ssl Then
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    EndIf
    ;Update settings
    $objEmail.Configuration.Fields.Update
    ; Set Email Importance
    Switch $s_Importance
        Case "High"
            $objEmail.Fields.Item ("urn:schemas:mailheader:Importance") = "High"
        Case "Normal"
            $objEmail.Fields.Item ("urn:schemas:mailheader:Importance") = "Normal"
        Case "Low"
            $objEmail.Fields.Item ("urn:schemas:mailheader:Importance") = "Low"
    EndSwitch
    $objEmail.Fields.Update
    ; Sent the Message
    $objEmail.Send
    If @error Then
        SetError(2)
        Return $oMyRet[1]
    EndIf
    $objEmail=""
EndFunc   ;==>_INetSmtpMailCom
;
;
; Com Error Handler
Func MyErrFunc()
    $HexNumber = Hex($oMyError.number, 8)
    $oMyRet[0] = $HexNumber
    $oMyRet[1] = StringStripWS($oMyError.description, 3)
    ConsoleWrite("### COM Error !  Number: " & $HexNumber & "   ScriptLine: " & $oMyError.scriptline & "   Description:" & $oMyRet[1] & @LF)
    SetError(1); something to check for when this function returns
    Return
EndFunc   ;==>MyErrFunc

