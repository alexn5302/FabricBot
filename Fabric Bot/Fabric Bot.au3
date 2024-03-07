#cs ----------------------------------------------------------------------------;
 Program: Fine Fabric Bot (Fragmentation Method)
 AutoIt Version: 3.3.14.5
 Author:  Alex
#ce ----------------------------------------------------------------------------

#RequireAdmin
#include <ImageSearch2015.au3>
#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>

; -----------------------------------------------[Read me]-------------------------------------------------------
; ******* Before editing, make sure you merge the required data files with your data folder in appdata.	*******	|
; Below are Global variables that have specific coordinates and hotkeys that you need to set yourself to have	|
; this bot working properly. Certain images may need to be taken again due to difference in screens.			|
; Hotkeys used:																									|
; ---------------------------																					|
; | [1] = Tailoring         |																					|
; | [2] = Rain Cast 		|																					|
; |	[3] = Fragmentation		|																					|
; | [4] = Harvest Song      |																					|
; | [I] = Inventory         |																					|
; | [U] = Homestead Window	|																					|
; | [END] = Stop Bot		|																					|
; ---------------------------																					|
; You must be in Auto POV for this to work																		|
; You must do this in your HS with your oven to the LEFT or RIGHT of your entrance								|
; You must turn off "Hightlight New Items" in your settings: Game -> Etc ->										|
; Make sure all bags are tagged except for a large bag where your sewing patterns will go to when unequipped	|
; You must have a Cylinder equipped and preferrably raincast CD boots in an equip switch						|
; If you do not have raincast CD boots Go To: Line 246 and Change the Sleep Delay to (23000)	                |
; Lasty, Compile the Fine Fabric Bot and open from the .exe file.												|
; ---------------------------------------------------------------------------------------------------------------

Global $x = 0, $y = 0							 ; Leave at 0
Global $RC = 0                               	 ; Reset RC Timer
Global $harvest = 0 							 ; Reset Harvest Timer - Only use this if not dan 3 tailor!!
Global $skillcancel = "{SPACE}"					 ; Cancel Skill
Global $char = "Lines"							 ; Set your IGN
Global $harvesttimer = 150000					 ; Harvest Duration. Keep at 150000
Global $RCtimer = 170000						 ; Duration of your Rain Cast. Default is 90000
Global $taskbar_x = 1233, $taskbar_y = 1067		 ; Coordinate to click your desktop taskbar (Make sure no opened programs are blocking)
Global $eqswitch = "{Y}"					 	 ; Hotkey to Eq. Switch to Cylinder
Global $wepswap = "{W}"						 	 ; Hotkey to Weapon swap
Global $fabric_x = 241, $fabric_y = 83       	 ; Coordinate for fabric bag
Global $thin_x = 336, $thin_y = 87           	 ; Coordinate for thin thread bag
Global $tail_int_x = 65, $tail_int_y = 583       ; Coordinate for tailoring slot
Global $tail_start_x = 153, $tail_start_y = 702  ; Coordinate for tailoring start button
Global $fbutton_x = 183, $fbutton_y = 281	 	 ; Coordinates for "Fragmentation" tab
Global $frag_int_x = 82, $frag_int_y = 401 		 ; Coordinates for dry oven slot
Global $frag_start_x = 182, $frag_start_y = 707  ; Coordinates for fragmentation start button
Global $color = 0x200C0C                     	 ; Color of broken item equipped
Global $hs_x = 1022, $hs_y = 952			 	 ; Coords for HS Entry Button

; Coordinates for kit and cylinder equipment slots (right hand)
Global $rhand_left = 1530, $rhand_top = 817, $rhand_right = 1575, $rhand_bot = 893

; Coordinates for pattern and guard equipment slots (left hand)
Global $lhand_left = 1630, $lhand_top = 819, $lhand_right = 1677, $lhand_bot = 893

; Coordinates for entire glove bag area
Global $bag_left = 748, $bag_top = 600, $bag_right = 1252, $bag_bottom = 963

;Coordinates for entire tailoring kit bag area
Global $kitbag_left = 422, $kitbag_top = 40, $kitbag_right = 583, $kitbag_bottom = 207

;Coordinates for entire sewing pattern bag area
Global $pattbag_left = 0, $pattbag_top = 316, $pattbag_right = 630, $pattbag_bottom = 1032

; Coordinates for last glove slot in bag
Global $bin_left = 1195, $bin_top = 910, $bin_right = 1256, $bin_bottom = 966

; Coordinates for tailoring mini game (does not need to be specific)
Global $g_start_x1 = 112, $g_start_y1 = 536  ; First pair
Global $g_start_x2 = 112, $g_start_y2 = 536  ; Second pair
Global $g_start_x3 = 112, $g_start_y3 = 536  ; Third pair

Global $g_end_x1 = 178, $g_end_y1 = 511 	 ; First pair
Global $g_end_x2 = 178, $g_end_y2 = 511		 ; Second Pair
Global $g_end_x3 = 178, $g_end_y3 = 511		 ; Third Pair

; Images used
Global $fab = 'fabric.png'                   ; Image for the word "fabric"
Global $thin = 'thin.png'				     ; Image for the word "thin"
Global $sewing = 'sewing.png' 				 ; Image for the word "sewing"
Global $pattern = 'pattern.png'				 ; Image for Pattern
Global $glove = 'glove.png'		             ; Image for glove
Global $tkit = 'tkit.png'				     ; Image for Tailoring Kit
Global $oven = 'oven.png'					 ; Image for Dry oven
Global $magnet = 'magnet.png'				 ; Image for item magnet
Global $drop = 'drop.png'					 ; Image for Drop button
Global $exit = 'exit.png'					 ; Image for Exit Button
Global $exit2 = 'exit2.png'					 ; Image for Exit Confirmation Button
Global $cancel = 'cancel.png'				 ; Image for Cancel Button
Global $close = 'close.png'					 ; Image for Close Button
Global $craft = 'craft.png'			 		 ; Image for active Tailor/Frag window
Global $out = 'out.png'						 ; Image for empty fabric bag
Global $start = 'start.png'					 ; Image for active start button
Global $notice = 'notice.png'				 ; Image for notice window
Global $silk = 'Silk.png'					 ; Image for the word "silk" to confirm glove is successfully placed
Global $fuckattendance = 'fuckattendance.png' ; Image for attendance X button
AutoItSetOption("MouseCoordMode", 1) ; Set coordinates relative to active window
HotKeySet("{END}", "Terminate")
Func Terminate()
	MsgBox(0, "Notice", "Stopped Fabric Bot")
    Exit
EndFunc

WinActivate($char)
WinWait($char)

While (1)
    Sleep(1000)
    MakeGloves()
    ;Raincast()
    FragmentGloves()
WEnd

Func MakeGloves()
	; Searches for full bag
	Local $bag = _ImageSearchArea($glove, 1, $bin_left, $bin_top, $bin_right, $bin_bottom, $x, $y, 70)
	Sleep(200)
    While $bag = 0
		Sleep(300)
		Local $emptybag = _ImageSearchArea($out, 1, 209, 34, 265, 114, $x, $y, 5)
		If $emptybag = 1 Then
			Local $restock = _ImageSearch($magnet, 1, $x, $y, 30)
			Sleep(500)
			MouseMove($x, $y, 2)
			Sleep(300)
			MouseClick("Left")
		EndIf

		; Check if both kit and pattern are broken
		Local $nodura = Pixelsearch($lhand_left, $lhand_top, $lhand_right, $lhand_bot, $color)
		Local $kit = Pixelsearch($rhand_left, $rhand_top, $rhand_right, $rhand_bot, $color)
		Sleep(200)
		If IsArray($kit) and IsArray($nodura) = 1 Then
			Sleep(500)
			DropPattern()
			;DropKit()
			GrabNewKit()
			GrabNewPattern()
		EndIf

		; Check for broken pattern
		If IsArray($nodura) = 1 Then
			Sleep(500)
			DropPattern()    		 ; Right Clicks pattern and drops it
			GrabNewPattern() 		 ; PixelSearches for pattern in bag and ctrl + click to equip
		EndIf

        ; Check for broken kit
		If IsArray($kit) = 1 Then
			Sleep(500)
			;DropKit()           	; Drops pattern and Kit
			GrabNewKit()        	; PixelSearches for kit in bag and ctrl + click to equip
			GrabNewPattern()    	; PixelSearches for pattern in bag and ctrl + click to equip
		EndIf

		; Hotkey for tailoring skill
        Send("{1}")
        Sleep(500)

		; Check for kit and pattern being in wrong hand
		Local $nokp = _ImageSearch($notice, 1, $x,$y, 5)
		If $nokp = 1 Then
			Do
			MouseMove(1058, 605, 3)
			Sleep(500)
			MouseClick("Left")       ; Close notice window
			Sleep(500)
			Local $nokp = _ImageSearch($notice, 1, $x,$y, 5)
			Until $nokp = 0
			Send($wepswap)		     ; Switch back to Kit & Pattern
			Sleep(500)
			Send("{1}")		         ; Re-use Tailoring to continue craft
		EndIf

		Sleep(300)
		Local $craftwindow = _ImageSearch($craft, 1, $x, $y, 10) ; Confirms that Tailoring window is present
		If $craftwindow = 1 Then
			; Searches for "fabric"
			Local $fabric = _ImageSearch($fab, 1, $x, $y, 1)
			Sleep(100)
			If $fabric = 1 Then
				InsertFabric()		; Inserts Fine Fabric
			EndIf

			Local $thinthread = _ImageSearch($thin, 1, $x, $y, 1)
			Sleep(100)
			If $thinthread = 1 Then
				InsertThin() 		; Inserts thin thread
				TailorGame() 		; Minigame to create gloves
			EndIf
		EndIf

		; Searches for full bag
		Sleep(500)
		Local $bag = _ImageSearchArea($glove, 1, $bin_left, $bin_top, $bin_right, $bin_bottom, $x, $y, 70)
	WEnd

	If $bag = 1 Then
		Sleep(200)
		Local $cbutton = _ImageSearch($cancel, 1, $x, $y, 30)
		If $cbutton = 1 Then
			Sleep(200)
			MouseMove($x, $y, 3)	; Moves to Cancel Button
			Sleep(500)
			MouseClick("Left")		; Clicks Cancel
		EndIf
		Sleep(1000)
		Local $restock = _ImageSearch($magnet, 1, $x, $y, 30)
		If $restock = 1 Then
			MouseMove($x, $y, 3)
			Sleep(500)
			MouseClick("Left")
			Sleep(300)
		EndIf
		Send($wepswap)          ; Weapon swaps to empty hand for fragmentation
	EndIf
EndFunc

Func Raincast()
    Sleep(700)
	MouseMove(234, 243, 1)
	Sleep(250)
	MouseClick("left")
	Sleep(300)
	MouseMove(1570, 844, 1)
	Sleep(300)
	MouseClick("Left")
    Sleep(700)
    Send("{2}")          	; Hotkey for Raincast
    Sleep(5100)             ; Sleep for Charge + Raincast animation
	$RC = 0
    $RC = TimerInIt()    	; Begins time for RC Duration
	Sleep(1500)
	MouseClick("Left")
	Sleep(500)
	MouseMove(234, 243, 1)
	Sleep(240)
	MouseClick("left")
	Sleep(400)
EndFunc

Func FragmentGloves()
	Local $glovebag = _ImageSearchArea($glove, 1, $bag_left, $bag_top, $bag_right, $bag_bottom, $x, $y, 70)
	Sleep(500)
    While $glovebag = 1
		Sleep(100)
		FindOven()

		#cs
		;Tracks RC Duration Time
        Local $RCend = TimerDiff($RC)
		If $RCend >= $RCtimer Then
			Sleep(2300)
			RainCast()
		EndIf
		#ce

		Local $craftwindow = _ImageSearch($craft, 1, $x, $y, 30)
		Sleep(300)
		If $craftwindow = 1 Then
			FragGlove()    ; Actually frags silk gloves
		EndIf

		Sleep(300)
		Local $glovebag = _ImageSearchArea($glove, 1, $bag_left, $bag_top, $bag_right, $bag_bottom, $x, $y, 70)
    WEnd

	If $glovebag = 0 Then
		Do
		Local $cbutton = _ImageSearch($cancel, 1, $x, $y, 30)
		If $cbutton = 1 Then
			MouseMove($x, $y, 3)	; Moves to Cancel Button
			Sleep(500)
			MouseClick("Left")		; Clicks Cancel
		EndIf
		Until $cbutton = 0
	EndIf

	Sleep(500)
	Local $restock = _ImageSearch($magnet, 1, $x, $y, 30)
	If $restock = 1 Then
		MouseMove($x, $y, 3)
		Sleep(500)
		MouseClick("Left")
		Sleep(800)
		Send($wepswap)
	EndIf
EndFunc

Func FindOven()
	Sleep(200)
	Send("{3}")
	Local $nokp = _ImageSearch($close, 1, $x,$y, 30)
	Sleep(200)
	If $nokp = 1 Then
		MouseMove($x, $y, 3)
		Sleep(500)
		MouseClick("Left")       ; Close notice window
		Sleep(500)
		Send($wepswap)		     ; Switch back to Empty Hands
		Sleep(500)
	EndIf
	#cs
	Local $findoven = _ImageSearch($oven, 1, $x, $y, 30)
	If $findoven = 1 Then
		Send("{3}")
	EndIf

		Sleep(200)
		Local $craftwindow = _ImageSearch($craft, 1, $x, $y, 5) ; Confirms that Fragmentation window is present
		If $craftwindow = 0 Then
			Local $findoven = _ImageSearch($oven, 1, $x, $y, 30)
			Sleep(200)
			If $findoven = 1 Then
				MouseMove($x, $y, 2)
				Sleep(300)
				MouseClick("Left")
			Else

		EndIf
	EndIf
	#ce
EndFunc

Func FragGlove()
	; Searches for existing gloves in ENTIRE bag
    Local $glovebag = _ImageSearchArea($glove, 1, $bag_left, $bag_top, $bag_right, $bag_bottom, $x, $y, 70)
	Sleep(150)
	If $glovebag = 1 Then
		MouseMove($x, $y, 1)                    ; Move to glove
		Sleep(300)
		MouseClick("Left")                      ; Pick up glove
		Sleep(300)
		MouseMove($frag_int_x, $frag_int_y, 1)  ; Coords for Fragmentation interface window
		Sleep(300)
		MouseClick("Left")                      ; Places glove in slot
		Sleep(500)
		Local $confirm = _ImageSearch($silk, 1, $x, $y, 20)
		If $confirm = 1 Then
			MouseClick("Left", $frag_start_x, $frag_start_y, 1, 2)
			Sleep(5100)                         ; Duration for Fragmentation Animation
		Else
			Local $craftwindow = _ImageSearch($craft, 1, $x, $y, 30)
			Sleep(300)
			If $craftwindow = 1 Then
				MouseClick("Right") 				; In case currently holding a glove - force it back to original slot
				Sleep(300)
				FragGlove()    ; Actually frags silk gloves
			EndIf
		EndIf
	EndIf
EndFunc


Func ExitHS()
	Send("{I}")		; Close Inventory to find Exit Button
	Sleep(1000)
	Local $findexit = _ImageSearch($exit, 1, $x, $y, 30)
	Sleep(200)
	If $findexit = 1 Then
		MouseMove($x, $y, 3)
		Sleep(500)
		MouseClick("Left", $x, $y, 2, 3)
		Sleep(500)
		Local $findpattern = _ImageSearchArea($sewing, 1, 831, 555, 960, 585, $x, $y, 30)
		If $findpattern = 1 Then
			Local $findexit2 = _ImageSearch($exit2, 1, $x, $y, 30)
			Sleep(1000)
			If $findexit2 = 1 Then
				MouseMove($x, $y, 3)
				Sleep(500)
				MouseClick("Left")
				Sleep(63000)
				EnterHS()
			EndIf
		Else
		Sleep(700)
		Local $cbutton = _ImageSearch($cancel, 1, $x, $y, 30)
		If $cbutton = 1 Then
			Sleep(200)
			MouseMove($x, $y, 3)	; Moves to Cancel Button
			Sleep(500)
			MouseClick("Left")		; Clicks Cancel
			Sleep(500)
		EndIf
			Send("{I}")
			Sleep(1000)
			Send($wepswap)
			Sleep(500)
		EndIf
	EndIf
EndFunc

Func EnterHS()
	Send("{U}")					; Open Homestead Window
	Sleep(500)
	MouseMove($hs_x, $hs_y, 3)	; Move to HS Entry button
	Sleep(500)
	MouseClick("Left")			; Click to enter
	Sleep(7000)				    ; Duration to enter HS with generous timing
	Send("{I}")					; Re-open inventory
	Sleep(1000)
	Send($wepswap)
	Sleep(500)
EndFunc

Func InsertFabric()
	Do
		MouseMove($fabric_x, $fabric_y, 1)
		Sleep(150)
		MouseClick("Left")                         		; Pick up Fabric Bag
		Sleep(150)
		MouseMove($tail_int_x, $tail_int_y, 1)       	; Move to Tailoring Slot
		Sleep(150)
		MouseClick("Left")                         		; Place in slot
		Sleep(400)
		Local $startbutton = _ImageSearch($start, 1, $x, $y, 5)
		Sleep(300)
	Until $startbutton = 1

	If $startbutton = 1 Then
		MouseMove($tail_start_x, $tail_start_y, 1)   	; Move to "Start" Button
		Sleep(200)
		MouseClick("Left")
		Sleep(4000)                                   	; Sleep for timed duration of tailoring
	EndIf
EndFunc

Func InsertThin()
	Do
		MouseMove($thin_x, $thin_y, 1)
		Sleep(150)
		MouseClick("Left")                         		; Pick up Thin Thread Bag
		Sleep(150)
		MouseMove($tail_int_x, $tail_int_y, 1)       	; Move to Tailoring Slot
		Sleep(150)
		MouseClick("Left")                        		; Place in slot
		Sleep(400)
		Local $startbutton = _ImageSearch($start, 1, $x, $y, 5)
		Sleep(300)
	Until $startbutton = 1

	If $startbutton = 1 Then
		MouseMove($tail_start_x, $tail_start_y, 1)  	; Move to "Start" Button
		Sleep(200)
		MouseClick("Left")
		Sleep(500)
	EndIf
EndFunc

Func TailorGame()
	Sleep(50)
    MouseMove($g_start_x1, $g_start_y1, 0)      ; Move Mouse to START Position
    Sleep(50)
    MouseDown($MOUSE_CLICK_LEFT)                ; Mouse Click "Down" for 1st START position
    Sleep(50)
    MouseMove($g_end_x1, $g_end_y1, 0)          ; Move Mouse to END Position
    Sleep(50)
    MouseUp($MOUSE_CLICK_LEFT)                  ; Mouse Click "UP" for 1st END position
    Sleep(50)

    MouseMove($g_start_x2, $g_start_y2, 0)      ; Move Mouse to START Position
    Sleep(50)
    MouseDown($MOUSE_CLICK_LEFT)                ; Mouse Click "Down" for 2nd START position
    Sleep(50)
    MouseMove($g_end_x2, $g_end_y2, 0)          ; Move Mouse to END Position
    Sleep(50)
    MouseUp($MOUSE_CLICK_LEFT)                  ; Mouse Click "UP" for 2nd END position
    Sleep(50)

    MouseMove($g_start_x3, $g_start_y3, 0)      ; Move Mouse to START Position
    Sleep(50)
    MouseDown($MOUSE_CLICK_LEFT)                ; Mouse Click "Down" for 3rd START position
    Sleep(50)
    MouseMove($g_end_x3, $g_end_y3, 0)          ; Move Mouse to END Position
    Sleep(50)
    MouseUp($MOUSE_CLICK_LEFT)                  ; Mouse Click "UP" for 3rd END position
    Sleep(50)
    MouseClick("Left")                          ; Clicks interface to complete minigame
	MouseClick("Left")
	Sleep(300)									; Delay before error check
	Local $gamefail = _ImageSearch($craft, 1, $x, $y, 30)
	Sleep(500)
	If $gamefail = 1 Then					    ; If script fails to click points,
		Sleep(10000)							    ; Sleep for 7s for minigame to end and craft glove
	EndIf
	Sleep(3500)									; Duration for Tailoring Animation
EndFunc

Func DropPattern()
	Sleep(500)
    Local $pdrop = Pixelsearch($lhand_left, $lhand_top, $lhand_right, $lhand_bot, $color)
	Sleep(500)
    If isArray($pdrop) Then
        MouseMove($pdrop[0] + 14, $pdrop[1] + 26, 2)    ; Moves to Broken Pattern
        Sleep(500)
        MouseClick("Right")                   			; Right Clicks Broken Pattern
		MouseClick("Right")
        Sleep(500)
		MouseMove(1808, 568, 3)				  			; Move mouse to not cover "Drop Button"
		Local $bdrop = _ImageSearch($drop, 1, $x, $y, 30)
		Sleep(1000)
		If $bdrop = 1 Then
			MouseMove($x, $y, 3)              			; Moves to "Drop" Button
			Sleep(500)
			MouseClick("Left")                			; Clicks Drop
			Sleep(500)
		EndIf
		Local $nokp = _ImageSearch($close, 1, $x,$y, 5)
		Sleep(200)
		If $nokp = 1 Then
			MouseMove($x, $y, 3)
			Sleep(500)
			MouseClick("Left")       ; Close notice window
			Sleep(500)
		EndIf
    EndIf
EndFunc

Func DropKit()
	Sleep(500)
    Local $kdrop = Pixelsearch($rhand_left, $rhand_top, $rhand_right, $rhand_bot, $color)
	Sleep(500)
    If isArray($kdrop) Then
        MouseMove($kdrop[0] + 14, $kdrop[1] + 26, 2)    ; Moves to Broken Kit
        Sleep(500)
        MouseClick("Right")                   			; Right Clicks Broken Kit
		MouseClick("Right")
        Sleep(500)
		MouseMove(1808, 568, 3)				  			; Move mouse to not cover "Drop Button"
		Local $bdrop = _ImageSearch($drop, 1, $x, $y, 30)
		Sleep(1000)
       If $bdrop = 1 Then
			MouseMove($x, $y, 3)              			; Moves to "Drop" Button
			Sleep(500)
			MouseClick("Left")                			; Clicks Drop
			Sleep(500)
	   EndIf
	   Local $nokp = _ImageSearch($close, 1, $x,$y, 5)
	   Sleep(500)
	   If $nokp = 1 Then
			MouseMove($x, $y, 3)
			Sleep(500)
			MouseClick("Left")       ; Close notice window
			Sleep(500)
	   EndIf
    EndIf
EndFunc

Func GrabNewKit()
	Sleep(500)
	Local $newkit = _ImageSearchArea($tkit, 1, $kitbag_left, $kitbag_top, $kitbag_right, $kitbag_bottom, $x, $y, 30)
	Sleep(500)
	If $newkit = 1 Then
		MouseMove($x, $y, 2)				  ; Move to Tailoring Kit
		Sleep(350)
		Send("{LCTRL down}")				  ; Ctrl + Click to Equip
		Sleep(500)
		MouseClick("Left")
		Sleep(300)
		Send("{LCTRL up}")
		Sleep(500)
		Send("{LCTRL up}")
		Sleep(200)
		MouseMove($taskbar_x, $taskbar_y, 3)  ; In case Ctrl gets stuck
		Sleep(300)
		MouseClick("left")
		Sleep(300)
	Else
		MsgBox(0, "Notice", "Ran out of Kits. Exiting Bot!")
		Exit								  ; This means you ran out of kits and need to restock. Script will end.
	EndIf
	WinActivate($char)
	WinWait($char)
EndFunc

Func GrabNewPattern()
	Sleep(500)
	Local $newpattern = _ImageSearchArea($pattern, 1, $pattbag_left, $pattbag_top, $pattbag_right, $pattbag_bottom, $x, $y, 30)
	Sleep(500)
	If $newpattern = 1 Then
		MouseMove($x, $y, 3)				  ; Move to Sewing Pattern Bag
		Sleep(350)
		Send("{LCTRL down}")				  ; Ctrl + Click to Equip
		Sleep(500)
		MouseClick("Left")
		Sleep(300)
		Send("{LCTRL up}")
		Sleep(500)
		Send("{LCTRL up}")
		Sleep(200)
		MouseMove($taskbar_x, $taskbar_y, 3) ; In case Ctrl gets stuck
		Sleep(300)
		MouseClick("left")
		Sleep(300)
	Else
		;Raincast()
		FragmentGloves()
		MsgBox(0, "Notice", "Ran out of Patterns. Exiting Bot!")
		Exit								; This means you ran out of sewing patterns and need to restock. Script will end.
	EndIf
	WinActivate($char)
	WinWait($char)
EndFunc
