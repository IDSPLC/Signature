'On Error resume next

Dim Ver
Dim InstalledVer
Dim TagLine
Dim ToDebug
Dim ArrayCountry
dim ArrayExemptDept
dim CheckReg
dim ArrEmailCountry
dim ArrFullSig
dim FullSig
dim managerscript
dim ArrDefaultFax(7,3)
dim ArrDefaultPhone(8,3)
dim arrEnvironText(3,2)
dim arrTagLine1(4,2)
dim arrTagLinkLine(4,2)
dim arrTagLink(3,2)



'********* USER OPTIONS ***************************************************************************************************
'* Increment this tag if you need to update the Signature to a new version
Ver=22

'*****************************************************************************
'* Set this to true if you want to print Marketing Taglines
TagLine=True

'* Use this Array to specify which Country gets the Tagline - Country Only!
ArrayCountry= array( "BE","DE","FR","UK","US","BR")

'* And to exclude a Department - Departments Only!
ArrayExemptDept = array("Directors")
'*****************************************************************************


'* Use this array to set the full signature text to both the New mail and Reply mail option - countries only!
ArrFullSig = array("DE")

'* Use this variable if you want the setting to apply regardless of version - Be careful as it will apply the script to all users!
'* DON'T Change this value, use debug option further down
CheckReg = true

'* this helps to finds where the script is going wrong - uncomment to use
ToDebug=False
'ToDebug=True
'bolOverrideVersion=TRUE
'bolOverrideUsername=TRUE
'strOverrideUsername="CN=Paul Young,OU=Users,OU=Group Finance,OU=IDS Boldon,OU=IDS,DC=idsltd,DC=com"
'ToOverride=True
'Overridecountry="BR"
'OverrideNumbers = True
'CheckReg = false

'* Location in the Registry of the Tagline - ADVANCED!
strRegistryLocation = "HKCU\Software\IDSPLC\Sig\Ver"
strRegistryLocationDate = "HKCU\Software\IDSPLC\Sig\Date"



'* Use this variable if you want the Manager tagline - this is an override so will stop it working for anyone.
bolEnableManagerLine = false

'* ArrEmailCountry = array("DE") - Not Sure What this is for


'******************************************************************************************************************************


'*** Country Codes **************************************
strCountryCodeFrance = "FR"
strCountryCodeUnitedKingdom = "UK"
strCountryCodeUnitedStates = "US"
strCountryCodeGermany = "DE"
strCountryCodeDenmark = "DK"
strCountryCodeBrazil = "BR"
strCountryCodeBelgium = "BE"
'*******************************************************

'* Lang codes ******************************************
strLangCodeFrench = "fr_fr"
strLangCodeBelgium = "fr_be"
strLangCodeBritish = "en_gb"
strLangCodeUS = "en_us"
strLangCodeBrazil = "po_br"
strLangCodeGerman = "de_de"

'** Raw AD Country Codes **************************
strRawCountryCodeUnitedKingdom = "GB"
'**************************************************



'* Array of Default Fax Numbers *************************
'* United Kingdom **
ArrDefaultFax(0,0) = strCountryCodeUnitedKingdom
ArrDefaultFax(0,1) = "NE35 9PD"
ArrDefaultFax(0,2) = "+441915190760"

'* Germany **
ArrDefaultFax(1,0) = strCountryCodeGermany
ArrDefaultFax(1,1) = "60329"
ArrDefaultFax(1,2) = "+49 69 3085 5125"

'* France Pouilly **
ArrDefaultFax(2,0) = strCountryCodeFrance
ArrDefaultFax(2,1) = "21 320"
ArrDefaultFax(2,2) = "+33(0)33.80.90.73.07"

'* France Paris **
ArrDefaultFax(3,0) = strCountryCodeFrance
ArrDefaultFax(3,1) = "75 013"
ArrDefaultFax(3,2) = "+33(0)1.40.77.04.55"

'* United States **
ArrDefaultFax(4,0) = strCountryCodeUnitedStates
ArrDefaultFax(4,1) = "20878"
ArrDefaultFax(4,2) = "+1(301)990-4236"

'* Belgium **
ArrDefaultFax(5,0) = strCountryCodeBelgium
ArrDefaultFax(5,1) = "4000"
ArrDefaultFax(5,2) = "+32(0) 4229 7160"

'* Brazil **
ArrDefaultFax(6,0) = strCountryCodeBrazil
ArrDefaultFax(6,1) = ""
ArrDefaultFax(6,2) = "+55 113 7406 105"
'*************************************************

'* Array of Default Phone Numbers *************************
'* United Kingdom **
ArrDefaultPhone(0,0) = strCountryCodeUnitedKingdom
ArrDefaultPhone(0,1) = "NE35 9PD"
ArrDefaultPhone(0,2) = "+441915190660"

'* Germany **
ArrDefaultPhone(1,0) = strCountryCodeGermany
ArrDefaultPhone(1,1) = "60329"
ArrDefaultPhone(1,2) = "+49 69 3085 5025"

'* France Pouilly **
ArrDefaultPhone(2,0) = strCountryCodeFrance
ArrDefaultPhone(2,1) = "21 320"
ArrDefaultPhone(2,2) = "+33(0)33.80.90.52.52"

'* France Paris **
ArrDefaultPhone(3,0) = strCountryCodeFrance
ArrDefaultPhone(3,1) = "75 013"
ArrDefaultPhone(3,2) = "+33(0)1.40.77.04.50"

'* United States - Maryland **
ArrDefaultPhone(4,0) = strCountryCodeUnitedStates
ArrDefaultPhone(4,1) = "20878"
ArrDefaultPhone(4,2) = "+1(877)852-6210"

'* Belgium **
ArrDefaultPhone(5,0) = strCountryCodeBelgium
ArrDefaultPhone(5,1) = "4000"
ArrDefaultPhone(5,2) = "+32(0)4.252.26.36"

'* Brazil **
ArrDefaultPhone(6,0) = strCountryCodeBrazil
ArrDefaultPhone(6,1) = ""
ArrDefaultPhone(6,2) = "+55 113 7406 100"

'* Denmark **
ArrDefaultPhone(7,0) = strCountryCodeDenmark
ArrDefaultPhone(7,1) = ""
ArrDefaultPhone(7,2) = "+45 44 84 00 91"
'************************************************************************************

'** US TechSupport Numbers *********************************************************
arrCustomPhoneOverrideDepts = array("Tech Services","Field Services")
strCustomPhoneUSTechSupport = "+1(877)852-6190"
'***********************************************************************************

'** Signature Wording ***************************************************************
bolNameTitleWordingBold = True
strNameTitleWordingFont = "Arial"
intNameTitleWordingSize = 9
intNameTitleWordingColour = 15

intTitleWordingSize = 7
strDualTitleSeperator = " / "

bolIDSNameWordingBold = TRUE
bolCoNameWordingBold = FALSE
intCoNamePartZeroWordingColour = intNameTitleWordingColour
strCoNamePartZeroWording = "IDS "
intCoNamePartOneWordingColour = RGB(225,14,73)
strCoNamePartOneWording = "Immuno"
intCoNamePartTwoWordingColour = intNameTitleWordingColour
strCoNamePartTwoWording = "diagnostic Systems"
intCoNamePartThreeWordingColour = intNameTitleWordingColour
strCoNamePartThreeWording = " Deutschland GmbH"
strCoNameSeperatorWording = " | "

intIconColour = intCoNamePartOneWordingColour

bolIconBold = TRUE
strIconPhoneFont = "Arial"
strIconPhoneSourceChar = "T: "
strIconMobileSourceChar = " M: "
strPhoneWordingFont = strNameTitleWordingFont
intPhoneWordingColour = intCoNamePartTwoWordingColour

strExtOpeningText = " (x"
strExtClosingText = ") "

strIconFaxFont = "Arial"
strIconFaxSourceChar = " F: "

strIconWebFont = "Arial"
strIconWebSourceChar = " W: "
strSigWebAddress = "www.idsplc.com"

intRegWordingSize = intTitleWordingSize
intRegWordingColour = intNameTitleWordingColour

strGermanyRegWordingLine1 = "Geschäftsführer | Managing Directors: Dr. Rudolf Schemer, Christopher Yates"
strGermanyRegWordingLine2 = "Sitz der Gesellschaft | Domicile: Frankfurt am Main"
strGermanyRegWordingLine3 = "Registergericht | Court of Registry: Amtsgericht Frankfurt am Main"
strGermanyRegWordingLine4 = "HRB-Nr. | Commercial Register No: 74484"

'************************************************************************************

intEnviroWordingSize = intTitleWordingSize
intEnviroWordingColour = 10   'Green
strEnvironFont = "Webdings"
strEnvironWebSourceChar = ""

arrEnvironText(0,0) = "fr_fr"
arrEnvironText(0,1) = "Avant d'imprimer cet email, pensez à l'environnement"
arrEnvironText(1,0) = "fr_be"
arrEnvironText(1,1) = "Pour protéger l’environnement, n’imprimez pas cet email s’il vous plait"
arrEnvironText(2,0) = "Other"
arrEnvironText(2,1) = "To protect the environment please think twice before printing this email"


'************************************************************************************

'** Phone Number Prefixes ***********************************************************
strPrefixInterNationalPhone = "+"
strPrefixExtPhone = "x"
'************************************************************************************

'** Phone Number Prefix Correction **************************************************
'* French Correction
strPrefixCorrectionFranceRaw = "+33."
strPrefixCorrectionFranceFix = "+33(0)"
'************************************************************************************

'* Signature Names ******************************************************************
strSigNameFull = "Full Signature"
strSigNameReply = "Reply Signature"
'***********************************************************************************

'******************************************************************************************************************************************

'* Marketing Tagline Settings **********************************************************

bolTagLineTitleBold = True
strTagLineFont = strNameTitleWordingFont 			'Currently Arial
intTagLineTitleColour = intNameTitleWordingColour

intTagLineTitleSize = 10
strTagLineTitle = "Expanding Automated Immunoassay Solutions"
intTagLineSize = 9
intTagLineColour = intNameTitleWordingColour
bolTagLineBold = True

arrTagLine1(0,0) = strCountryCodeUnitedKingdom
arrTagLine1(0,1) = "Join us at the AACC 2015 (stand #1425) 28th – 30th July, Atlanta, USA"
arrTagLine1(1,0) = strCountryCodeFrance
arrTagLine1(1,1) = "Join us at the AACC 2015 (stand #1425) 28th – 30th July, Atlanta, USA"
arrTagLine1(2,0) = strCountryCodeGermany
arrTagLine1(2,1) = "Join us at the AACC 2015 (stand #1425) 28th – 30th July, Atlanta, USA"
arrTagLine1(3,0) = strCountryCodeBelgium
arrTagLine1(3,1) = "Join us at the AACC 2015 (stand #1425) 28th – 30th July, Atlanta, USA"

arrTagLinkLine(0,0) = strCountryCodeUnitedKingdom
arrTagLinkLine(0,1) = "Click here for more IDS AACC information"
arrTagLinkLine(1,0) = strCountryCodeFrance
arrTagLinkLine(1,1) = "Cliquez ici pour plus d'informations IDS AACC"
arrTagLinkLine(2,0) = strCountryCodeGermany
arrTagLinkLine(2,1) = "Click here for more IDS AACC information"
arrTagLinkLine(3,0) = strCountryCodeBelgium
arrTagLinkLine(3,1) = "Cliquez ici pour plus d'informations IDS AACC"

strTagLink = "http://www.idsplc.com/join-ids-at-the-aacc-2015-annual-meeting-georgia-world-congress-center-atlanta-usa/"
arrTagLink(0,0) = strCountryCodeUnitedKingdom
arrTagLink(0,1) = strTagLink
arrTagLink(1,0) = strCountryCodeFrance
arrTagLink(1,1) = strTagLink
arrTagLink(2,0) = strCountryCodeGermany
arrTagLink(2,1) = strTagLink
arrTagLink(2,0) = strCountryCodeBelgium
arrTagLink(2,1) = strTagLink

intTagLinkColour = 2
intTagLinkSize = 9

'***************************************************************************************


InstalledVer=ReadReg(strRegistryLocation)
if ToDebug then wscript.echo "Running Program - good luck!"


if CheckReg = false then
	SetSig()
Else
	If InstalledVer < Ver then
		if ToDebug then wscript.echo "Version Upgrade Required: " & Ver & " is newer than InstalledVer: " & InstalledVer

		SetSig()
	Else
		if ToDebug then 
			wscript.echo "Program Complete, Version: " & Ver & " Matches InstalledVer: " & InstalledVer
			if bolOverrideVersion then
				wscript.echo "Overriding Version Check: " & Ver
				SetSig()
			end if
		end if
	End If
End If

Function WriteReg(RegPath, Value, RegType)

      Dim objRegistry, Key
      Set objRegistry = CreateObject("Wscript.shell")

      Key = objRegistry.RegWrite(RegPath, Value, RegType)
      WriteReg = Key
	  
End Function

Function ReadReg(RegPath)
		
      Dim objRegistry, Key
      On Error Resume Next
	  Set objRegistry = CreateObject("Wscript.shell")
      Key = objRegistry.RegRead(RegPath)
      if err.number <> 0 then
		 ReadReg = 0
	else
		ReadReg = Key
	end if
	if ToDebug then wscript.echo "DEBUG: Current Registry value is " & Key
	
End Function



Function SetSig()
	
		if ToDebug then wscript.echo "DEBUG: Setting Signature"
		
		
		if ToOverride then wscript.echo "DEBUG: Overriding values"
		
		
		
	
		Set objSysInfo = CreateObject("ADSystemInfo")

		Set WshShell = CreateObject("WScript.Shell")
	
		strUser = objSysInfo.UserName
		if bolOverrideUsername then
			strUser = strOverrideUsername
		end if
		Set objUser = GetObject("LDAP://" & strUser)

		if ToDebug then wscript.echo "DEBUG: Setting Signature for " & strUser
		
		
		StrName = objUser.displayName
		
		strCountry = objUser.c
		
		if ToOverride and todebug then 
			wscript.echo "DEBUG: Override - Old Value: " & strCountry & " - New Value: " & Overridecountry
			strCountry = Overridecountry
		End if
		
		if strCountry = strCountryCodeFrance then 
			strTitle = Replace(objUser.Description, strDualTitleSeperator, Chr(11))
		else
			strTitle = objUser.title
		end if
		
		if strCountry = strRawCountryCodeUnitedKingdom then
				strCountry = strCountryCodeUnitedKingdom
		end if
		
		
			
		FilteredFullSig = Filter(ArrFullSig,strCountry)
			lengthC = UBound(FilteredFullSig)
			if lengthC = -1 then
				FullSig = False
			elseif lengthC > -1 then
				FullSig = True
			end if
			
			
		strManager = objUser.Manager
		if ToDebug then wscript.echo "DEBUG: Manager value is " & strManager
		
		
		
		strCred = objUser.info
		strStreet = objUser.StreetAddress
		strState = objUser.St
		strDept = objUser.department
		if ToDebug then wscript.echo "DEBUG: Dept value is " & strDept

		
		strLocation = objUser.l
		strPostCode = objUser.PostalCode
		if isArray(objUser.otherTelephone) then
			arrOtherTelephone = objUser.otherTelephone
			for x = lbound(arrOtherTelephone) to ubound(arrOtherTelephone)
				if(left(arrotherTelephone(x),1) = strPrefixInterNationalPhone)then
				'*response.write(" otherTelephone(" & x & ") [" & arrotherTelephone(x) & "] /")#
					
					strPhone = arrotherTelephone(x)
					exit for
				end if
            next
		else
			strPhone = objUser.otherTelephone
        end if
		if ToOverride and OverrideNumbers and todebug then
			wscript.echo "DEBUG: Override Phone #s- Old Value: " & strPhone
			strPhone = ""
		end if
		
		strMobile = objUser.Mobile
		strFax = objUser.FacsimileTelephoneNumber
		if ToOverride and OverrideNumbers then
			wscript.echo "DEBUG: Override Fax #s- Old Value: " & strFax
			strFax = ""
		end if
		
		strEmail = objUser.mail
		strCompany = objUser.company
		strExten = objUser.ipPhone
		strPostCode = objUser.PostalCode
		
		if strCountry = "" then
			MsgBox "Please contact IT, your contact details are incorrect"
			WScript.Quit
		end if
		
		if strCountry = strCountryCodeFrance then 
			strPhone = Replace(strPhone, strPrefixCorrectionFranceRaw, strPrefixCorrectionFranceFix)
			strMobile = Replace(strMobile, strPrefixCorrectionFranceRaw, strPrefixCorrectionFranceFix)
			strFax = Replace(strFax, strPrefixCorrectionFranceRaw, strPrefixCorrectionFranceFix)
		end if
		
		Set objWord = CreateObject("Word.Application")

		Set objDoc = objWord.Documents.Add()
		Set objSelection = objWord.Selection

		Set objEmailOptions = objWord.EmailOptions
		Set objSignatureObject = objEmailOptions.EmailSignature

		Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

			objSelection.Font.bold = bolNameTitleWordingBold
			objSelection.Font.Name = strNameTitleWordingFont
			objSelection.Font.Size = intNameTitleWordingSize
			objSelection.Font.ColorIndex = intNameTitleWordingColour
			objSelection.TypeText StrName & Chr(11)
			objSelection.Font.Size = intTitleWordingSize
			objSelection.TypeText strTitle & Chr(11)
			objSelection.TypeText Chr(11)
			objSelection.Font.bold = bolIDSNameWordingBold
			if strCountry = "DE" then 
				objSelection.Font.ColorIndex = intCoNamePartZeroWordingColour
				objSelection.TypeText strCoNamePartZeroWording
			end if
			objSelection.Font.Color = intCoNamePartOneWordingColour
			objSelection.TypeText strCoNamePartOneWording
			objSelection.Font.ColorIndex = intCoNamePartTwoWordingColour
			objSelection.TypeText strCoNamePartTwoWording 
			if strCountry = "DE" then 
				objSelection.Font.ColorIndex = intCoNamePartThreeWordingColour
				objSelection.TypeText strCoNamePartThreeWording
			end if
			objSelection.Font.bold = bolCoNameWordingBold
			objSelection.TypeText strCoNameSeperatorWording
			
			'***********************************************************************************************************************************
			'** This Section Controls the format of the address, in some countries the postcode proceeds the town, others the reverse is true **
			'***********************************************************************************************************************************
			
			if strCountry = strCountryCodeUnitedKingdom OR strCountry = strCountryCodeUnitedStates then
				objSelection.TypeText strStreet & strCoNameSeperatorWording & _ 
										strLocation & strCoNameSeperatorWording & _ 
										strState & strCoNameSeperatorWording & _ 
										strPostCode & strCoNameSeperatorWording & _
										strCountry 
			
			elseif strCountry = strCountryCodeGermany then 
				objSelection.TypeText strStreet & strCoNameSeperatorWording & _										
										strPostCode & " " & _
										strState & strCoNameSeperatorWording & _
										strCountry
										
			elseif strCountry = strCountryCodeFrance or strCountry = strCountryCodeBelgium then 
				objSelection.TypeText strStreet & strCoNameSeperatorWording & _
										strPostCode & strCoNameSeperatorWording & _ 
										strLocation & strCoNameSeperatorWording & _
										strCountry 				
			else 
				objSelection.TypeText strStreet & strCoNameSeperatorWording & _
										strLocation & strCoNameSeperatorWording & _
										strState & strCoNameSeperatorWording & _
										strPostCode & strCoNameSeperatorWording & _
										strCountry 
			end if
			
			'**************************************************************************************************************************************
			'This section controls the telephone output
			'**************************************************************************************************************************************
			
			objSelection.TypeText Chr(11)
			
			objSelection.Font.bold = bolIconBold
			objSelection.Font.Color = intIconColour
			objSelection.Font.Name = strIconPhoneFont
			objSelection.TypeText strIconPhoneSourceChar
			objSelection.Font.bold = false
			objSelection.Font.ColorIndex = intPhoneWordingColour
			objSelection.Font.Name = strPhoneWordingFont
			if ToDebug then wscript.echo "DEBUG: Phone value is " & strPhone
			if ToDebug then wscript.echo "DEBUG: Country value is " & strCountry
			if strPhone = "" and strCountry = strCountryCodeUnitedKingdom then
				strPhone = ArrDefaultPhone(0,2)
			end if
			if strPhone = "" and strCountry = strCountryCodeFrance and strPostCode = ArrDefaultPhone(2,1) then
				strPhone = ArrDefaultPhone(2,2)
			end if
			if strPhone = "" and strCountry = strCountryCodeFrance then 
				strPhone = ArrDefaultPhone(3,2)
			end if
			if strPhone = "" and strCountry = strCountryCodeBelgium then 
				strPhone = ArrDefaultPhone(5,2)
			end if
			if strPhone = "" and strCountry = strCountryCodeGermany then 
				strPhone = ArrDefaultPhone(1,2)
			end if
			if strPhone = "" and strCountry = strCountryCodeBrazil then 
				strPhone = ArrDefaultPhone(6,2)
			end if
			if ToDebug then wscript.echo "DEBUG: US Phone value is " & ArrDefaultPhone(4,2)
			if ToDebug then wscript.echo "DEBUG: strPhone value is " & strPhone
		
			'if strPhone = "" and strCountry = strCountryCodeUnitedStates then 
			'	if strDept = "Tech Services" OR strDept = "Field Services" then 
			'		strPhone = strCustomPhoneUSTechSupport
			'	else
			'		
			'		strPhone = ArrDefaultPhone(4,2)
			'	end if
			'end if
						
			if strCountry = strCountryCodeUnitedStates then
				if strDept = "Tech Services" OR strDept = "Field Services" then 
					strPhone = strCustomPhoneUSTechSupport
				else 
					if strPhone = "" then
						strPhone = ArrDefaultPhone(4,2)
					end if
				end if
			end if
			
			objSelection.TypeText strPhone
			
			'This section outputs the phone extension			
			if strExten <> "" then 
				if strCountry = strCountryCodeUnitedStates then
					if strDept <> "Tech Services" AND strDept <> "Field Services" then
						objSelection.TypeText strExtOpeningText & strExten & strExtClosingText
					end if
				else objSelection.TypeText strExtOpeningText & strExten & strExtClosingText
				end if
			end if
			
			objSelection.TypeText strCoNameSeperatorWording
			
			'***************************This section outputs the mobile number**************************************
			
			'this section sets the mobile number to blank if user is part of US Tech or Field Services
			if strCountry = strCountryCodeUnitedStates then
					if strDept = "Tech Services" OR strDept = "Field Services" then 
						strMobile = ""
					end if
			end if
			
			if strMobile <> "" then
				objSelection.Font.bold = bolIconBold
				objSelection.Font.Color = intIconColour
				objSelection.Font.Name = strIconPhoneFont
				objSelection.TypeText strIconMobileSourceChar
				objSelection.Font.bold = false
				objSelection.Font.ColorIndex = intPhoneWordingColour
				objSelection.Font.Name = strPhoneWordingFont
				objSelection.TypeText strMobile & strCoNameSeperatorWording
			end if
			
			'*************************This section outputs the fax number**********************************************
			
			objSelection.Font.bold = bolIconBold
			objSelection.Font.Color = intIconColour
			objSelection.Font.Name = strIconFaxFont
			objSelection.TypeText strIconFaxSourceChar
			objSelection.Font.bold = false
			objSelection.Font.ColorIndex = intPhoneWordingColour
			objSelection.Font.Name = strPhoneWordingFont
			
			if strFax = "" and strCountry = strCountryCodeUnitedKingdom then strFax = ArrDefaultFax(0,2)
			if strFax = "" and strPostCode = ArrDefaultFax(2,1)  then strFax = ArrDefaultFax(2,2)
			if strFax = "" and strCountry = strCountryCodeFrance then strFax = ArrDefaultFax(3,2)			
			if strFax = "" and strCountry = strCountryCodeUnitedStates then strFax = ArrDefaultFax(4,2)
			if strFax = "" and strCountry = strCountryCodeGermany then strFax = ArrDefaultFax(1,2)
			if strFax = "" and strCountry = strCountryCodeBrazil then strFax = ArrDefaultFax(6,2)
			if strFax = "" and strCountry = strCountryCodeBelgium then strFax = ArrDefaultFax(5,2)
	
			
			
			objSelection.TypeText strFax & strCoNameSeperatorWording
			
			objSelection.Font.bold = bolIconBold
			objSelection.Font.Color = intIconColour
			objSelection.Font.Name = strIconWebFont
			objSelection.TypeText strIconWebSourceChar
			objSelection.Font.bold = false
			objSelection.Font.ColorIndex = intPhoneWordingColour
			objSelection.Font.Name = strNameTitleWordingFont
			objSelection.TypeText strSigWebAddress
			objSelection.TypeText Chr(11)
			
			'* passes over string and object variables to ContactManager sub
			if bolEnableManagerLine = true then
				if len(strManager) > 1 then
					ContactManager objSelection, objDoc, StrManager, strCountry
					objSelection.TypeText Chr(11)
				end if
			end if
			
			
			
			if strCountry = strCountryCodeGermany then
				'*****************************************************************************************
				'* This Section is required to comply with German Law, the text is above in user options *
				'*****************************************************************************************
				
				objSelection.Font.Size = intRegWordingSize
				objSelection.Font.ColorIndex = intRegWordingColour
				objSelection.TypeText chr(11)
				objSelection.TypeText strGermanyRegWordingLine1
				objSelection.TypeText chr(11)
				objSelection.TypeText strGermanyRegWordingLine2
				objSelection.TypeText chr(11)
				objSelection.TypeText strGermanyRegWordingLine3
				objSelection.TypeText chr(11)
				objSelection.TypeText strGermanyRegWordingLine4
				objSelection.TypeText chr(11)
			end if

			



			'* this is to determine if you are going to need the tagline
			if TagLine = true then
			    '* this filters out any departments that dont need the tagline
				FilteredExemptDept = Filter( ArrayExemptDept,StrDept)
					FilteredExemptDeptlength = UBound(FilteredExemptDept)
						if ToDebug = True then wscript.echo "DEBUG: Filtered Departments Array length: " & FilteredExemptDeptlength
					if FilteredExemptDeptlength = -1 then
						if ToDebug = True then wscript.echo "DEBUG: Looking through the countries that need to apply tag"
						
							'* this decides if your country if correct for the tagline
							FilteredCountry = Filter(ArrayCountry,StrCountry)
								FilteredCountryLength = UBound(FilteredCountry)
								if ToDebug = True then wscript.echo "DEBUG: Department for tag: " & FilteredCountry(0)
									
								if FilteredCountryLength > -1 then
									if ToDebug = true then wscript.echo "DEBUG: Running Marketing Tag Sub Routine"
									
									'* this passes over the objSelection (to create an object) and objDoc( to add a word document) to MarketingTagLine sub routine
									MarketingTagLine objSelection, objDoc, strCountry
										objSelection.TypeText chr(11)
			
										if ToDebug = True then wscript.echo "DEBUG: Completed the Marketing tag Sub Routine"
											
										
								end if
				    end if
			End if
			
			objSelection.Font.Size = intEnviroWordingSize
			objSelection.Font.ColorIndex = intEnviroWordingColour
			objSelection.Font.Name = strEnvironFont
			objSelection.TypeText Chr(11)
			objSelection.TypeText strEnvironWebSourceChar
			objSelection.Font.Name = strNameTitleWordingFont
			
			if strCountry = strCountryCodeFrance then
				objSelection.TypeText arrEnvironText(0,1)
			elseif strCountry = strCountryCodeBelgium then
				objSelection.TypeText arrEnvironText(1,1)
			Else
				objSelection.TypeText arrEnvironText(2,1)
			End if
			
			
		Set objSelection = objDoc.Range()

		objSignatureEntries.Add strSigNameFull, objSelection
		objSignatureObject.NewMessageSignature = strSigNameFull

		objDoc.Saved = True
		objWord.Quit

		Set objWord = CreateObject("Word.Application")

		Set objDoc = objWord.Documents.Add()
		Set objSelection = objWord.Selection

		Set objEmailOptions = objWord.EmailOptions
		Set objSignatureObject = objEmailOptions.EmailSignature

		Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

			if FullSig = false then
				if ToDebug then wscript.echo "DEBUG: Reply Bold Value is " & bolNameTitleWordingBold
				objSelection.Font.bold = bolNameTitleWordingBold
				objSelection.Font.Name = strNameTitleWordingFont
				objSelection.Font.Size = intNameTitleWordingSize
				objSelection.Font.ColorIndex = intNameTitleWordingColour
				objSelection.TypeText StrName & Chr(11)
				objSelection.Font.Size = intTitleWordingSize
				objSelection.TypeText strTitle & Chr(11) & Chr(11)
				objSelection.Font.ColorIndex = intEnviroWordingColour
				objSelection.Font.Name = strEnvironFont
				objSelection.TypeText strEnvironWebSourceChar
				objSelection.Font.Name = strNameTitleWordingFont
				
				if strCountry = strCountryCodeFrance then
					objSelection.TypeText arrEnvironText(0,1)
				elseif strCountry = strCountryCodeBelgium then
					objSelection.TypeText arrEnvironText(1,1)
				Else
					objSelection.TypeText arrEnvironText(2,1)
				End if
			end if

		Set objSelection = objDoc.Range()

		objSignatureEntries.Add strSigNameReply, objSelection

		if FullSig = true then
			objSignatureObject.ReplyMessageSignature = strSigNameFull
			if todebug then	wscript.echo "DEBUG: Applying Full Signature to Reply Msg - as per german law"
			
		elseif FullSig = false then
			objSignatureObject.ReplyMessageSignature = strSigNameReply
			if todebug then	wscript.echo "DEBUG: Applying Reply Signature to Reply Msg"
			
		end if
		
		objDoc.Saved = True
		objWord.Quit
		
		Temp = WriteReg(strRegistryLocation,Ver,"REG_DWORD")
		Temp = WriteReg(strRegistryLocationDate,FormatDateTime(Now, 1),"REG_SZ")
		
		if ToDebug then wscript.echo "Program Complete, Version " & Ver & " saved!"
	
	end Function
	
Sub MarketingTagLine(objSelection,objDoc,strCountry)
    
	'************************** this is the code to run the tagline that is allowed earlier on ***********************
	objSelection.Font.bold = bolTagLineTitleBold
	objSelection.TypeText Chr(11) & Chr(11)
	objSelection.Font.Name = strTagLineFont
	'objSelection.Font.Size = intTagLineTitleSize
	'objSelection.Font.ColorIndex = intTagLineTitleColour
	'objSelection.TypeText strTagLineTitle
	objSelection.Font.Size = intTagLineSize
	objSelection.Font.ColorIndex = intTagLineColour
	objSelection.TypeText Chr(11)
	objSelection.Font.bold = bolTagLineBold
	
	select case strCountry
		case "DE"
			objSelection.TypeText arrTagLine1(2,1)
		case "BE"
			objSelection.TypeText arrTagLine1(1,1)
		case strCountryCodeFrance
			objSelection.TypeText arrTagLine1(1,1)
		case else
			objSelection.TypeText arrTagLine1(0,1)
		End Select
	
	objSelection.TypeText Chr(11)
	
	
	select case strCountry
		case strCountryCodeGermany
			set objLink = objSelection.Hyperlinks.Add(objSelection.Range, arrTagLink(2,1),,,arrTagLinkLine(2,1))
		case strCountryCodeBelgium
			set objLink = objSelection.Hyperlinks.Add(objSelection.Range, arrTagLink(1,1),,,arrTagLinkLine(1,1))
		case strCountryCodeFrance
			set objLink = objSelection.Hyperlinks.Add(objSelection.Range, arrTagLink(1,1),,,arrTagLinkLine(1,1))
		case else
			set objLink = objSelection.Hyperlinks.Add(objSelection.Range, arrTagLink(0,1),,,arrTagLinkLine(0,1))
		End Select
	objLink.Range.Font.ColorIndex = intTagLinkColour
	objLink.Range.Font.Size = intTagLinkSize
	
	objSelection.TypeText Chr(11) & Chr(11)
	
	if ToDebug then wscript.echo "DEBUG: Hyperlinks have been added."
	
		

End Sub

Sub ContactManager (objSelection,objDoc,strManager,strCountry)

	Set objManager = GetObject("LDAP://" & strManager)
	strManagerName = objManager.displayName
	strManagerTitle = objManager.title
	strManagerEmail = objManager.mail
	objSelection.Font.Name = "Arial"
	objSelection.Font.Size = 7
	if ToDebug then wscript.echo "begin case" & strCountry
	
	select case strCountry
		case "DE"
			objSelection.TypeText "Wenn Sie Probleme haben, mailen Sie mein Manager :"
		case "BE"
			objSelection.TypeText "Si vous avez des problèmes, envoyez mon manager :"
		case strCountryCodeFrance
			objSelection.TypeText "Si vous avez des problèmes, envoyez mon manager :"
		case else
			objSelection.TypeText "If you have any problems, email my manager :"
		End Select
	set  objlink = objSelection.Hyperlinks.Add(objSelection.Range,"mailto:" & strManagerEmail,,,strManagerName)
	objLink.Range.Font.Size = 7
	objLink.Range.Font.ColorIndex = 15
	objSelection.TypeText ", " & strManagerTitle

end sub           
