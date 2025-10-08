## Enable PS scripting for this user account:
# > Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
# NB: This is *not* sufficient if you downloaded this script from the web, e.g., from Slack or alike!
#
## Tested on the following systems and PS versions (Get-Host), respectively:
# - Win 11: 5.1.22621.3958
# - Win 10: 5.1.19041.1682

$pathPartDB = '..\altium_lib\UZgen\jlcpcb.csv'
$BOMbase    = 'Project Outputs for UZ_A_MAX11331\Rev03\BOM\BOM_JLC-UZ_A_MAX11331(Differential_Input)'


# File paths (internal)
$pathBOMcfg = $BOMbase + '-UZgenPartsMap.csv'
$pathBOM    = $BOMbase + '.csv'
$pathBOMnew = $BOMbase + '-UZgenResolved.csv'
# Row labels (internal)
$rl_lcsc = 'LCSC Part Number'
$rl_mpn = 'Manufacturer Part Number'
$rl_UZgen = 'UZgenMPN'
# CSV extras (internal)
$csv_enc = 'iso-8859-15'
$csv_del = ','
# MPN prefix (internal)
$mpn_prefix_nowarn = 'UZgen_'
$mpn_suffix_parare = '^' + $mpn_prefix_nowarn + '.*Para[RLC]$'



'Importing map table "' + $pathPartDB + '"'
# Read file
$PartDB = Import-Csv -Delimiter $csv_del -LiteralPath $pathPartDB | Where-Object { ($_.'' -ne $rl_lcsc) -and ($_.$rl_mpn -ne '') }
# Retrieve all registered LCSC-MPNs and store to collection
$knownMPNs = $PartDB | Foreach-Object { $_ | Select -ExpandProperty $rl_lcsc }
' Done, found ' + $knownMPNs.count + " LCSC/MPN mappings`n"

'Verifying map table for uniqueness ...'
$duplicates = $PartDB | Group-Object -Property $rl_lcsc | Where-Object { ($_.Count -ne 1) }
if ($duplicates) {
	' ... found duplicates:'
	$duplicates

	"`n"
	'Aborting, please fix the DB and rerun...'

	exit
}
" Done, all good`n"


'Importing part conf "' + $pathBOMcfg + '"'
# Read file
$MapCFG = Import-Csv -Delimiter $csv_del -LiteralPath $pathBOMcfg

# Check whether columns exist as expected
'Verifying header of part conf ...'
if ( !( $MapCFG[0] | Get-Member $rl_uzgen ) -or !( $MapCFG[0] | Get-Member $rl_lcsc ) ) {
	' ... could not read columns, please ensure that the required header (i.e., "UZgenMPN,LCSC Part Number") is at the top of the file'

	''
	'Aborting, please fix the map file and rerun...'

	exit
}
$MapCFG = $MapCFG | Where-Object { ($_.$rl_uzgen -ne '') -and ($_.$rl_lcsc -ne '') }
# Retrieve all registered MPNs and store to collection
$knownMaps = $MapCFG | Foreach-Object { $_ | Select -ExpandProperty $rl_uzgen }
' Done, found ' + $knownMaps.count + " MPNs with mapping`n"


'Importing BOM "' + $pathBOM + '"'
# This version of PS seems to lack configurability w.r.t. UTF8 BOMs (not as in "bill of materials", as one would associate!), great...
# Get-Content only supports "Unknown, String, Unicode, Byte, BigEndianUnicode, UTF8, UTF7, UTF32, Ascii, Default, Oem, BigEndianUTF32"
# Import-Csv only supports Unicode,UTF7,UTF8,ASCII,UTF32,BigEndianUnicode,Default,OEM as -Encoding here, so jump through some hoops...
$BOM = [IO.File]::ReadAllText($pathBOM, [Text.Encoding]::GetEncoding($csv_enc)) | ConvertFrom-Csv -Delimiter $csv_del | Where-Object {$_.Quantity -ne ''}
" Done`n"


$lcscisnull=0
$unresolved=0
$unresolved_lcsc=''
$nongeneric=0
$mpnisempty=0
$mpnhaspara=0
$BOM | Foreach-Object {
	$mpn_generic = $_ | Select -ExpandProperty $rl_mpn
	'Current MPN: ' + $mpn_generic

	$addline=0

	if ( $mpn_generic -match $mpn_suffix_parare ) {
		' WARNING: Part MPN indicates a parametric part (cf. $mpn_suffix_parare), this should not happen!'
		$mpnhaspara++
		$addline=1
	}

	if ( $mpn_generic -in $knownMaps ) {

		$PartMap = $MapCFG | Where-Object { $_.$rl_uzgen -eq $mpn_generic }
		$mpn_real = $PartMap | Select -ExpandProperty $rl_lcsc

		if ( $mpn_real -in $knownMPNs ) {

			' Match, replacing'

			# We rely on our PartDB listing each $rl_lcsc exactly once 
			$newPart = $PartDB | Where-Object { $_.$rl_lcsc -eq $mpn_real }
			$properties = $newPart.PSObject.Properties

			foreach ( $property in $properties ) {
				$property = $property.name

				$value_orig = $_ | Select -ExpandProperty $property
				$value_new = $newPart | Select -ExpandProperty $property

				if ( ( $property -eq $rl_mpn ) -and ( -not $value_orig.StartsWith($mpn_prefix_nowarn) ) ) {
					'  WARNING: Replacing a part with a non-generic MPN - Beware that this might not be what you want!'
					$nongeneric++

					if ( $value_orig -ne $value_new ) {
						'  WARNING: Original and new MPNs are different - Beware even more whether this is what you want!!'
					}

					''
				}

				'  old "' + $property + '" of "' + $value_orig + '" with "' +  $value_new + '"'

				$_.$property = $value_new
			}
		} else {
			' WARNING: Mapped part could not be found in DB (' + $pathPartDB + ') - Leaving it as is, please extend the DB and rerun...'
			$unresolved++
			$unresolved_lcsc += " '" + $mpn_real + "'"
		}

		$addline=1
	} else {
		if ( $mpn_generic -eq '' ) {
			' WARNING: Part with empty MPN - Make sure that the resulting BoM entry yields the part you want!'
			$mpnisempty++
			$addline=1
		}


		$mpn_lcsc = $_ | Select -ExpandProperty $rl_lcsc

		if ( $mpn_lcsc -eq '' ) {
			' WARNING: Unmapped part without LCSC number'
			$lcscisnull++
			$addline=1
		}
	}

	if ( $addline ) {
		''
	}
}
"`n"


'Writing new BOM "' + $pathBOMnew + '"'
$BOMstring = $BOM | ConvertTo-Csv -NoTypeInformation -Delimiter $csv_del
$Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $False
[System.IO.File]::WriteAllLines($pathBOMnew, $BOMstring, $Utf8NoBomEncoding)
" Done`n"


if ( $mpnisempty -ne 0 ) {
	'WARNING: There were ' + $mpnisempty + ' part(s) (of any part type) without an MPN!'
	"`n"
}

if ( $mpnhaspara -ne 0 ) {
	'WARNING: There were ' + $mpnhaspara + ' part(s) that appear to be (un-)parametric!'
	"`n"
}

if ( $lcscisnull -ne 0 ) {
	'WARNING: There were ' + $lcscisnull + ' unmapped part(s) with an empty LCSC field!'
	"`n"
}

if ( $unresolved -ne 0 ) {
	'WARNING: There were ' + $unresolved + ' mapped part(s) that could not be resolved!'
	' Associated ' + $rl_lcsc + 's:' + $unresolved_lcsc
	"`n"
}

if ( $nongeneric -ne 0 ) {
	'WARNING: There were ' + $nongeneric + ' non-generic part(s) that were overridden!'
	"`n"
}