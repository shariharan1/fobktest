# Accept parameters for your script
param(
[Parameter()]
[Int32]$month,

[Parameter()]
[Int32]$year,

[Parameter()]
[bool]$photos
)

# Temporarily set the execution policy for this script
$originalPolicy = Get-ExecutionPolicy

# Set the execution policy to Bypass for the duration of this script only
try {
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction Stop
} catch {
    Write-Host "Failed to set execution policy. Please run PowerShell as Administrator." -ForegroundColor Red
    exit
}

# Main script
try {
	
    # Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

    $winTitle = "FoBK - OP7 Planting - Download Data"

    $originalTitle = $Host.UI.RawUI.WindowTitle
    $Host.UI.RawUI.WindowTitle = $winTitle

    $currFolder = $PWD | select -Expand Path

	$kobo_token = $env:KOBO_TOKEN
	$op7PlantingAssetId = $env:ASSET_ID_OP7PLANT_MONITORING
	$queryBase = '{"$and": [{"today":{"$gte": "START_DATE"}},{"today":{"$lt":"END_DATE"}}]}'
	$op7PlantingURL = $env:KOBO_URL_BASE

	function initialize-globals
	{

        $status = $true

        if ([string]::IsNullOrEmpty($kobo_token)) {
            Write-Host "KOBO Access Token is not set or is empty!!!"
            $status = $false
        } else {
            Write-Host "KOBO Access Token is set. [$kobo_token]"
        }

        if ([string]::IsNullOrEmpty($op7PlantingAssetId)) {
            Write-Host "OP7 Planting Asset ID is NOT Set!!!"
            $status = $false
        } else {
            Write-Host "op7PlantingAssetId is set. [$op7PlantingAssetId]"
        }
		
        if ([string]::IsNullOrEmpty($op7PlantingURL)) {
            Write-Host "KOBO URL is NOT Set!!!"
            $status = $false
        } else {
    		$op7PlantingURL = $op7PlantingURL.replace("!!ASSET_ID!!", $op7PlantingAssetId)
            Write-Host "op7PlantingURL is set. [$op7PlantingURL]"
        }

		$script:headers = @{
			Authorization = "Token " + $kobo_token
		}

        return $status

	}


	function Get-Op7Planting-Row([pscustomobject]$koboObject)
	{

		#Write-Host("in Get-Op7Planting-Row")

		# "fields": [
		#     "today",
		#     "Scientific_name_per_tag",
		#     "Plant_tag_ID",
		#     "Is_the_plant_ALIVE",
		#     "Does_the_plant_appear_HEALTHY",
		#     "Height_of_plant_in_cm",
		#     "Other_observations_of_the_plant_status",
		#     "Photo_of_the_plant_1",
		#     "Photo_of_the_plant_2",
		#     "Photo_of_the_plant_3",
		#     "Observer_s_name",
		#     "_uuid",
		#     "_submission_time"
		# ],

		$item = [pscustomobject]@{
			koboId = $koboObject._id
			observationDate = $koboObject.today
			scientificName = $koboObject.Scientific_name_per_tag
			plantTagId = $koboObject.Plant_tag_ID
			plantAlive = $koboObject.Is_the_plant_ALIVE
			plantHealthy = $koboObject.Does_the_plant_appear_HEALTHY
			plantHeight = $koboObject.Height_of_plant_in_cm
			otherObservations = $koboObject.Other_observations_of_the_plant_status
			observerName = $koboObject.Observer_s_name

			photo1Name = $koboObject.Photo_of_the_plant_1
			photo1Url = ($koboObject._attachments | Where-Object { $_.question_xpath -eq "Photo_of_the_plant_1" } ).download_url
			photo1NewName = ""

			photo2Name = $koboObject.Photo_of_the_plant_2
			photo2Url = ($koboObject._attachments | Where-Object { $_.question_xpath -eq "Photo_of_the_plant_2" } ).download_url
			photo2NewName = ""

			photo3Name = $koboObject.Photo_of_the_plant_3
			photo3Url = ($koboObject._attachments | Where-Object { $_.question_xpath -eq "Photo_of_the_plant_3" } ).download_url
			photo3NewName = ""

			photoCount = 0
			
			status = $koboObject._status
			validationStatus = $koboObject._validation_status
			formVersion = $koboObject.__version__ 
			submissionTime = $koboObject._submission_time 
		}

		$obsDate = $koboObject.today -replace "-", ""
		$photoCount = 0

		if (-not [string]::IsNullOrEmpty($item.photo1Name)) {
			$item.photo1NewName = "$($item.plantTagId)_$($obsDate)_01" + [System.IO.Path]::GetExtension($item.photo1Name)
			$photoCount++
		}
		if (-not [string]::IsNullOrEmpty($item.photo2Name)) {
			$item.photo2NewName = "$($item.plantTagId)_$($obsDate)_02" + [System.IO.Path]::GetExtension($item.photo2Name)
			$photoCount++
		}
		if (-not [string]::IsNullOrEmpty($item.photo3Name)) {
			$item.photo3NewName = "$($item.plantTagId)_$($obsDate)_03" + [System.IO.Path]::GetExtension($item.photo3Name)
			$photoCount++
		}

		$item.photoCount = $photoCount

		return $item
	}

	function Download-And-Process-OP7-Planting-Data([DateTime]$startDt, [DateTime]$endDt)
	{
		Set-Location $op7PlantingDir

		$script:formItem = "OP7Planting"

		$ts = New-TimeSpan -Start $startDt -End $endDt
		$days = 0
		$totalDays = $ts.Days
		if ($totalDays -le 0) { $totalDays = 1 }


		Write-Host "$($formItem) >> Start Date [$($startDt.ToString("yyyy-MM-dd"))] End Date [$($endDt.ToString("yyyy-MM-dd"))]"

		Write-Progress -Id 10 "Working on $($formItem) data" -PercentComplete 0

		$startDtStr = $startDt.ToString("yyyy-MM-dd")
		$endDtStr = $endDt.ToString("yyyy-MM-dd")

		$fullList = New-Object Collections.Generic.List[pscustomobject]
        $fullListFileName = "op7planting-" + $startDt.ToString("yyyyMM") + ".csv"

		$currDate = $startDt
		$logFile = "$($formItem)_log.txt"

		"`r`nStarting $($formItem) .... $(Get-Date) " | Out-File $logFile -Append ascii
		"$($formItem) for Start Date [$($startDt.ToString("yyyy-MM-dd"))] End Date [$($endDt.ToString("yyyy-MM-dd"))]" | Out-File $logFile -Append ascii
		if (-not $downloadPhotos) {
			"Photos will not be downloaded." | Out-File $logFile -Append ascii
		}

		while ($currDate -le $endDt) {

			$currDateStr = $currDate.ToString("yyyy-MM-dd")

			$queryJson = $script:queryBase.Replace("START_DATE", $currDate.ToString("yyyy-MM-dd") ).Replace("END_DATE", $currDate.AddDays(1).ToString("yyyy-MM-dd") )
			$koboUrl = $script:op7PlantingURL.Replace("JSON_QUERY", $queryJson)

			#Write-Host "Query Json [$queryJson]" 
			#Write-Host "Kobo URL [$koboUrl]" 

			Write-Progress -Id 20 -Activity "Downloading..." "Raw data for [$currDateStr] " -PercentComplete -1
			$response = Invoke-RestMethod -Uri $koboUrl -Method Get -Headers $script:headers 
			Write-Progress -Id 20 -Activity "Downloading..." "Raw data for [$currDateStr] " -Completed -PercentComplete 100
            #Write-Progress -Id 20 -Activity "Downloading..." "Raw data for [$currDateStr] " -Completed

			if ($response.count -gt 0) {

				$photoBase = "Photos\"
                #$csvFileName = $currDateStr + "-" + $formItem.ToLower() + ".csv"
				$csvFileName = "$($formItem.ToLower()).csv"

				if (-not $downloadPhotos) {
                    $csvFileName = $currDateStr + "-" + $formItem.ToLower() + "-nophotos.csv"
					$csvFileName = "$($formItem.ToLower())-nophotos.csv"
				}

				Write-Host "Found [$($response.count)] records for [$currDateStr]"
				# "Found [$($response.count)] records for [$currDateStr] Target CSV File [$csvFileName]" | Out-File $logFile -Append ascii 
                "Found [$($response.count)] records for [$currDateStr] " | Out-File $logFile -Append ascii 

				if ( -not( Test-Path -Path $photoBase) ) {
				  $null = New-Item $photoBase -ItemType Directory
				}

				#$list = New-Object Collections.Generic.List[pscustomobject]

				Write-Progress -Id 30 -Activity "Working on [$currDateStr]..." -PercentComplete 0
				
                $totalPhotos = ($response.results | ForEach-Object {$_._attachments.Count} ) | Measure-Object -Sum

                $photosDownloaded = 0

				$recCount = 0
				Foreach ($item in $response.results) { 

					$currRow = Get-Op7Planting-Row $item
					if ($downloadPhotos -and $currRow.photoCount -gt 0) { 

						Write-Progress -Id 30 -Activity "Working on [$currDateStr]..."  "Downloading [$($totalPhotos.Sum)] Photos Remaining [$($totalPhotos.Sum - $photosDownloaded)]" -PercentComplete ( $recCount / $response.count * 100 )
						$phIdx = 1
						
						if (-not [string]::IsNullOrEmpty($currRow.photo1Name)) {
							Write-Progress -Id 35 -Activity "Downloading [$($phIdx)] of [$($currRow.photoCount)] Photos for [$($currRow.plantTagId)]..." -PercentComplete ( ($phIdx-1) / $currRow.photoCount * 100 )
							Invoke-WebRequest -Uri $currRow.photo1Url -Method Get -OutFile ($photoBase + $currRow.photo1NewName) -Headers $headers
                            Write-Progress -Id 35 -Activity "Downloading [$phIdx] of [$($currRow.photoCount)] Photos for [$($currRow.plantTagId)]..." -PercentComplete ( ($phIdx) / $currRow.photoCount * 100 )
                            $phIdx++
                            $photosDownloaded++
                            Write-Progress -Id 30 -Activity "Working on [$currDateStr]..."  "Downloading [$($totalPhotos.Sum)] Photos Remaining [$($totalPhotos.Sum - $photosDownloaded)]" -PercentComplete ( $recCount / $response.count * 100 )
						}
						if (-not [string]::IsNullOrEmpty($currRow.photo2Name)) {
							Write-Progress -Id 35 -Activity "Downloading [$($phIdx)] of [$($currRow.photoCount)] Photos for [$($currRow.plantTagId)]..." -PercentComplete ( ($phIdx-1) / $currRow.photoCount * 100 )
							Invoke-WebRequest -Uri $currRow.photo2Url -Method Get -OutFile ($photoBase + $currRow.photo2NewName) -Headers $headers
                            Write-Progress -Id 35 -Activity "Downloading [$phIdx] of [$($currRow.photoCount)] Photos for [$($currRow.plantTagId)]..." -PercentComplete ( ($phIdx) / $currRow.photoCount * 100 )
                            $phIdx++
                            $photosDownloaded++
                            Write-Progress -Id 30 -Activity "Working on [$currDateStr]..."  "Downloading [$($totalPhotos.Sum)] Photos Remaining [$($totalPhotos.Sum - $photosDownloaded)]" -PercentComplete ( $recCount / $response.count * 100 )
						}
						if (-not [string]::IsNullOrEmpty($currRow.photo3Name)) {
							Write-Progress -Id 35 -Activity "Downloading [$($phIdx)] of [$($currRow.photoCount)] Photos for [$($currRow.plantTagId)]..." -PercentComplete ( ($phIdx-1) / $currRow.photoCount * 100 )
							Invoke-WebRequest -Uri $currRow.photo3Url -Method Get -OutFile ($photoBase + $currRow.photo3NewName) -Headers $headers
							Write-Progress -Id 35 -Activity "Downloading [$phIdx] of [$($currRow.photoCount)] Photos for [$($currRow.plantTagId)]..." -PercentComplete ( ($phIdx) / $currRow.photoCount * 100 )
                            $phIdx++
                            $photosDownloaded++
                            Write-Progress -Id 30 -Activity "Working on [$currDateStr]..."  "Downloading [$($totalPhotos.Sum)] Photos Remaining [$($totalPhotos.Sum - $photosDownloaded)]" -PercentComplete ( $recCount / $response.count * 100 )
						}
						
                        # Write-Progress -Id 30 -Activity "Working on [$currDateStr]..."  "Downloaded [$photosDownloaded] of [$($totalPhotos.Sum)] Photos " -PercentComplete ( $recCount / $response.count * 100 )
					}

                    Write-Progress -Id 35 -Activity "Clearing..." -Completed -PercentComplete 100
					#$list.Add($currRow)
                    $fullList.Add($currRow)

					$recCount++ 

					Write-Progress -Id 30 -Activity "Working on [$currDateStr]..." -PercentComplete ( $recCount / $response.count * 100 )
				}

				Write-Progress -Id 30 -Activity "Working on [$currDateStr]..."  -PercentComplete 100 -Completed 

				# $list | Select-Object observationDate, scientificName, plantTagId, plantAlive, plantHealthy, plantHeight, observerName, otherObservations, photo1NewName, photo2NewName, photo3NewName, koboId | ConvertTo-Csv -NoTypeInformation | Out-File $csvFileName -Encoding Ascii 

			} else {
				"No Records found for [$currDateStr] " | Out-File $logFile -Append ascii 
			}


			$days = $days + 1 
			if ($days -gt $ts.days) { $days = $totalDays }

			Write-Progress -Id 10 "Working on $($formItem) data" -PercentComplete ( $days/$totalDays*100 ) 

			$currDate = $currDate.AddDays(1) 

		}

		if ( $fullList.Count -gt 0 ) {
			# $fullList | Sort-Object -Property SurveyDate, WayPoint | ConvertTo-Csv -NoTypeInformation | Out-File $fullListFileName -Encoding Ascii 
            $fullList | Select-Object observationDate, scientificName, plantTagId, plantAlive, plantHealthy, plantHeight, observerName, otherObservations, photo1NewName, photo2NewName, photo3NewName, koboId | ConvertTo-Csv -NoTypeInformation | Out-File $fullListFileName -Encoding Ascii 
            "CSV file [$fullListFileName] Generated." | Out-File $logFile -Append ascii 
		}

        
		"Completed  $($formItem) .... $(Get-Date)`r`n" | Out-File $logFile -Append ascii
		Write-Progress -Id 10 "$($formItem) data - Completed!" -PercentComplete 100 -Completed
	}

    function Get-UserInputSimple {
        param (
            [string]$PromptMessage = "Enter a value",
            [string]$DefaultValue = ""
        )

        # Display the prompt with the default value in parentheses
        $fullPrompt = "$PromptMessage (default $DefaultValue): "

        # Read user input
        $userInput = Read-Host -Prompt $fullPrompt

        # If no input is provided, use the default value
        if ([string]::IsNullOrWhiteSpace($userInput)) {
            $userInput = $DefaultValue
        }

        return $userInput
    }

    function Get-UserInput {
        param (
            [string]$PromptMessage = "Enter a value",
            [string]$DefaultValue = "",
            [ValidateSet("String", "Number", "Boolean")] [string]$DataType = "String",
            [int]$MinValue = [int]::MinValue,       # Minimum value for numeric inputs
            [int]$MaxValue = [int]::MaxValue        # Maximum value for numeric inputs
        )

        function ConvertTo-Boolean {
            param (
                [string]$inValue
            )
        
            $inValue = $inValue.Trim().ToLower()
        
            switch ($inValue) {
                "true"  { return $true }
                "false" { return $false }
                "y"     { return $true }
                "n"     { return $false }
                "t"     { return $true }
                "f"     { return $false }
                "1"     { return $true }
                "0"     { return $false }
                default { return $null }
            }
        }
    
        function ConvertTo-Number {
            param (
                [string]$inValue
            )
            if ([int]::TryParse($inValue, [ref]$null)) {
                return [int]$inValue
            } elseif ([double]::TryParse($inValue, [ref]$null)) {
                return [double]$inValue
            } else {
                return $null
            }
        }

        do {
            # Display the prompt with the default value in parentheses
            $fullPrompt = "$PromptMessage (default $DefaultValue): "
        
            # Read user input
            $userInput = Read-Host -Prompt $fullPrompt
        
            # If no input is provided, use the default value
            if ([string]::IsNullOrWhiteSpace($userInput)) {
                $userInput = $DefaultValue
            }

            # Validate the input based on the specified data type
            switch ($DataType) {
                "String" {
                    $validInput = $userInput
                }
                "Number" {
                    $validInput = ConvertTo-Number $userInput
                    #Write-Host $validInput
                    if ($null -eq $validInput -or $validInput -lt $MinValue -or $validInput -gt $MaxValue) {
                        Write-Host "Invalid number. Please enter a value between $MinValue and $MaxValue."
                        $validInput = $null
                    }
                }
                "Boolean" {
                    $validInput = ConvertTo-Boolean $userInput
                    if ($null -eq $validInput) {
                        Write-Host "Invalid boolean. Please enter 'true', 'false', 'y', 'n', 't', 'f', '1' or '0'. "
                        $validInput = $null
                    }
                }
            }

        } until ($null -ne $validInput)

        return $validInput
    }

	function start-main()
	{
	 
		$baseDir = $PSScriptRoot 	# Default to the current folder 

		$op7PlantingDir = Join-Path -Path $PSScriptRoot -ChildPath "OP7PlantingData" 
		
		if ( -not( Test-Path -Path $op7PlantingDir) ) {
			$null = New-Item $op7PlantingDir -ItemType Directory
		}

		Set-Location $baseDir

		#
		# Use the $month and $year to determine the Start and End Dates for the month/year 
		$startDt = Get-Date (Get-Date -Hour 0 -Minute 0 -Second 0 -Month $month -Year $year -Day 1)
		$endDt = (Get-Date $startDt.AddMonths(1).AddSeconds(-1))

		# Download OP7 Planting Data
		Download-And-Process-OP7-Planting-Data $startDt $endDt

	}

	$currDate = (Get-Date).AddMonths(-1)  # added for testing
    #$currDate = Get-Date

    Write-Host
    Write-Host $winTitle
    Write-Host

	#
	# If no month provided, default to last month, and last month's Year
	if (-not ($month) ) { 
		$month = $currDate.AddMonths(-1).Month 
		$year = $currDate.AddMonths(-1).Year
	}

	#
	# If no year provided, default to last month's year
	if (-not ($year) ) { 
		$year = $currDate.AddMonths(-1).Year
	}

	if (-not $PSBoundParameters.ContainsKey("photos") ) { 
		$downloadPhotos = $true
	} else {
		$downloadPhotos = $photos
	}

	# $downloadPhotos = $false

	if ( -not (initialize-globals)  ) {
        Write-Host "Environment Variables NOT SET PROPERLY!"
        exit 1      # Environment Variables NOT set properly!!!
    }

	Write-Host "Processing for Month [$month] Year [$year] Photos [$downloadPhotos]"

	#start-main

	Write-Host "Processing for Month [$month] Year [$year] Photos [$downloadPhotos] - COMPLETED!" 

    Write-Host
    Write-Host $winTitle  " - COMPLETED!" 
    Write-Host
    Write-Host


} finally {
    # Restore the original policy
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy $originalPolicy -Force

    Set-Location $currFolder

    $Host.UI.RawUI.WindowTitle = $originalTitle 
}
