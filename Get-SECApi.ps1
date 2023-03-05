# documentation
# https://www.sec.gov/edgar/sec-api-documentation
# https://www.powershellgallery.com/packages/ImportExcel/6.0.0/Content/Export-Excel.ps1
# https://www.powershellgallery.com/packages/ImportExcel/5.3.2/Content/New-ExcelChart.ps1

# should be permanent parameter
[PSCustomObject]$defaultHeader = @{"User-Agent" =  "EZ Financial Services ezadminAAB@ezfinancials.com"} #modify as needed => must comply SEC & Laws
[string]$defaultTaxonomy = "us-gaap" # only tested with us-gaap
[string]$defaultCurrecy = "USD" # only tested with USD
[string]$defaultCIK = "0001640147" #  => https://www.sec.gov/edgar/searchedgar/cik (0001579914 = random CIK as test)

# should be variable parameter
[string]$defaultFileName = "Report" # will append: "-cik000123456"
[int]$defaultAppendDateToFilename = 1 # 0=no 1=DDMMYYYY 2=DDMMYYYY-hhmm // will append: "-20042022-0815"
[string]$defaultTitle = "Title"
[string]$defaultWorksheetName = "Report"
[string]$defaultTableStyle = "Medium6"
[string]$defaultChartStyle = "Line"
[int]$defaultTitleSize = 30
[int]$defaultStartRow = 20
[bool]$defaultdeleteExcelFileBeforeFirstRun = 0 # 0=off 1=on //recommanded:0

# functions

# "The company-concept API returns all the XBRL disclosures from a single company (CIK) and concept (a taxonomy and tag) into a single JSON file"
# get specific tag for specific company
function getCompanyConcept(){

    param(
    # default header by SEC
    [Parameter(Mandatory = $False)]
    [PSCustomObject]$header = $defaultHeader,
    # cik of company
    [Parameter(Mandatory = $False)]
    [string]$cik = $defaultCIK,
    # tested with us-gaap
    [Parameter(Mandatory = $False)]
    [string]$taxonomy = $defaultTaxonomy,
    # statement you're looking for
    [Parameter (Mandatory = $False)]
    [string]$tag = $listOfAllFilings.One
    )
    

# modify tag
    [string]$newTag = $tag + ".json"

# modify cik
    [string]$newCIK = "CIK" + $cik

# request
    [PSCustomObject]$responseGetCompanyConcept = Invoke-RestMethod "https://data.sec.gov/api/xbrl/companyconcept/$newCIK/$taxonomy/$newTag" -Method "GET" -Headers $header

# the end
    return $responseGetCompanyConcept
}

# compare (end date), get most recent (filing date) per date, drop old values
# this helps to filter out re-reports of the same numbers by only keeping the newest one
# idea is: the most recent number should be the most correct one, as "errors by calculation" only get caught afterwards (and rarely to the up-side)
function getUnique{

    param(        
        [Parameter(Mandatory = $True)]
        [PSCustomObject]$inputObject
        )

    # create dummy
    $sortedInputObject = [PSCustomObject]@()
    $tempObject = [PSCustomObject]@()
    $exportObject = [PSCustomObject]@()

    # group and sort
    $sortedInputObject = $inputObject | Group-Object end | sort -Descending Count

    # get newest result per line
    foreach($tempObject in $sortedInputObject){

        # sort values by date
        $tempObject = $tempObject.Group | sort -Descending filed

        # get newest date
        $exportObject += $tempObject[0]

    }

    return $exportObject
}

# take $exportData and export to Excel Sheet
function exportToExcel {
# properties
    param(
        # export Data
        [Parameter(Mandatory = $True)]
            [PSCustomObject]$exportData,
        # default Title
        [Parameter(Mandatory = $False)]
            [string]$exportTitle = $defaultTitle,
        # default Worksheet Name
        [Parameter(Mandatory = $False)]
            [string]$exportWorksheetName = $defaultWorksheetName,
        # default Table Style
        [Parameter(Mandatory = $False)]
            [string]$exportTableStyle = $defaultTableStyle,
        # default Title Size
        [Parameter(Mandatory = $False)]
            [string]$exportTitleSize = $defaultTitleSize,
        # default Start Row
        [Parameter(Mandatory = $False)]
            [int]$exportStartRow = $defaultStartRow,
        # cik of company (for filename)
        [Parameter(Mandatory = $False)]
            [string]$exportFileName = $defaultFileName,
        # get default company cik
            [string]$cik = $defaultCIK,
        [Parameter(Mandatory = $False)]
            [string]$exportChartType = $defaultChartStyle    
         )

# export to excel
    $exportChartData = New-ExcelChartDefinition -XRange end -YRange val -YAxisNumberformat '#,##0 $;-#,##0 $' -ChartType $exportChartType -Title $exportTitle -ChartTrendLine MovingAvgerage -SeriesHeader Price -YAxisTitleText Price -XAxisTitleText Time -Column 0 -Row 0 -Width 790
    Export-Excel $xlTempFile -WorksheetName $exportWorksheetName -TargetData $exportData -StartRow $exportStartRow -Title $exportTitle -TableStyle $exportTableStyle -TitleSize $exportTitleSize -AutoSize -ExcelChartDefinition $exportChartData -AutoNameRange -ClearSheet

}

# settings
## All Tags needed
$listOfAllTags = @{
  "UnamortizedDebtIssuanceExpense" = "Unamortized Debt Issuance Expense"
  "StockholdersEquity" = "Stockholders Equity"
  "ShareBasedCompensation" = "ShareBased Compensation"
  "SellingGeneralAndAdministrativeExpense" = "Selling General And Administrative (Expense)"
  "Revenues" = "Revenue"
  "ResearchAndDevelopmentExpense" = "Research And Development (or Expense)"
  "ProductWarrantyExpense" = "Product Warranty Expense"
  "OtherNonoperatingIncomeExpense" = "Other Nonoperating Income (or Expense)"
  "OtherNoncashIncomeExpense" = "Other Noncash Income (or Expense)"
  "OperatingIncomeLoss" = "Operating Income (or Loss)"
  "OperatingExpenses" = "Operating Expenses"
  "NetIncomeLoss" = "Net Income (or Loss)"
  "NetCashProvidedByUsedInOperatingActivities" = "Net Cash Provided By (Used In) Operating Activities"
  "NetCashProvidedByUsedInContinuingOperations" = "Net Cash Provided By (Used In) Continuing Operations"
  "MarketingAndAdvertisingExpense" = "Marketing And Advertising Expense"
  "LitigationSettlementExpense" = "Litigation Settlement Expense"
  "InterestPaid" = "Interest Paid"
  "InterestExpense" = "Interest Expense"
  "IncomeTaxExpenseBenefit" = "Income Tax Expense (Benefit)"
  "EmployeeServiceShareBasedCompensationAllocationOfRecognizedPeriodCostsCapitalizedAmount" = "Employee Service ShareBased Compensation (Allocation Of) Recognized Period Costs Capitalized Amount"
  "CommonStockValue" = "Common Stock Value"
  "AllocatedShareBasedCompensationExpense" = "Allocated ShareBased Compensation (Expense)"
}

## all Filing Formats
$listOfAllFilings = @{
  "One" = "10-K"
  #"Two" = "10-Q"
  #"Three" = "10-k/A"
  #"Four" = "10-A"
}

# options
# option: remove local file
    if($defaultdeleteExcelFileBeforeFirstRun -eq $True){
    # remove file
        Remove-Item $xlTempFile -ErrorAction SilentlyContinue
    }
    if($defaultdeleteExcelFileBeforeFirstRun -eq $False){

    # do nothing
        continue

    }

# option: append date to filename
    if($defaultAppendDateToFilename -eq 2){
        [string]$xlTempFile = "$PWD\$defaultFileName-" + "cik" + $defaultCIK + "-" + (Get-Date -Format "ddMMyyyy-HHmm") + ".xlsx"
        }
    if($defaultAppendDateToFilename -eq 1){
        [string]$xlTempFile = "$PWD\$defaultFileName-" + "cik" + $defaultCIK + "-" + (Get-Date -Format "ddMMyyyy") + ".xlsx"
        }
    if($defaultAppendDateToFilename -eq 0){
        [string]$xlTempFile = "$PWD\$defaultFileName-" + "cik" + $defaultCIK + ".xlsx"
        }

# Runtime

# create dummy
$resultOfAllTags = [PSCustomObject]{}
$allResultOfAllTags = [PSCustomObject]@()

# request all data
foreach($anyTag in $listOfAllTags.GetEnumerator()){

    # create dummy
    $getCompanyConceptResult = [PSCustomObject]@()

    # get data
    $getCompanyConceptResult = getCompanyConcept -tag ($anyTag.key)

    # add to final object
    Add-Member -InputObject $resultOfAllTags -NotePropertyName ($getCompanyConceptResult.tag) -NotePropertyValue ($getCompanyConceptResult)

    }

# get all results
$allResultOfAllTags = (Get-Member -InputObject $resultOfAllTags -MemberType NoteProperty)

# export everything per page
foreach($oneResultOfAllTags in $allResultOfAllTags){

    # create dummy
    $sortedObject = [PSCustomObject]@()

    # sort data by filing date
    $sortedObject = $resultOfAllTags.($oneResultOfAllTags.Name).units.$defaultCurrecy | Sort-Object -Descending filed

    # filter out unique
    $uniqueResult = [PSCustomObject]@()
    $uniqueResult = getUnique -inputObject $sortedObject

    # sort data by end date ascending (for chart to work well in excel)
    $exportData = ($uniqueResult | Select -Property end, val, accn, fy, fp, form, filed | sort end)

    # shorten Name if necessary
    if($oneResultOfAllTags.Name.Length -ge 21){
        [array]$chars = "abcdefghijkmnopqrstuvwxyzABCEFGHJKLMNPQRSTUVWXYZ1234567890".ToCharArray()
        [string]$randomString = ((Get-Random -InputObject $chars)+(Get-Random -InputObject $chars))
        [string]$oneNameForThisResult = ($oneResultOfAllTags.Name.Substring(0,15)) + "-" + $randomString
        }
    else{
        [string]$oneNameForThisResult = ($oneResultOfAllTags.Name)
        }

    # export to excel
    exportToExcel -exportFileName $xlTempFile -exportTitle ($resultOfAllTags.($oneResultOfAllTags.Name).label) -exportWorksheetName ($oneNameForThisResult) -exportData $exportData

    # done
    Invoke-Expression "Write-Host -ForegroundColor green "ok for:" , ($resultOfAllTags.($oneResultOfAllTags.Name).label)"

    }
