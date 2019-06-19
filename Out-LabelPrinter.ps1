param (
    [Parameter(Mandatory=$True)][string] $TicketNumber
)

[void][System.Reflection.Assembly]::LoadFile($PSScriptRoot+'\Interop.bpac.dll')

$connectionString = "Data Source=DATASOURCE; " +
    "Initial Catalog=DATABASE;" +
    "User Id=SQLLoginUsername;" +
    "Password=SQLLoginPassword"

$Query = @"
SELECT [_TELMASTE_].[SEQUENCE]
    ,[DESCRIPTION]
    ,CONCAT([_CUSTOMER_].[FNAME], ' ', [_CUSTOMER_].[NAME]) as NAME
FROM [TRACKIT].[_SMDBA_].[_TELMASTE_]
LEFT JOIN [_SMDBA_].[_CUSTOMER_] ON [_TELMASTE_].[CLIENT] = [_CUSTOMER_].[SEQUENCE]
WHERE [_TELMASTE_].[SEQUENCE] = $TicketNumber
"@

$connection = new-object system.data.SqlClient.SQLConnection($connectionString)

$command = new-object system.data.sqlclient.sqlcommand($Query,$connection)
$connection.Open()

$adapter = New-Object System.Data.sqlclient.sqlDataAdapter $command
$dataset = New-Object System.Data.DataSet
$adapter.Fill($dataSet) | Out-Null

$connection.Close()
$dataSet.Tables

$Ticket = $dataset.Tables[0]

#$Printers = New-Object bpac.PrinterClass
#$Printers.GetInstalledPrinters()
#$Printers.IsPrinterSupported('Brother QL-580N')
#Exit

$Label = New-Object bpac.DocumentClass

$Filename = $PSScriptRoot+'\TicketLabel.lbx'

If ($Label.Open($Filename)) {
    $Label.GetObject('Title').Text = $Ticket.SEQUENCE
    $Label.GetObject('Name').Text = $Ticket.NAME
    $Label.GetObject('Body').Text = $Ticket.DESCRIPTION

    try {
        # Print Preview:
        #    $Label.Export(4, ($PSScriptRoot+'\test.bmp'), 180)

        $Label.SetPrinter('Brother QL-580N', 0)
        $Label.StartPrint('',0)
        $Label.PrintOut(1, 0)
        $Label.Close()
        $Label.EndPrint()
    } catch {
        Write-Output 'Failed'
        Write-Output $Label.ErrorCode
    }
} Else {
    Write-Output 'Failed to open label file'
}
