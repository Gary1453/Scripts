
<#
function beginForm
{
    Add-Type -AssemblyName System.Windows.Forms

    $Form = New-Object system.Windows.Forms.Form
    $Form.StartPosition = "CenterScreen"
    $Form.Text = "BGCheck Archive"

    $Form.Width = 500
    $Form.Height = 500 

    #Defino la caja de texto
    $TextBox = New-Object System.Windows.Forms.TextBox
    #Defino la posición
    $TextBox.Location = New-Object System.Drawing.Size(135,70)
    #Defino el texto que viene por defecto
    $TextBox.Text = ""
    $TextBox.Width = 300
    $Form.Controls.Add($TextBox)

    $Form.ShowDialog() 

}
#>

function validatePath( $type )
{
    do
    {
        $value = Read-Host -Prompt "Ingrese Ruta $type"

        if( (Test-Path $value) -eq $False )
        {
            Write-Host "Ruta $value no existe, por favor ingrese nuevamente la ruta"
        }


    } while ( !( Test-Path $value ) )

    $value
}

function createStructure( $ruta_input, $ruta_output , $files )
{
    $output = "Output_" + (get-date -uformat "%d_%m_%Y_%H_%M_%S")
    $fuentes = "SUNAT","OSCE","LINKEDIN","SENTINEL","GOOGLE ADVANCED","WORLD COMPLIANCE","SUNARP Persona Jurídica"
    
    $lista = dir $rutaOrigen | Select-Object Name

    $temp = $files.custodianTarget
    $custodians = $temp | select -Unique
    
    mkdir $ruta_output/"$output"
    
    
    foreach( $element in $custodians )
    {
        mkdir $ruta_output/"$output"/"$element"    
        
        foreach( $fuente in $fuentes)
        {
            mkdir $ruta_output/"$output"/"$element"/$fuente
        }
    }
    
    foreach( $file in $files )
    {
        if( $file)
        {
            $name = ($file.nombreOriginal).Trim()
            $custodian = ($file.custodianTarget).Trim()
            $fuente = ($file.fuenteOrigen).Trim()
            $nombreFinal = ($file.nombreFinal).Trim()
        
            $var1 = "$ruta_input\$name"
            $var2 = "$ruta_output\$output\$custodian\$fuente\$nombreFinal"
         
            Copy-Item $var1 -Destination $var2
        }
    }
    
    $varOutput = dir $rutaDestino\$output -Recurse  *.pdf | Select-Object Name
    
    <#validateCopyFiles $lista $varOutput#>
    
}

function validateCopyFiles( $var1 , $var2 )
{
    


}



function processExcel( $ruta )
{
    $ruta = $ruta + '\Fuente.xlsx'
    $directorio = @()
    
    $objExcel = New-Object -ComObject Excel.Application
    $objExcel.Visible = $true
    
    $wb = $objExcel.Workbooks.Open($ruta)
    $ws = $wb.Sheets.Item(1)
    
    $rowMax = ( $ws.UsedRange.Rows ).count 

    for ( $i=2; $i -le $rowMax-1; $i++)
    {

        $codigo = $ws.Cells.Item( $rowName + $i, 3 ).text
        $custodianTarget = $ws.Cells.Item( $rowName + $i, 4 ).text        
        $fuenteOrigen = $ws.Cells.Item( $rowName + $i, 5 ).text        
        $nombreOriginal = $ws.Cells.Item( $rowName + $i, 8 ).text
        $nombreFinal = $ws.Cells.Item( $rowName + $i, 10 ).text

        $archivo = @{ 
                        codigo = $codigo;
                        custodianTarget = $custodianTarget; 
                        fuenteOrigen = $fuenteOrigen; 
                        nombreOriginal = $nombreOriginal; 
                        nombreFinal = $nombreFinal;
                    } 

        $directorio += $archivo;
     }
    
    $wb.Close()            
    $objExcel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject( $objExcel )

    $directorio
}

$rutaOrigen = validatePath("origen")
$rutaDestino = validatePath("destino")


$tempFiles = processExcel $rutaOrigen

createStructure $rutaOrigen $rutaDestino $tempFiles

$tempFiles.Clear()
