<#

#>

function RobocopyEY( $rutaOrigen , $rutaDestino , $isInRed )
{
    if( $isInRed -eq "yes" )
    {
        $robType = "COPY:DATOS"
    }
    elseif( $isInRed -eq "no" )
    {
        $robType = "COPYALL"
    }
    
    $date = Get-Date  -UFormat "%y_%m_%d_%h"
    
    $logRobocopy = "$rutaDestino\robocopy.txt"

    Robocopy "$rutaOrigen" "$rutaDestino" /MT:16 /FP /S /E /COPYALL /NP /LOG:"$logRobocopy"
    
}

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

    $rutaOrigen = validatePath("origen")
    $rutaDestino = validatePath("destino")
    $isInRed = Read-Host -Prompt "¿La operación a realizar es en red (yes/no)?"
    

    RobocopyEY $rutaOrigen $rutaDestino  $isInRed   
    
    
     