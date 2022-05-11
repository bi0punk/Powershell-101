#requires -version 2

function TreeSizeHtml {
<#
.SYNOPSIS
A Powershell clone of the classic TreeSize administrators tool. Works on local volumes or network shares.
Outputs the report to one or more interactive HTML files, and optionally zips them into a single zip file.
Requires Powershell 2. For Windows 2003 servers, install http://support.microsoft.com/kb/968930	
Author: James Weakley (jameswillisweakley@gmail.com)

Un clon de Powershell de la clásica herramienta de administración TreeSize. Funciona en volúmenes locales o recursos compartidos de red.
Envía el informe a uno o más archivos HTML interactivos y, opcionalmente, los comprime en un solo archivo zip.
Requiere Powershell 2. Para servidores Windows 2003, instale http://support.microsoft.com/kb/968930
Autor: James Weakley (jameswillisweakley@gmail.com)



.DESCRIPTION

Recursively iterates a folder structure and reports on the space consumed below each individual folder.
Outputs to a single HTML file which, with the help of a couple of third party javascript libraries,
displays in a web browser as an expandable tree, sorted by largest first.

Itera recursivamente una estructura de carpetas e informa sobre el espacio consumido debajo de cada carpeta individual.
Salidas a un solo archivo HTML que, con la ayuda de un par de bibliotecas javascript de terceros,
se muestra en un navegador web como un árbol expandible, ordenado por el más grande primero.



.PARAMETER paths
One or more comma separated locations to report on.
A report on each of these locations will be output to a single HTML file per location, defined by htmlOutputFilenames
Pass in the value "ALL" to report on all fixed disks.


Una o más ubicaciones separadas por comas sobre las que informar.
Se generará un informe sobre cada una de estas ubicaciones en un solo archivo HTML por ubicación, definido por htmlOutputFilenames
Pase el valor "TODOS" para informar sobre todos los discos fijos.




.PARAMETER reportOutputFolder
The folder location to output the HTML report(s) and zip file. This folder must exist already.

La ubicación de la carpeta para generar los informes HTML y el archivo zip. Esta carpeta ya debe existir.


.PARAMETER htmlOutputFilenames

One or more comma separated filenames to output the HTML reports to. There must be one of these to correspond with each path specified.
If "ALL" is specified for paths, then this parameter is ignored and the reports use the filenames "C_Drive.html","D_Drive.html", and so on

Uno o más nombres de archivo separados por comas para generar los informes HTML. Debe haber uno de estos para corresponder con cada ruta especificada.
Si se especifica "TODO" para las rutas, este parámetro se ignora y los informes usan los nombres de archivo "C_Drive.html","D_Drive.html", etc.


.PARAMETER zipOutputFilename
Name of zip file to place all generated HTML reports in. If this value is empty, HTML files are not zipped up.

Nombre del archivo zip para colocar todos los informes HTML generados. Si este valor está vacío, los archivos HTML no se comprimen.

.PARAMETER topFilesCountPerFolder
Setting this parameter filters the number of files shown at each level.
For example, setting it to 10 will mean that at each folder level, only the largest 10 files will be displayed in the report.
The count and sum total size of all other files will be shown as one item.
The default value is 20.
Setting the value to -1 disables filtering and always displays all files. Note that this may generate HTML files large enough to crash your web browser!



La configuración de este parámetro filtra la cantidad de archivos que se muestran en cada nivel.
Por ejemplo, establecerlo en 10 significará que en cada nivel de carpeta, solo se mostrarán los 10 archivos más grandes en el informe.
El conteo y el tamaño total de la suma de todos los demás archivos se mostrarán como un elemento.
El valor predeterminado es 20.
Establecer el valor en -1 deshabilita el filtrado y siempre muestra todos los archivos. ¡Tenga en cuenta que esto puede generar archivos HTML lo suficientemente grandes como para bloquear su navegador web!



.PARAMETER folderSizeFilterDepthThreshold
Enables a folder size filter which, in conjunction with folderSizeFilterMinSize, excludes from the report sections of the tree that are smaller than a particular size.
This value determines how many subfolders deep to travel before applying the filter.
The default value is 8
Note that this filter does not affect the accuracy of the report. The total size of the filtered out branches are still displayed in the report, you just can't drill down any further.
Setting the value to -1 disables filtering and always displays all files. Note that this may generate HTML files large enough to crash your web browser!



folderSizeFilterMinSize, excluye del informe las secciones del árbol que son más pequeñas que un tamaño particular.
Este valor determina a cuántas subcarpetas se debe viajar antes de aplicar el filtro.
El valor predeterminado es 8
Tenga en cuenta que este filtro no afecta la precisión del informe. El tamaño total de las ramas filtradas todavía se muestra en el informe, simplemente no puede profundizar más.
Establecer el valor en -1 deshabilita el filtrado y siempre muestra todos los archivos. ¡Tenga en cuenta que esto puede generar archivos HTML lo suficientemente grandes como para bloquear su navegador web!

.PARAMETER folderSizeFilterMinSize
Used in conjunction with folderSizeFilterDepthThreshold to excludes from the report sections of the tree that are smaller than a particular size.
This value is in bytes.
The default value is 104857600 (100MB)


Se usa junto con folderSizeFilterDepthThreshold para excluir del informe las secciones del árbol que son más pequeñas que un tamaño particular.
Este valor está en bytes.
El valor predeterminado es 104857600 (100 MB)


.PARAMETER displayUnits
A string which must be one of "B","KB","MB","GB","TB". This is the units to display in the report.
The default value is MB

Una cadena que debe ser una de "B","KB","MB","GB","TB". Estas son las unidades que se mostrarán en el informe.
El valor predeterminado es MB



.EXAMPLE
TreeSizeHtml -paths "C:\" -reportOutputFolder "C:\temp" -htmlOutputFilenames "c_drive.html"
This will output a report on C:\ to C:\temp\c_drive.html using the default filter settings.

TreeSizeHtml -paths "C:\" -reportOutputFolder "C:\temp" -htmlOutputFilenames "c_drive.html"
Esto generará un informe en C:\ a C:\temp\c_drive.html utilizando la configuración de filtro predeterminada.



.EXAMPLE
TreeSizeHtml -paths "C:\,D:\" -reportOutputFolder "C:\temp" -htmlOutputFilenames "c_drive.html,d_drive.html" -zipOutputFilename "report.zip"
This will output two size reports:
- A report on C:\ to C:\temp\c_drive.html
- A report on D:\ to C:\temp\d_drive.html
Both reports will be placed in a zip file at "C:\temp\report.zip"

TreeSizeHtml -paths "C:\,D:\" -reportOutputFolder "C:\temp" -htmlOutputFilenames "c_drive.html,d_drive.html" -zipOutputFilename "report.zip"
Esto generará informes de dos tamaños:
- Un informe de C:\ a C:\temp\c_drive.html
- Un informe de D:\ a C:\temp\d_drive.html
Ambos informes se colocarán en un archivo zip en "C:\temp\report.zip"



.EXAMPLE
TreeSizeHtml -paths "\\nas\ServerBackups" -reportOutputFolder "C:\temp" -htmlOutputFilenames "nas_server_backups.html" -topFilesCountPerFolder -1 -folderSizeFilterDepthThreshold -1
This will output a report on \\nas\ServerBackups to c:\temp\nas_server_backups.html
The report will include all files and folders, no matter how many or how small

TreeSizeHtml -paths "\\nas\ServerBackups" -reportOutputFolder "C:\temp" -htmlOutputFilenames "nas_server_backups.html" -topFilesCountPerFolder -1 -folderSizeFilterDepthThreshold -1
Esto generará un informe sobre \\nas\ServerBackups en c:\temp\nas_server_backups.html
El informe incluirá todos los archivos y carpetas, sin importar cuántos o cuán pequeños


.EXAMPLE
TreeSizeHtml -paths "E:\" -reportOutputFolder "C:\temp" -htmlOutputFilenames "e_drive_summary.html" -folderSizeFilterDepthThreshold 0 -folderSizeFilterMinSize 1073741824
This will output a report on E:\ to c:\temp\e_drive_summary.html
As soon as a branch accounts for less than 1GB of space, it is excluded from the report.

TreeSizeHtml -paths "E:\" -reportOutputFolder "C:\temp" -htmlOutputFilenames "e_drive_summary.html" -folderSizeFilterDepthThreshold 0 -folderSizeFilterMinSize 1073741824
Esto generará un informe en E:\ a c:\temp\e_drive_summary.html
Tan pronto como una sucursal cuenta con menos de 1 GB de espacio, se excluye del informe.



.NOTES
You need to run this function as a user with permission to traverse the tree, otherwise you'll have sections of the tree labeled 'Permission Denied'

Debe ejecutar esta función como un usuario con permiso para recorrer el árbol; de lo contrario, tendrá secciones del árbol etiquetadas como "Permiso denegado".
#>
#>
    param (
       [Parameter(Mandatory=$true)][String] $paths,
       [Parameter(Mandatory=$true)][String] $reportOutputFolder,
       [Parameter(Mandatory=$false)][String] $htmlOutputFilenames = $null,
       [Parameter(Mandatory=$false)][String] $zipOutputFilename = $null,
       [Parameter(Mandatory=$false)][int] $topFilesCountPerFolder = 5,
       [Parameter(Mandatory=$false)][int] $folderSizeFilterDepthThreshold = 2,
       [Parameter(Mandatory=$false)][long] $folderSizeFilterMinSize = 104857600,
       [Parameter(Mandatory=$false)][String] $displayUnits = "GB"
    )
    $ErrorActionPreference = "Stop"
   
    $pathsArray = @();
    $htmlFilenamesArray = @();
   

    # check output folder exists
    # comprobar que existe la carpeta de salida
    if (!($reportOutputFolder.EndsWith("\")))
    {
        $reportOutputFolder = $reportOutputFolder + "\"
    }

    $reportOutputFolderInfo = New-Object System.IO.DirectoryInfo $reportOutputFolder
    if (!$reportOutputFolderInfo.Exists)
    {
        Throw "Report output folder $reportOutputFolder does not exist"
    }
   

    # passing in "ALL" means that all fixed disks are to be included in the report

    # pasar "TODOS" significa que todos los discos fijos se incluirán en el informe

    if ($paths -eq "E:\")
    {
        gwmi win32_logicaldisk -filter "drivetype = 3" | % {
            $pathsArray += $_.DeviceID+"\"
            $htmlFilenamesArray += $_.DeviceID.replace(":","_Drive.html");
        }
       
    }
    else
    {
        if ($htmlOutputFilenames -eq $null -or $htmlOutputFilenames -eq '')
        {
            throw "las rutas no eran 'TODAS', pero htmlOutputFilenames no estaba definido. Si las rutas están definidas, entonces se debe especificar el mismo número de htmlOutputFileNames"
        }
        # split up the paths and htmlOutputFilenames parameters by comma
        # dividir las rutas y los parámetros htmlOutputFilenames por comas
        $pathsArray = $paths.split(",");
        $htmlFilenamesArray = $htmlOutputFilenames.split(",");
        if (!($pathsArray.Length -eq $htmlFilenamesArray.Length))
        {
            Throw "$($pathsArray.Length) se especificaron rutas pero $($htmlFilenamesArray.Length) htmlOutputFilenames. The number of HTML output filenames must be the same as the number of paths specified"
        }
    }
    for ($i=0;$i -lt $htmlFilenamesArray.Length; $i++)
    {
        $htmlFilenamesArray[$i] = ($reportOutputFolderInfo.FullName)+$htmlFilenamesArray[$i]
    }
    if (!($zipOutputFilename -eq $null -or $zipOutputFilename -eq ''))
    {
        $zipOutputFilename = ($reportOutputFolderInfo.FullName)+$zipOutputFilename
    }
   
    write-host "Parámetros de informe"
    write-host "-----------------"
    write-host "Ubicaciones para incluir:"
    for ($i=0;$i -lt $pathsArray.Length;$i++)
    {
        write-host "- $($pathsArray[$i]) to $($htmlFilenamesArray[$i])"        
    }
    if ($zipOutputFilename -eq $null -or $zipOutputFilename -eq '')
    {
        write-host "Omitir la creación del archivo zip"
    }
    else
    {
        write-host "Informar sobre los archivos HTML que se van a comprimir $zipOutputFilename"
    }
   
    write-host
    write-host "Filtros:"
    if ($topFilesCountPerFolder -eq -1)
    {
        write-host "- Mostrar todos los archivos"
    }
    else
    {
        write-host "- Visualización de los archivos $topFilesCountPerFolder más grandes por carpeta"
    }
   
    if ($folderSizeFilterDepthThreshold -eq -1)
    { write-host "- Visualización de toda la estructura de carpetas"
       
    }
    else
    {
        write-host "- Después de una profundidad de carpetas de $folderSizeFilter Depth Threshold, se excluyen las ramas con un tamaño total inferior a $folderSizeFilterMinSize bytes"
    }    
       
    write-host
   
    for ($i=0;$i -lt $pathsArray.Length; $i++){
   
        $_ = $pathsArray[$i];
        # get the Directory info for the root directory
        # obtener la información del directorio para el directorio raíz
        $dirInfo = New-Object System.IO.DirectoryInfo $_
        # test that it exists, throw error if it doesn't
        # probar que existe, arrojar error si no existe
        if (!$dirInfo.Exists)
        {
            Throw "La ruta $dirInfo no existe"
        }
       
       
        write-host "Árbol de objetos de construcción para la ruta $_"
        # traverse the folder structure and build an in-memory tree of objects
        # recorrer la estructura de carpetas y construir un árbol de objetos en memoria
        $treeStructureObj = @{}
        buildDirectoryTree_Recursive $treeStructureObj $_
        $treeStructureObj.Name = $dirInfo.FullName; #.replace("\","\\");        
       
        write-host "Building HTML output"
       
        # initialise a StringBuffer. The HTML will be written to here
        # inicializar un StringBuffer. El HTML se escribirá aquí.
        $sb = New-Object -TypeName "System.Text.StringBuilder";
        $fecha = Get-Date -Format D
        $fecha2 = Get-Date



        $partitions= Get-WmiObject -Class Win32_LogicalDisk -Filter 'DriveType = 3' |Select-Object PSComputerName, Caption,@{N='Capacity_GB'; E={[math]::Round(($_.Size / 1GB), 2)}},@{N='FreeSpace_GB'; E={[math]::Round(($_.FreeSpace / 1GB), 2)}},@{N='PercentUsed'; E={[math]::Round(((($_.Size - $_.FreeSpace) / $_.Size) * 100), 2) }},@{N='PercentFree'; E={[math]::Round((($_.FreeSpace / $_.Size) * 100), 2) }}
 
        foreach($z in $partitions)
        {  
            $particion =  "Unidad: $($z.Caption)" ;    
            $c_total =  "Capacidad Total:     $($z.Capacity_GB)  GB" ;    
            $e_libre =  "Espacio Libre:       $($z.FreeSpace_GB) GB" ;    
            $per_usado =  "Porcentaje Usado:  $($z.PercentUsed)  %" ;    
            $per_libre =  "Porcentaje Libre:  $($z.PercentFree)  %" ;   
        }

       
        # output the HTML and javascript for the report page to the StringBuffer
        # below here are mostly comments for the javascript code, which  
        # runs in the browser of the user viewing this report
        
        # enviar el HTML y javascript para la página del informe al StringBuffer
        # a continuación aquí hay principalmente comentarios para el código javascript, que
        # se ejecuta en el navegador del usuario que está viendo este informe


        sbAppend "<!DOCTYPE html>"
        sbAppend "<html>"
        sbAppend "<head>"
        sbAppend    "<meta charset='utf-8'>"


        # jquery javascript src (from web)
        sbAppend "<link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/jstree/3.3.12/themes/default/style.min.css'/>"
     
        sbAppend "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@4.1.3/dist/css/bootstrap.min.css' integrity='sha384-MCw98/SFnGE8fJT3GXwEOngsV7Zt27NXFoaoApmYm81iuXoPkFOJwJ8ERdknLPMO' crossorigin='anonymous'>"
        sbAppend "<script src='https://cdnjs.cloudflare.com/ajax/libs/jquery/1.12.1/jquery.min.js' type='text/javascript'></script>"

        # jstree javascript src (from web)
        sbAppend "<script src='https://cdnjs.cloudflare.com/ajax/libs/jstree/3.2.1/jstree.min.js' type='text/javascript'></script>"
        sbAppend "<script type='text/javascript'>`$(document).ready(function(){var a=`$('#jstree'),b=`$('#loading');a.on('loading.jstree',function(){return b.removeClass('hidden')}).on('loaded.jstree',function(){return b.addClass('hidden')}),a.jstree()})</script>"
        sbAppend "<script typw='text/javascript'>`$('#jstree').jstree({'core' : {'themes': {'name': 'default-dark','dots': true,'icons': true, 'checkbox':true}</script>"
        sbAppend "<meta name='theme-color' content='#7952b3'>"

        sbAppend "<style>"
        sbAppend ".hidden {"
        sbAppend "  visibility: hidden;"
        sbAppend "}"
        sbAppend "</style>"

        sbAppend "</head>"
        sbAppend "<body>"
        sbAppend "<style>"
        sbAppend "card {"
        sbAppend "  margin: 20px;"
        sbAppend "}"
        sbAppend "</style>"

        sbAppend "<style>"
        sbAppend "  .bd-placeholder-img {"
        sbAppend "  font-size: 1.125rem;"
        sbAppend "  text-anchor: middle;"
        sbAppend "  -webkit-user-select: none;"
        sbAppend "  -moz-user-select: none;"
        sbAppend "  user-select: none;"
        sbAppend "}"

        sbAppend  "@media (min-width: 768px) {"
        sbAppend  ".bd-placeholder-img-lg {"
        sbAppend  "font-size: 3.5rem;"
        sbAppend "}"
        sbAppend "}"

        sbappend "</style>"
        sbAppend "<style>"
        sbAppend ".b-example-divider {"
        sbAppend "height: 3rem;"
        sbAppend "background-color: rgba(0, 0, 0, .1);"
        sbAppend "border: solid rgba(0, 0, 0, .15);"
        sbAppend "border-width: 1px 0;"
        sbAppend "box-shadow: inset 0 .5em 1.5em rgba(0, 0, 0, .1), inset 0 .125em .5em rgba(0, 0, 0, .15);"
        sbAppend "}"

        sbAppend "@media (min-width: 992px) {"
        sbAppend ".rounded-lg-3 { border-radius: .3rem; }"
        sbAppend "}"
        sbAppend "</style>"

        sbAppend "<div id='kd' class='card'>"
        sbAppend    "<center><h4 class='card-header'>Reporte Uso de Discos - $fecha - $fecha2</h4></center>"

        sbAppend "<main>"
        

        sbAppend "<div class='px-4 py-5 my-5 text-center'>"
        sbAppend "<img class='d-block mx-auto mb-4' src='https://upload.wikimedia.org/wikipedia/commons/thumb/b/bd/Kyndryl_logo.svg/1920px-Kyndryl_logo.svg.png' alt='' width='140' height='45'>"

        sbAppend "<h3 class='display-5 fw-bold'>Centered hero</h3>"

        sbAppend "<div class='col-lg-6 mx-auto'>"
        sbAppend "<p class='lead mb-4'></p>"
        sbAppend "<div class='d-grid gap-2 d-sm-flex justify-content-sm-center'>"
        sbAppend "<table class='table table-striped'>"
        sbAppend    "<thead>"
        sbAppend        "<tr>"
        sbAppend            "<th scope='col'></th>"
        sbAppend            "<th scope='col'>First</th>"
        sbAppend            "<th scope='col'>Last</th>"
        sbAppend        "</tr>"
        sbAppend    "</thead>"
        sbAppend    "<tbody>"
        sbAppend        "<tr>"
        sbAppend            "<th scope='row'></th>"
        sbappend            "<td>Mark</td>"
        sbAppend            "<td>Otto</td>"
        sbAppend        "</tr>"
        sbAppend        "<tr>"
        sbAppend            "<th scope='row'></th>"
        sbAppend            "<td>Jacob</td>"
        sbAppend            "<td>Thornton</td>"
        sbAppend        "</tr>"
        sbAppend        "<tr>"
        sbAppend            "<th scope='row'></th>"
        sbAppend            "<td>Larry</td>"
        sbAppend            "<td>the Bird</td>"
        sbAppend         "</tr>"
        sbAppend    "</tbody>"
        sbAppend "</table>"
        sbAppend "</div>"
        sbAppend "</div>"
        sbAppend "</div>"


        sbAppend "<script src='https://cdn.jsdelivr.net/npm/@popperjs/core@2.9.2/dist/umd/popper.min.js' integrity='sha384-IQsoLXl5PILFhosVNubq5LC7Qb9DXgDA9i+tQ8Zj3iwWAwPtgFTxbJ8NT4GN1R8p' crossorigin='anonymous'></script"
        sbAppend "<script src='https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.min.js' integrity='sha384-cVKIPhGWiC2Al4u+LWgxfKTRIcfu0JTxR+EQDz/bgldoEyl4H0zUF0QKbrJ0EcQF' crossorigin='anonymous'></script>"
        sbAppend "</main>"
        sbAppend "</div>"
        sbAppend "<div id='header' class='card'>"
        sbAppend    "<div class='card-body'>"
       
                    $machine = hostname
        sbAppend        "<br>"           
        sbAppend        "<h5 class='card-title'>Servidor: $machine</h5>"
        sbAppend        "<h5>$DiscInfo</h5>"
        sbAppend        "<h5 class='card-title'>Directorio Raiz: ($($dirInfo.FullName))</h5>"

        sbAppend        "<p>$c_total</p>"
        sbAppend        "<p>$e_libre</p>"
        sbAppend        "<p>$per_usado</p>"
        sbAppend        "<p>$per_libre</p>"
        sbAppend        "<p class='card-text'></p>"
        sbAppend    "</div>"



        sbAppend    "<div class='card-footer bg-transparent'>Filtros de informes"
        sbAppend        "<p class='card-text'></p>"

        sbAppend    "</div>"

        sbAppend "<div class='card'>"
        sbAppend "<ul>"
       
        if ($topFilesCountPerFolder -eq -1)
        {
            sbAppend "<li>Visualización de todos los archivos</li>"
        }
        else
        {
            sbAppend "<li>Visualizacion de los $topFilesCountPerFolder archivos mas grandes por directorio</li>"
        }
       
        if ($folderSizeFilterDepthThreshold -eq -1)
        {
            sbAppend "<li>Visualizacion de la estructura de carpetas completa</li>"
        }
        else
        {
            sbAppend "<li>Despues de una profundidad de directorios de $folderSizeFilterDepthThreshold, se excluyen las ramas con un size total inferior a $folderSizeFilterMinSize bytes</li>"
        }    
       
        sbAppend "</ul>"
        sbAppend "</div>"
        sbAppend "<div id='error'></div>"
        # include a loading message and spinny icon while jsTree initialises
        # incluir un mensaje de carga y un ícono giratorio mientras jsTree se inicializa
        sbAppend "<div id='loading'>Loading...<img src='https://cdnjs.cloudflare.com/ajax/libs/jstree/3.3.3/themes/default/throbber.gif'/></div>"
        sbAppend "<div id='jstree'>"
        sbAppend "<ul id='tree'>"
       
        $size = bytesFormatter $treeStructureObj.SizeBytes $displayUnits
        $name = $treeStructureObj.Name.replace("'","\'")        
        # output the name and total size of the root folder
        # muestra el nombre y el tamaño total de la carpeta raíz
        sbAppend "   <li><span class='folder'>$name ($size)</span>"
        sbAppend "     <ul>"
        # recursively build the javascript object in the format that jsTree uses
        # construye recursivamente el objeto javascript en el formato que usa jsTree
        outputNode_Recursive $treeStructureObj $sb $topFilesCountPerFolder $folderSizeFilterDepthThreshold $folderSizeFilterMinSize 1;
        sbAppend "     </ul>"
        sbAppend "   </li>"
        sbAppend "</ul>"
        sbAppend "</div>"
       
       
       
       
        sbAppend "</body>"
        sbAppend "</html>"
       
       
        # finally, output the contents of the StringBuffer to the filesystem
        # finalmente, envíe el contenido del StringBuffer al sistema de archivos
        $outputFileName = $htmlFilenamesArray[$i]
        write-host "Writing HTML to file $outputFileName"
       
        Out-file -InputObject $sb.ToString() $outputFileName -encoding "UTF8"
    }
   
    if ($zipOutputFilename -eq $null -or $zipOutputFilename -eq '')
    {
        write-host "Skipping zip file creation"
    }
    else
    {
        # create zip file
        # crear archivo zip
    set-content $zipOutputFilename ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
    (dir $zipOutputFilename).IsReadOnly = $false
       
        for ($i=0;$i -lt $htmlFilenamesArray.Length; $i++){
           
            write-host "Copying $($htmlFilenamesArray[$i]) to zip file $zipOutputFilename"
            $shellApplication = new-object -com shell.application
        $zipPackage = $shellApplication.NameSpace($zipOutputFilename)
           
        $zipPackage.CopyHere($htmlFilenamesArray[$i])
           
            # the zip is asynchronous, so we have to wait and keep checking (ugly)
            # use a DirectoryInfo object to retrieve just the file name within the path,
            # this is what we check for every second

            # el zip es asíncrono, así que tenemos que esperar y seguir revisando (feo)
            # usar un objeto DirectoryInfo para recuperar solo el nombre del archivo dentro de la ruta,
            # esto es lo que verificamos cada segundo


            $fileInfo = New-Object System.IO.DirectoryInfo $htmlFilenamesArray[$i]
           
            $size = $zipPackage.Items().Item($fileInfo.Name).Size
            while($zipPackage.Items().Item($fileInfo.Name) -Eq $null)
            {
                start-sleep -seconds 1
                write-host "." -nonewline
            }
        }
        $inheritance = get-acl $zipOutputFilename
        $inheritance.SetAccessRuleProtection($false,$false)
        set-acl $zipOutputFilename -AclObject $inheritance
    }
   
}

#.SYNOPSIS
#
# Used internally by the TreeSizeHtml function.
# Used to perform Depth-First (http://en.wikipedia.org/wiki/Depth-first_search) search of the entire folder structure.
# This allows the cumulative total of space used to be added up during backtracking.
#
# Utilizado internamente por la función TreeSizeHtml.
# Se utiliza para realizar una búsqueda en profundidad (http://en.wikipedia.org/wiki/Depth-first_search) de toda la estructura de carpetas.
# Esto permite que el total acumulativo de espacio utilizado se sume durante el retroceso.


#.PARAMETER currentNode
#
# The current node object, a temporary custom object which represents the current folder in the tree.
# El objeto de nodo actual, un objeto personalizado temporal que representa la carpeta actual en el árbol.
#
#.PARAMETER currentPath
#
# The path to the current folder in the tree
# La ruta a la carpeta actual en el árbol

function buildDirectoryTree_Recursive {  
        param (  
            [Parameter(Mandatory=$true)][Object] $currentParentDirInfo,  
            [Parameter(Mandatory=$true)][String] $currentDirInfo
        )  
    $substDriveLetter = $null
   
    # if the current directory length is too long, try to work around the feeble Windows size limit by using the subst command
    # si la longitud del directorio actual es demasiado larga, intente sortear el débil límite de tamaño de Windows usando el comando subst
    if ($currentDirInfo.Length -gt 248)
    {
        Write-Host "$currentDirInfo has a length of $($currentDirInfo.Length), greater than the maximum 248, invoking workaround"
        $substDriveLetter = ls function:[d-z]: -n | ?{ !(test-path $_) } | select -First 1
        $parentFolder = ($currentDirInfo.Substring(0,$currentDirInfo.LastIndexOf("\")))
        $relative = $substDriveLetter+($currentDirInfo.Substring($currentDirInfo.LastIndexOf("\")))
        write-host "Mapping $substDriveLetter to $parentFolder for access via $relative"
        subst $substDriveLetter $parentFolder

        $dirInfo = New-Object System.IO.DirectoryInfo $relative

    }
    else
    {
        $dirInfo = New-Object System.IO.DirectoryInfo $currentDirInfo
    }

    # add its details to the currentParentDirInfo object
    # agregar sus detalles al objeto currentParentDirInfo
    $currentParentDirInfo.Files = @()
    $currentParentDirInfo.Folders = @()
    $currentParentDirInfo.SizeBytes = 0;
    $currentParentDirInfo.Name = $dirInfo.Name;
    $currentParentDirInfo.Type = "Folder";
   
   

    # iterate all subdirectories
    # iterar todos los subdirectorios
    try
    {
        $dirs = $dirInfo.GetDirectories() | where {!$_.Attributes.ToString().Contains("ReparsePoint")}; #don't include reparse points
        $files = $dirInfo.GetFiles();
        # remove any drive mappings created via subst above
        # eliminar las asignaciones de unidades creadas a través de subst arriba
        if (!($substDriveLetter -eq $null))
        {
            write-host "eliminando la unidad sustituta $substDriveLetter"
            subst $substDriveLetter /D
            $substDriveLetter = $null
        }

        $dirs | % {
            # create a new object for the subfolder to pass in
            # crear un nuevo objeto para que la subcarpeta pase
            $subFolder = @{}
            if ($_.Name.length -lt 1)
            {
                return;
            }
            # call this function in the subfolder. It will return after the entire branch from here down is traversed
            # llamar a esta función en la subcarpeta. Regresará después de que se atraviese toda la rama desde aquí hacia abajo.
            buildDirectoryTree_Recursive $subFolder ($currentDirInfo + "\" + $_.Name);
            # add the subfolder object to the list of folders at this level
            # agregar el objeto de subcarpeta a la lista de carpetas en este nivel
            $currentParentDirInfo.Folders += $subFolder;
            # the total size consumed from the subfolder down is now available.
            # Add it to the running total for the current folder
            $currentParentDirInfo.SizeBytes= $currentParentDirInfo.SizeBytes + $subFolder.SizeBytes;
           
        }
        # iterate all files
        $files | % {
            # create a custom object for each file, containing the name and size

            # crear un objeto personalizado para cada archivo, que contenga el nombre y el tamaño
            $htmlFileObj = @{};
            $htmlFileObj.Type = "File";
            $htmlFileObj.Name = $_.Name;
            $htmlFileObj.SizeBytes = $_.Length
            # add the file object to the list of files at this level
            # agregar el objeto de archivo a la lista de archivos en este nivel

            $currentParentDirInfo.Files += $htmlFileObj;
            # add the file's size to the running total for the current folder
            # agregue el tamaño del archivo al total acumulado de la carpeta actual
            $currentParentDirInfo.SizeBytes= $currentParentDirInfo.SizeBytes + $_.Length
        }
    }
    catch [Exception]
    {
        if ($_.Exception.Message.StartsWith('Access to the path'))
        {
            $currentParentDirInfo.Name = $currentParentDirInfo.Name + " (Acceso denegado)"
        }
        else
        {
            Write-Host $_.Exception.ToString()
        }
    }
}

function bytesFormatter{
<#
.SYNOPSIS
Used internally by the TreeSizeHtml function.
Takes a number in bytes, and a string which must be one of B,KB,MB,GB,TB and returns a nicely formatted converted string.

Utilizado internamente por la función TreeSizeHtml.
Toma un número en bytes y una cadena que debe ser una de B, KB, MB, GB, TB y devuelve una cadena convertida con un formato agradable.

.EXAMPLE
bytesFormatter -bytes 102534233454 -notation "MB"
returns "97,784 MB"

bytesFormatter -bytes 102534233454 -notación "MB"
devuelve "97.784 MB"
#>
param (
        [Parameter(Mandatory=$true)][decimal][AllowNull()] $bytes,
        [Parameter(Mandatory=$true)][String] $notation
    )
    if ($bytes -eq $null)
    {
        return "unknown size";
    }
    $notation = $notation.ToUpper();
    if ($notation -eq 'B')
    {
        return ($bytes.ToString())+" B";
    }
    if ($notation -eq 'KB')
    {
        return (roundOffAndAddCommas($bytes/1024)).ToString() + " KB"
    }
    if ($notation -eq 'MB')
    {
        return (roundOffAndAddCommas($bytes/1048576)).ToString() + " MB"
    }
    if ($notation -eq 'GB')
    {
        return (roundOffAndAddCommas($bytes/1073741824)).ToString() + " GB"
    }
    if ($notation -eq 'TB')
    {
        return (roundOffAndAddCommas($bytes/1099511627776)).ToString() + " TB"
    }
    Throw "Unrecognised notation: $notation. Must be one of B,KB,MB,GB,TB"
}

function roundOffAndAddCommas{
<#
.SYNOPSIS
Used internally by the TreeSizeHtml function.
Takes a number and returns it as a string with commas as thousand separators, rounded to 2dp

Utilizado internamente por la función TreeSizeHtml.
Toma un número y lo devuelve como una cadena con comas como separadores de miles, redondeado a 2dp
#>
param(
    [Parameter(Mandatory=$true)][decimal] $number)

    $value = "{0:N2}" -f $number;
    return $value.ToString();
}

function sbAppend{
<#
.SYNOPSIS
Used internally by the TreeSizeHtml function.
Shorthand function to append a string to the sb variable

Utilizado internamente por la función TreeSizeHtml.
Función abreviada para agregar una cadena a la variable sb
#>
param(
    [Parameter(Mandatory=$true)][string] $stringToAppend)
    $sb.Append($stringToAppend) | out-null;
}

function outputNode_Recursive{
<#
.SYNOPSIS
Used internally by the TreeSizeHtml function.
Utilizado internamente por la función TreeSizeHtml.


Used to output the folder tree to a StringBuffer in the format of an HTML unordered list which the TreeView library can display.
Se utiliza para enviar el árbol de carpetas a un StringBuffer en el formato de una lista desordenada HTML que puede mostrar la biblioteca TreeView.


.PARAMETER node
The current node object, a temporary custom object which represents the current folder in the tree.
El objeto de nodo actual, un objeto personalizado temporal que representa la carpeta actual en el árbol.
#>
    param (
        [Parameter(Mandatory=$true)][Object] $node,
        [Parameter(Mandatory=$true)][System.Text.StringBuilder] $sb,
        [Parameter(Mandatory=$true)][int] $topFilesCountPerFolder,
        [Parameter(Mandatory=$true)][int] $folderSizeFilterDepthThreshold,
        [Parameter(Mandatory=$true)][long] $folderSizeFilterMinSize,
        [Parameter(Mandatory=$true)][int] $CurrentDepth
    )
   
    # If there is more than one subfolder from this level, sort by size, largest first
    # Si hay más de una subcarpeta de este nivel, ordenar por tamaño, la más grande primero
    if ($node.Folders.Length -gt 1)
    {
        $folders = $node.Folders | Sort -Descending {$_.SizeBytes}
    }
    else
    {
        $folders = $node.Folders
    }
    # iterate each subfolder
    for ($i = 0; $i -lt $node.Folders.Length; $i++)
    {
        $_ = $folders[$i];
        # append to the string buffer a HTML List Item which represents the properties of this folder
        # iterar cada subcarpeta
       
        $size = bytesFormatter $_.SizeBytes $displayUnits
        $name = $_.Name.replace("'","\'")
        sbAppend "<li><span class='folder'>$name ($size)</span>"
        sbAppend "<ul>"
       
        if ($name -eq "winsxs")
        {
            sbAppend "<li><span class='folder'>Contents of folder hidden as <a href='http://support.microsoft.com/kb/2592038'>winsxs</a> commonly contains tens of thousands of files</span></li>"
        }
        elseif ($folderSizeFilterDepthThreshold -le $CurrentDepth -and $_.SizeBytes -lt $folderSizeFilterMinSize)
        {
            sbAppend "<li><span class='folder'>Contents of folder hidden via size filter</span></li>"
        }
        else
        {
            # call this function in the subfolder. It will return after the entire branch from here down is output to the string buffer
            # llamar a esta función en la subcarpeta. Regresará después de que toda la rama desde aquí hacia abajo se envíe al búfer de cadena
            outputNode_Recursive $_ $sb $topFilesCountPerFolder $folderSizeFilterDepthThreshold $folderSizeFilterMinSize ($CurrentDepth+1);
        }
       
        sbAppend "</ul>"        
        sbAppend "</li>"
       
    }
    # If there is more than one file on level, sort by size, largest first
    # Si hay más de un archivo en el nivel, ordenar por tamaño, primero el más grande
    if ($node.Files.Length -gt 1)
    {
        $files = $node.Files | Sort -Descending {$_.SizeBytes}
    }
    else
    {
        $files = $node.Files
    }

    # iterar cada archivo
    for ($i = 0; $i -lt $node.Files.Length; $i++)
    {
        if ($i -lt $topFilesCountPerFolder)
        {
            $_ = $files[$i];
            # append to the string buffer a HTML List Item which represents the properties of this file
            # agregar al búfer de cadena un elemento de lista HTML que representa las propiedades de este archivo
            $size = bytesFormatter $_.SizeBytes $displayUnits
            $name = $_.Name.replace("'","\'")
            sbAppend "<li><span class='file'>$name ($size)</span></li>"
        }
        else
        {
            $remainingFilesSize = 0;
            while ($i -lt $node.Files.Length)
            {
                $remainingFilesSize += $files[$i].SizeBytes
                $i++;
            }
            $size = bytesFormatter $_.SizeBytes $displayUnits
            $name = "..."+($node.Files.Length-$topFilesCountPerFolder)+" more files"
            sbAppend "<li><span class='file'>$name ($size)</span></li>"
        }
    }
}

TreeSizeHtml -paths "E:\" -reportOutputFolder "C:\Users\EnriqueSuarez\Desktop" -htmlOutputFilenames "script-disco-1.html"