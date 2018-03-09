#https://channel9.msdn.com/Blogs/MVP-Azure/Work-with-SharePoint-Online-lists-with-PNP-PowerShell
#Información de las opciones del campo FIELD
#    https://msdn.microsoft.com/en-us/library/office/aa979575.aspx
#https://olafd.wordpress.com/2017/05/09/create-fields-from-xml-for-sharepoint-online/
#Formulas de campos calculados
#    https://msdn.microsoft.com/es-es/library/office/bb862071(v=office.14).aspx
#Ejemplo de campo calculado. CURSO. Según la fecha de la incidencia calcula el curso de la incidencia.
#       =IF(MONTH([Fecha Incidencia])<9;YEAR([Fecha Incidencia])-1&"/"&YEAR([Fecha Incidencia]);YEAR([Fecha Incidencia])&"/"&YEAR([Fecha Incidencia])+1)
# uninstall-module SharePointPnPPowerShellOnline
# Install-Module SharePointPnPPowerShellOnline 

$credential = get-credential
connect-pnponline -url https://educantabria.sharepoint.com/iesalisal/tic/IncInformaticas/ -Credential $credential

Write-Host "       INICIO" -ForegroundColor yellow -BackgroundColor red
$CogerXML= Get-PnPField -List "Incidencias Informáticas" -Identity "FechaIncidencia"
$CogerXML.SchemaXml

$CogerXML= Get-PnPField -List "Incidencias Informáticas" -Identity "Estado"
$CogerXML.SchemaXml

###############################################################################
#RECREACION de todas las listas
###############################################################################
$QueLista       = "IncTecnico"
$QueListaTitulo = "Incidencias Técnico"
###############################################################################
Write-Host "Se va a crear la lista: "$QueLista -ForegroundColor Red -BackgroundColor Yellow
    new-pnplist -title $QueListaTitulo -Template GenericList -OnQuickLaunch -Url $QueLista
    Start-Sleep -s 5
    Write-Host "               Se añaden los campos a la lista: "$QueLista -ForegroundColor Red -BackgroundColor green
    Add-PnPListItem -List $QueLista -Values @{"Title" = "Técnico 01"}
    Add-PnPListItem -List $QueLista -Values @{"Title" = "Técnico 02"}
    Add-PnPListItem -List $QueLista -Values @{"Title" = "Técnico 03"}
    Add-PnPListItem -List $QueLista -Values @{"Title" = "Técnico 04"}

    Start-Sleep -s 5
Write-Host "               Se cambian algunos valores de los campos de la lista: "$QueLista -ForegroundColor Red -BackgroundColor green
    for ($i=0;$i -le 1;$i++) 
    { 
        $ColumnaTitle= Get-PnPField -List $QueLista  -Identity Title
        $ColumnaTitle.Description = "Nombre del Técnico"
        $ColumnaTitle.Title = "Nombre del Técnico"
        $ColumnaTitle.Update() 
    }
    $QueListaPropiedades=Get-PnPList |  Where-Object title -eq $QueListaTitulo
    $ID_IncTecnico=$QueListaPropiedades.Id
    Write-Host "                    Fin lista: "$QueLista". Tiene el ID= "$ID_IncTecnico -ForegroundColor Red -BackgroundColor Yellow

###############################################################################
$QueLista       = "IncUbicaciones"
$QueListaTitulo = "Incidencias Ubicaciones"
###############################################################################
Write-Host "Se va a crear la lista: "$QueLista -ForegroundColor Red -BackgroundColor Yellow
    new-pnplist -title $QueListaTitulo -Template GenericList -OnQuickLaunch -Url $QueLista
    Start-Sleep -s 5
    Write-Host "               Se añaden los campos a la lista: "$QueLista -ForegroundColor Red -BackgroundColor green
    Add-PnPfield -list $QueLista -Type Text      -DisplayName "Ubicación Completa"        -InternalName "UbicacionCompleta"           -AddtoDefaultView

    Add-PnPListItem -List $QueLista -Values @{"Title" = "Aula202"        ; "UbicacionCompleta"  = "Aula 202"}
    Add-PnPListItem -List $QueLista -Values @{"Title" = "SalaProfesores" ; "UbicacionCompleta"  = "Sala de Profesores"}
    Add-PnPListItem -List $QueLista -Values @{"Title" = "Compe"          ; "UbicacionCompleta"  = "Compensatoria"}

    Start-Sleep -s 5
Write-Host "               Se cambian algunos valores de los campos de la lista: "$QueLista -ForegroundColor Red -BackgroundColor green
    for ($i=0;$i -le 1;$i++) 
    { 
        $ColumnaTitle= Get-PnPField -List $QueLista  -Identity Title
        $ColumnaTitle.Description = "Ubicación donde ocurre la incidencia"
        $ColumnaTitle.Title = "Ubicación"
        $ColumnaTitle.Update() 
    }
    $QueListaPropiedades=Get-PnPList |  Where-Object title -eq $QueListaTitulo
    $ID_IncUbicaciones=$QueListaPropiedades.Id
Write-Host "                    Fin lista: "$QueLista". Tiene el ID= "$ID_IncUbicaciones -ForegroundColor Red -BackgroundColor Yellow

###############################################################################
$QueLista       = "IncTipos"
$QueListaTitulo = "Incidencias Tipos"
###############################################################################
Write-Host "Se va a crear la lista: "$QueLista -ForegroundColor Red -BackgroundColor Yellow
    new-pnplist -title $QueListaTitulo -Template GenericList -OnQuickLaunch -Url $QueLista
    Start-Sleep -s 5
    Write-Host "               Se añaden los campos a la lista: "$QueLista -ForegroundColor Red -BackgroundColor green
    Add-PnPListItem -List $QueLista -Values @{"Title" = "Error Hardware"}
    Add-PnPListItem -List $QueLista -Values @{"Title" = "Problema con usuario"}
    Add-PnPListItem -List $QueLista -Values @{"Title" = "Solicitud de compra"}
    Add-PnPListItem -List $QueLista -Values @{"Title" = "Actualización de software"}
    Add-PnPListItem -List $QueLista -Values @{"Title" = "Error software"}

    Start-Sleep -s 5
Write-Host "               Se cambian algunos valores de los campos de la lista: "$QueLista -ForegroundColor Red -BackgroundColor green
    for ($i=0;$i -le 1;$i++) 
    { 
        $ColumnaTitle= Get-PnPField -List $QueLista  -Identity Title
        $ColumnaTitle.Description = "Tipos de incidencias"
        $ColumnaTitle.Title = "Tipo Incidencia"
        $ColumnaTitle.Update() 
    }
    $QueListaPropiedades=Get-PnPList |  Where-Object title -eq $QueListaTitulo
    $ID_IncTipos=$QueListaPropiedades.Id
    Write-Host "                    Fin lista: "$QueLista". Tiene el ID= "$ID_IncTipos -ForegroundColor Red -BackgroundColor Yellow


###############################################################################
$QueLista       = "IncInformaticas"
$QueListaTitulo = "Incidencias Informáticas"
###############################################################################
Write-Host "Se va a crear la lista: "$QueLista -ForegroundColor Red -BackgroundColor Yellow
    new-pnplist -Title $QueListaTitulo -Template GenericList -OnQuickLaunch -Url $QueLista
    Start-Sleep -s 5
    $QueListaPropiedades=Get-PnPList |  Where-Object title -eq $QueListaTitulo
    $ID_IncInformaticas=$QueListaPropiedades.Id
    
Write-Host "               Se añaden los campos a la lista: "$QueLista -ForegroundColor Red -BackgroundColor green
#-------------------------------------------------------------------------------
    Add-PnPfield -list $QueLista -Type Note     -DisplayName "Descripción Incidencia"   -InternalName "DescripIncidencia"      -AddToDefaultView -Required
#-------------------------------------------------------------------------------
    $xml = '<Field Type="DateTime" Name="FechaIncidencia" DisplayName="Fecha Incidencia" Description="Fecha de la Incidencia."
			    Required="TRUE" 
			    SourceID="'+$ID_IncInformaticas+'" 
			    StaticName="FechaIncidencia" 
			    Format="DateOnly"
                FriendlyDisplayFormat="Disabled">
                <Default>[today]</Default>
            </Field>'
    Add-PnPFieldFromXml -FieldXml $xml -List $QueLista 
#-------------------------------------------------------------------------------
    $xml = '<Field Type="Lookup" Name="Ubicacion" DisplayName="Ubicación" Description="Ubicación donde ocurre la incidencia."
                Viewable="TRUE"
	            Required="TRUE" 
	            List="'+$ID_IncUbicaciones+'"
	            ShowField="Title" 
	            SourceID="'+$ID_IncInformaticas+'" 
	            StaticName="Ubicacion" 
	          
            />'
    Add-PnPFieldFromXml -FieldXml $xml -List $QueLista     
#-------------------------------------------------------------------------------
    Add-PnPField -List $QueLista -Type Choice     -DisplayName "Prioridad"             -InternalName "Prioridad" -AddToDefaultView -Choices "1-Media","2-Alta","3-Baja"
#-------------------------------------------------------------------------------
    Add-PnPField -List $QueLista -Type text       -DisplayName "Nº Serie Consejería"        -InternalName "NumSerieConsejeria"      -AddToDefaultView 
#-------------------------------------------------------------------------------
#    Add-PnPField -List $QueLista -Type user       -DisplayName "Incidencia Asignada A"      -InternalName "IncidenciaAsignadaA"     -AddToDefaultView 
    $xml = '<Field Type="Lookup" Name="IncidenciaAsignadaA" DisplayName="Incidencia Asignada A" Description="Persona a la que le asigna la resolución de la incidencia. (A Rellenar por el TIC)." 
                Viewable="TRUE"
	            Required="FALSE" 
	            List="'+$ID_IncTecnico+'"
	            ShowField="Title" 
	            SourceID="'+$ID_IncInformaticas+'" 
	            StaticName="IncidenciaAsignadaA" 

            />'
    Add-PnPFieldFromXml -FieldXml $xml -List $QueLista 
#-------------------------------------------------------------------------------
    $xml = '<Field Type="Lookup" Name="TipoIncidencia" DisplayName="Tipo Incidencia" Description="Tipo de incidencia. (A rellenar por el TIC)." 
                Viewable="TRUE"
	            Required="FALSE" 
	            List="'+$ID_IncTipos+'"
	            ShowField="Title" 
	            SourceID="'+$ID_IncInformaticas+'" 
	            StaticName="TipoIncidencia" 

            />'
    Add-PnPFieldFromXml -FieldXml $xml -List $QueLista 
#-------------------------------------------------------------------------------
    Add-PnPField -List $QueLista -Type Choice     -DisplayName "Estado"                -InternalName "Estado"    -AddToDefaultView -Choices "En Progreso" ,"En Espera","Cerrada","Nuevo"
#-------------------------------------------------------------------------------
    $xml = '<Field Type="DateTime" Name="FechaSolucion" DisplayName="Fecha Solución" Description="Fecha solución.  (A rellenar por el TIC)." 
			SourceID="'+$ID_IncInformaticas+'" 
			StaticName="FechaSolucion" 
			Format="DateOnly"
		    />'
    Add-PnPFieldFromXml -FieldXml $xml -List $QueLista 
#-------------------------------------------------------------------------------
    Add-PnPfield -list $QueLista -Type Note       -DisplayName "Descripción Solución"       -InternalName "DescripSolucion"         -AddToDefaultView 
#-------------------------------------------------------------------------------
    Add-PnPfield -list $QueLista -Type Number     -DisplayName "Duración(min)"              -InternalName "Duracion"                -AddToDefaultView 
#-------------------------------------------------------------------------------
    Add-PnPfield -list $QueLista -Type Boolean    -DisplayName "Llamada Servicio Técnico"   -InternalName "LlamadaServicioTecnico"  -AddToDefaultView
#-------------------------------------------------------------------------------
    $xml = '<Field Type="DateTime" Name="FechaLlamadaServTecnico" DisplayName="Fecha Llamada" Description="Fecha de llamada al servicio técnico. (A rellenar por el TIC)."
			SourceID="'+$ID_IncInformaticas+'" 
			StaticName="FechaLlamadaServTecnico" 
			Format="DateOnly"
		    />'
    Add-PnPFieldFromXml -FieldXml $xml -List $QueLista 
#-------------------------------------------------------------------------------
    $xml = '<Field Type="Calculated" Name="Curso" DisplayName="Curso" Description="Curso escolar." 
			    Format="DateOnly" 
			    ResultType="Text" 
			    ReadOnly="TRUE" 
			    SourceID="'+$ID_IncInformaticas+'" 
			    StaticName="Curso">
			    <Formula>
				    =IF(MONTH([Fecha Incidencia])&lt;9,YEAR([Fecha Incidencia])-1&amp;"/"&amp;YEAR([Fecha Incidencia]),YEAR([Fecha Incidencia])&amp;"/"&amp;YEAR([Fecha Incidencia])+1)
			    </Formula>
			    <FieldRefs>
					    <FieldRef Name="[Fecha Incidencia]" />
			    </FieldRefs>
		    </Field>'
    Add-PnPFieldFromXml -FieldXml $xml -List $QueLista
#-------------------------------------------------------------------------------
    #Otros posibles tipos de campos URL/Moneda...    
    #Add-PnPfield -list $QueLista -Type URL        -DisplayName "Página Web"                  -InternalName "PaginaWeb"       -AddToDefaultView 
    #Add-PnPfield -list $QueLista -Type Currency   -DisplayName "Precio"                      -InternalName "Precio"          -AddToDefaultView 
#-------------------------------------------------------------------------------
    Start-Sleep -s 5
Write-Host "               Se cambian algunos valores de los campos de la lista: "$QueLista -ForegroundColor Red -BackgroundColor green
    for ($i=0;$i -le 1;$i++) 
    { 
            #$Columna= Get-PnPField -List $QueLista  -Identity FechaIncidencia 
            #$Columna.DefaultValue="[today]"
            #$Columna.Update()

            $Columna= Get-PnPField -List $QueLista  -Identity Prioridad 
            $Columna.DefaultValue="1-Media"
            $Columna.Update()
            
            $Columna= Get-PnPField -List $QueLista  -Identity Estado 
            $Columna.DefaultValue="Nuevo"
            $Columna.Update()

            $Columna= Get-PnPField -List $QueLista  -Identity Duracion 
            $Columna.DefaultValue=0
            $Columna.Update()

            $ColumnaTitle= Get-PnPField -List $QueLista  -Identity Title
            $ColumnaTitle.Description = "Descripción corta de la incidencia."
            $ColumnaTitle.Title = "Asunto"
            $ColumnaTitle.Update() 
                        
            $Columna= Get-PnPField -List $QueLista  -Identity DescripIncidencia
            $Columna.Description = "Descripción larga. Se puede reflejar en este campo todo lo necesario para facilitar la resolución de la incidencia."
            $Columna.Update() 

            $Columna= Get-PnPField -List $QueLista  -Identity NumSerieConsejeria
            $Columna.Description = "Nº de Serie de la Consejería de Educación. Rellenar si es posible (NO ES OBLIGATORIO)."
            $Columna.Update() 

            $Columna= Get-PnPField -List $QueLista  -Identity IncidenciaAsignadaA
            $Columna.Description = "Persona a la que le asigna la resolución de la incidencia. (A rellenar por el TIC)."
            $Columna.Update() 

            $Columna= Get-PnPField -List $QueLista  -Identity Duracion
            $Columna.Description = "Tiempo (minutos) que se ha tardado en resolver la incidencia. (A rellenar por el TIC)."
            $Columna.Update() 

            $Columna= Get-PnPField -List $QueLista  -Identity LlamadaServicioTecnico
            $Columna.Description = "Sólo habilitar si se llama a la empresa encargada del servicio técnico. (A rellenar por el TIC)."
            $Columna.Update() 

    }
Write-Host "                    Fin lista: "$QueLista -ForegroundColor Red -BackgroundColor Yellow

Write-Host "       FINNNNNNN" -ForegroundColor yellow -BackgroundColor red