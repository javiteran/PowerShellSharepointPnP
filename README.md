# PowerShellSharepointPnP
Power ShellSharepoint PnP. Creación y modificación de listas, campos e inserción de datos


Es necesario tener instalado el módulo SharePointPnPPowerShellOnline
    # uninstall-module SharePointPnPPowerShellOnline
      Install-Module SharePointPnPPowerShellOnline 

La información se ha obtenido de las siguiente fuentes
https://channel9.msdn.com/Blogs/MVP-Azure/Work-with-SharePoint-Online-lists-with-PNP-PowerShell


#Información de las opciones del campo FIELD
    https://msdn.microsoft.com/en-us/library/office/aa979575.aspx

    https://olafd.wordpress.com/2017/05/09/create-fields-from-xml-for-sharepoint-online/

#Formulas de campos calculados
    https://msdn.microsoft.com/es-es/library/office/bb862071(v=office.14).aspx

#Ejemplo de campo calculado. CURSO. Según la fecha de la incidencia calcula el curso de la incidencia.
   =IF(MONTH([Fecha Incidencia])<9;YEAR([Fecha Incidencia])-1&"/"&YEAR([Fecha Incidencia]);YEAR([Fecha Incidencia])&"/"&YEAR([Fecha Incidencia])+1)


