call "env.cmd"
cd ..
cd ..
cd ..
cd ..
C:
cd %foldersign%

SignTool sign /f %foldercert% %folder%netstat.vbs
SignTool sign /f %foldercert% %folder%netstat_MORE.vbs
SignTool sign /f %foldercert% %folder%procesosmonitor.vbs
SignTool sign /f %foldercert% %folder%procesosmonitorClient.vbs

SignTool sign /f %foldercert% %folder%procesosmonitorMDMS.vbs
SignTool sign /f %foldercert% %folder%procesosmonitorPARALELO.vbs

SignTool sign /f %foldercert% %folder%RUNMACRO_MBA.vbs
SignTool sign /f %foldercert% %folder%RUNMACRO_MONITOREDMS.vbs
SignTool sign /f %foldercert% %folder%RUNMACRO_PARALELO.vbs
pause