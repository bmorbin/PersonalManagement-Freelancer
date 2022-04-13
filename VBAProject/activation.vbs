Set objExcel = CreateObject("Excel.Application")
objExcel.Application.Run "'pathFile'!iniciar.iniciar"
objExcel.DisplayAlerts = False
objExcel.Application.Quit
Set objExcel = Nothing