Attribute VB_Name = "Module1"
Public numfactura As Double
Public nombre As String
Public direccion As String
Public t As Double
Public tiva As Double
Public STI As Double 'SUMA TOTAL DE IVA
Public recuperada As Boolean
Public ivarecuperado As Double
Public Anno As Integer ' A�o de la factura

'Variables para controlar la empresa con la que estamos facturando

Public NombreEmpresa As String
Public Empresa As Integer 'Va a contener el n�mero de empresa asociado a �sta                         'con la que estamos facturando
'1 ser� la empresa "Metalisteria M&B"
'2 ser� la empresa "Martin-Metal"
'3 ser� la empresa "Metalister�a Bonilla"
