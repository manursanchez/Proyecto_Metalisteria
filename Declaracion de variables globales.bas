Attribute VB_Name = "Module1"
Public numfactura As Double
Public nombre As String
Public direccion As String
Public t As Double
Public tiva As Double
Public STI As Double 'SUMA TOTAL DE IVA
Public recuperada As Boolean
Public ivarecuperado As Double
Public Anno As Integer ' Año de la factura

'Variables para controlar la empresa con la que estamos facturando

Public NombreEmpresa As String
Public Empresa As Integer 'Va a contener el número de empresa asociado a ésta                         'con la que estamos facturando
'1 será la empresa "Metalisteria M&B"
'2 será la empresa "Martin-Metal"
'3 será la empresa "Metalistería Bonilla"
