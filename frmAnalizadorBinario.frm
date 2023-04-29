VERSION 5.00
Begin VB.Form frmAnalizadorBinario 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form1"
   ClientHeight    =   10320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12645
   LinkTopic       =   "Form1"
   ScaleHeight     =   10320
   ScaleWidth      =   12645
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8760
      Left            =   5760
      TabIndex        =   3
      Top             =   120
      Width           =   5295
   End
   Begin VB.CommandButton cmdAnalizarBinario 
      BackColor       =   &H00FF8080&
      Caption         =   "ANALIZAR BINARIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label lblDigitosSignificativos 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblBinario 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmAnalizadorBinario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'* PROYECTO      : ANALISIS BINARIO 4X4
'* CONTENIDO     : EXPLORAR SOLUCIONES DEL SUDOKUS DE 4X4 CON ANALISIS BINARIO
'* VERSION       : 1.1
'* AUTORES       : MIGUEL QUINTEIRO PIÑERO / MIGUEL QUINTEIRO FERNANDEZ
'* INICIO        : 02 DE FEBRERO DE 2013
'* ACTUALIZACION : 29 DE MARZO DE 2013
'****************************************************************************************
Option Explicit

' LAS 288 SOLUCIONES QUE EXISTEN PARA EL SUDOKU DE 4X4
Private Type misSoluciones
    miNumero As Long
    miCasilla(1 To 16) As Integer
End Type

' DECLARACION DE VARIABLES
Dim miSolucion(1 To 288) As misSoluciones
Dim miPlanteamiento(1 To 288) As misSoluciones

' DECLARACION DE VARIABLES
Dim miLineInput As String
Dim miLineOutput As String


Private Sub Form_Load()
    ' CARGA VALORES INICIALES PARA LAS SOLUCIONES
    ' ABRE EL ARCHIVO CON LAS 288 SOLUCIONES
    Dim x As Integer
    Dim miNumero As Integer
    Open "miSolucionesTotal.txt" For Input As #10
    Do Until EOF(10)
        Line Input #10, miLineInput
        miNumero = Val(Mid(miLineInput, 32, 3))
        ' CARGA EL NUMERO DE LA GRILLA
        miSolucion(miNumero).miNumero = Val(Mid(miLineInput, 32, 3))
        ' CARGA LOS VALORES DE LOS DÍGITOS QUE COMPONEN LA SOLUCION
        For x = 1 To 16
            miSolucion(miNumero).miCasilla(x) = Val(Mid(miLineInput, x, 1))
        Next x
    Loop
    Close #10
End Sub


' BOTON PARA ANALIZAR BINARIOS
Private Sub cmdAnalizarBinario_Click()
    Dim i As Long
    Dim j As Integer
    Dim k As Integer
    
    Dim miNumeroCasos As Long
    Dim miMascaraBinaria As String
    
    ' RECORRE LOS 65534
    For i = 3 To 3 ' hasta 65534
        miNumeroCasos = i
                
        ' APLICAR LA MASCARA BINARIA PARA OBTENER LOS PLANTENAMIENTOS
        miNumeroCasos = i
        miMascaraBinaria = DecimalBinario(miNumeroCasos)
             
        ' RECORRE LOS 288
        For j = 1 To 288 ' hasta 288
            miPlanteamiento(j).miNumero = miSolucion(j).miNumero
            ' RECORRE LOS 16
            For k = 1 To 16
                If Mid(miMascaraBinaria, k, 1) = "1" Then
                    miPlanteamiento(j).miCasilla(k) = miSolucion(j).miCasilla(k)
                Else
                    miPlanteamiento(j).miCasilla(k) = 0
                End If
            Next k
            
            ' MUESTRA DATOS CON MASCARA EN EL LISTBOX
            miLineOutput = ""
            For k = 1 To 16
                miLineOutput = miLineOutput + Str(miPlanteamiento(j).miCasilla(k))
            Next k
            List1.AddItem miLineOutput
                    
        Next j
        
        miNumeroCasos = i
        lblBinario = DecimalBinario(miNumeroCasos)
        DoEvents
    Next i
End Sub

' CONVIERTE DECIMAL ( 1 - 65534) A BINARIO EN FORMATO DE 16 DIGITOS
Private Function DecimalBinario(miNumeroDecimal As Long) As String
    'Declaramos el Residuo
    Dim Residuo As String
    Dim miDigitosSignificativos As Integer
    Dim miCompletaCeros As String
    
    'Seteamos el Resultado a vacio
    DecimalBinario = ""
    miDigitosSignificativos = 0
    miCompletaCeros = "0000000000000000"
    
    Do
        'Obtenemos el Residuo de la division
        Residuo = miNumeroDecimal Mod 2
        
        If Residuo = 1 Then
            miDigitosSignificativos = miDigitosSignificativos + 1
        End If
        'concatenamos el Residuo al final con lo acumulado en el resultado, recuerden que las cadenas se concatenan con el ampersan "&" no con el "+"
        DecimalBinario = DecimalBinario & Trim(Str(Residuo))
        'Obtenemos el entero de la division
        miNumeroDecimal = Int(miNumeroDecimal / 2)
    'Seguimos haciendo la operación hasta que el numero sea 0 o 1
    Loop Until miNumeroDecimal < 2
    'verificamos que valor tenemos como ultimo residuo o mejor dicho como ultimo numero
    If (miNumeroDecimal = 1) Then
        'le agregamos el ultimo valor al inicio ya que el valor anterior lo vamos a revertir
        DecimalBinario = "1" & StrReverse(DecimalBinario)
        miDigitosSignificativos = miDigitosSignificativos + 1
    Else
        'como no hay nada que concatenar, simplemente revertimos
        DecimalBinario = StrReverse(DecimalBinario)
    End If
    
    If Len(DecimalBinario) < 16 Then
        DecimalBinario = Mid(miCompletaCeros, 1, (16 - Len(DecimalBinario))) & DecimalBinario
    End If
    
    lblDigitosSignificativos = miDigitosSignificativos
End Function

