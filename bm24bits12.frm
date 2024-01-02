VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} bm24bits12 
   Caption         =   "Imagem BMP 24 bits"
   ClientHeight    =   3570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2910
   OleObjectBlob   =   "bm24bits12.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "bm24bits12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Licenciado sob a licença MIT.
' Copyright (C) 2012 - 2024 @Fabasa-Pro. Todos os direitos reservados.
' Consulte LICENSE.TXT na raiz do projeto para obter informações.

Option Explicit

Private Sub CommandButton1_Click()

    ' Declarações gerais:
    
    Dim HX As String    ' Dados (hexadecimal)
    Dim BT As String    ' Bytes
    Dim i As Integer    ' Índices
    
    ' Primeira estrutura 'Bitmap File Header' contém informações sobre o tipo,
    ' tamanho e layout de um bitmap e ocupa 14 bytes (padrão).
        
    HX = HX & "424D"        ' BitmapFileType         WORD               4D42 = 19778, 42 = 66 4D = 77 "BM"    O tipo de arquivo ("BM").
    HX = HX & "B2040000"    ' BitmapFileSize         DOUBLE WORD    000004B2 = 14 + 12 + 1176 = 1202 bytes    O tamanho do arquivo bitmap.
    HX = HX & "0000"        ' BitmapFileReserved1    WORD               0000 = 0 byte                         Reservados (0 byte)
    HX = HX & "0000"        ' BitmapFileReserved2    WORD               0000 = 0 byte                         Reservados (0 byte)
    HX = HX & "1A000000"    ' BitmapFileOffBits      DOUBLE WORD    0000001A = 14 + 12 = 26 bytes             O deslocamento desde o início da estrutura BITMAPFILEHEADER até os bits de bitmap.
    
    ' Segunda estrutura 'Bitmap Core Header' é semelhante à primeira, porém
    ' contém dados reduzidos, apenas informações sobre as dimensões e formato de
    ' cores de um bitmap e ocupa 12 bytes (padrão).
    
    HX = HX & "0C000000"    ' BitmapCoreSize         DOUBLE WORD    0000000C = 12 bytes     Especifica o número de bytes exigidos pela estrutura.
    HX = HX & "1200"        ' BitmapCoreWidth        WORD           00000012 = 18 pixels    Especifica a largura do bitmap.
    HX = HX & "1500"        ' BitmapCoreHeight       WORD           00000015 = 21 pixels    Especifica a altura do bitmap.
    HX = HX & "0100"        ' BitmapCorePlanes       WORD               0001 = 1 plano      Especifica o número de planos para o dispositivo de destino. (1 plano)
    HX = HX & "1800"        ' BitmapCoreBitCoun      WORD               0018 = 24 bpp       Especifica o número de bits por pixel.
       
        
    ' Terceira estrutura 'Palette' não é necessária para o bitmaps, aqui temos
    ' já a quarta estrutura 'Bitmap' contém todos os pixels extrudados em uma
    ' matriz de coluna e linha, onde temos linhas de 0 a 20 = 21 de altura e 18
    ' na largura, em partes de 32 bits, por esse motivo completamos com 0 (zero)
    ' até obter os completos 32 bits, ela ocupa 21 linhas * 56 bytes = 1176
    ' bytes.
    
    '               32 bits                     32 bits                     32 bits                     32 bits                     32 bits
    '      32 bits --------- 32 bits   32 bits --------- 32 bits   32 bits --------- 32 bits   32 bits --------- 32 bits   32 bits ---------
    '     ---------         --------- ---------         --------- ---------         --------- ---------         --------- ---------
    '  0: FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF 0000
    '  1: FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF 0000
    '  2: FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF 000000 000000 000000 000000 FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF 0000
    '  3: FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF 000000 000000 00FFFF 00FFFF 00FFFF 00FFFF 000000 000000 FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF 0000
    '  4: FFFFFF FFFFFF FFFFFF FFFFFF 000000 00FFFF 00FFFF 00FFFF 00FFFF 00FFFF 00FFFF 00FFFF 00FFFF 000000 FFFFFF FFFFFF FFFFFF FFFFFF 0000
    '  5: FFFFFF FFFFFF FFFFFF FFFFFF 000000 00FFFF 00FFFF 00FFFF 00FFFF 00FFFF 00FFFF 00FFFF 00FFFF 000000 FFFFFF FFFFFF FFFFFF FFFFFF 0000
    '  6: FFFFFF FFFFFF FFFFFF 000000 00FFFF FFFFFF 00FFFF 00FFFF 00FFFF 00FFFF 00FFFF 00FFFF 00FFFF 00FFFF 000000 FFFFFF FFFFFF FFFFFF 0000
    '  7: FFFFFF FFFFFF FFFFFF 000000 FFFFFF FFFFFF FFFFFF FFFFFF 00FFFF 00FFFF 00FFFF 00FFFF 00FFFF FFFFFF 000000 FFFFFF FFFFFF FFFFFF 0000
    '  8: FFFFFF FFFFFF FFFFFF 000000 FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF 000000 FFFFFF FFFFFF FFFFFF 0000
    '  9: FFFFFF FFFFFF FFFFFF FFFFFF 000000 FFFFFF FFFFFF 000000 FFFFFF FFFFFF 000000 FFFFFF FFFFFF 000000 FFFFFF FFFFFF FFFFFF FFFFFF 0000
    ' 10: FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF 000000 FFFFFF 000000 FFFFFF FFFFFF 000000 FFFFFF 000000 FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF 0000
    ' 11: FFFFFF FFFFFF FFFFFF FFFFFF 000000 0000FF 000000 FFFFFF FFFFFF FFFFFF FFFFFF 000000 0000FF 000000 FFFFFF FFFFFF FFFFFF FFFFFF 0000
    ' 12: FFFFFF FFFFFF FFFFFF 000000 FFFFFF 0000FF 0000FF 000000 000000 000000 000000 0000FF 0000FF FFFFFF 000000 FFFFFF FFFFFF FFFFFF 0000
    ' 13: FFFFFF FFFFFF 000000 FFFFFF FFFFFF 000000 0000FF 0000FF 0000FF 0000FF 0000FF 0000FF 000000 FFFFFF FFFFFF 000000 FFFFFF FFFFFF 0000
    ' 14: FFFFFF FFFFFF 000000 FFFFFF FFFFFF 000000 0000FF 0000FF 0000FF 0000FF 0000FF 0000FF 000000 FFFFFF FFFFFF 000000 FFFFFF FFFFFF 0000
    ' 15: FFFFFF FFFFFF FFFFFF 000000 000000 FFFF00 000000 000000 000000 000000 000000 000000 FFFF00 000000 000000 FFFFFF FFFFFF FFFFFF 0000
    ' 16: FFFFFF FFFFFF FFFFFF FFFFFF 000000 FFFF00 FFFF00 FFFF00 FFFF00 FFFF00 FFFF00 FFFF00 FFFF00 000000 FFFFFF FFFFFF FFFFFF FFFFFF 0000
    ' 17: FFFFFF FFFFFF FFFFFF FFFFFF 000000 FF0000 FF0000 FF0000 000000 000000 FF0000 FF0000 FF0000 000000 FFFFFF FFFFFF FFFFFF FFFFFF 0000
    ' 18: FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF 000000 000000 000000 FFFFFF FFFFFF 000000 000000 000000 FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF 0000
    ' 19: FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF 0000
    ' 20: FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF FFFFFF 0000
                                                                                                                                        
    HX = HX & "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0000"    ' 20:                                                                                                                               0000
    HX = HX & "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0000"    ' 19:                                                                                                                               0000
    HX = HX & "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF000000000000000000FFFFFFFFFFFF000000000000000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0000"    ' 18:                                    000000 000000 000000               000000 000000 000000                                    0000
    HX = HX & "FFFFFFFFFFFFFFFFFFFFFFFF000000FF0000FF0000FF0000000000000000FF0000FF0000FF0000000000FFFFFFFFFFFFFFFFFFFFFFFF0000"    ' 17:                             000000 FF0000 FF0000 FF0000 000000 000000 FF0000 FF0000 FF0000 000000                             0000
    HX = HX & "FFFFFFFFFFFFFFFFFFFFFFFF000000FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00000000FFFFFFFFFFFFFFFFFFFFFFFF0000"    ' 16:                             000000 FFFF00 FFFF00 FFFF00 FFFF00 FFFF00 FFFF00 FFFF00 FFFF00 000000                             0000
    HX = HX & "FFFFFFFFFFFFFFFFFF000000000000FFFF00000000000000000000000000000000000000FFFF00000000000000FFFFFFFFFFFFFFFFFF0000"    ' 15:                      000000 000000 FFFF00 000000 000000 000000 000000 000000 000000 FFFF00 000000 000000                      0000
    HX = HX & "FFFFFFFFFFFF000000FFFFFFFFFFFF0000000000FF0000FF0000FF0000FF0000FF0000FF000000FFFFFFFFFFFF000000FFFFFFFFFFFF0000"    ' 14:               000000               000000 0000FF 0000FF 0000FF 0000FF 0000FF 0000FF 000000               000000               0000
    HX = HX & "FFFFFFFFFFFF000000FFFFFFFFFFFF0000000000FF0000FF0000FF0000FF0000FF0000FF000000FFFFFFFFFFFF000000FFFFFFFFFFFF0000"    ' 13:               000000               000000 0000FF 0000FF 0000FF 0000FF 0000FF 0000FF 000000               000000               0000
    HX = HX & "FFFFFFFFFFFFFFFFFF000000FFFFFF0000FF0000FF0000000000000000000000000000FF0000FFFFFFFF000000FFFFFFFFFFFFFFFFFF0000"    ' 12:                      000000        0000FF 0000FF 000000 000000 000000 000000 0000FF 0000FF        000000                      0000
    HX = HX & "FFFFFFFFFFFFFFFFFFFFFFFF0000000000FF000000FFFFFFFFFFFFFFFFFFFFFFFF0000000000FF000000FFFFFFFFFFFFFFFFFFFFFFFF0000"    ' 11:                             000000 0000FF 000000                             000000 0000FF 000000                             0000
    HX = HX & "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF000000FFFFFF000000FFFFFFFFFFFF000000FFFFFF000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0000"    ' 10:                                    000000        000000               000000        000000                                    0000
    HX = HX & "FFFFFFFFFFFFFFFFFFFFFFFF000000FFFFFFFFFFFF000000FFFFFFFFFFFF000000FFFFFFFFFFFF000000FFFFFFFFFFFFFFFFFFFFFFFF0000"    '  9:                             000000               000000               000000               000000                             0000
    HX = HX & "FFFFFFFFFFFFFFFFFF000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF000000FFFFFFFFFFFFFFFFFF0000"    '  8:                      000000                                                                       000000                      0000
    HX = HX & "FFFFFFFFFFFFFFFFFF000000FFFFFFFFFFFFFFFFFFFFFFFF00FFFF00FFFF00FFFF00FFFF00FFFFFFFFFF000000FFFFFFFFFFFFFFFFFF0000"    '  7:                      000000                             00FFFF 00FFFF 00FFFF 00FFFF 00FFFF        000000                      0000
    HX = HX & "FFFFFFFFFFFFFFFFFF00000000FFFFFFFFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF000000FFFFFFFFFFFFFFFFFF0000"    '  6:                      000000 00FFFF        00FFFF 00FFFF 00FFFF 00FFFF 00FFFF 00FFFF 00FFFF 00FFFF 000000                      0000
    HX = HX & "FFFFFFFFFFFFFFFFFFFFFFFF00000000FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF000000FFFFFFFFFFFFFFFFFFFFFFFF0000"    '  5:                             000000 00FFFF 00FFFF 00FFFF 00FFFF 00FFFF 00FFFF 00FFFF 00FFFF 000000                             0000
    HX = HX & "FFFFFFFFFFFFFFFFFFFFFFFF00000000FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF000000FFFFFFFFFFFFFFFFFFFFFFFF0000"    '  4:                             000000 00FFFF 00FFFF 00FFFF 00FFFF 00FFFF 00FFFF 00FFFF 00FFFF 000000                             0000
    HX = HX & "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF00000000000000FFFF00FFFF00FFFF00FFFF000000000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0000"    '  3:                                    000000 000000 00FFFF 00FFFF 00FFFF 00FFFF 000000 000000                                    0000
    HX = HX & "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF000000000000000000000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0000"    '  2:                                                  000000 000000 000000 000000                                                  0000
    HX = HX & "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0000"    '  1:                                                                                                                               0000
    HX = HX & "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0000"    '  0:                                                                                                                               0000
    
    ' Salvar arquivo bitmap 16.777.216 cores (*.bmp;*.dib).
    
    Open Project.ThisDocument.Path & "\~$bm24bits12.bmp" For Binary Access Write As #1
        For i = 0 To Len(HX) - 1 Step 2
            BT = BT & Chr(Val("&H" & Mid(HX, i + 1, 2)))
        Next
        Put #1, , BT
    Close #1
    
    ' Visualizar o arquivo bitmap.
    
    Me.Image1.Picture = LoadPicture(Project.ThisDocument.Path & "\~$bm24bits12.bmp")
    
End Sub
