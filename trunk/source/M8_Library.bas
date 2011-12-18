Attribute VB_Name = "M8_Library"
Option Explicit
Dim mu_tmp As Double, sd_tmp As Double
'Listado de funciones, rutinas a usar i.e. random generations

Public Function Cumd_Norm(x)
Attribute Cumd_Norm.VB_ProcData.VB_Invoke_Func = " \n14"
'This function calculates the cumulative of a standardized (u=0,sd=1) normal curve
Dim u, b1, b2, b3, b4, b5, p, w, y, pcul As Double
  b1 = 0.31938153
  b2 = -0.356563782
  b3 = 1.781477937
  b4 = -1.821255978
  b5 = 1.330274429
  p = 0.2316419
  
  If x >= 0 Then
    u = 1 / (1 + p * x)
    y = ((((b5 * u + b4) * u + b3) * u + b2) * u + b1) * u
    pcul = 1 - 0.3989422804 * Exp(-0.5 * x * x) * y

  Cumd_Norm = pcul
  
  Else
  
     w = -x
     u = 1 / (1 + p * w)
     y = ((((b5 * u + b4) * u + b3) * u + b2) * u + b1) * u
     pcul = 0.3989422804 * Exp(-0.5 * x * x) * y
    
  Cumd_Norm = pcul
End If

End Function

Public Function normal(mean, stdev) As Double
Attribute normal.VB_ProcData.VB_Invoke_Func = " \n14"

Dim iset As Long
Dim v1, v2, r, gset, fac, gasdev As Double
    iset = 0
    If iset = 0 Then
stm1:           v1 = 2# * Rnd - 1#
                v2 = 2# * Rnd - 1#
                r = v1 ^ 2 + v2 ^ 2
                
        If r > 1 Then GoTo stm1
                fac = Sqr(-2# * Log(r) / r)
                gset = v1 * fac
                gasdev = v2 * fac
                iset = 1
    Else
        gasdev = gset
        iset = 0
    End If
    normal = mean + gasdev * stdev
End Function

Sub Norm(Area, age)
Attribute Norm.VB_ProcData.VB_Invoke_Func = " \n14"

Dim ilen As Integer, intfact As Double
    
mu_tmp = muTmp(Area, age)
sd_tmp = sdTmp(Area, age)

intfact = 0
For ilen = 1 To Nilens
    pLage(Area, age, ilen) = Exp(-0.5 * ((l(ilen) - mu_tmp) / sd_tmp) ^ 2)
    intfact = intfact + pLage(Area, age, ilen)
Next
    
For ilen = 1 To Nilens
    pLage(Area, age, ilen) = pLage(Area, age, ilen) / intfact
Next

End Sub

Sub Trunc_Norm(Area, age)
Attribute Trunc_Norm.VB_ProcData.VB_Invoke_Func = " \n14"
Dim ilen As Integer, ifrac As Integer, ilenlast As Integer
Dim nextL As Double, intfact As Double

mu_tmp = muTmp(Area, age)
sd_tmp = sdTmp(Area, age)

For ilen = 1 To Nilens
    pLage(Area, age, ilen) = Exp(-0.5 * ((l(ilen) - mu_tmp) / sd_tmp) ^ 2)
Next

ifrac = 1
ilenlast = iLfull(Area)

While ifrac < Nfracs(Area, age) And ilen <= Nilens
    
    ilen = ilenlast
               
     nextL = sd_tmp * Z(Area, age, ifrac + 1) + mu_tmp
     
     While l(ilen) <= nextL
        pLage(Area, age, ilen) = pLage(Area, age, ilen) * frac(Area, age, ifrac)
        ilen = ilen + 1
     Wend
     ilenlast = ilen
   
     ifrac = ifrac + 1

Wend

' para ifrac = Nfracs(Area, age) - agregado con Ines
If (Nfracs(Area, age) > 0) Then
   For ilen = ilenlast To Nilens
   
     pLage(Area, age, ilen) = pLage(Area, age, ilen) * frac(Area, age, ifrac)

   Next
End If


intfact = 0
For ilen = 1 To Nilens
    intfact = intfact + pLage(Area, age, ilen)
Next ilen

For ilen = 1 To Nilens
    pLage(Area, age, ilen) = pLage(Area, age, ilen) / intfact
Next ilen

End Sub

Sub QuickSort(List() As Integer)
Attribute QuickSort.VB_ProcData.VB_Invoke_Func = " \n14"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'receives Vector and sorts it if lenght of vector <9 does a Bubble sort, if >9 does QuickSort

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim i As Integer, j As Integer, b As Integer
    Dim l As Integer, t As Integer, r As Integer, d As Integer, k As Integer, comp As Long, swic As Integer
    Dim oldx1 As Variant, oldy1, oldx2, oldy2, newx1, newx2, newy1, newy2 As Variant

    Dim p(1 To 100) As Integer
    Dim w(1 To 100) As Integer

    k = 1
    p(k) = LBound(List)
    w(k) = UBound(List)
    l = 1
    d = 1
    r = UBound(List)
    Do
toploop:
        If r - l < 9 Then GoTo bubsort
        i = l
        j = r
        While j > i
           comp = comp + 1
           If List(i) > List(j) Then
               swic = swic + 1
               t = List(j)
               oldx1 = List(j)
               oldy1 = j
               List(j) = List(i)
               oldx2 = List(i)
               oldy2 = i
               newx1 = List(j)
               newy1 = j
               List(i) = t
               newx2 = List(i)
               newy2 = i
               d = -d
           End If
           If d = -1 Then
               j = j - 1
                Else
                    i = i + 1
           End If
       Wend
           j = j + 1
           k = k + 1
            If i - l < r - j Then
                p(k) = j
                w(k) = r
                r = i
                Else
                    p(k) = l
                    w(k) = i
                    l = j
            End If
            d = -d
            GoTo toploop
bubsort:
    If r - l > 0 Then
        For i = l To r
            b = i
            For j = b + 1 To r
                comp = comp + 1
                If List(j) <= List(b) Then b = j
            Next j
            If i <> b Then
                swic = swic + 1
                t = List(b)
                oldx1 = List(b)
                oldy1 = b
                List(b) = List(i)
                oldx2 = List(i)
                oldy2 = i
                newx1 = List(b)
                newy1 = b
                List(i) = t
                newx2 = List(i)
                newy2 = i
            End If
        Next i
    End If
    l = p(k)
    r = w(k)
    k = k - 1
    Loop Until k = 0
End Sub
Sub order(List, indices) 'Ordena List de mayor a menor

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'receives Vector and creates a sorting index - if lenght of vector <9 does a Bubble sort, if >9 does QuickSort

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim i As Integer, j As Integer, b As Integer
    Dim l As Integer, r As Integer, d As Integer, k As Integer
    Dim t As Double
    Dim indt As Integer
        
    Dim p(1 To 100) As Integer ' ¿Restringidas las dimensiones a 100?
    Dim w(1 To 100) As Integer

    k = 1
    p(k) = LBound(List, 1) 'Devuelve el menor subíndice disponible para la dimensión indicada de una matriz.
    w(k) = UBound(List, 1) 'Devuelve el mayor subíndice disponible para la dimensión indicada de una matriz.
    l = 1
    d = 1
    r = UBound(List, 1) 'Dimensiones del vector
    
      
    Do
toploop:
        If r - l < 9 Then GoTo bubsort 'Si vector <9 salta esta parte y ve al bubsort
        i = l
        j = r
        While j > i                     'Va del primero al último
           If List(i) > List(j) Then  'Mira si el elemento del principio es más pequeño que el del final y:
               indt = indices(j)            'Ultimo elemento se coloca en t y en oldx1
               t = List(j)
               indices(j) = indices(i)      'El elemento j pasa a ser el elemento i
               List(j) = List(i)
               indices(i) = indt            'El elemento j antiguo pasa a ser el nuevo elemento i
               List(i) = t
               d = -d                   'd cambia de signo
           End If
           If d = -1 Then
               j = j - 1                'si el último elemento era menor que el primero, entonces pasas al penúltimo elemento
                Else
                    i = i + 1           'sino pasas al segundo
           End If
       Wend                             '''''''''''''''''''''''''
           j = j + 1
           k = k + 1
            If i - l < r - j Then   'Si se igualaron en los primeros elementos (hubo muchos cambios)
                p(k) = j            'j pasa a ser p(k) y r w(k). Se van a volver a ordenar los primeros elementos
                w(k) = r
                r = i               'r (que antes era la dimensión del vector) pasa a ser i
                Else                'Si se igualaron en los ultimos elementos (hubo pocos cambios
                    p(k) = l        'l pasa a ser j. Se igualan los ultimos elementos
                    w(k) = i
                    l = j
            End If
            d = -d                  'invertir de nuevo el signo de d
            GoTo toploop            'Se repite el proceso hasta que es subvector es menor que 9
bubsort:
    If r - l > 0 Then               'si hay más de un elemento en el subvector
        For i = l To r              'para todos los elementos del subvector
            b = i                   'donde b es el primer elemento y sucesivos
            For j = b + 1 To r      'desde el elemento siguiente a b hasta el final del subvector
               If List(j) >= List(b) Then b = j  'Mira si el orden está invertido
            Next j
            If i <> b Then          'Si b está descolocado(invertido), entonces lo coloca el su indice correspondiente 'j'
                indt = indices(b)
                t = List(b)
                indices(b) = indices(i)
                List(b) = List(i)
                indices(i) = indt
                List(i) = t
            End If
        Next i
    End If
    l = p(k)
    r = w(k)
    k = k - 1
    Loop Until k = 0
      
End Sub




Sub RandomizeVector(Xj() As Integer)
Attribute RandomizeVector.VB_ProcData.VB_Invoke_Func = " \n14"

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Randomizes a Vector - requires passing a vector
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim p As Integer, a() As Double, i As Integer

NI = UBound(Xj)

ReDim a(NI)

    For i = NI To 1 Step -1
        p = Int(i * Rnd + 1)
        a(i) = Xj(i)
        Xj(i) = Xj(p)
        Xj(p) = a(i)
    Next i

End Sub



Public Function Multiplyv2(Mat1() As Variant, Mat2() As Variant) As Variant
Attribute Multiplyv2.VB_ProcData.VB_Invoke_Func = " \n14"
        
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Multiply two matrices, their dimensions should be compatible!, Only the matrices have to be passed (not their dimensions)
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim l, i, j As Integer
        Dim OptiString As String
        Dim sol() As Double, MulAdd As Double
        
        MulAdd = 0
        
        Rows1 = UBound(Mat1, 1)
        Rows2 = UBound(Mat2, 1)
        Cols1 = UBound(Mat1, 2)
        Cols2 = UBound(Mat2, 2)
        
        ReDim sol(Rows1, Cols2)

        For i = 1 To Rows1
            For j = 1 To Cols2
                For l = 1 To Cols1
                    MulAdd = MulAdd + Mat1(i, l) * Mat2(l, j)
                Next l
                sol(i, j) = MulAdd
                MulAdd = 0
            Next j
        Next i

    Multiplyv2 = sol
        
End Function


