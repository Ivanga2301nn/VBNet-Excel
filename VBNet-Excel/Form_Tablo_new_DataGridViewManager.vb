Imports System.Drawing
Imports System.Windows.Forms

Public Class DataGridViewManager
    ' Всички технически каталози
    Private ReadOnly _disconnectorCatalog As DisconnectorCatalog
    Private ReadOnly _breakerCatalog As BreakerCatalog
    Private ReadOnly _cableCatalog As CableCatalog
    Private ReadOnly _rcdCatalog As RCDCatalog

    ''' <summary>
    ''' Конструктор на мениджъра за DataGridView.
    ''' </summary>
    Public Sub New(ByVal disconnectorCat As DisconnectorCatalog,
                   ByVal breakerCat As BreakerCatalog,
                   ByVal cableCat As CableCatalog,
                   ByVal rcdCat As RCDCatalog)

        ' Записваме референциите към данните и каталозите
        Me._disconnectorCatalog = disconnectorCat
        Me._breakerCatalog = breakerCat
        Me._cableCatalog = cableCat
        Me._rcdCatalog = rcdCat
    End Sub




End Class