VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Reversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type data
    id As Integer
    etapa As String
    serie As String
    uso As String
    estado As String
    proyecto As String
    numeroPartida As String
    resolucion As String
    expedienteHojaRuta As String
    anioExpendiente As Integer
    administrado As String
    dnis As String
    zona As String
    sector As String
    barrio As String
    grupoResidencial As Integer
    mz As String
    lote As Integer
    asuntoUtimoDocumento As String
    numeroFolio As Integer
    paquete As String
    ubicacionExpediente As String
    observacion As String
    profesional As String
    fechaActualizacion As Date
    rubroComercioActividad As String
    area As String
    contacto As String
    metro As Variant
End Type

Private this As data

Public Property Get id() As Integer
    id = this.id
End Property

Public Property Let id(ByVal id As Integer)
    this.id = id
End Property

Public Property Get etapa() As String
    etapa = this.etapa
End Property

Public Property Let etapa(ByVal etapa As String)
    this.etapa = UCase(etapa)
End Property
Public Property Get serie() As String
    serie = this.serie
End Property

Public Property Let serie(ByVal serie As String)
    this.serie = UCase(serie)
End Property

Public Property Get uso() As String
    uso = this.uso
End Property

Public Property Let uso(ByVal uso As String)
    this.uso = UCase(uso)
End Property

Public Property Get estado() As String
    estado = this.estado
End Property

Public Property Let estado(ByVal estado As String)
    this.estado = UCase(estado)
End Property

Public Property Get proyecto() As String
    proyecto = this.proyecto
End Property

Public Property Let proyecto(ByVal proyecto As String)
    this.proyecto = UCase(proyecto)
End Property

Public Property Get numeroPartida() As String
    numeroPartida = this.numeroPartida
End Property

Public Property Let numeroPartida(ByVal numeroPartida As String)
    this.numeroPartida = UCase(numeroPartida)
End Property

Public Property Get resolucion() As String
    resolucion = this.resolucion
End Property

Public Property Let resolucion(ByVal resolucion As String)
    this.resolucion = UCase(resolucion)
End Property

Public Property Get expedienteHojaRuta() As String
    expedienteHojaRuta = this.expedienteHojaRuta
End Property

Public Property Let expedienteHojaRuta(ByVal expedienteHojaRuta As String)
    this.expedienteHojaRuta = UCase(expedienteHojaRuta)
End Property

Public Property Get anioExpendiente() As Integer
    anioExpendiente = this.anioExpendiente
End Property

Public Property Let anioExpendiente(ByVal anioExpendiente As Integer)
    this.anioExpendiente = UCase(anioExpendiente)
End Property

Public Property Get administrado() As String
    administrado = this.administrado
End Property

Public Property Let administrado(ByVal administrado As String)
    this.administrado = UCase(administrado)
End Property

Public Property Get dnis() As String
    dnis = this.dnis
End Property

Public Property Let dnis(ByVal dnis As String)
    this.dnis = dnis
End Property

Public Property Get zona() As String
    zona = this.zona
End Property

Public Property Let zona(ByVal zona As String)
    this.zona = UCase(zona)
End Property

Public Property Get sector() As String
    sector = this.sector
End Property

Public Property Let sector(ByVal sector As String)
    this.sector = left(UCase(sector), 1)
End Property

Public Property Get barrio() As String
    barrio = this.barrio
End Property

Public Property Let barrio(ByVal barrio As String)
    this.barrio = left(UCase(barrio), 4)
End Property

Public Property Get grupoResidencial() As Integer
    grupoResidencial = this.grupoResidencial
End Property

Public Property Let grupoResidencial(ByVal grupoResidencial As Integer)
    this.grupoResidencial = grupoResidencial
End Property

Public Property Get mz() As Integer
    mz = this.mz
End Property

Public Property Let mz(ByVal mz As Integer)
    this.mz = mz
End Property

Public Property Get lote() As Integer
    lote = this.lote
End Property

Public Property Let lote(ByVal lote As Integer)
    this.lote = lote
End Property

Public Property Get asuntoUtimoDocumento() As String
    asuntoUtimoDocumento = this.asuntoUtimoDocumento
End Property

Public Property Let asuntoUtimoDocumento(ByVal asuntoUtimoDocumento As String)
    this.asuntoUtimoDocumento = UCase(asuntoUtimoDocumento)
End Property

Public Property Get numeroFolio() As Integer
    numeroFolio = this.numeroFolio
End Property

Public Property Let numeroFolio(ByVal numeroFolio As Integer)
    this.numeroFolio = UCase(numeroFolio)
End Property

Public Property Get paquete() As String
    paquete = this.paquete
End Property

Public Property Let paquete(ByVal paquete As String)
    this.paquete = UCase(paquete)
End Property

Public Property Get ubicacionExpediente() As String
    ubicacionExpediente = this.ubicacionExpediente
End Property

Public Property Let ubicacionExpediente(ByVal ubicacionExpediente As String)
    this.ubicacionExpediente = UCase(ubicacionExpediente)
End Property

Public Property Get observacion() As String
    observacion = this.observacion
End Property

Public Property Let observacion(ByVal observacion As String)
    this.observacion = UCase(observacion)
End Property

Public Property Get profesional() As String
    profesional = this.profesional
End Property

Public Property Let profesional(ByVal profesional As String)
    this.profesional = UCase(profesional)
End Property

Public Property Get fechaActualizacion() As Date
    fechaActualizacion = this.fechaActualizacion
End Property

Public Property Let fechaActualizacion(ByVal fechaActualizacion As Date)
    this.fechaActualizacion = fechaActualizacion
End Property

Public Property Get rubroComercioActividad() As String
    rubroComercioActividad = this.rubroComercioActividad
End Property

Public Property Let rubroComercioActividad(ByVal rubroComercioActividad As String)
    this.rubroComercioActividad = UCase(rubroComercioActividad)
End Property

Public Property Get area() As String
    area = this.area
End Property

Public Property Let area(ByVal area As String)
    this.area = UCase(area)
End Property

Public Property Get contacto() As String
    contacto = this.contacto
End Property

Public Property Let contacto(ByVal contacto As String)
    this.contacto = UCase(contacto)
End Property

Public Property Get metro() As Variant
    metro = this.metro & "M2"
End Property

Public Property Let metro(ByVal metro As Variant)
    this.metro = metro
End Property







