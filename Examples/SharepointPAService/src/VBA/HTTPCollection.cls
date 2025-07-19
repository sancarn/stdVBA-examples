VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HTTPCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type TThis
    col as Collection
    responseMapper as stdICallable
End Type
Private This as TThis

'Create a new HTTPCollection object.
'@return - A new HTTPCollection object
Public Function Create() as HTTPCollection
  set Create = new HTTPCollection
  Call Create.protInit()
End Function

'Initialize the HTTPCollection object.
'@protected
Public Sub protInit()
  set this.col = new Collection
End Sub

'Add a new HTTP object to the collection.
Public Sub add(ByVal http as stdHTTP)
  Call this.col.add(http)
End Sub

'Await for all HTTP requests in the collection to complete.
'@return - The HTTPCollection object itself
Public Function Await() as HTTPCollection
  Dim http as stdHTTP
  For each http in this.col
    Call http.Await()
  Next http
  Set Await = Me
End Function

'Get/Set the `responseMapper` which is utilised by the `responseMapped` property.
'@return - The responseMapper callable object
Public Property Get responseMapper() as stdICallable
  set responseMapper = this.responseMapper
End Property
Public Property Set responseMapper(v as stdICallable)
  set this.responseMapper = v
End Property


'Uses reesponseMapper to return some mapped value from the collection.
'@return - The mapped value from the collection
Public Property Get ResponseMapped() as variant
  Call CopyVariant(ResponseMapped, this.responseMapper.run(this.col))
End Property

'Copies the value of src to dest. If src is an object, it sets dest to the same object reference.
'@param dest - The destination to copy the value to
'@param src - The source to copy the value from
Private Sub CopyVariant(ByRef dest, ByVal src)
  if isObject(src) then
    set dest = src
  else
    let dest = src
  end if
End Sub
