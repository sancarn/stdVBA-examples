VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MyMap 
   Caption         =   "UserForm4"
   ClientHeight    =   7065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12840
   OleObjectBlob   =   "MyMap.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MyMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private wv As stdWebView

Private Sub UserForm_Initialize()
    EnsureFeaturesTable

    Set wv = stdWebView.CreateFromUserform(Me)

    Dim html As String
    html = ""
    html = html & "<!doctype html>"
    html = html & "<html>"
    html = html & "<head>"
    html = html & "<meta charset=""utf-8"">"
    html = html & "<title>Leaflet Map Editor</title>"
    html = html & "<meta name=""viewport"" content=""width=device-width, initial-scale=1"">"
    html = html & "<link rel=""stylesheet"" href=""https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"">"
    html = html & "<link rel=""stylesheet"" href=""https://unpkg.com/leaflet-draw@1.0.4/dist/leaflet.draw.css"">"
    html = html & "<style>"
    html = html & "html,body{height:100%;margin:0;padding:0;background:#1e1e1e;color:#eee;font:14px Segoe UI,Arial}"
    html = html & "#map{position:absolute;top:0;left:0;right:0;bottom:0}"
    html = html & ".leaflet-container{background:#1e1e1e}"
    html = html & ".popup-btn{display:inline-block;margin-top:6px;padding:6px 10px;background:#c62828;color:#fff;border:none;border-radius:3px;cursor:pointer}"
    html = html & "</style>"
    html = html & "</head>"
    html = html & "<body>"
    html = html & "<div id=""map""></div>"
    html = html & "<script src=""https://unpkg.com/leaflet@1.9.4/dist/leaflet.js""></script>"
    html = html & "<script src=""https://unpkg.com/leaflet-draw@1.0.4/dist/leaflet.draw.js""></script>"
    html = html & "<script>"
    html = html & "const host = chrome.webview.hostObjects.map;"
    html = html & "const mapView = L.map('map',{zoomControl:true}).setView([51.505,-0.09],13);"
    html = html & "L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{maxZoom:19,attribution:'&copy; OpenStreetMap contributors'}).addTo(mapView);"
    html = html & "const drawnItems = new L.FeatureGroup();"
    html = html & "mapView.addLayer(drawnItems);"
    html = html & "const drawControl = new L.Control.Draw({edit:{featureGroup:drawnItems,remove:true},draw:{circle:false,circlemarker:false}});"
    html = html & "mapView.addControl(drawControl);"

    html = html & "function popupHtml(id){"
    html = html & "  return '<div><strong>Feature</strong><br>id: ' + id + '<br><button class=""popup-btn"" data-role=""delete"">Delete</button></div>';"
    html = html & "}"

    html = html & "function bindLayer(layer, feature){"
    html = html & "  layer.feature = feature;"
    html = html & "  const id = feature.id || '';"
    html = html & "  layer.bindPopup(popupHtml(id));"
    html = html & "  layer.on('popupopen', function(){"
    html = html & "    const btn = document.querySelector('[data-role=""delete""]');"
    html = html & "    if(btn){"
    html = html & "      btn.onclick = function(){"
    html = html & "        if(!layer.feature || !layer.feature.id){ alert('Missing id'); return; }"
    html = html & "        const ok = host.DeleteFeatureById(String(layer.feature.id));"
    html = html & "        if(ok){"
    html = html & "          mapView.closePopup();"
    html = html & "          drawnItems.removeLayer(layer);"
    html = html & "        } else {"
    html = html & "          alert('Delete failed');"
    html = html & "        }"
    html = html & "      };"
    html = html & "    }"
    html = html & "  });"
    html = html & "}"

    html = html & "async function loadAll(){"
    html = html & "  try{"
    html = html & "    const raw = await host.GetAllFeaturesJson();"
    html = html & "    console.log({raw});"
    html = html & "    if(!raw){ return; }"
    html = html & "    const fc = JSON.parse(raw);"
    html = html & "    L.geoJSON(fc,{"
    html = html & "      onEachFeature:function(feature, layer){"
    html = html & "        drawnItems.addLayer(layer);"
    html = html & "        bindLayer(layer, feature);"
    html = html & "      }"
    html = html & "    });"
    html = html & "  }catch(err){"
    html = html & "    console.error(err);"
    html = html & "    alert('Failed to load features: ' + err.message);"
    html = html & "  }"
    html = html & "}"

    html = html & "loadAll();"

    html = html & "mapView.on(L.Draw.Event.CREATED, function(e){"
    html = html & "  try{"
    html = html & "    const layer = e.layer;"
    html = html & "    const feature = layer.toGeoJSON();"
    html = html & "    const newId = String(host.AddFeature(JSON.stringify(feature)));"
    html = html & "    feature.id = newId;"
    html = html & "    feature.properties = feature.properties || {};"
    html = html & "    host.UpdateFeature(newId, JSON.stringify(feature));"
    html = html & "    drawnItems.addLayer(layer);"
    html = html & "    bindLayer(layer, feature);"
    html = html & "  }catch(err){"
    html = html & "    console.error(err);"
    html = html & "    alert('Add failed: ' + err.message);"
    html = html & "  }"
    html = html & "});"

    html = html & "mapView.on(L.Draw.Event.EDITED, function(e){"
    html = html & "  e.layers.eachLayer(function(layer){"
    html = html & "    try{"
    html = html & "      if(!layer.feature || !layer.feature.id){ return; }"
    html = html & "      const feature = layer.toGeoJSON();"
    html = html & "      feature.id = String(layer.feature.id);"
    html = html & "      feature.properties = feature.properties || {};"
    html = html & "      host.UpdateFeature(feature.id, JSON.stringify(feature));"
    html = html & "      bindLayer(layer, feature);"
    html = html & "    }catch(err){"
    html = html & "      console.error(err);"
    html = html & "    }"
    html = html & "  });"
    html = html & "});"

    html = html & "mapView.on(L.Draw.Event.DELETED, function(e){"
    html = html & "  e.layers.eachLayer(function(layer){"
    html = html & "    try{"
    html = html & "      if(layer.feature && layer.feature.id){"
    html = html & "        host.DeleteFeatureById(String(layer.feature.id));"
    html = html & "      }"
    html = html & "    }catch(err){"
    html = html & "      console.error(err);"
    html = html & "    }"
    html = html & "  });"
    html = html & "});"

    html = html & "</script>"
    html = html & "</body>"
    html = html & "</html>"

    wv.html = html
    wv.AddHostObject "map", Me
    stdWindow.CreateFromIUnknown(Me).isResizable = True
End Sub

Private Sub UserForm_Resize()
    If Not wv Is Nothing Then wv.Resize
End Sub

Public Function GetAllFeaturesJson() As String
    Dim lo As ListObject
    Dim r As ListRow
    Dim sb As String
    Dim first As Boolean

    Set lo = GetFeaturesListObject()

    sb = "{""type"":""FeatureCollection"",""features"":["
    first = True

    For Each r In lo.ListRows
        Dim fj As String
        fj = Trim$(CStr(r.Range(1, 2).value))

        If Len(fj) > 0 Then
            If Not first Then sb = sb & ","
            sb = sb & fj
            first = False
        End If
    Next

    sb = sb & "]}"
    GetAllFeaturesJson = sb
End Function

Public Function AddFeature(ByVal featureJson As String) As String
    Dim lo As ListObject
    Dim r As ListRow
    Dim idv As String

    Set lo = GetFeaturesListObject()
    idv = NewId()

    Set r = lo.ListRows.Add
    r.Range(1, 1).value = idv
    r.Range(1, 2).value = featureJson

    AddFeature = idv
End Function

Public Function UpdateFeature(ByVal idv As String, ByVal featureJson As String) As Boolean
    Dim row As ListRow

    Set row = FindRowById(idv)

    If row Is Nothing Then
        UpdateFeature = False
    Else
        row.Range(1, 2).value = featureJson
        UpdateFeature = True
    End If
End Function

Public Function DeleteFeatureById(ByVal idv As String) As Boolean
    Dim row As ListRow

    Set row = FindRowById(idv)

    If row Is Nothing Then
        DeleteFeatureById = False
    Else
        Application.EnableEvents = False
        row.Delete
        Application.EnableEvents = True
        DeleteFeatureById = True
    End If
End Function

Private Sub EnsureFeaturesTable()
    Dim ws As Worksheet
    Dim lo As ListObject

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Features")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = "Features"
    End If

    On Error Resume Next
    Set lo = ws.ListObjects("Features")
    On Error GoTo 0

    If lo Is Nothing Then
        ws.Range("A1").value = "Id"
        ws.Range("B1").value = "FeatureJson"
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B2"), , xlYes)
        lo.name = "Features"
    End If
End Sub

Private Function GetFeaturesListObject() As ListObject
    Set GetFeaturesListObject = ThisWorkbook.Worksheets("Features").ListObjects("Features")
End Function

Private Function FindRowById(ByVal idv As String) As ListRow
    Dim lo As ListObject
    Dim r As ListRow

    Set lo = GetFeaturesListObject()

    For Each r In lo.ListRows
        If CStr(r.Range(1, 1).value) = idv Then
            Set FindRowById = r
            Exit Function
        End If
    Next

    Set FindRowById = Nothing
End Function

Private Function NewId() As String
    Randomize
    NewId = Replace(CStr(Now * 86400000#) & "-" & CLng(Rnd() * 1000000000#), " ", "")
End Function
