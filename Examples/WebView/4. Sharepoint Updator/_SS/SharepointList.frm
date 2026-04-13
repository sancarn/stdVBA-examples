VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sharepoint 
   Caption         =   "Sharepoint Updator"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6030
   OleObjectBlob   =   "SharepointList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SharepointList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const DEFAULT_READY_TIMEOUT_MS As Long = 180000
Private Const READY_STABLE_MS As Long = 500
Private Const READY_POLL_MS As Long = 200

Private Type TThis
    wv As stdWebView
    site As String
    siteBase As String
    listTitle As String
    readyTimeoutMs As Long
    runtimeInjected As Boolean
End Type
Private This As TThis

'Creates an authenticated SharePoint list client userform.
'@constructor
'@param site - SharePoint URL (site/list/allitems) used for interactive sign-in.
'@param listTitle - Optional list title override. If omitted, runtime infers from page context.
'@param readyTimeoutMs - Max milliseconds to wait for stable authenticated page readiness.
'@returns - Hidden-but-live instance retaining authenticated browser session.
'@example `set sharepoint = Sharepoint.Create("https://contoso.sharepoint.com/sites/MySite/Lists/MyList/AllItems.aspx")`
Public Function Create( _
    ByVal site As String, _
    Optional ByVal listTitle As String = vbNullString, _
    Optional ByVal readyTimeoutMs As Long = DEFAULT_READY_TIMEOUT_MS) As SharepointList

    Dim instance As SharepointList
    Set instance = New SharepointList
    Call instance.protInit(site, listTitle, readyTimeoutMs)
    Set Create = instance
End Function

'Initialise the userform and block until authenticated page readiness.
'@protected
'@param site - SharePoint URL (site/list/allitems) used for interactive sign-in.
'@param listTitle - Optional list title override.
'@param readyTimeoutMs - Max milliseconds to wait for stable authenticated page readiness.
Public Sub protInit( _
    ByVal site As String, _
    Optional ByVal listTitle As String = vbNullString, _
    Optional ByVal readyTimeoutMs As Long = DEFAULT_READY_TIMEOUT_MS)
    This.readyTimeoutMs = IIf(readyTimeoutMs > 0, readyTimeoutMs, DEFAULT_READY_TIMEOUT_MS)

    Set This.wv = stdWebView.CreateFromUserform(Me)
    Call NavigateToSite(site, listTitle)
End Sub

'Navigate the existing authenticated WebView session to another site/list URL.
'@param site - Target SharePoint URL.
'@param listTitle - Optional list title override for the new context.
Public Sub NavigateToSite(ByVal site As String, Optional ByVal listTitle As String = vbNullString)
    If LenB(Trim$(site)) = 0 Then
        Err.Raise 5, "SharepointList::NavigateToSite", "site cannot be blank"
    End If
    If This.wv Is Nothing Then
        Err.Raise 5, "SharepointList::NavigateToSite", "WebView session is not initialized. Call Create first."
    End If
    Call ActivateSiteContext(site, listTitle)
End Sub

'Release hosted WebView resources.
'@returns - None.
Public Sub Quit()
    On Error Resume Next
    If Not This.wv Is Nothing Then This.wv.Quit
End Sub

'Create a single SharePoint list item.
'@param payload as Variant<stdJSON<{colName1:value1,...}> | Array<header1:string,value1:variant,header2:string,value2:variant,...>> - Item fields payload.
'@param listTitle - Optional list title override for this call.
'@returns as stdJSON<{...}> - Parsed SharePoint response payload.
Public Function ItemCreate(ByVal payload As Variant, Optional ByVal listTitle As String = vbNullString) As stdJSON
    Dim response As stdJSON
    Set response = ExecuteRuntimeMethod("itemCreate", BuildArgsJson(NormalizeObjectPayloadJson(payload), JsonNullableString(listTitle)))
    Set ItemCreate = UnwrapResponseDataJson(response)
End Function

'Update a single SharePoint list item by ID.
'@param itemId - SharePoint list item ID.
'@param payload as Variant<stdJSON<{colName1:value1,...}> | Array<header1:string,value1:variant,header2:string,value2:variant,...>> - Item fields payload.
'@param listTitle - Optional list title override for this call.
'@returns as stdJSON<{id:number,status:number}> - Update status envelope.
Public Function ItemUpdate(ByVal itemId As Long, ByVal payload As Variant, Optional ByVal listTitle As String = vbNullString) As stdJSON
    Dim response As stdJSON
    Set response = ExecuteRuntimeMethod("itemUpdate", BuildArgsJson(CStr(itemId), NormalizeObjectPayloadJson(payload), JsonNullableString(listTitle)))
    Set ItemUpdate = UnwrapResponseDataJson(response)
End Function

'Delete a single SharePoint list item by ID.
'@param itemId - SharePoint list item ID.
'@param listTitle - Optional list title override for this call.
'@returns as stdJSON<{id:number,status:number}> - Delete status envelope.
Public Function ItemDelete(ByVal itemId As Long, Optional ByVal listTitle As String = vbNullString) As stdJSON
    Dim response As stdJSON
    Set response = ExecuteRuntimeMethod("itemDelete", BuildArgsJson(CStr(itemId), JsonNullableString(listTitle)))
    Set ItemDelete = UnwrapResponseDataJson(response)
End Function

'Create many SharePoint list items.
'@param payloads as stdJSON<Array<stdJSON<{colName1:value1,...}>>> - Create payloads.
'@param listTitle - Optional list title override for this call.
'@returns as stdJSON<Array<{ok:boolean,type:"create",data?:object,error?:string,payload?:any}>> - Per-item create results.
Public Function BatchItemsCreate(ByVal payloads As stdJSON, Optional ByVal listTitle As String = vbNullString) As stdJSON
    Dim response As stdJSON
    Set response = ExecuteRuntimeMethod("batchItemsCreate", BuildArgsJson(RequireStdJsonArrayPayload(payloads, "BatchItemsCreate"), JsonNullableString(listTitle)))
    Set BatchItemsCreate = UnwrapResponseDataJson(response)
End Function

'Update many SharePoint list items.
'@param payloads as stdJSON<Array<{id:number,data:stdJSON<{colName1:value1,...}>}>> - Update payloads.
'@param listTitle - Optional list title override for this call.
'@returns as stdJSON<Array<{ok:boolean,type:"update",id?:number,data?:object,error?:string,payload?:any}>> - Per-item update results.
Public Function BatchItemsUpdate(ByVal payloads As stdJSON, Optional ByVal listTitle As String = vbNullString) As stdJSON
    Dim response As stdJSON
    Set response = ExecuteRuntimeMethod("batchItemsUpdate", BuildArgsJson(RequireStdJsonArrayPayload(payloads, "BatchItemsUpdate"), JsonNullableString(listTitle)))
    Set BatchItemsUpdate = UnwrapResponseDataJson(response)
End Function

'Delete many SharePoint list items.
'@param ids as stdJSON<Array<number>> - Item IDs to delete.
'@param listTitle - Optional list title override for this call.
'@returns as stdJSON<Array<{ok:boolean,type:"delete",id?:number,data?:object,error?:string}>> - Per-item delete results.
Public Function BatchItemsDelete(ByVal ids As stdJSON, Optional ByVal listTitle As String = vbNullString) As stdJSON
    Dim response As stdJSON
    Set response = ExecuteRuntimeMethod("batchItemsDelete", BuildArgsJson(RequireStdJsonArrayPayload(ids, "BatchItemsDelete"), JsonNullableString(listTitle)))
    Set BatchItemsDelete = UnwrapResponseDataJson(response)
End Function

'Process grouped commit payloads.
'@param payloads as stdJSON<{update:Array<{id:number,data:stdJSON<{colName1:value1,...}>}>,create:Array<stdJSON<{colName1:value1,...}>>,delete:Array<number | {id:number}>}> - Commit payload object.
'@param listTitle - Optional list title override for this call.
'@returns as stdJSON<{create:Array<any>,update:Array<any>,delete:Array<any>}> - Grouped operation results.
Public Function ProcessCommitPayloads(ByVal payloads As stdJSON, Optional ByVal listTitle As String = vbNullString) As stdJSON
    Dim response As stdJSON
    Set response = ExecuteRuntimeMethod("processCommitPayloads", BuildArgsJson(RequireStdJsonObjectPayload(payloads, "ProcessCommitPayloads"), JsonNullableString(listTitle)))
    Set ProcessCommitPayloads = UnwrapResponseDataJson(response)
End Function

Private Sub UserForm_Terminate()
    On Error Resume Next
    Call Quit
End Sub

Private Sub ActivateSiteContext(ByVal site As String, Optional ByVal listTitle As String = vbNullString)
    This.site = Trim$(site)
    This.siteBase = InferSiteBase(This.site)
    This.listTitle = Trim$(listTitle)
    This.runtimeInjected = False

    Call This.wv.Navigate(This.site)
    Call Me.Show(False)
    If Not This.wv.WaitForDocumentReady(This.siteBase, This.readyTimeoutMs, READY_STABLE_MS, READY_POLL_MS) Then
        Me.Hide
        Err.Raise 5, "SharepointList::ActivateSiteContext", "Timed out waiting for authenticated SharePoint page readiness"
    End If

    Call RefreshContextFromPage
    Call EnsureRuntimeInjected
    Me.Hide
End Sub

Private Function ExecuteRuntimeMethod(ByVal methodName As String, Optional ByVal argsJson As String = "[]") As stdJSON
    Dim script As String
    Dim responseRaw As String
    Dim responseText As String
    Dim envelope As stdJSON

    Call EnsureRuntimeInjected

    script = "(async function(){" _
        & "try{" _
            & "if(!window.__spUpdater||typeof window.__spUpdater." & methodName & "!=='function'){throw new Error('SharePoint runtime method not available: " & methodName & "');}" _
            & "const __args=" & argsJson & ";" _
            & "const __result=await window.__spUpdater." & methodName & ".apply(window.__spUpdater,__args);" _
            & "return JSON.stringify({ok:true,data:__result});" _
        & "}catch(e){" _
            & "return JSON.stringify({ok:false,error:(e&&e.message)?e.message:String(e)});" _
        & "}" _
    & "})();"

    responseRaw = This.wv.JavaScriptRunSync(script)
    responseText = JsonUnquoteString(responseRaw)
    Set envelope = stdJSON.CreateFromString(responseText)
    If Not envelope.Exists("ok") Or Not CBool(envelope("ok")) Then
        Err.Raise 5, "SharepointList::" & methodName, CStr(envelope("error"))
    End If
    Set ExecuteRuntimeMethod = envelope
End Function

Private Sub EnsureRuntimeInjected()
    If This.runtimeInjected Then Exit Sub
    Call This.wv.JavaScriptRunSync(BuildRuntimeBootstrapScript())
    This.runtimeInjected = True
End Sub

Private Sub RefreshContextFromPage()
    Dim responseText As String
    Dim contextJson As stdJSON
    Dim script As String

    script = "(function(){" _
        & "const p=window._spPageContextInfo||{};" _
        & "return JSON.stringify({" _
            & "webAbsoluteUrl:p.webAbsoluteUrl||'',siteAbsoluteUrl:p.siteAbsoluteUrl||'',listTitle:p.listTitle||''" _
        & "});" _
    & "})()"

    responseText = JsonUnquoteString(This.wv.JavaScriptRunSync(script))
    Set contextJson = stdJSON.CreateFromString(responseText)

    If contextJson.Exists("webAbsoluteUrl") Then
        If LenB(Trim$(CStr(contextJson("webAbsoluteUrl")))) > 0 Then This.siteBase = Trim$(CStr(contextJson("webAbsoluteUrl")))
    End If
    If LenB(This.siteBase) = 0 And contextJson.Exists("siteAbsoluteUrl") Then
        If LenB(Trim$(CStr(contextJson("siteAbsoluteUrl")))) > 0 Then This.siteBase = Trim$(CStr(contextJson("siteAbsoluteUrl")))
    End If

    If LenB(This.listTitle) = 0 And contextJson.Exists("listTitle") Then
        This.listTitle = Trim$(CStr(contextJson("listTitle")))
    End If
End Sub

Private Function BuildRuntimeBootstrapScript() As String
    Dim siteBaseJson As String
    Dim listTitleJson As String
    Dim script As String
    siteBaseJson = JsonNullableString(This.siteBase)
    listTitleJson = JsonNullableString(This.listTitle)
    Call ScriptAppendLine(script, "(function(){")
    Call ScriptAppendLine(script, "const cfg={siteBase:" & siteBaseJson & ",listTitle:" & listTitleJson & "};")
    Call ScriptAppendLine(script, "if(window.__spUpdater&&window.__spUpdater.__version==='1.0.0'){window.__spUpdater.config=cfg;return 'ok';}")
    Call ScriptAppendLine(script, "const digestCache={value:'',expiresAt:0};")
    Call ScriptAppendLine(script, "const ensureUserCache=new Map();")
    Call ScriptAppendLine(script, "const nowMs=()=>Date.now();")
    Call ScriptAppendLine(script, "const toObject=(payload)=>{if(payload===null||payload===undefined)return {};if(Array.isArray(payload)){if(payload.length%2!==0)throw new Error('Pair-array payload must contain key/value pairs');const o={};for(let i=0;i<payload.length;i+=2){o[String(payload[i])]=payload[i+1];}return o;}if(typeof payload==='object')return payload;throw new Error('Payload must be an object or key/value pair array');};")
    Call ScriptAppendLine(script, "const ensureUserId=async(userText)=>{if(userText===null||userText===undefined)return null;const raw=String(userText).trim();if(!raw)return null;const cacheKey=raw.toLowerCase();if(ensureUserCache.has(cacheKey))return ensureUserCache.get(cacheKey);const res=await api('/_api/web/ensureuser',{method:'POST',headers:{'Content-Type':'application/json;odata=verbose'},body:JSON.stringify({logonName:raw})});const id=Number(res.json&&res.json.d&&res.json.d.Id);if(!isFinite(id)||id<=0)throw new Error('Unable to ensure user: '+raw);ensureUserCache.set(cacheKey,id);return id;};")
    Call ScriptAppendLine(script, "const ensureUserIdFromAny=async(v)=>{if(v===null||v===undefined)return null;if(typeof v==='number')return v;if(typeof v==='string')return ensureUserId(v);if(typeof v==='object'){if(isFinite(Number(v.Id)))return Number(v.Id);if(isFinite(Number(v.id)))return Number(v.id);const probe=v.Email||v.EMail||v.LoginName||v.loginName||v.Name||v.name||v.Title||v.title;if(probe!==undefined&&probe!==null&&String(probe).trim())return ensureUserId(String(probe));}throw new Error('Unsupported person value');};")
    Call ScriptAppendLine(script, "const ensureUserIdArray=async(value)=>{if(value===null||value===undefined)return [];const arr=Array.isArray(value)?value:[value];const ids=[];for(const v of arr){const ensured=await ensureUserIdFromAny(v);if(ensured!==null&&ensured!==undefined)ids.push(ensured);}return ids;};")
    Call ScriptAppendLine(script, "const transformPayloadForSharePoint=async(payload)=>{const src=toObject(payload);const out={};for(const key of Object.keys(src)){const value=src[key];if(key.endsWith('Email')){const base=key.slice(0,-5);const uid=await ensureUserIdFromAny(value);out[base+'Id']=uid;continue;}if(key.endsWith('Emails')){const base=key.slice(0,-6);out[base+'Id']={results:await ensureUserIdArray(value)};continue;}if(key.endsWith('Id')&&Array.isArray(value)){out[key]={results:await ensureUserIdArray(value)};continue;}if(key.endsWith('Id')&&(typeof value==='string'||(value&&typeof value==='object'))){out[key]=await ensureUserIdFromAny(value);continue;}out[key]=value;}return out;};")
    Call ScriptAppendLine(script, "const resolveSiteBase=()=>cfg.siteBase||((window._spPageContextInfo&&(window._spPageContextInfo.webAbsoluteUrl||window._spPageContextInfo.siteAbsoluteUrl))||location.origin);")
    Call ScriptAppendLine(script, "const resolveListTitle=(override)=>{const t=(override||cfg.listTitle||(window._spPageContextInfo&&window._spPageContextInfo.listTitle)||'').trim();if(!t)throw new Error('List title is required. Pass it to Sharepoint.Create or per call.');return t;};")
    Call ScriptAppendLine(script, "const listPath=(listTitle)=>{const e=String(listTitle).replace(/'/g,""''"");return ""/_api/web/lists/GetByTitle('""+e+""')/items"";};")
    Call ScriptAppendLine(script, "const unwrapOData=(json)=>json&&json.d?json.d:json;")
    Call ScriptAppendLine(script, "const parseJsonSafe=(text)=>{if(!text)return null;try{return JSON.parse(text);}catch(_){return null;}};")
    Call ScriptAppendLine(script, "const api=async(relativePath,options)=>{const opts=options||{};const headers=Object.assign({'Accept':'application/json;odata=verbose'},opts.headers||{});const url=(/^https?:\/\//i.test(relativePath)?relativePath:resolveSiteBase()+relativePath);const resp=await fetch(url,{method:opts.method||'GET',headers:headers,body:opts.body,credentials:'include'});const text=await resp.text();const json=parseJsonSafe(text);if(!resp.ok){const serverMessage=json&&json.error&&json.error.message&&json.error.message.value;throw new Error(serverMessage||text||('HTTP '+resp.status));}return {status:resp.status,text:text,json:json};};")
    Call ScriptAppendLine(script, "const getDigest=async()=>{if(digestCache.value&&digestCache.expiresAt>nowMs())return digestCache.value;const res=await api('/_api/contextinfo',{method:'POST'});const info=res.json&&res.json.d&&res.json.d.GetContextWebInformation;const digest=info&&info.FormDigestValue;if(!digest)throw new Error('Unable to obtain SharePoint request digest');const timeoutSec=Number((info&&info.FormDigestTimeoutSeconds)||1200);digestCache.value=digest;digestCache.expiresAt=nowMs()+Math.max(10000,(timeoutSec-30)*1000);return digest;};")
    Call ScriptAppendLine(script, "const itemCreate=async(payload,listTitle)=>{const title=resolveListTitle(listTitle);const digest=await getDigest();const body=JSON.stringify(await transformPayloadForSharePoint(payload));const res=await api(listPath(title),{method:'POST',headers:{'Content-Type':'application/json;odata=nometadata','X-RequestDigest':digest},body:body});return unwrapOData(res.json);};")
    Call ScriptAppendLine(script, "const itemUpdate=async(itemId,payload,listTitle)=>{const id=Number(itemId);if(!isFinite(id)||id<=0)throw new Error('itemId must be a positive number');const title=resolveListTitle(listTitle);const digest=await getDigest();const body=JSON.stringify(await transformPayloadForSharePoint(payload));const res=await api(listPath(title)+'('+id+')',{method:'POST',headers:{'Content-Type':'application/json;odata=nometadata','X-RequestDigest':digest,'X-HTTP-Method':'MERGE','IF-MATCH':'*'},body:body});return {id:id,status:res.status};};")
    Call ScriptAppendLine(script, "const itemDelete=async(itemId,listTitle)=>{const id=Number(itemId);if(!isFinite(id)||id<=0)throw new Error('itemId must be a positive number');const title=resolveListTitle(listTitle);const digest=await getDigest();const res=await api(listPath(title)+'('+id+')',{method:'POST',headers:{'X-RequestDigest':digest,'X-HTTP-Method':'DELETE','IF-MATCH':'*'}});return {id:id,status:res.status};};")
    Call ScriptAppendLine(script, "const normalizeArray=(v,name)=>{if(!Array.isArray(v))throw new Error(name+' must be an array');return v;};")
    Call ScriptAppendLine(script, "const batchItemsCreate=async(payloads,listTitle)=>{const arr=normalizeArray(payloads,'payloads');const out=[];for(const payload of arr){try{const data=await itemCreate(payload,listTitle);out.push({ok:true,type:'create',data:data});}catch(e){out.push({ok:false,type:'create',error:(e&&e.message)?e.message:String(e),payload:payload});}}return out;};")
    Call ScriptAppendLine(script, "const batchItemsUpdate=async(payloads,listTitle)=>{const arr=normalizeArray(payloads,'payloads');const out=[];for(const entry of arr){try{if(!entry||typeof entry!=='object')throw new Error('Each update entry must be an object {id,data}');const data=await itemUpdate(entry.id,entry.data,listTitle);out.push({ok:true,type:'update',id:entry.id,data:data});}catch(e){out.push({ok:false,type:'update',id:entry&&entry.id,error:(e&&e.message)?e.message:String(e),payload:entry});}}return out;};")
    Call ScriptAppendLine(script, "const batchItemsDelete=async(ids,listTitle)=>{const arr=normalizeArray(ids,'ids');const out=[];for(const id of arr){try{const data=await itemDelete(id,listTitle);out.push({ok:true,type:'delete',id:id,data:data});}catch(e){out.push({ok:false,type:'delete',id:id,error:(e&&e.message)?e.message:String(e)});}}return out;};")
    Call ScriptAppendLine(script, "const processCommitPayloads=async(payloads,listTitle)=>{if(!payloads||typeof payloads!=='object'||Array.isArray(payloads))throw new Error('payloads must be an object with create/update/delete arrays');const creates=Array.isArray(payloads.create)?payloads.create:[];const updates=Array.isArray(payloads.update)?payloads.update:[];const deletesRaw=Array.isArray(payloads.delete)?payloads.delete:[];const deletes=[];for(const entry of deletesRaw){if(entry&&typeof entry==='object'&&entry.id!==undefined){deletes.push(entry.id);}else{deletes.push(entry);}}return {create:await batchItemsCreate(creates,listTitle),update:await batchItemsUpdate(updates,listTitle),delete:await batchItemsDelete(deletes,listTitle)};};")
    Call ScriptAppendLine(script, "window.__spUpdater={__version:'1.0.0',config:cfg,itemCreate,itemUpdate,itemDelete,batchItemsCreate,batchItemsUpdate,batchItemsDelete,processCommitPayloads};")
    Call ScriptAppendLine(script, "return 'ok';")
    Call ScriptAppendLine(script, "})();")
    BuildRuntimeBootstrapScript = script
End Function

Private Sub ScriptAppendLine(ByRef script As String, ByVal lineText As String)
    If LenB(script) = 0 Then
        script = lineText
    Else
        script = script & vbLf & lineText
    End If
End Sub

Private Function NormalizeObjectPayloadJson(ByVal payload As Variant) As String
    If IsArray(payload) And IsPairArray(payload) Then
        NormalizeObjectPayloadJson = PairArrayToObjectJson(payload)
    Else
        NormalizeObjectPayloadJson = NormalizeJsonLiteral(payload)
    End If
End Function

Private Function RequireStdJsonArrayPayload(ByVal payload As Variant, ByVal methodName As String) As String
    If Not IsObject(payload) Then
        Err.Raise 5, "SharepointList::" & methodName, methodName & " requires a stdJSON array payload"
    End If
    If TypeName(payload) <> "stdJSON" Then
        Err.Raise 5, "SharepointList::" & methodName, methodName & " requires a stdJSON array payload"
    End If
    If payload.JsonType <> eJSONArray Then
        Err.Raise 5, "SharepointList::" & methodName, methodName & " requires stdJSON of type array"
    End If
    RequireStdJsonArrayPayload = payload.ToString
End Function

Private Function RequireStdJsonObjectPayload(ByVal payload As Variant, ByVal methodName As String) As String
    If Not IsObject(payload) Then
        Err.Raise 5, "SharepointList::" & methodName, methodName & " requires a stdJSON object payload"
    End If
    If TypeName(payload) <> "stdJSON" Then
        Err.Raise 5, "SharepointList::" & methodName, methodName & " requires a stdJSON object payload"
    End If
    If payload.JsonType <> eJSONObject Then
        Err.Raise 5, "SharepointList::" & methodName, methodName & " requires stdJSON of type object"
    End If
    RequireStdJsonObjectPayload = payload.ToString
End Function

Private Function NormalizeJsonLiteral(ByVal value As Variant) As String
    If IsObject(value) Then
        If value Is Nothing Then
            NormalizeJsonLiteral = "null"
        ElseIf TypeName(value) = "stdJSON" Then
            NormalizeJsonLiteral = value.ToString
        Else
            NormalizeJsonLiteral = stdJSON.CreateFromVariant(value).ToString
        End If
        Exit Function
    End If

    If IsArray(value) Then
        NormalizeJsonLiteral = stdJSON.CreateFromVariant(value).ToString
        Exit Function
    End If

    Select Case VarType(value)
        Case vbEmpty, vbNull
            NormalizeJsonLiteral = "null"
        Case vbBoolean
            NormalizeJsonLiteral = IIf(CBool(value), "true", "false")
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            NormalizeJsonLiteral = Replace(CStr(value), ",", ".")
        Case vbString
            If IsLikelyJson(CStr(value)) Then
                NormalizeJsonLiteral = stdJSON.CreateFromString(CStr(value)).ToString
            Else
                NormalizeJsonLiteral = JsonQuote(CStr(value))
            End If
        Case Else
            NormalizeJsonLiteral = JsonQuote(CStr(value))
    End Select
End Function

Private Function UnwrapResponseDataJson(ByVal response As stdJSON) As stdJSON
    If response Is Nothing Then
        Set UnwrapResponseDataJson = stdJSON.Create(eJSONObject)
        Exit Function
    End If
    If Not response.Exists("data") Then
        Set UnwrapResponseDataJson = stdJSON.Create(eJSONObject)
        Exit Function
    End If
    If IsObject(response("data")) Then
        If TypeName(response("data")) = "stdJSON" Then
            Set UnwrapResponseDataJson = response("data")
            Exit Function
        End If
    End If

    Set UnwrapResponseDataJson = stdJSON.Create(eJSONObject)
    Call UnwrapResponseDataJson.Add("value", response("data"))
End Function

Private Function BuildArgsJson(ParamArray args() As Variant) As String
    Dim i As Long
    Dim parts() As String
    On Error GoTo noArgs
    If UBound(args) < LBound(args) Then
        BuildArgsJson = "[]"
        Exit Function
    End If

    ReDim parts(LBound(args) To UBound(args))
    For i = LBound(args) To UBound(args)
        parts(i) = CStr(args(i))
    Next
    BuildArgsJson = "[" & Join(parts, ",") & "]"
    Exit Function
noArgs:
    BuildArgsJson = "[]"
End Function

Private Function JsonNullableString(ByVal value As String) As String
    value = Trim$(value)
    If LenB(value) = 0 Then
        JsonNullableString = "null"
    Else
        JsonNullableString = JsonQuote(value)
    End If
End Function

Private Function JsonQuote(ByVal value As String) As String
    Dim s As String
    s = value
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    JsonQuote = """" & s & """"
End Function

Private Function JsonUnquoteString(ByVal s As String) As String
    Dim t As String
    Dim i As Long
    Dim c As String
    Dim out As String

    t = Trim$(s)
    If Len(t) < 2 Then
        JsonUnquoteString = t
        Exit Function
    End If
    If Left$(t, 1) <> """" Or Right$(t, 1) <> """" Then
        JsonUnquoteString = t
        Exit Function
    End If

    t = Mid$(t, 2, Len(t) - 2)
    i = 1
    Do While i <= Len(t)
        c = Mid$(t, i, 1)
        If c = "\" And i < Len(t) Then
            i = i + 1
            Select Case Mid$(t, i, 1)
                Case """": out = out & """"
                Case "\": out = out & "\"
                Case "n": out = out & vbLf
                Case "r": out = out & vbCr
                Case "t": out = out & vbTab
                Case Else: out = out & Mid$(t, i, 1)
            End Select
        Else
            out = out & c
        End If
        i = i + 1
    Loop
    JsonUnquoteString = out
End Function

Private Function IsLikelyJson(ByVal value As String) As Boolean
    value = Trim$(value)
    If LenB(value) = 0 Then Exit Function
    IsLikelyJson = (Left$(value, 1) = "{" And Right$(value, 1) = "}") _
        Or (Left$(value, 1) = "[" And Right$(value, 1) = "]")
End Function

Private Function IsPairArray(ByVal values As Variant) As Boolean
    On Error GoTo cleanFail
    Dim lb As Long
    Dim ub As Long
    lb = LBound(values)
    ub = UBound(values)
    If ((ub - lb + 1) Mod 2) <> 0 Then Exit Function
    IsPairArray = True
    Exit Function
cleanFail:
    IsPairArray = False
End Function

Private Function PairArrayToObjectJson(ByVal values As Variant) As String
    Dim json As stdJSON
    Dim i As Long
    Dim lb As Long
    Dim ub As Long

    If Not IsPairArray(values) Then
        Err.Raise 5, "SharepointList::PairArrayToObjectJson", "Pair-array payload must contain an even number of values"
    End If

    Set json = stdJSON.Create(eJSONObject)
    lb = LBound(values)
    ub = UBound(values)
    For i = lb To ub Step 2
        Call json.Add(CStr(values(i)), values(i + 1))
    Next
    PairArrayToObjectJson = json.ToString
End Function

Private Function InferSiteBase(ByVal rawUrl As String) As String
    Dim s As String
    Dim protocolEnd As Long
    Dim pathStart As Long
    Dim hostPart As String
    Dim pathPart As String
    Dim firstSeg As String
    Dim secondSlash As Long

    s = Trim$(rawUrl)
    If LenB(s) = 0 Then Exit Function

    protocolEnd = InStr(1, s, "://", vbTextCompare)
    If protocolEnd <= 0 Then
        InferSiteBase = s
        Exit Function
    End If

    pathStart = InStr(protocolEnd + 3, s, "/")
    If pathStart <= 0 Then
        InferSiteBase = s
        Exit Function
    End If

    hostPart = Left$(s, pathStart - 1)
    pathPart = Mid$(s, pathStart)
    If Left$(LCase$(pathPart), 7) = "/sites/" Or Left$(LCase$(pathPart), 7) = "/teams/" Then
        secondSlash = InStr(8, pathPart, "/")
        If secondSlash > 0 Then
            firstSeg = Left$(pathPart, secondSlash - 1)
        Else
            firstSeg = pathPart
        End If
        InferSiteBase = hostPart & firstSeg
    Else
        InferSiteBase = hostPart
    End If
End Function
