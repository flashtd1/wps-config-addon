
function openOfficeFileFromSystemDemo(param){
    let jsonObj = (typeof(param)=='string' ? JSON.parse(param) : param)
    alert("从业务系统传过来的参数为：" + JSON.stringify(jsonObj))
    return {wps加载项项返回: jsonObj.filepath + ", 这个地址给的不正确"}
}

function InvokeFromSystemDemo(param){
    let jsonObj = (typeof(param)=='string' ? JSON.parse(param) : param)
    let handleInfo = jsonObj.Index
    switch (handleInfo){
        case "getDocumentName":{
            let docName = ""
            if (wps.EtApplication().ActiveWorkbook){
                docName = wps.EtApplication().ActiveWorkbook.Name
            }

            return {当前打开的文件名为:docName}
        }

        case "newDocument":{
            let newDocName=""
            let doc = wps.EtApplication().Workbooks.Add()
            newDocName = doc.Name
            
            return {操作结果:"新建文档成功，文档名为：" + newDocName}
        }
    }

    return {其它xxx:""}
}


export default{
    openOfficeFileFromSystemDemo,
    InvokeFromSystemDemo
}