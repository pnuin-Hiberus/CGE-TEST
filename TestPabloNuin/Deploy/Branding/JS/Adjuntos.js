(function () {

    // Create object that have the context information about the field that we want to change it output render  

    var linkFiledContext = {};

    linkFiledContext.Templates = {};

    linkFiledContext.Templates.Fields = {
         
        "Attachments": { "View": AttachmentsFiledTemplate }
    };

    SPClientTemplates.TemplateManager.RegisterTemplateOverrides(linkFiledContext);

})();


// This function provides the rendering logic for list view 
function AttachmentsFiledTemplate(ctx) {
    var itemId = ctx.CurrentItem.ID;
    var listName = ctx.ListTitle;       
    return getAttachments(listName, itemId);
}

function getAttachments(listName,itemId) {
  
    var url = _spPageContextInfo.webAbsoluteUrl;
    var requestUri = url + "/_api/web/lists/getbytitle('" + listName + "')/items(" + itemId + ")/AttachmentFiles";
    var str = "";
    // execute AJAX request
    $.ajax({
        url: requestUri,
        type: "GET",
        headers: { "ACCEPT": "application/json;odata=verbose" },
        async: false,
        success: function (data) {
            for (var i = 0; i < data.d.results.length; i++) {
                str += "<a target='_blank' href='" + data.d.results[i].ServerRelativeUrl + "'><img width='16' height='16' title='" + data.d.results[i].FileName + "' src='/_layouts/15/images/attach16.png?rev=44' border='0'></a>";
                if (i != data.d.results.length - 1) {
                    str += "<br/>";
                }                
            }          
        },
        error: function (err) {
            //alert(err);
        }
    });
    return str;
}