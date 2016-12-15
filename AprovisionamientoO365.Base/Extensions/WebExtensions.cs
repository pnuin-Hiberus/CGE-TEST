using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Utilities;

namespace AprovisionamientoO365.Base.Extensions
{
    public static class WebExtensions
    {
        /// <summary>
        /// Create a content type based on the classic feature framework structure.
        /// *Nuevo. Ahora hace el removefields también
        /// </summary>
        /// <param name="web">Web to operate against</param>
        /// <param name="xDocument">Actual XML document</param>
        public static ContentType CreateContentTypeFromXMLHiberus(this Web web, XDocument xDocument)
        {
            ContentType returnCT = null;
            var ns = xDocument.Root.Name.Namespace;

            var contentTypes = from cType in xDocument.Descendants(ns + "ContentType") select cType;

            foreach (var ct in contentTypes)
            {
                string ctid = ct.Attribute("ID").Value;
                string name = ct.Attribute("Name").Value;

                if (!web.ContentTypeExistsByName(name))
                {
                    var description = ct.Attribute("Description") != null ? ct.Attribute("Description").Value : string.Empty;
                    var group = ct.Attribute("Group") != null ? ct.Attribute("Group").Value : string.Empty;

                    // Create CT
                    web.CreateContentType(name, description, ctid, group);

                    // Add fields to content type 
                    var fieldRefs = from fr in ct.Descendants(ns + "FieldRefs").Elements(ns + "FieldRef") select fr;
                    foreach (var fieldRef in fieldRefs)
                    {
                        var frid = fieldRef.Attribute("ID").Value;
                        var required = fieldRef.Attribute("Required") != null ? bool.Parse(fieldRef.Attribute("Required").Value) : false;
                        var hidden = fieldRef.Attribute("Hidden") != null ? bool.Parse(fieldRef.Attribute("Hidden").Value) : false;
                        web.AddFieldToContentTypeById(ctid, frid, required, hidden);
                    }

                    //Remove fields from content type
                    var removeFieldRefs = from fr in ct.Descendants(ns + "FieldRefs").Elements(ns + "RemoveFieldRef") select fr;
                    foreach (var removeFieldRef in removeFieldRefs)
                    {
                        var rfrid = removeFieldRef.Attribute("ID").Value;
                        web.RemoveFieldFromContentTypeById(ctid, rfrid);
                    }

                    returnCT = web.GetContentTypeById(ctid);
                }
                else
                {
                    return null;
                }

            }

            return returnCT;
        }

        /// <summary>
        /// Quita field del content type
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="contentTypeID">String representation of the id of the content type to add the field to</param>
        /// <param name="fieldId">String representation of the field ID (=guid)</param>
        public static void RemoveFieldFromContentTypeById(this Web web, string contentTypeID, string fieldId)
        {
            try
            {
                // Get content type
                var ct = web.GetContentTypeById(contentTypeID);

                ct.EnsureProperties(c => c.Id, c => c.SchemaXml, c => c.FieldLinks.Include(fl => fl.Id, fl => fl.Required, fl => fl.Hidden));

                // Get field
                var fld = ct.FieldLinks.Cast<FieldLink>().Where(c => c.Id.Equals(new Guid(fieldId))).SingleOrDefault();
                if (fld == null) return;

                //Borro
                fld.DeleteObject();

                //Actualizo
                ct.Update(true);
                web.Context.ExecuteQueryRetry();
            }
            catch (Exception ex)
            {

            }
        }

        /// <summary>
        /// Private method to support all kinds of file uploads to the master page gallery
        /// </summary>
        /// <param name="web">Web as the root site of the publishing site collection</param>
        /// <param name="sourceFilePath">Full path to the file which will be uploaded</param>
        /// <param name="title">Title for the page layout</param>
        /// <param name="description">Description for the page layout</param>
        /// <param name="associatedContentTypeID">Associated content type ID</param>
        /// <param name="itemContentTypeId">Content type id for the item.</param>
        /// <param name="folderHierarchy">Folder hierarchy where the file will be uploaded</param>
        public static void DeployHtmlMasterPage(this Web web, string sourceFilePath, string title, string description, string associatedContentTypeID, string folderHierarchy = "")
        {
            if (string.IsNullOrEmpty(sourceFilePath))
            {
                throw new ArgumentNullException("sourceFilePath");
            }

            if (!System.IO.File.Exists(sourceFilePath))
            {
                throw new System.IO.FileNotFoundException("File for param sourceFilePath file does not exist", sourceFilePath);
            }

            string fileName = System.IO.Path.GetFileName(sourceFilePath);
            //Log.Info(Constants.LOGGING_SOURCE, CoreResources.BrandingExtension_DeployPageLayout, fileName, web.Context.Url);

            // Get the path to the file which we are about to deploy
            List masterPageGallery = web.GetCatalog((int)ListTemplateType.MasterPageCatalog);
            Folder rootFolder = masterPageGallery.RootFolder;
            web.Context.Load(masterPageGallery);
            web.Context.Load(rootFolder);
            web.Context.ExecuteQueryRetry();

            // Create folder structure inside master page gallery, if does not exists
            // For e.g.: _catalogs/masterpage/contoso/
            web.EnsureFolder(rootFolder, folderHierarchy);

            var fileBytes = System.IO.File.ReadAllBytes(sourceFilePath);

            // Use CSOM to upload the file in
            FileCreationInformation newFile = new FileCreationInformation();
            newFile.Content = fileBytes;
            newFile.Url = UrlUtility.Combine(rootFolder.ServerRelativeUrl, folderHierarchy, fileName);
            newFile.Overwrite = true;

            Microsoft.SharePoint.Client.File uploadFile = rootFolder.Files.Add(newFile);
            web.Context.Load(uploadFile);
            web.Context.ExecuteQueryRetry();

            // Check out the file if needed
            if (masterPageGallery.ForceCheckout || masterPageGallery.EnableVersioning)
                if (uploadFile.CheckOutType == CheckOutType.None)
                    uploadFile.CheckOut();

            // Get content type for ID to assign associated content type information
            ContentType associatedCt = web.GetContentTypeById(associatedContentTypeID, true);

            var listItem = uploadFile.ListItemAllFields;
            listItem["Title"] = title;
            listItem["MasterPageDescription"] = description;
            // set the item as page layout
            listItem["ContentTypeId"] = associatedContentTypeID; //associatedCt.Id;
            // Set the associated content type ID property
            listItem["PublishingAssociatedContentType"] = string.Format(";#{0};#{1};#", associatedCt.Name, associatedCt.Id);
            listItem["UIVersion"] = Convert.ToString(15);
            listItem.Update();

            // Check in the page layout if needed
            if (masterPageGallery.ForceCheckout || masterPageGallery.EnableVersioning)
                uploadFile.CheckIn(string.Empty, CheckinType.MajorCheckIn);
            if (masterPageGallery.EnableModeration)
                listItem.File.Publish(string.Empty);
            web.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Comprueba si un fichero existe
        /// </summary>
        /// <param name="web">Web</param>
        /// <param name="serverRelativeUrl">Ruta relativa al fichero</param>
        /// <returns>True si el fichero existe, False en caso contrario</returns>
        public static bool FileExists(this Web web, string serverRelativeUrl)
        {
            try
            {
                web.Context.Load(web, w => w.ServerRelativeUrl);
                web.Context.ExecuteQueryRetry();

                Microsoft.SharePoint.Client.File file = web.GetFileByServerRelativeUrl(string.Format("{0}/{1}", web.ServerRelativeUrl, serverRelativeUrl));

                web.Context.Load(file);
                web.Context.ExecuteQueryRetry();

                return file.ServerObjectIsNull.HasValue && !file.ServerObjectIsNull.Value;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
