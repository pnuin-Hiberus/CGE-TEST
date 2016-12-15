using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Hiberus.Aprovisionamiento.Schema;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.DocumentSet;
using OfficeDevPnP.Core.Utilities;

namespace AprovisionamientoO365.Base
{
    public static class Apoyo
    {
        public const string ResolveSiteCollection = "~ResolveSiteCollection";

        public static char[] trimChars = new char[] { '/' };

        /// <summary>
        /// Ejecuta un método capturando la excepción para registrarla en el log del sistema, junto con los tiempos de ejecución
        /// </summary>
        /// <param name="action">Método a ejecutar</param>
        public static void ExecuteWithTryCatch(Action action, string logInicial = "")
        {
            var watch = Stopwatch.StartNew();

            try
            {
                Console.Write(logInicial);

                action();

                Console.Write(" - ");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.Write("OK");
                Console.ForegroundColor = ConsoleColor.Gray;
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("ERROR: ");
                Console.Write(ex.Message);
                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.Gray;
                return;
            }
            finally
            {
                watch.Stop();
                var elapsedMS = watch.ElapsedMilliseconds;
                Console.Write(string.Format(" Tiempo: {0} ms", elapsedMS));
                Console.WriteLine();
            }
        }

        /// <summary>
        /// Ejecuta una función capturando la excepción para registrarla en el log del sistema, junto con los tiempos de ejecución
        /// </summary>
        /// <typeparam name="T">Tipo devuelto por la función</typeparam>
        /// <param name="function">Función a ejecutar</param>
        /// <returns>Resultado de la función</returns>
        public static T ExecuteWithTryCatch<T>(Func<T> function, string logInicial)
        {
            var watch = Stopwatch.StartNew();

            try
            {
                Console.Write(logInicial);

                T resultado = function();

                Console.Write(" - ");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.Write("OK");
                Console.ForegroundColor = ConsoleColor.Gray;

                return resultado;
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("ERROR: ");
                Console.Write(ex.Message);
                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.Gray;

                return default(T);
            }
            finally
            {
                watch.Stop();
                var elapsedMS = watch.ElapsedMilliseconds;
                Console.Write(string.Format(" Tiempo: {0} ms", elapsedMS));
                Console.WriteLine();
            }
        }

        public static Microsoft.SharePoint.Client.File AddFile(string rootUrl, Web web, string filePath, string fileName, string serverPath, string serverFolder)
        {
            var fileUrl = string.Concat(serverPath, serverFolder, (string.IsNullOrEmpty(serverFolder) ? string.Empty : "/"), fileName);
            var folder = web.GetFolderByServerRelativeUrl(string.Concat(serverPath, serverFolder));

            FileCreationInformation spFile = new FileCreationInformation()
            {
                Content = System.IO.File.ReadAllBytes(filePath + fileName.Replace("/", "\\")),
                Url = fileUrl,
                Overwrite = true
            };
            var uploadFile = folder.Files.Add(spFile);
            web.Context.Load(uploadFile, f => f.CheckOutType, f => f.Level);
            web.Context.ExecuteQuery();

            return uploadFile;
        }

        public static Folder EnsureFolder(Web web, string listUrl, string folderUrl, Folder parentFolder)
        {
            Folder folder = null;
            var folderServerRelativeUrl = parentFolder == null ? listUrl.TrimEnd(trimChars) + "/" + folderUrl : parentFolder.ServerRelativeUrl.TrimEnd(trimChars) + "/" + folderUrl;

            if (string.IsNullOrEmpty(folderUrl))
            {
                return null;
            }

            var lists = web.Lists;
            web.Context.Load(web);
            web.Context.Load(lists, l => l.Include(ll => ll.DefaultViewUrl));
            web.Context.ExecuteQuery();

            ExceptionHandlingScope scope = new ExceptionHandlingScope(web.Context);
            using (scope.StartScope())
            {
                using (scope.StartTry())
                {
                    folder = web.GetFolderByServerRelativeUrl(folderServerRelativeUrl);
                    web.Context.Load(folder);
                }

                using (scope.StartCatch())
                {
                    var list = lists.Where(l => l.DefaultViewUrl.IndexOf(listUrl, StringComparison.CurrentCultureIgnoreCase) >= 0).FirstOrDefault();

                    if (parentFolder == null)
                    {
                        parentFolder = list.RootFolder;
                    }


                    folder = parentFolder.Folders.Add(folderUrl);
                    web.Context.Load(folder);
                }

                using (scope.StartFinally())
                {
                    folder = web.GetFolderByServerRelativeUrl(folderServerRelativeUrl);
                    web.Context.Load(folder);
                }
            }

            web.Context.ExecuteQuery();
            return folder;
        }

        public static void SubirFichero(Site currentSite, string sourcePath, string targetPath, string fileName, bool resolveSiteCollection)
        {
            //Creamos la ruta de carpetas
            Folder folder = currentSite.RootWeb.EnsureFolderPath(targetPath);

            //Obtenemos la ruta relativa del fichero
            string filePath = UrlUtility.Combine(folder.ServerRelativeUrl, fileName);

            //Obtenemos el fichero para comprobar si ya existe
            Microsoft.SharePoint.Client.File file = currentSite.RootWeb.GetFileByServerRelativeUrl(filePath);

            //CheckOut
            currentSite.RootWeb.CheckOutFile(filePath);

            //Resolvemos los nombres ResolveSiteCollection por la ruta relativa de currentSite
            if (resolveSiteCollection)
            {
                string fileData = System.IO.File.ReadAllText(string.Format("{0}\\{1}", sourcePath, fileName));

                fileData = fileData.Contains(ResolveSiteCollection) ? fileData.Replace(ResolveSiteCollection, currentSite.RootWeb.ServerRelativeUrl) : fileData;

                byte[] bytes = new byte[fileData.Length * sizeof(char)];
                System.Buffer.BlockCopy(fileData.ToCharArray(), 0, bytes, 0, bytes.Length);

                file = folder.UploadFile(fileName, new MemoryStream(bytes), true);
            }
            else
                //Upload
                file = folder.UploadFile(fileName, Path.Combine(sourcePath, fileName), true);

            //CheckIn
            currentSite.RootWeb.CheckInFile(filePath, CheckinType.MajorCheckIn, string.Empty);

            //Publish
            currentSite.RootWeb.PublishFile(filePath, string.Empty);
        }

        public static void VC_CreateContentType(TenantSiteWebsRootWebContentTypesContentType contentType, TenantSite site, Site currentSite, List<Field> fields)
        {
            Console.WriteLine();

            ExecuteWithTryCatch(() =>
            {
                XDocument contentTypeDocument = null;
                XElement contentTypeNode = null;

                string contentTypeXMLPath = string.Format("{0}\\{1}\\{2}", Environment.CurrentDirectory, site.Webs.RootWeb.ContentTypes.SourcePath, contentType.SourceXML);
                contentTypeDocument = XDocument.Load(contentTypeXMLPath);
                contentTypeNode = contentTypeDocument.Descendants(contentTypeDocument.Root.Name.Namespace + "ContentType").FirstOrDefault();

                //Creamos el contentType a partir del XML
                ContentType ct = VC_CreateContentTypeFromXML(currentSite.RootWeb, contentTypeNode, fields.ToList());

                if (ct == null)
                {
                    Console.Write(" [Ya existe]");
                    return;
                }

                #region Configuración Document Set

                //Comprobamos si el Tipo de contenido parte de un conjunto a partir de la herencia del Identificador
                if (contentTypeNode.Attribute("ID").Value.StartsWith("0x0120D520"))
                {
                    ExecuteWithTryCatch(() =>
                    {
                        DocumentSetTemplate documentSetTemplate = DocumentSetTemplate.GetDocumentSetTemplate(currentSite.RootWeb.Context, ct);
                        currentSite.Context.Load(documentSetTemplate, w => w.AllowedContentTypes, w => w.SharedFields);
                        currentSite.Context.ExecuteQueryRetry();

                        //Eliminamos el tipo de documento base (Documento)
                        ContentType documentoCT = currentSite.RootWeb.GetContentTypeById("0x0101");
                        documentSetTemplate.AllowedContentTypes.Remove(documentoCT.Id);
                        documentSetTemplate.Update(true);

                        //List<XElement> AllowedContentTypes = contentTypeDocument.Descendants(contentTypeDocument.Root.Name.Namespace + "AllowedContentType").ToList();
                        List<XElement> SharedFields = contentTypeDocument.Descendants(contentTypeDocument.Root.Name.Namespace + "SharedField").ToList();
                        List<XElement> WelcomePageFields = contentTypeDocument.Descendants(contentTypeDocument.Root.Name.Namespace + "WelcomePageField").ToList();

                        List<ContentType> ContentTypes = new List<ContentType>();

                        ExecuteWithTryCatch(() =>
                        {
                            List<string> AllowedContentTypesIDs = contentTypeDocument.Descendants(contentTypeDocument.Root.Name.Namespace + "AllowedContentType").Select(w => w.Attribute("ID").Value).ToList();

                            foreach (string allowedContentTypeID in AllowedContentTypesIDs)
                            {
                                ContentType allowedContentType = currentSite.RootWeb.AvailableContentTypes.GetById(allowedContentTypeID);
                                currentSite.Context.Load(allowedContentType, w => w.Id);
                                ContentTypes.Add(allowedContentType);
                            }

                            currentSite.Context.ExecuteQueryRetry();
                        },
                        "Obtención AssociatedContentTypes");

                        if (ContentTypes.Count > 0) //xmlDocument.GetElementsByTagName("AllowedContentType").Count > 0)
                            foreach (ContentType associatedContentType in ContentTypes)
                            {
                                documentSetTemplate.AllowedContentTypes.Add(associatedContentType.Id);
                                documentSetTemplate.Update(true);
                            }

                        if (SharedFields.Count > 0)
                            foreach (XElement associatedField in SharedFields)
                            {
                                Field rootWebField = currentSite.RootWeb.GetFieldById<Field>(new Guid(associatedField.Attribute("ID").Value));
                                documentSetTemplate.SharedFields.Add(rootWebField);
                                documentSetTemplate.Update(true);
                            }

                        if (WelcomePageFields.Count > 0)
                            foreach (XElement associatedWelcomePageField in WelcomePageFields)
                            {
                                Field rootWebWelcomeField = currentSite.RootWeb.GetFieldById<Field>(new Guid(associatedWelcomePageField.Attribute("ID").Value));
                                documentSetTemplate.WelcomePageFields.Add(rootWebWelcomeField);
                                documentSetTemplate.Update(true);
                            }

                        currentSite.Context.ExecuteQueryRetry();

                        Console.WriteLine();
                        Console.Write("> Fin configuración DocumentSet <");
                    },
                    string.Format("> Inicio configuración DocumentSet <", contentType.SourceXML));
                }

                #endregion

                Console.WriteLine();
                Console.Write(string.Format("Fin creando contentType: [SourceXML: {0}]", contentType.SourceXML));
            },
            string.Format("Inicio creando contentType: [SourceXML: {0}]", contentType.SourceXML));

            Console.WriteLine();
        }

        /// <summary>
        /// Create a content type based on the classic feature framework structure.
        /// *Nuevo. Ahora hace el removefields también
        /// </summary>
        /// <param name="web">Web to operate against</param>
        /// <param name="xDocument">Actual XML document</param>
        public static ContentType VC_CreateContentTypeFromXML(Web web, XElement xmlNode, List<Field> fields)
        {
            ContentType createdContentType = null;
            List<FieldLink> ctfieldLinks = null;
            List<XElement> fieldRefs = null;
            List<XElement> removeFieldRefs = null;

            //Namespace
            var ns = xmlNode.Document.Root.Name.Namespace;

            string ctid = xmlNode.Attribute("ID").Value;
            string name = xmlNode.Attribute("Name").Value;

            if (!web.ContentTypeExistsByName(name))
            {
                var description = xmlNode.Attribute("Description") != null ? xmlNode.Attribute("Description").Value : string.Empty;
                var group = xmlNode.Attribute("Group") != null ? xmlNode.Attribute("Group").Value : string.Empty;

                // Create CT
                ExecuteWithTryCatch(() =>
                {
                    createdContentType = web.CreateContentType(name, description, ctid, group);
                },
                string.Format("web.CreateContentType"));

                ExecuteWithTryCatch(() =>
                {
                    web.Context.Load(createdContentType);
                    web.Context.Load(createdContentType.FieldLinks, w => w.Include(a => a.Id, a => a.Required, a => a.Hidden));
                    web.Context.ExecuteQueryRetry();

                    ctfieldLinks = createdContentType.FieldLinks.ToList();
                },
                string.Format("Cargando propiedades CT"));

                ExecuteWithTryCatch(() =>
                {
                    fieldRefs = xmlNode.Descendants(ns + "FieldRefs").Elements(ns + "FieldRef").ToList();
                },
                string.Format("Recuperando nodos a agregar"));

                ExecuteWithTryCatch(() =>
                {
                    foreach (XElement fieldRef in fieldRefs)
                    {
                        var fieldID = fieldRef.Attribute("ID").Value;
                        var required = fieldRef.Attribute("Required") != null ? bool.Parse(fieldRef.Attribute("Required").Value) : false;
                        var hidden = fieldRef.Attribute("Hidden") != null ? bool.Parse(fieldRef.Attribute("Hidden").Value) : false;

                        // Get field
                        Field fld = fields.Where(w => w.Id.ToString("B").Equals(fieldID, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();

                        // Get the field if already exists in content type, else add field to content type
                        // This will help to customize (required or hidden) any pre-existing field, also to handle existing field of Parent Content type
                        FieldLink flink = ctfieldLinks.Where(w => w.Id == fld.Id).FirstOrDefault();

                        if (flink == null)
                        {
                            XElement fieldElement = XElement.Parse(fld.SchemaXml);
                            fieldElement.SetAttributeValue("AllowDeletion", "TRUE"); // Default behavior when adding a field to a CT from the UI.
                            fld.SchemaXml = fieldElement.ToString();

                            var fldInfo = new FieldLinkCreationInformation();
                            fldInfo.Field = fld;
                            flink = createdContentType.FieldLinks.Add(fldInfo);
                            flink.Required = required;
                            flink.Hidden = hidden;
                            createdContentType.Update(true);
                        }
                        else
                        {
                            // Update FieldLink
                            flink.Required = required;
                            flink.Hidden = hidden;
                            createdContentType.Update(true);
                        }
                    }
                },
                string.Format("Asociando FieldLinks ({0})", fieldRefs.Count));

                ExecuteWithTryCatch(() =>
                {
                    removeFieldRefs = xmlNode.Descendants(ns + "FieldRefs").Elements(ns + "RemoveFieldRef").ToList();
                },
                string.Format("Recuperando nodos eliminados"));

                ExecuteWithTryCatch(() =>
                {
                    foreach (XElement removeFieldRef in removeFieldRefs)
                    {
                        var RFID = removeFieldRef.Attribute("ID").Value;

                        // Get field
                        var fld = createdContentType.FieldLinks.GetById(new Guid(RFID));

                        if (fld != null)
                        {
                            Console.Write(string.Format("Eliminando campo: {0}", RFID));

                            fld.DeleteObject();
                            createdContentType.Update(true);
                        }
                    }
                },
                string.Format("Eliminando FieldLinks ({0})", removeFieldRefs.Count));
            }

            ExecuteWithTryCatch(() =>
            {
                web.Context.ExecuteQueryRetry();
            },
            "ExecuteQueryRetry");

            Console.WriteLine();

            return createdContentType;
        }
    }
}
