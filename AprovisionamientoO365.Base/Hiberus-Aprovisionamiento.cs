using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using Hiberus.Aprovisionamiento.Schema;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.DocumentSet;
using Microsoft.SharePoint.Client.Publishing;
using Microsoft.SharePoint.Client.Publishing.Navigation;
using Microsoft.SharePoint.Client.Search.Administration;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Utilities;
using AprovisionamientoO365.Base.Extensions;

namespace AprovisionamientoO365.Base
{
    public class AprovisionamientoBase
    {
        static void Main(string[] args)
        {
            
        }

        /// <summary>
        /// Despliega una solución a partir del nombre del fichero XML
        /// </summary>
        /// <param name="settingsXMLName">Nombre del fichero XML. (Ej. SettingsAprovisionamiento.xml)</param>
        /// <param name="nombreFicheroLog">Nombre del fichero que se generará en C. (Ej. O365Planasa)</param>
        public static void Empieza(string settingsXMLName, string nombreFicheroLog = "Hiberus-Aprovisionamiento")
        {
            try
            {
                if (string.IsNullOrEmpty(settingsXMLName))
                    throw new ArgumentException("Valores de entrada no válidos");

                #region Obtención directorio actual

                var dir = System.IO.Directory.GetCurrentDirectory();
                dir = dir.Substring(0, dir.IndexOf("\\bin"));
                System.IO.Directory.SetCurrentDirectory(dir);

                #endregion

                //Lee el fichero xml para deploy de la solución
                System.IO.StreamReader xml = new System.IO.StreamReader(settingsXMLName);
                if (xml == null)
                    throw new Exception(String.Format("No se ha podido recuperar el fichero '{0}'", settingsXMLName));

                XmlSerializer xSerializer = new XmlSerializer(typeof(Tenant));

                Tenant tenant = (Tenant)xSerializer.Deserialize(xml);

                #region Credenciales

                SecureString password = new SecureString();

                foreach (char c in tenant.Credentials.Password.ToCharArray())
                    password.AppendChar(c);

                SharePointOnlineCredentials credentials = new SharePointOnlineCredentials(tenant.Credentials.Account, password);

                #endregion

                FileStream filestream = new FileStream(string.Format(@"C:\{0}_{1}.txt", nombreFicheroLog, DateTime.Now.ToString("yyyy-MM-dd_hh-mm-ss")), FileMode.Create);
                var streamwriter = new StreamWriter(filestream);
                streamwriter.AutoFlush = true;
                Console.SetOut(streamwriter);
                Console.SetError(streamwriter);

                Console.WriteLine("Inicio: " + DateTime.Now.ToString("yyyy-MM-dd_hh-mm-ss"));

                using (ClientContext clientContext = new ClientContext(tenant.AdminUrl))
                {
                    //credenciales del usuario introducido
                    clientContext.Credentials = credentials;

                    Microsoft.Online.SharePoint.TenantAdministration.Tenant o365Tenant = new Microsoft.Online.SharePoint.TenantAdministration.Tenant(clientContext);

                    foreach (TenantSite site in tenant.Sites)
                    {
                        #region Creación Site Collection (Si ya existe se omite el paso)

                        //Creamos la Site Collection si no existe o está en la papelera de reciclaje
                        if (!o365Tenant.SiteExists(string.Concat(tenant.Url, site.Url)) || o365Tenant.CheckIfSiteExists(string.Concat(tenant.Url, site.Url), "Recycled"))
                        {
                            Apoyo.ExecuteWithTryCatch(() =>
                            {
                                SiteEntity siteEntity = new SiteEntity()
                                {
                                    Title = site.Title,
                                    Description = site.Description,
                                    Url = string.Concat(tenant.Url, site.Url),
                                    Lcid = site.LCID,
                                    CurrentResourceUsage = site.CurrentResourceUsage,
                                    StorageUsage = site.StorageUsage,
                                    SiteOwnerLogin = site.SiteOwnerLogin,
                                    Template = site.Template,
                                    TimeZoneId = site.TimeZoneId,
                                };

                                o365Tenant.CreateSiteCollection(siteEntity, true, true);
                            },
                            string.Format("Creando siteCollection: [URL: {0}]", site.Url));
                        }

                        #endregion

                        #region Obtención del Site Collection

                        Site currentSite = Apoyo.ExecuteWithTryCatch(() =>
                        {
                            return o365Tenant.GetSiteByUrl(string.Concat(tenant.Url, site.Url));
                        },
                        string.Format("Obteniendo SiteCollection: [Url: {0}]", site.Url));

                        #endregion

                        #region Site Collection Features

                        if (site.Features != null && site.Features.Provision)
                            foreach (TenantSiteFeaturesFeature feature in site.Features.Feature)
                                if (!currentSite.IsFeatureActive(new Guid(feature.ID)))
                                {
                                    Apoyo.ExecuteWithTryCatch(() =>
                                    {
                                        currentSite.ActivateFeature(new Guid(feature.ID));
                                    },
                                    string.Format("Activando feature: [ID: {0}]", feature.ID));
                                }

                        #endregion

                        #region RootWeb

                        if (site.Webs.RootWeb != null && site.Webs.RootWeb.Provision)
                        {
                            Console.WriteLine(string.Empty);
                            Console.WriteLine(">> Provisionando la RootWeb <<");
                            Console.WriteLine(string.Empty);

                            clientContext.Load(currentSite.RootWeb, w => w.ServerRelativeUrl);
                            clientContext.Load(currentSite.RootWeb, w => w.Url);
                            clientContext.ExecuteQueryRetry();

                            #region Ficheros

                            //Configuramos el siteCollection
                            if (site.Webs.RootWeb.Files != null && site.Webs.RootWeb.Files.Provision)
                            {
                                //Imágenes
                                if (site.Webs.RootWeb.Files.Images != null && site.Webs.RootWeb.Files.Images.Provision)
                                    foreach (TenantSiteWebsRootWebFilesImagesFile image in site.Webs.RootWeb.Files.Images.File)
                                    {
                                        Apoyo.ExecuteWithTryCatch(() =>
                                        {
                                            Apoyo.SubirFichero(
                                                currentSite,
                                                string.IsNullOrWhiteSpace(image.SourcePath) ? site.Webs.RootWeb.Files.Images.SourcePath : Path.Combine(site.Webs.RootWeb.Files.Images.SourcePath, image.SourcePath),
                                                string.IsNullOrWhiteSpace(image.TargetPath) ? site.Webs.RootWeb.Files.Images.TargetPath : string.Concat(site.Webs.RootWeb.Files.Images.TargetPath, image.TargetPath),
                                                image.Name,
                                                false);
                                        },
                                        string.Format("Subiendo imagen: [Nombre: {0}]", image.Name));
                                    }

                                //CSS
                                if (site.Webs.RootWeb.Files.Css != null && site.Webs.RootWeb.Files.Css.Provision)
                                    foreach (TenantSiteWebsRootWebFilesCssFile css in site.Webs.RootWeb.Files.Css.File)
                                    {
                                        Apoyo.ExecuteWithTryCatch(() =>
                                        {
                                            Apoyo.SubirFichero(
                                                currentSite,
                                                string.IsNullOrWhiteSpace(css.SourcePath) ? site.Webs.RootWeb.Files.Css.SourcePath : Path.Combine(site.Webs.RootWeb.Files.Css.SourcePath, css.SourcePath),
                                                string.IsNullOrWhiteSpace(css.TargetPath) ? site.Webs.RootWeb.Files.Css.TargetPath : string.Concat(site.Webs.RootWeb.Files.Css.TargetPath, css.TargetPath),
                                                css.Name,
                                                false);
                                        },
                                        string.Format("Subiendo css: [Nombre: {0}]", css.Name));
                                    }

                                //JS
                                if (site.Webs.RootWeb.Files.Js != null && site.Webs.RootWeb.Files.Js.Provision)
                                    foreach (TenantSiteWebsRootWebFilesJSFile js in site.Webs.RootWeb.Files.Js.File)
                                    {
                                        Apoyo.ExecuteWithTryCatch(() =>
                                        {
                                            Apoyo.SubirFichero(
                                                currentSite,
                                                string.IsNullOrWhiteSpace(js.SourcePath) ? site.Webs.RootWeb.Files.Js.SourcePath : Path.Combine(site.Webs.RootWeb.Files.Js.SourcePath, js.SourcePath),
                                                string.IsNullOrWhiteSpace(js.TargetPath) ? site.Webs.RootWeb.Files.Js.TargetPath : string.Concat(site.Webs.RootWeb.Files.Js.TargetPath, js.TargetPath),
                                                js.Name,
                                                false);
                                        },
                                        string.Format("Subiendo js: [Nombre: {0}]", js.Name));
                                    }

                                //XSL
                                if ( site.Webs.RootWeb.Files.xsl != null && site.Webs.RootWeb.Files.xsl.Provision)
                                    foreach (TenantSiteWebsRootWebFilesXslFile xsl in site.Webs.RootWeb.Files.xsl.File)
                                    {
                                        Apoyo.ExecuteWithTryCatch(() =>
                                        {
                                            Apoyo.SubirFichero(
                                                currentSite,
                                                string.IsNullOrWhiteSpace(xsl.SourcePath) ? site.Webs.RootWeb.Files.xsl.SourcePath : Path.Combine(site.Webs.RootWeb.Files.xsl.SourcePath, xsl.SourcePath),
                                                string.IsNullOrWhiteSpace(xsl.TargetPath) ? site.Webs.RootWeb.Files.xsl.TargetPath : string.Concat(site.Webs.RootWeb.Files.xsl.TargetPath, xsl.TargetPath),
                                                xsl.Name,
                                                false);
                                        },
                                        string.Format("Subiendo xsl: [Nombre: {0}]", xsl.Name));
                                    }

                                //DisplayTemplates
                                if (site.Webs.RootWeb.Files.DisplayTemplates != null && site.Webs.RootWeb.Files.DisplayTemplates.Provision)
                                    foreach (TenantSiteWebsRootWebFilesDisplayTemplatesFile displayTemplate in site.Webs.RootWeb.Files.DisplayTemplates.File)
                                    {
                                        Apoyo.ExecuteWithTryCatch(() =>
                                        {
                                            Apoyo.SubirFichero(
                                                currentSite,
                                                string.IsNullOrEmpty(displayTemplate.SourcePath) ? site.Webs.RootWeb.Files.DisplayTemplates.SourcePath : Path.Combine(site.Webs.RootWeb.Files.DisplayTemplates.SourcePath, displayTemplate.SourcePath),
                                                string.IsNullOrWhiteSpace(displayTemplate.TargetPath) ? site.Webs.RootWeb.Files.DisplayTemplates.TargetPath : string.Concat(site.Webs.RootWeb.Files.DisplayTemplates.TargetPath, displayTemplate.TargetPath),
                                                displayTemplate.Name,
                                                false);

                                        },
                                        string.Format("Subiendo Display template: [Nombre: {0}]", displayTemplate.Name));
                                    }

                                //Fonts
                                if (site.Webs.RootWeb.Files.Fonts != null && site.Webs.RootWeb.Files.Fonts.Provision)
                                    foreach (TenantSiteWebsRootWebFilesFontsFile font in site.Webs.RootWeb.Files.Fonts.File)
                                    {
                                        Apoyo.ExecuteWithTryCatch(() =>
                                        {
                                            Apoyo.SubirFichero(
                                                currentSite,
                                                string.IsNullOrWhiteSpace(font.SourcePath) ? site.Webs.RootWeb.Files.Fonts.SourcePath : Path.Combine(site.Webs.RootWeb.Files.Fonts.SourcePath, font.SourcePath),
                                                string.IsNullOrWhiteSpace(font.TargetPath) ? site.Webs.RootWeb.Files.Fonts.TargetPath : string.Concat(site.Webs.RootWeb.Files.Fonts.TargetPath, font.TargetPath),
                                                font.Name,
                                                false);
                                        },
                                        string.Format("Subiendo fuente: [Nombre: {0}]", font.Name));
                                    }

                                //LanguageFiles
                                if(site.Webs.RootWeb.Files.LanguageFiles != null && site.Webs.RootWeb.Files.LanguageFiles.Provision)
                                    foreach(TenantSiteWebsRootWebFilesLanguageFilesFile languageFile in site.Webs.RootWeb.Files.LanguageFiles.File)
                                    {
                                        Apoyo.ExecuteWithTryCatch(() => 
                                        {
                                            Apoyo.SubirFichero(
                                              currentSite,
                                              string.IsNullOrWhiteSpace(languageFile.SourcePath) ? site.Webs.RootWeb.Files.LanguageFiles.SourcePath : Path.Combine(site.Webs.RootWeb.Files.LanguageFiles.SourcePath, languageFile.SourcePath),
                                              string.IsNullOrWhiteSpace(languageFile.TargetPath) ? site.Webs.RootWeb.Files.LanguageFiles.TargetPath : string.Concat(site.Webs.RootWeb.Files.LanguageFiles.TargetPath, languageFile.TargetPath),
                                              languageFile.Name,
                                              false);

                                        }, string.Format("Subiendo fichero de lenguage: [Nombre:{0}, Idioma:{1}]", languageFile.Name, languageFile.SourcePath));
                                    }
                            }

                            #endregion

                            #region CustomActions

                            if(site.Webs.RootWeb.CustomActions != null && site.Webs.RootWeb.CustomActions.Provision)
                            foreach (TenantSiteWebsRootWebCustomActionsCustomAction customAction in site.Webs.RootWeb.CustomActions.CustomAction)
                                Apoyo.ExecuteWithTryCatch(() => 
                                {
                                    if (!currentSite.AddCustomActionFromXMLFile_Hiberus(customAction.SourcePath))
                                        throw new Exception(string.Format("No se ha añadido el custom action: [SourcePath: {0}]", customAction.SourcePath));
                                }, string.Format("Añadido custom action: [SourcePath: {0}]", customAction.SourcePath));

                            #endregion

                            #region Catalogs

                            if (site.Webs.RootWeb.Catalogs != null && site.Webs.RootWeb.Catalogs.Provision)
                            {
                                Folder folder = currentSite.RootWeb.GetCatalog((int)ListTemplateType.MasterPageCatalog).RootFolder;
                                clientContext.Load(folder);
                                clientContext.ExecuteQueryRetry();

                                //MasterPages
                                if (site.Webs.RootWeb.Catalogs.Masterpages != null && site.Webs.RootWeb.Catalogs.Masterpages.Provision)
                                    foreach (TenantSiteWebsRootWebCatalogsMasterpagesMasterpage masterPage in site.Webs.RootWeb.Catalogs.Masterpages.Masterpage)
                                    {
                                        Apoyo.ExecuteWithTryCatch(() =>
                                        {
                                            //html con tipo de contenido HTML Principal para que se genere automáticamente el fichero .master
                                            currentSite.RootWeb.DeployHtmlMasterPage(
                                                 Path.Combine(site.Webs.RootWeb.Catalogs.Masterpages.SourcePath, masterPage.Name),
                                                 masterPage.Title,
                                                 masterPage.Description,
                                                 masterPage.ContentType,
                                                 String.Empty);

                                            //Prewview
                                            folder.UploadFile(masterPage.Preview,
                                                                    Path.Combine(site.Webs.RootWeb.Catalogs.Masterpages.SourcePath, masterPage.Preview),
                                                                    true);
                                        },
                                        string.Format("Subiendo MasterPage: [Nombre: {0}]", masterPage.Name));
                                    }

                                //PageLayouts
                                if (site.Webs.RootWeb.Catalogs.PageLayouts != null && site.Webs.RootWeb.Catalogs.PageLayouts.Provision)
                                    foreach (TenantSiteWebsRootWebCatalogsPageLayoutsPageLayout pageLayout in site.Webs.RootWeb.Catalogs.PageLayouts.PageLayout)
                                    {
                                        Apoyo.ExecuteWithTryCatch(() =>
                                        {
                                            currentSite.RootWeb.DeployHtmlPageLayout(
                                               Path.Combine(site.Webs.RootWeb.Catalogs.PageLayouts.SourcePath, pageLayout.Name),
                                               pageLayout.Title,
                                               pageLayout.Description,
                                               pageLayout.ContentType,
                                               site.Webs.RootWeb.Catalogs.PageLayouts.TargetPath);
                                        },
                                        string.Format("Subiendo PageLayout: [Nombre: {0}]", pageLayout.Name));
                                    }

                                //Colors
                                if (site.Webs.RootWeb.Catalogs.Colors != null && site.Webs.RootWeb.Catalogs.Colors.Provision)
                                    foreach (TenantSiteWebsRootWebCatalogsColorsColor color in site.Webs.RootWeb.Catalogs.Colors.Color)
                                    {
                                        Apoyo.ExecuteWithTryCatch(() =>
                                        {
                                            currentSite.RootWeb.UploadThemeFile(Path.Combine(site.Webs.RootWeb.Catalogs.Colors.SourcePath, color.Name));
                                        },
                                        string.Format("Subiendo color de tema: [Nombre: {0}]", color.Name));
                                    }

                                //Fonts
                                if (site.Webs.RootWeb.Catalogs.Fonts != null && site.Webs.RootWeb.Catalogs.Fonts.Provision)
                                    foreach (TenantSiteWebsRootWebCatalogsFontsFont font in site.Webs.RootWeb.Catalogs.Fonts.Font)
                                    {
                                        Apoyo.ExecuteWithTryCatch(() =>
                                        {
                                            currentSite.RootWeb.UploadThemeFile(Path.Combine(site.Webs.RootWeb.Catalogs.Fonts.SourcePath, font.Name));
                                        },
                                        string.Format("Subiendo fuente de tema: [Nombre: {0}]", font.Name));
                                    }

                                //Backgrounds
                                if (site.Webs.RootWeb.Catalogs.Backgrounds != null && site.Webs.RootWeb.Catalogs.Backgrounds.Provision)
                                    foreach (TenantSiteWebsRootWebCatalogsBackgroundsBackground background in site.Webs.RootWeb.Catalogs.Backgrounds.Background)
                                    {
                                        Apoyo.ExecuteWithTryCatch(() =>
                                        {
                                            currentSite.RootWeb.UploadThemeFile(Path.Combine(site.Webs.RootWeb.Catalogs.Backgrounds.SourcePath, background.Name));
                                        },
                                        string.Format("Subiendo background de tema: [Nombre: {0}]", background.Name));
                                    }
                            }

                            #endregion

                            #region Web Features

                            if (site.Webs.RootWeb.Features != null && site.Webs.RootWeb.Features.Provision)
                                foreach (TenantSiteWebsRootWebFeaturesFeature feature in site.Webs.RootWeb.Features.Feature)
                                {
                                    if (!currentSite.RootWeb.IsFeatureActive(new Guid(feature.ID)))
                                    {
                                        Apoyo.ExecuteWithTryCatch(() =>
                                        {
                                            currentSite.RootWeb.ActivateFeature(new Guid(feature.ID));
                                        },
                                        string.Format("Activando feature: [ID: {0}]", feature.ID));
                                    }
                                }

                            #endregion

                            #region SiteColumns

                            if (site.Webs.RootWeb.SiteColumns != null && site.Webs.RootWeb.SiteColumns.Provision)
                            {
                                foreach (TenantSiteWebsRootWebSiteColumnsField field in site.Webs.RootWeb.SiteColumns.Field)
                                {
                                    Apoyo.ExecuteWithTryCatch(() =>
                                    {
                                        //Leemos el XML
                                        string xmlString = System.IO.File.ReadAllText(string.Format("{0}\\{1}", site.Webs.RootWeb.SiteColumns.SourcePath, field.SourceXML));

                                        //Generamos un objeto XMLReader
                                        XmlDocument xmlDocument = new XmlDocument();
                                        xmlDocument.LoadXml(xmlString);

                                        //Obtenemos el Nodo Field
                                        XmlNode fieldNode = xmlDocument.GetElementsByTagName("Field").Item(0);

                                        //Si ya está creado el campo en SharePoint continuamos
                                        if (currentSite.RootWeb.FieldExistsById(fieldNode.Attributes["ID"].Value))
                                        {
                                            Console.Write(" [Ya existe]");

                                            return;
                                        }

                                        Field createdField = null;

                                        #region LookupField

                                        if (fieldNode.Attributes["Type"].Value.Contains("Lookup"))
                                        {
                                            List LookupList = null;
                                            Web requiredWeb = null;

                                            if (field.RequiredWeb != null)
                                            {
                                                //Creamos la web requerida
                                                #region Creacion de Web (si ya existe se omite el paso)

                                                if (!currentSite.RootWeb.WebExists(field.RequiredWeb.Url))
                                                {
                                                    Apoyo.ExecuteWithTryCatch(() =>
                                                    {
                                                        SiteEntity newWeb = new SiteEntity()
                                                        {
                                                            Title = field.RequiredWeb.Title,
                                                            Description = field.RequiredWeb.Description,
                                                            Url = field.RequiredWeb.Url,
                                                            Lcid = field.RequiredWeb.LCID,
                                                            SiteOwnerLogin = field.RequiredWeb.SiteOwnerLogin,
                                                            Template = field.RequiredWeb.Template
                                                        };

                                                        requiredWeb = currentSite.RootWeb.CreateWeb(newWeb, true, true);
                                                    },
                                                    string.Format("Creando Web: [URL: {0}]", field.RequiredWeb.Url));
                                                }

                                                if (requiredWeb == null)
                                                    requiredWeb = currentSite.RootWeb.GetWeb(field.RequiredWeb.Url);

                                                clientContext.Load(requiredWeb, w => w.ServerRelativeUrl);
                                                clientContext.ExecuteQueryRetry();

                                                #endregion
                                            }
                                            else
                                            {
                                                requiredWeb = currentSite.RootWeb;
                                            }

                                            #region Creación de la lista requerida

                                            //Creación de la lista requerida (Si es necesario)
                                            if (field.RequiredList != null)
                                                if (!requiredWeb.ListExists(field.RequiredList.Name))
                                                    Apoyo.ExecuteWithTryCatch(() =>
                                                    {
                                                        LookupList = requiredWeb.CreateList(
                                                                            (ListTemplateType)Enum.Parse(typeof(ListTemplateType), field.RequiredList.TemplateType, true),
                                                                            field.RequiredList.Name,
                                                                            field.RequiredList.EnableVersioning,
                                                                            true,
                                                                            field.RequiredList.UrlPath.Replace("\\", "/"),
                                                                            field.RequiredList.EnableContentTypes);
                                                    },
                                                    string.Format("Creando lista requerida: [Name: {0}]", field.RequiredList.Name));

                                            if (LookupList == null)
                                                LookupList = requiredWeb.GetListByUrl(field.RequiredList.UrlPath);

                                            #endregion

                                            clientContext.Load(LookupList, w => w.Id);
                                            clientContext.ExecuteQueryRetry();

                                            Web web = currentSite.OpenWeb(fieldNode.Attributes["WebId"].Value);
                                            clientContext.Load(web, w => w.Id);
                                            clientContext.ExecuteQuery();

                                            foreach (XmlNode fieldLookupField in xmlDocument.GetElementsByTagName("Field"))
                                            {
                                                #region Asignación del ID de lista

                                                fieldLookupField.Attributes["List"].Value = LookupList.Id.ToString("B");

                                                #endregion

                                                #region Asignación del ID de la web

                                                fieldLookupField.Attributes["WebId"].Value = web.Id.ToString("D");

                                                #endregion

                                                //Creamos el SPField a partir del XML
                                                createdField = currentSite.RootWeb.CreateField(fieldLookupField.OuterXml, true);

                                                //Actualización de visibilidad del campo
                                                if (fieldNode.Attributes["ShowInNewForm"] != null ||
                                                    fieldNode.Attributes["ShowInEditForm"] != null ||
                                                    fieldNode.Attributes["ShowInViewForm"] != null)
                                                {
                                                    Apoyo.ExecuteWithTryCatch(() =>
                                                    {
                                                        //ShowInNewForm
                                                        if (fieldNode.Attributes["ShowInNewForm"] != null)
                                                            createdField.SetShowInNewForm(bool.Parse(fieldNode.Attributes["ShowInNewForm"].Value));

                                                        //ShowInEditForm
                                                        if (fieldNode.Attributes["ShowInEditForm"] != null)
                                                            createdField.SetShowInEditForm(bool.Parse(fieldNode.Attributes["ShowInEditForm"].Value));

                                                        //ShowInViewForm
                                                        if (fieldNode.Attributes["ShowInViewForm"] != null)
                                                            createdField.SetShowInEditForm(bool.Parse(fieldNode.Attributes["ShowInViewForm"].Value));

                                                        createdField.Update();
                                                        clientContext.ExecuteQueryRetry();
                                                    }, "Actualizada visibilidad del campo");
                                                }
                                            }

                                            return;
                                        }

                                        #endregion

                                        #region TaxonomyField

                                        //Comprobamos si el campo es de tipo taxonomía para crearlo a partir de la función de Office PnP
                                        if (fieldNode.Attributes["Type"].Value == "TaxonomyFieldType" || fieldNode.Attributes["Type"].Value == "TaxonomyFieldTypeMulti")
                                        {
                                            //Obtenemos el almacén de términos del site collection
                                            TermStore termStore = currentSite.GetDefaultSiteCollectionTermStore();

                                            //Obtenemos el grupo de términos asociado a la columna
                                            TermGroup termGroup = termStore.Groups.GetByName(field.TermGroupName);

                                            //Obtenemos el conjunto de términos asociado a la columna
                                            TermSet termSet = termGroup.TermSets.GetByName(field.TermSetName);

                                            clientContext.Load(termStore);
                                            clientContext.Load(termSet);
                                            clientContext.ExecuteQueryRetry();

                                            //Atributos adicionales (LCID)
                                            List<KeyValuePair<string, string>> additionalAttributes = new List<KeyValuePair<string, string>>();
                                            additionalAttributes.Add(new KeyValuePair<string, string>("ShowField", string.Format("Term{0}", field.LCID)));

                                            TaxonomyFieldCreationInformation taxonomyFieldCreation = new TaxonomyFieldCreationInformation()
                                            {
                                                Id = new Guid(fieldNode.Attributes["ID"].Value),
                                                InternalName = fieldNode.Attributes["Name"].Value,
                                                DisplayName = fieldNode.Attributes["DisplayName"].Value,
                                                Group = fieldNode.Attributes["Group"].Value,
                                                TaxonomyItem = termSet,
                                                AdditionalAttributes = additionalAttributes,
                                                MultiValue = (fieldNode.Attributes["Type"].Value == "TaxonomyFieldTypeMulti") ? true : false
                                            };

                                            createdField = currentSite.RootWeb.CreateTaxonomyField(taxonomyFieldCreation);

                                            //Actualización de visibilidad del campo
                                            if (fieldNode.Attributes["ShowInNewForm"] != null ||
                                                fieldNode.Attributes["ShowInEditForm"] != null ||
                                                fieldNode.Attributes["ShowInViewForm"] != null)
                                            {
                                                Apoyo.ExecuteWithTryCatch(() =>
                                                {
                                                    //ShowInNewForm
                                                    if (fieldNode.Attributes["ShowInNewForm"] != null)
                                                        createdField.SetShowInNewForm(bool.Parse(fieldNode.Attributes["ShowInNewForm"].Value));

                                                    //ShowInEditForm
                                                    if (fieldNode.Attributes["ShowInEditForm"] != null)
                                                        createdField.SetShowInEditForm(bool.Parse(fieldNode.Attributes["ShowInEditForm"].Value));

                                                    //ShowInViewForm
                                                    if (fieldNode.Attributes["ShowInViewForm"] != null)
                                                        createdField.SetShowInEditForm(bool.Parse(fieldNode.Attributes["ShowInViewForm"].Value));

                                                    createdField.Update();
                                                    clientContext.ExecuteQueryRetry();
                                                }, "Actualizada visibilidad del campo");
                                            }

                                            return;
                                        }

                                        #endregion

                                        //Creamos el SPField a partir del XML
                                        createdField = currentSite.RootWeb.CreateField(fieldNode.OuterXml, true);

                                        //Actualización de visibilidad del campo
                                        if (fieldNode.Attributes["ShowInNewForm"] != null ||
                                            fieldNode.Attributes["ShowInEditForm"] != null ||
                                            fieldNode.Attributes["ShowInViewForm"] != null)
                                        {
                                            Apoyo.ExecuteWithTryCatch(() => 
                                            {
                                                //ShowInNewForm
                                                if (fieldNode.Attributes["ShowInNewForm"] != null)
                                                    createdField.SetShowInNewForm(bool.Parse(fieldNode.Attributes["ShowInNewForm"].Value));

                                                //ShowInEditForm
                                                if (fieldNode.Attributes["ShowInEditForm"] != null)
                                                    createdField.SetShowInEditForm(bool.Parse(fieldNode.Attributes["ShowInEditForm"].Value));

                                                //ShowInViewForm
                                                if (fieldNode.Attributes["ShowInViewForm"] != null)
                                                    createdField.SetShowInEditForm(bool.Parse(fieldNode.Attributes["ShowInViewForm"].Value));

                                                createdField.Update();
                                                clientContext.ExecuteQueryRetry();
                                            }, "Actualizada visibilidad del campo");
                                        }
                                    },
                                    string.Format("Creando columna: [SourceXML: {0}]", field.SourceXML));
                                }
                            }

                            #endregion

                            #region ContentTypes

                            if (site.Webs.RootWeb.ContentTypes != null && site.Webs.RootWeb.ContentTypes.Provision)
                            {
                                FieldCollection fields = null;
                                Apoyo.ExecuteWithTryCatch(() =>
                                {
                                    fields = currentSite.RootWeb.Fields;
                                    clientContext.Load(fields, w => w.Include(a => a.Id, a => a.SchemaXml));
                                    clientContext.ExecuteQueryRetry();
                                },
                                string.Format("Obtención de los campos de la web"));

                                foreach (TenantSiteWebsRootWebContentTypesContentType contentType in site.Webs.RootWeb.ContentTypes.ContentType)
                                {
                                    Apoyo.VC_CreateContentType(contentType, site, currentSite, fields.ToList());
                                }
                            }

                            #endregion

                            #region Groups

                            //Creación de grupos RootWeb
                            if (site.Webs.RootWeb.Groups != null && site.Webs.RootWeb.Groups.Provision)
                                foreach (TenantSiteWebsRootWebGroupsGroup grupo in site.Webs.RootWeb.Groups.Group)
                                {
                                    if (!currentSite.RootWeb.GroupExists(grupo.Name))
                                    {
                                        Apoyo.ExecuteWithTryCatch(() =>
                                        {
                                            //Creamos el grupo
                                            Group group = currentSite.RootWeb.AddGroup(grupo.Name, grupo.Description, grupo.Name == grupo.Owner, true);

                                            if (grupo.Name != grupo.Owner)
                                            {
                                                clientContext.Load(currentSite.RootWeb.SiteGroups);
                                                clientContext.ExecuteQuery();

                                                Group ownerGroup = currentSite.RootWeb.SiteGroups.GetByName(grupo.Owner);

                                                //Establecemos como propietario del grupo el grupo de propietarios de la site collection
                                                group.Owner = ownerGroup;
                                            }
                                            group.Update();

                                            clientContext.ExecuteQueryRetry();
                                        },
                                        string.Format("Creando grupo: [Nombre: {0}]", grupo.Name));
                                    }
                                }

                            #endregion

                            #region Permissions

                            //Asignación de permisos de grupo
                            if (site.Webs.RootWeb.Permissions != null && site.Webs.RootWeb.Permissions.Provision)
                            {

                            }

                            #endregion

                            #region Tema

                            if (site.Webs.RootWeb.Theme != null && site.Webs.RootWeb.Theme.Provision)
                            {
                                //Crear
                                Apoyo.ExecuteWithTryCatch(() =>
                                {
                                    //currentSite.RootWeb.CreateComposedLookByName(
                                    //            site.Webs.RootWeb.Theme.Titulo,
                                    //            site.Webs.RootWeb.Theme.Colors,
                                    //            site.Webs.RootWeb.Theme.Fonts,
                                    //            site.Webs.RootWeb.Theme.BackgroundImage,
                                    //            site.Webs.RootWeb.Theme.MasterPage);

                                    currentSite.RootWeb.CreateComposedLookByUrl(
                                        site.Webs.RootWeb.Theme.Titulo,
                                        String.Format("{0}_catalogs/Theme/15/{1}",
                                        UrlUtility.EnsureTrailingSlash(currentSite.RootWeb.ServerRelativeUrl),
                                        site.Webs.RootWeb.Theme.Colors),
                                        String.Format("{0}_catalogs/Theme/15/{1}",
                                        UrlUtility.EnsureTrailingSlash(currentSite.RootWeb.ServerRelativeUrl),
                                        site.Webs.RootWeb.Theme.Fonts),
                                        string.IsNullOrEmpty(site.Webs.RootWeb.Theme.BackgroundImage) ? string.Empty : String.Format("{0}_catalogs/Theme/15/{1}", UrlUtility.EnsureTrailingSlash(currentSite.RootWeb.ServerRelativeUrl), site.Webs.RootWeb.Theme.BackgroundImage),
                                        String.Format("{0}_catalogs/masterpage/{1}",
                                        UrlUtility.EnsureTrailingSlash(currentSite.RootWeb.ServerRelativeUrl),
                                        site.Webs.RootWeb.Theme.MasterPage),
                                        1, true);

                                }, string.Format("Creado Theme: [Nombre: {0}]", site.Webs.RootWeb.Theme.Titulo));

                                //Set
                                Apoyo.ExecuteWithTryCatch(() =>
                                {
                                    currentSite.RootWeb.SetComposedLookByUrl(site.Webs.RootWeb.Theme.Titulo);

                                }, string.Format("Cambio a Theme: [Nombre: {0}]", site.Webs.RootWeb.Theme.Titulo));

                                //Cambio pagina maestra de sistema
                                if (!string.IsNullOrEmpty(site.Webs.RootWeb.Theme.SystemMasterPage))
                                    Apoyo.ExecuteWithTryCatch(() =>
                                    {
                                        Folder folder = currentSite.RootWeb.GetCatalog((int)ListTemplateType.MasterPageCatalog).RootFolder;
                                        clientContext.Load(folder);
                                        clientContext.Load(currentSite.RootWeb, w => w.MasterUrl);
                                        clientContext.ExecuteQueryRetry();

                                        currentSite.RootWeb.MasterUrl = String.Concat(UrlUtility.EnsureTrailingSlash(folder.ServerRelativeUrl), site.Webs.RootWeb.Theme.SystemMasterPage);
                                        currentSite.RootWeb.Update();
                                        clientContext.ExecuteQueryRetry();
                                    }, string.Format("Cambio a pagina de sistema: [Nombre: {0}]", site.Webs.RootWeb.Theme.SystemMasterPage));
                            }

                            #endregion

                            #region Logo

                            if (site.Webs.RootWeb.Logo != null && site.Webs.RootWeb.Logo.Provision)
                            {
                                if (!string.IsNullOrWhiteSpace(site.Webs.RootWeb.Logo.Url))
                                {
                                    clientContext.Load(currentSite.RootWeb, w => w.ServerRelativeUrl);
                                    clientContext.ExecuteQueryRetry();

                                    Apoyo.ExecuteWithTryCatch(() =>
                                    {
                                        currentSite.RootWeb.SiteLogoUrl = string.Concat(currentSite.RootWeb.ServerRelativeUrl, site.Webs.RootWeb.Logo.Url);
                                        currentSite.RootWeb.Update();
                                        currentSite.RootWeb.Context.ExecuteQuery();
                                    },
                                    string.Format("Actualizando logo: [Url: {0}]", site.Webs.RootWeb.Logo.Url));
                                }
                            }

                            #endregion

                            #region Lists

                            if (site.Webs.RootWeb.Lists != null && site.Webs.RootWeb.Lists.Provision)
                                foreach (TenantSiteWebsRootWebListsList list in site.Webs.RootWeb.Lists.List)
                                {
                                    Apoyo.ExecuteWithTryCatch(() =>
                                    {
                                        List newList = null;

                                        #region Creación / Obtención de la lista
                                        //Creación de la lista
                                        if (!currentSite.RootWeb.ListExists(list.Name)) 
                                        {
                                            
                                            newList = currentSite.RootWeb.CreateList(
                                                                (ListTemplateType)Enum.Parse(typeof(ListTemplateType), list.TemplateType, true),
                                                                list.Name,
                                                                list.EnableVersioning,
                                                                true,
                                                                list.UrlPath.Replace("\\", "/"),
                                                                list.EnableContentTypes);
                                        }

                                        if (newList == null)
                                            newList = currentSite.RootWeb.GetListByTitle(list.Name);

                                        clientContext.Load(newList, w => w.ContentTypes);
                                        clientContext.ExecuteQueryRetry();

                                        #endregion

                                        #region Asociación de Content Types

                                        if (list.ContentTypes != null && list.ContentTypes.Provision)
                                            foreach (TenantSiteWebsRootWebListsListContentTypesContentType contentType in list.ContentTypes.ContentType)
                                            {
                                                if (!newList.ContentTypeExistsByName(contentType.Name))
                                                {
                                                    Apoyo.ExecuteWithTryCatch(() =>
                                                    {
                                                        newList.AddContentTypeToListByName(contentType.Name, contentType.SetAsDefault, contentType.SearchContentTypeInSiteHierarchy);
                                                    },
                                                    string.Format("Asociando contentType: [Nombre: {0}]", contentType.Name));
                                                }
                                            }

                                        #endregion

                                        #region Eliminación de Content Types

                                        if (list.ContentTypes != null && list.ContentTypes.Provision)
                                            foreach (TenantSiteWebsRootWebListsListContentTypesRemoveContentType removeContentType in list.ContentTypes.RemoveContentType)
                                            {
                                                Apoyo.ExecuteWithTryCatch(() =>
                                                {
                                                    //Asociamos el Content-Type a la lista
                                                    newList.RemoveContentTypeByName(removeContentType.Name);
                                                    clientContext.ExecuteQueryRetry();
                                                },
                                                string.Format("Eliminando contentType de la lista: [Nombre: {0}]", removeContentType.Name));
                                            }

                                        #endregion

                                        #region Asignación de permisos a la lista

                                        if (list.Permissions != null)
                                        {
                                            Apoyo.ExecuteWithTryCatch(() =>
                                            {
                                                Console.WriteLine();

                                                Apoyo.ExecuteWithTryCatch(() =>
                                                {
                                                    newList.BreakRoleInheritance(list.Permissions.CopyRoleAssignments, list.Permissions.ClearSubscopes);
                                                    clientContext.ExecuteQueryRetry();
                                                },
                                                "Rompiendo herencia");

                                                //Cargamos todos los grupos
                                                clientContext.Load(currentSite.RootWeb.SiteGroups);
                                                clientContext.ExecuteQueryRetry();

                                                #region Agregar permisos

                                                Apoyo.ExecuteWithTryCatch(() =>
                                                {
                                                    //Agregamos permisos
                                                    foreach (TenantSiteWebsRootWebListsListPermissionsAddRoleAssignment addRoleAssignment in list.Permissions.AddRoleAssignment)
                                                    {
                                                        //Obtenemos el grupo
                                                        Principal principal = currentSite.RootWeb.SiteGroups.Where(w => w.Title == addRoleAssignment.Name).FirstOrDefault();

                                                        //Almacenamos el rol del grupo
                                                        RoleDefinition roleDefinition = currentSite.RootWeb.RoleDefinitions.GetByType((RoleType)Enum.Parse(typeof(RoleType), addRoleAssignment.RoleType, true));

                                                        if (addRoleAssignment.RemoveExistingRoleDefinitions)
                                                        {
                                                            newList.RoleAssignments.GetByPrincipal(principal).DeleteObject();
                                                            newList.Context.ExecuteQueryRetry();
                                                        }

                                                        //Creamos una nueva definición de rol
                                                        RoleDefinitionBindingCollection rdc = new RoleDefinitionBindingCollection(clientContext);
                                                        rdc.Add(roleDefinition);
                                                        newList.RoleAssignments.Add(principal, rdc);
                                                    }

                                                    newList.Context.ExecuteQueryRetry();
                                                },
                                                "Agregando RoleDefinitions");

                                                #endregion

                                                #region Eliminar permisos

                                                Apoyo.ExecuteWithTryCatch(() =>
                                                {
                                                    //Agregamos permisos
                                                    foreach (TenantSiteWebsRootWebListsListPermissionsRemoveRoleAssignment removeRoleAssignment in list.Permissions.RemoveRoleAssignment)
                                                    {
                                                        //Obtenemos el grupo
                                                        Principal principal = currentSite.RootWeb.SiteGroups.Where(w => w.Title == removeRoleAssignment.Name).FirstOrDefault();

                                                        //Cargamos los roleAssignments
                                                        clientContext.Load(newList.RoleAssignments);
                                                        clientContext.ExecuteQueryRetry();

                                                        foreach (RoleAssignment rol in newList.RoleAssignments)
                                                            if (rol.PrincipalId == principal.Id)
                                                            {
                                                                rol.DeleteObject();
                                                                newList.Context.ExecuteQueryRetry();
                                                                break;
                                                            }
                                                    }
                                                },
                                                "Eliminando RoleDefinitions");

                                                #endregion
                                            },
                                            string.Format("Asignando permisos a la lista"));
                                        }

                                        #endregion

                                        #region Creación de carpetas

                                        if (list.Folders != null && list.Folders.Provision)
                                            foreach (TenantSiteWebsRootWebListsListFoldersFolder folder in list.Folders.Folder)
                                            {
                                                if (!newList.RootFolder.FolderExists(folder.Name))
                                                {
                                                    Apoyo.ExecuteWithTryCatch(() =>
                                                    {
                                                        newList.RootFolder.CreateFolder_Hiberus(folder.Name);
                                                    },
                                                    string.Format("Creando carpeta: [Nombre: {0}]", folder.Name));
                                                }
                                            }

                                        #endregion

                                        #region Creación de vistas

                                        if (!string.IsNullOrEmpty(list.ViewsSourcePath) )
                                            Apoyo.ExecuteWithTryCatch(() =>
                                            {
                                                currentSite.RootWeb.CreateViewsFromXMLFile_Hiberus(string.Format("{0}\\{1}", site.Url.Replace("/", "\\"), list.UrlPath), list.ViewsSourcePath);
                                            }, string.Format("Se ha añadido la vista {0}", list.ViewsSourcePath));

                                        #endregion

                                        #region Encuestas: Creación de Preguntas

                                        if (list.TemplateType == "Survey")
                                        {
                                            if (list.Questions != null && list.Questions.Provision)
                                            {
                                                foreach (TenantSiteWebsRootWebListsListQuestionsQuestion question in list.Questions.Question)
                                                {
                                                    Console.WriteLine();

                                                    List surveyList = currentSite.RootWeb.GetListByTitle(list.Name);

                                                    Apoyo.ExecuteWithTryCatch(() =>
                                                    {
                                                        Extensions.ListExtensions.CreateQuestion_Hiberus(question.FieldSourcePath, surveyList);
                                                    });
                                                }
                                            }
                                        }

                                        #endregion

                                    },
                                    string.Format("Creando lista: [Nombre: {0}] [Template: {1}]", list.Name, list.TemplateType));
                                }

                            #endregion

                            #region Apps

                            if (site.Webs.RootWeb.Apps != null && site.Webs.RootWeb.Apps.Provision)
                            {
                                Guid developmentFeatureId = new Guid("e374875e-06b6-11e0-b0fa-57f5dfd72085");

                                try
                                {
                                    //Activamos la característica de desarrollo para poder activar Apps desde CSOM
                                    if (!currentSite.IsFeatureActive(developmentFeatureId))
                                        currentSite.ActivateFeature(developmentFeatureId);

                                    Site siteAppCatalog = o365Tenant.GetSiteByUrl(string.Concat(tenant.Url, site.Webs.RootWeb.Apps.AppCatalogUrl));

                                    foreach (TenantSiteWebsRootWebAppsApp app in site.Webs.RootWeb.Apps.App)
                                    {
                                        //Comprobamos si ya existe alguna instancia instalada
                                        ClientObjectList<AppInstance> instances = currentSite.RootWeb.GetAppInstancesByProductId(new Guid(app.ProductId));
                                        clientContext.Load(instances);
                                        clientContext.ExecuteQueryRetry();

                                        if (instances.Count <= 0)
                                        {
                                            Apoyo.ExecuteWithTryCatch(() =>
                                            {
                                                Microsoft.SharePoint.Client.File fileApp = siteAppCatalog.RootWeb.GetFileByServerRelativeUrl(
                                                    string.Format("/{0}/{1}/{2}", site.Webs.RootWeb.Apps.AppCatalogUrl, site.Webs.RootWeb.Apps.AppCatalogListUrl, app.AppFileName));

                                                clientContext.Load(fileApp);
                                                clientContext.ExecuteQueryRetry();

                                                ClientResult<Stream> data = fileApp.OpenBinaryStream();
                                                clientContext.ExecuteQueryRetry();

                                                //Instalación de la APP
                                                AppInstance instance = currentSite.RootWeb.LoadAndInstallAppInSpecifiedLocale(data.Value, 3082);

                                                #region Esperamos a que se haya instalado correctamente la APP

                                                //Obtenemos el ID de instancia
                                                clientContext.Load(instance, w => w.Id, w => w.Status);
                                                clientContext.ExecuteQueryRetry();

                                                Guid instanceID = instance.Id;

                                                int maxTry = 15;
                                                int count = 0;
                                                do
                                                {
                                                    System.Threading.Thread.Sleep(2000);
                                                    instance = currentSite.RootWeb.GetAppInstanceById(instanceID);
                                                    clientContext.Load(instance, w => w.Status);
                                                    clientContext.ExecuteQueryRetry();
                                                    count++;
                                                }
                                                while (instance != null && instance.Status != AppInstanceStatus.Installed && count < maxTry);

                                                #endregion
                                            },
                                            string.Format("Instalando App: [ProductID: {0}]", app.ProductId));
                                        }
                                    }
                                }
                                finally
                                {
                                    if (currentSite.IsFeatureActive(developmentFeatureId))
                                        currentSite.DeactivateFeature(developmentFeatureId);
                                }
                            }

                            #endregion

                            #region Pages

                            if (site.Webs.RootWeb.Pages != null && site.Webs.RootWeb.Pages.Provision)
                                foreach (TenantSiteWebsRootWebPagesPublishingPage page in site.Webs.RootWeb.Pages.PublishingPage)
                                {
                                    Apoyo.ExecuteWithTryCatch(() =>
                                    {
                                        string paginaURL = string.Format("Paginas/{0}.aspx", page.Name);

                                        //Si tiene la opción de borrar, la borramos
                                        if (currentSite.RootWeb.FileExists(paginaURL) && page.RemoveIfExists)
                                        {
                                            //Obtenemos la URL relativa del sitio
                                            clientContext.Load(currentSite, w => w.RootWeb.ServerRelativeUrl);
                                            clientContext.ExecuteQueryRetry();

                                            Microsoft.SharePoint.Client.File pagina = currentSite.RootWeb.GetFileByServerRelativeUrl(string.Format("{0}/{1}", currentSite.RootWeb.ServerRelativeUrl, paginaURL));
                                            pagina.DeleteObject();
                                            currentSite.RootWeb.Update();
                                            clientContext.ExecuteQueryRetry();
                                        }

                                        //Si la página no existe la creamos
                                        if (!currentSite.RootWeb.FileExists(paginaURL))
                                        {
                                            #region Creación de la página

                                            //Obtenemos la URL relativa del sitio
                                            clientContext.Load(currentSite, w => w.ServerRelativeUrl);
                                            clientContext.ExecuteQueryRetry();

                                            //Obtenemos el PageLayout
                                            Microsoft.SharePoint.Client.File pageFromPageLayout = currentSite.RootWeb.GetFileByServerRelativeUrl(String.Format("{0}_catalogs/masterpage/{1}.aspx",
                                            UrlUtility.EnsureTrailingSlash(currentSite.ServerRelativeUrl),
                                            page.Layout));

                                            ListItem pageLayoutItem = pageFromPageLayout.ListItemAllFields;
                                            clientContext.Load(pageLayoutItem);
                                            clientContext.ExecuteQueryRetry();

                                            //Obtenemos el objeto web de publicación
                                            PublishingWeb publishingWeb = PublishingWeb.GetPublishingWeb(clientContext, currentSite.RootWeb);
                                            clientContext.Load(publishingWeb);

                                            PublishingPage mPage = publishingWeb.AddPublishingPage(new PublishingPageInformation
                                            {
                                                Name = string.Format("{0}.aspx", page.Name),
                                                PageLayoutListItem = pageLayoutItem
                                            });

                                            #endregion

                                            #region Asociación de webparts
                                         
                                            //Asociamos los webparts
                                            if (page.WebParts != null)
                                            foreach (TenantSiteWebsRootWebPagesPublishingPageWebPart webpart in page.WebParts)
                                            {
                                                Apoyo.ExecuteWithTryCatch(() =>
                                                {
                                                    //Obtenemos el XML del webpart
                                                    string webPartXml = System.IO.File.ReadAllText(string.Format("{0}\\{1}", webpart.SourcePath, webpart.Name));

                                                    webPartXml = webPartXml.Contains(Apoyo.ResolveSiteCollection) ? webPartXml.Replace(Apoyo.ResolveSiteCollection, currentSite.RootWeb.ServerRelativeUrl) : webPartXml;

                                                    //Creamos la entidad de webpart
                                                    OfficeDevPnP.Core.Entities.WebPartEntity webPartEntity = new OfficeDevPnP.Core.Entities.WebPartEntity()
                                                    {
                                                        WebPartTitle = webpart.Title,
                                                        WebPartXml = webPartXml,
                                                        WebPartZone = webpart.Zone,
                                                        WebPartIndex = webpart.Index
                                                    };

                                                    //Agregamos el webpart
                                                    currentSite.RootWeb.AddWebPartToWebPartPage(webPartEntity, paginaURL);
                                                },
                                                string.Format("Asociando webpart: [Nombre: {0}]", webpart.Name));
                                            }

                                            #endregion

                                            #region Actualización del título de página

                                            //Cargamos la página para establecer el título y hacer el check-in
                                            clientContext.Load(mPage.ListItem, w => w.ParentList);
                                            clientContext.ExecuteQueryRetry();

                                            //Actualización del título de la página
                                            ListItem pageItem = mPage.ListItem;
                                            pageItem["Title"] = page.Title;
                                            pageItem.Update();

                                            #endregion

                                            #region Check-in y Publish

                                            clientContext.Load(pageItem, p => p.File.CheckOutType);
                                            clientContext.ExecuteQueryRetry();

                                            //Check-in
                                            if (pageItem.File.CheckOutType != CheckOutType.None)
                                                pageItem.File.CheckIn(String.Empty, CheckinType.MajorCheckIn);

                                            //Publicación
                                            pageItem.File.Publish(String.Empty);
                                            if (pageItem.ParentList.EnableModeration)
                                                pageItem.File.Approve(String.Empty);

                                            clientContext.ExecuteQueryRetry();

                                            #endregion
                                        }
                                    },
                                    string.Format("Creando la página: [Nombre: {0}] [Layout: {1}]", page.Name, page.Layout));
                                }

                            #endregion

                            #region Search

                            if (site.Webs.RootWeb.Search != null && site.Webs.RootWeb.Search.Provision)
                            {
                                //Página principal de búsqueda
                                Apoyo.ExecuteWithTryCatch(() =>
                                {
                                    //Ruta por defecto del sitio web
                                    String searchDefaultPath = String.Format("{{\"Inherit\":{0},\"ResultsPageAddress\":\"~site/{1}.aspx\",\"ShowNavigation\":{2}}}",
                                                                                                        site.Webs.RootWeb.Search.Inherit.ToString().ToLower(),
                                                                                                        site.Webs.RootWeb.Search.DefaultResultsPage,
                                                                                                        site.Webs.RootWeb.Search.ShowNavigation.ToString().ToLower());

                                    currentSite.RootWeb.SetPropertyBagValue("SRCH_SB_SET_WEB", searchDefaultPath);

                                },
                                String.Format("Cambiando ruta de búsqueda predeterminada: [Nueva pagina: {0}] ]", site.Webs.RootWeb.Search.DefaultResultsPage));

                                //Nodos de navegación de búsqueda
                                if (site.Webs.RootWeb.Search.SearchNodes != null && site.Webs.RootWeb.Search.SearchNodes.Provision)
                                {
                                    currentSite.RootWeb.DeleteAllNavigationNodes(OfficeDevPnP.Core.Enums.NavigationType.SearchNav);
                                    foreach (TenantSiteWebsRootWebSearchSearchNodesSearchNode searchNode in site.Webs.RootWeb.Search.SearchNodes.SearchNode)
                                    {
                                        Apoyo.ExecuteWithTryCatch(() =>
                                        {
                                            //Ruta final del nodo de navegación
                                            String targetPath = String.Format("{0}{1}{2}", tenant.Url,
                                                                                            UrlUtility.EnsureTrailingSlash(currentSite.RootWeb.ServerRelativeUrl.Remove(0, 1).ToString()),
                                                                                            searchNode.TargetPath);

                                            //Añado el nodo a la búsqueda del sitio
                                            currentSite.RootWeb.AddNavigationNode(searchNode.Title,
                                                                                  new Uri(targetPath),
                                                                                  searchNode.ParentNodeTitle,
                                                                                  OfficeDevPnP.Core.Enums.NavigationType.SearchNav);

                                        },
                                        String.Format("Añadido nuevo nodo de búsqueda: [Nuevo nodo: {0}] ]", searchNode.Title));
                                    }
                                }
                            }

                            #endregion

                            #region Search Settings

                            if (!String.IsNullOrEmpty(site.Webs.RootWeb.SearchSettingsSourcePath))
                                Apoyo.ExecuteWithTryCatch(() =>
                                {
                                    using (ClientContext rootClientContext = new ClientContext(currentSite.RootWeb.Url))
                                    {
                                        rootClientContext.Credentials = credentials;
                                        rootClientContext.ImportSearchSettings(site.Webs.RootWeb.SearchSettingsSourcePath, SearchObjectLevel.SPSite);
                                    }
                                }, String.Format("Se ha aplicado la configuración de búsqueda"));

                            #endregion
                        }

                        #endregion

                        #region Webs

                        if (site.Webs.Provision && site.Webs.Web != null)

                            foreach (TenantSiteWebsWeb web in site.Webs.Web)
                            {
                                if (!web.Provision)
                                    continue;

                                Console.WriteLine(string.Empty);
                                Console.WriteLine(">> Provisionando la Web {0} <<", web.Title);
                                Console.WriteLine(string.Empty);

                                Web currentWeb = null;

                                #region Creacion de Web (si ya existe se omite el paso)

                                if (!currentSite.RootWeb.WebExists(web.Url))
                                {
                                    Apoyo.ExecuteWithTryCatch(() =>
                                    {
                                        SiteEntity newWeb = new SiteEntity()
                                        {
                                            Title = web.Title,
                                            Description = web.Description,
                                            Url = web.Url,
                                            Lcid = web.lcid,
                                            SiteOwnerLogin = web.SiteOwnerLogin,
                                            Template = web.Template,
                                        };

                                        currentWeb = currentSite.RootWeb.CreateWeb(newWeb, web.InheritPermissions, web.InheritNavigation);
                                    },
                                    string.Format("Creando Web: [URL: {0}]", web.Url));
                                }

                                if (currentWeb == null)
                                    currentWeb = currentSite.RootWeb.GetWeb(web.Url);

                                clientContext.Load(currentWeb, w => w.ServerRelativeUrl);
                                clientContext.ExecuteQueryRetry();

                                #endregion

                                #region Web Features

                                if (web.Features != null && web.Features.Provision)
                                    foreach (TenantSiteWebsWebFeaturesFeature feature in web.Features.Feature)
                                    {
                                        if (!currentWeb.IsFeatureActive(new Guid(feature.ID)))
                                        {
                                            Apoyo.ExecuteWithTryCatch(() =>
                                            {
                                                currentWeb.ActivateFeature(new Guid(feature.ID));
                                            },
                                            string.Format("Activando feature: [ID: {0}]", feature.ID));
                                        }
                                    }

                                #endregion

                                #region SiteColumns

                                if (web.SiteColumns != null && web.SiteColumns.Provision)
                                {
                                    foreach (TenantSiteWebsWebSiteColumnsField field in web.SiteColumns.Field)
                                    {
                                        Apoyo.ExecuteWithTryCatch(() =>
                                        {
                                            //Leemos el XML
                                            string xmlString = System.IO.File.ReadAllText(string.Format("{0}\\{1}", web.SiteColumns.SourcePath, field.SourceXML));

                                            //Generamos un objeto XMLReader
                                            XmlDocument xmlDocument = new XmlDocument();
                                            xmlDocument.LoadXml(xmlString);

                                            //Obtenemos el Nodo Field
                                            XmlNode fieldNode = xmlDocument.GetElementsByTagName("Field").Item(0);

                                            //Si ya está creado el campo en SharePoint continuamos
                                            if (currentWeb.FieldExistsById(fieldNode.Attributes["ID"].Value))
                                            {
                                                Console.Write(" [Ya existe]");

                                                return;
                                            }

                                            #region LookupField

                                            if (fieldNode.Attributes["Type"].Value.Contains("Lookup"))
                                            {
                                                List LookupList = null;

                                                #region Creación de la lista requerida

                                                //Creación de la lista requerida (Si es necesario)
                                                if (field.RequiredList != null)
                                                    if (!currentWeb.ListExists(field.RequiredList.Name))
                                                        Apoyo.ExecuteWithTryCatch(() =>
                                                        {
                                                            LookupList = currentWeb.CreateList(
                                                                                (ListTemplateType)Enum.Parse(typeof(ListTemplateType), field.RequiredList.TemplateType, true),
                                                                                field.RequiredList.Name,
                                                                                field.RequiredList.EnableVersioning,
                                                                                true,
                                                                                field.RequiredList.UrlPath.Replace("\\", "/"),
                                                                                field.RequiredList.EnableContentTypes);
                                                        },
                                                        string.Format("Creando lista requerida: [Name: {0}]", field.RequiredList.Name));

                                                if (LookupList == null)
                                                    LookupList = currentWeb.GetListByUrl(field.RequiredList.UrlPath);

                                                #endregion

                                                clientContext.Load(LookupList, w => w.Id);
                                                clientContext.ExecuteQueryRetry();

                                                foreach (XmlNode fieldLookupField in xmlDocument.GetElementsByTagName("Field"))
                                                {
                                                    #region Asignación del ID de lista

                                                    fieldLookupField.Attributes["List"].Value = LookupList.Id.ToString("B");

                                                    #endregion

                                                    #region Asignación del ID de la web

                                                    clientContext.Load(currentWeb, w => w.Id);
                                                    clientContext.ExecuteQueryRetry();

                                                    fieldLookupField.Attributes["WebId"].Value = currentWeb.Id.ToString("D");

                                                    #endregion

                                                    //Creamos el SPField a partir del XML
                                                    currentWeb.CreateField(fieldLookupField.OuterXml, true);
                                                }

                                                return;
                                            }

                                            #endregion

                                            #region TaxonomyField

                                            //Comprobamos si el campo es de tipo taxonomía para crearlo a partir de la función de Office PnP
                                            if (fieldNode.Attributes["Type"].Value == "TaxonomyFieldType" || fieldNode.Attributes["Type"].Value == "TaxonomyFieldTypeMulti")
                                            {
                                                //Obtenemos el almacén de términos del site collection
                                                TermStore termStore = currentSite.GetDefaultSiteCollectionTermStore();

                                                //Obtenemos el grupo de términos asociado a la columna
                                                TermGroup termGroup = termStore.Groups.GetByName(field.TermGroupName);

                                                //Obtenemos el conjunto de términos asociado a la columna
                                                TermSet termSet = termGroup.TermSets.GetByName(field.TermSetName);

                                                clientContext.Load(termStore);
                                                clientContext.Load(termSet);
                                                clientContext.ExecuteQueryRetry();

                                                //Atributos adicionales (LCID)
                                                List<KeyValuePair<string, string>> additionalAttributes = new List<KeyValuePair<string, string>>();
                                                additionalAttributes.Add(new KeyValuePair<string, string>("ShowField", string.Format("Term{0}", field.LCID)));

                                                TaxonomyFieldCreationInformation taxonomyFieldCreation = new TaxonomyFieldCreationInformation()
                                                {
                                                    Id = new Guid(fieldNode.Attributes["ID"].Value),
                                                    InternalName = fieldNode.Attributes["Name"].Value,
                                                    DisplayName = fieldNode.Attributes["DisplayName"].Value,
                                                    Group = fieldNode.Attributes["Group"].Value,
                                                    TaxonomyItem = termSet,
                                                    AdditionalAttributes = additionalAttributes,
                                                    MultiValue = (fieldNode.Attributes["Type"].Value == "TaxonomyFieldTypeMulti") ? true : false
                                                };

                                                currentWeb.CreateTaxonomyField(taxonomyFieldCreation);

                                                return;
                                            }

                                            #endregion

                                            //Creamos el SPField a partir del XML
                                            currentWeb.CreateField(fieldNode.OuterXml, true);
                                        },
                                        string.Format("Creando columna: [SourceXML: {0}]", field.SourceXML));
                                    }
                                }

                                #endregion

                                #region ContentTypes

                                if (web.ContentTypes != null && web.ContentTypes.Provision)
                                {
                                    foreach (TenantSiteWebsWebContentTypesContentType contentType in web.ContentTypes.ContentType)
                                    {
                                        Apoyo.ExecuteWithTryCatch(() =>
                                        {
                                            //Obtenemos la ruta del XML del contentType
                                            string contentTypeXMLPath = string.Format("{0}\\{1}\\{2}", Environment.CurrentDirectory, web.ContentTypes.SourcePath, contentType.SourceXML);

                                            //Leemos el XML
                                            string xmlString = System.IO.File.ReadAllText(contentTypeXMLPath);

                                            //Generamos un objeto XMLReader
                                            XmlDocument xmlDocument = new XmlDocument();
                                            xmlDocument.LoadXml(xmlString);

                                            //Obtenemos el Nodo Field
                                            XmlNode contentTypeNode = xmlDocument.GetElementsByTagName("ContentType").Item(0);

                                            if (currentWeb.ContentTypeExistsByName(contentTypeNode.Attributes["Name"].Value))
                                            {
                                                Console.Write(" [Ya existe]");
                                                return;
                                            }

                                            //Creamos el contentType a partir del XML
                                            ContentType ct = currentWeb.CreateContentTypeFromXMLHiberus(XDocument.Load(contentTypeXMLPath));

                                            #region Configuración Document Set

                                            //Comprobamos si el Tipo de contenido parte de un conjunto a partir de la herencia del Identificador
                                            if (contentTypeNode.Attributes["ID"].Value.StartsWith("0x0120D520"))
                                            {
                                                DocumentSetTemplate documentSetTemplate = DocumentSetTemplate.GetDocumentSetTemplate(currentWeb.Context, ct);
                                                clientContext.Load(documentSetTemplate, w => w.AllowedContentTypes, w => w.SharedFields);
                                                clientContext.ExecuteQueryRetry();

                                                //Eliminamos el tipo de documento base (Documento)
                                                ContentType documentoCT = currentSite.RootWeb.GetContentTypeById("0x0101");
                                                documentSetTemplate.AllowedContentTypes.Remove(documentoCT.Id);

                                                documentSetTemplate.Update(true);
                                                clientContext.Load(documentSetTemplate);
                                                clientContext.ExecuteQueryRetry();

                                                if (xmlDocument.GetElementsByTagName("AllowedContentType").Count > 0)
                                                    foreach (XmlNode associatedContentType in xmlDocument.GetElementsByTagName("AllowedContentType"))
                                                    {
                                                        //Asociamos los tipos de contenido
                                                        ContentType rootWebCT = currentWeb.GetContentTypeById(associatedContentType.Attributes["ID"].Value);
                                                        if (rootWebCT == null)
                                                            rootWebCT = currentSite.RootWeb.GetContentTypeById(associatedContentType.Attributes["ID"].Value);

                                                        documentSetTemplate.AllowedContentTypes.Add(rootWebCT.Id);

                                                        documentSetTemplate.Update(true);
                                                        clientContext.Load(documentSetTemplate);
                                                        clientContext.ExecuteQueryRetry();
                                                    }

                                                if (xmlDocument.GetElementsByTagName("SharedField").Count > 0)
                                                    foreach (XmlNode associatedField in xmlDocument.GetElementsByTagName("SharedField"))
                                                    {
                                                        Field rootWebField = currentSite.RootWeb.GetFieldById<Field>(new Guid(associatedField.Attributes["ID"].Value));
                                                        if (rootWebField == null)
                                                            rootWebField = currentWeb.GetFieldById<Field>(new Guid(associatedField.Attributes["ID"].Value));

                                                        //Asociamos las columnas compartidas
                                                        documentSetTemplate.SharedFields.Add(rootWebField);

                                                        documentSetTemplate.Update(true);
                                                        clientContext.Load(documentSetTemplate);
                                                        clientContext.ExecuteQueryRetry();
                                                    }

                                                if (xmlDocument.GetElementsByTagName("WelcomePageField").Count > 0)
                                                    foreach (XmlNode associatedWelcomePageField in xmlDocument.GetElementsByTagName("WelcomePageField"))
                                                    {
                                                        Field rootWebWelcomeField = currentSite.RootWeb.GetFieldById<Field>(new Guid(associatedWelcomePageField.Attributes["ID"].Value));
                                                        if (rootWebWelcomeField == null)
                                                            rootWebWelcomeField = currentWeb.GetFieldById<Field>(new Guid(associatedWelcomePageField.Attributes["ID"].Value));

                                                        //Asociamos los campos de la página de inicio
                                                        documentSetTemplate.WelcomePageFields.Add(rootWebWelcomeField);

                                                        documentSetTemplate.Update(true);
                                                        clientContext.Load(documentSetTemplate);
                                                        clientContext.ExecuteQueryRetry();
                                                    }
                                            }

                                            #endregion
                                        },
                                        string.Format("Creando contentType: [SourceXML: {0}]", contentType.SourceXML));
                                    }
                                }

                                #endregion

                                #region Catalogs

                                if (web.Catalogs != null && web.Catalogs.Provision)
                                {
                                    Folder folder = currentWeb.GetCatalog((int)ListTemplateType.MasterPageCatalog).RootFolder;
                                    clientContext.Load(folder);
                                    clientContext.ExecuteQueryRetry();

                                    //MasterPages
                                    if (web.Catalogs.Masterpages != null && web.Catalogs.Masterpages.Provision)
                                        foreach (TenantSiteWebsWebCatalogsMasterpagesMasterpage masterPage in web.Catalogs.Masterpages.Masterpage)
                                        {
                                            Apoyo.ExecuteWithTryCatch(() =>
                                            {
                                                //El _catalogs de un subsitio es diferente al del root web, no crea el fichero .master a partir del html
                                                //No se puede referenciar al rootweb el tema, porque se necesita el fichero .preview para poder cambiar el aspecto 
                                                //y poder seleccionar un background predeterminado para el subsitio. 
                                                //Por lo tanto se opta por copiar el fichero .master y el preview y moverlos a la carpeta _catalogs del subsitio
                                                Microsoft.SharePoint.Client.File masterFile = currentSite.RootWeb.GetFileByServerRelativeUrl(String.Format("{0}_catalogs/masterpage/{1}",
                                                                                                                                          UrlUtility.EnsureTrailingSlash(currentSite.RootWeb.ServerRelativeUrl),
                                                                                                                                          masterPage.Name));

                                                ClientResult<Stream> data = masterFile.OpenBinaryStream();
                                                clientContext.Load(masterFile);
                                                clientContext.ExecuteQueryRetry();

                                                //Master Upload
                                                folder.UploadFile(masterPage.Name, data.Value, true);

                                                //Preview Upload
                                                folder.UploadFile(masterPage.Preview,
                                                                    Path.Combine(web.Catalogs.Masterpages.SourcePath, masterPage.Preview),
                                                                    true);
                                            },
                                            string.Format("Subiendo MasterPage: [Nombre: {0}]", masterPage.Name));
                                        }
                                }

                                #endregion

                                #region Permissions

                                //Asignación de permisos de grupo
                                if (web.Permissions != null && web.Permissions.Provision)
                                {
                                    //Rompemos herencia
                                    if (web.Permissions.BreakRoleInheritance)
                                    {
                                        currentWeb.BreakRoleInheritance(true, false);
                                        clientContext.ExecuteQueryRetry();
                                    }

                                    //Agregamos permisos
                                    foreach (TenantSiteWebsWebPermissionsAddPermission addPermission in web.Permissions.AddPermission)
                                    {
                                        Apoyo.ExecuteWithTryCatch(() =>
                                        {
                                            //Obtenemos el grupo
                                            var group = currentSite.RootWeb.SiteGroups.GetByName(addPermission.Name);

                                            //Asignamos permisos de contribución al grupo correspondiente de la web correspondiente
                                            currentWeb.AddPermissionLevelToGroup(addPermission.Name, (RoleType)Enum.Parse(typeof(RoleType), addPermission.RoleType, true), true);

                                            //Establecemos el grupo como grupo por defecto asociado de integrantes
                                            if (addPermission.IsAssociatedMemberGroup)
                                                currentWeb.AssociateDefaultGroups(null, group, null);

                                            if (addPermission.IsAssociatedVisitorGroup)
                                                currentWeb.AssociateDefaultGroups(null, null, group);

                                            if (addPermission.IsAssociatedOwnerGroup)
                                                currentWeb.AssociateDefaultGroups(group, null, null);
                                        },
                                        string.Format("Asignando permiso: [Nombre: {0}] [Rol: {1}]", addPermission.Name, addPermission.RoleType));
                                    }
                                }

                                #endregion

                                #region Tema

                                if (web.Theme != null && web.Theme.Provision)
                                {
                                    //Crear
                                    Apoyo.ExecuteWithTryCatch(() =>
                                    {
                                        clientContext.Load(currentWeb, w => w.ServerRelativeUrl);
                                        clientContext.ExecuteQueryRetry();

                                        currentWeb.CreateComposedLookByUrl(
                                                    web.Theme.Titulo,
                                                    String.Format("{0}_catalogs/Theme/15/{1}",
                                                    UrlUtility.EnsureTrailingSlash(currentSite.RootWeb.ServerRelativeUrl),
                                                    web.Theme.Colors),
                                                    String.Format("{0}_catalogs/Theme/15/{1}",
                                                    UrlUtility.EnsureTrailingSlash(currentSite.RootWeb.ServerRelativeUrl),
                                                    web.Theme.Fonts),
                                                    string.IsNullOrEmpty(web.Theme.BackgroundImage) ? string.Empty : String.Format("{0}_catalogs/Theme/15/{1}", UrlUtility.EnsureTrailingSlash(currentSite.RootWeb.ServerRelativeUrl), web.Theme.BackgroundImage),
                                                    String.Format("{0}_catalogs/masterpage/{1}",
                                                    UrlUtility.EnsureTrailingSlash(currentWeb.ServerRelativeUrl),
                                                    web.Theme.MasterPage),
                                                    1, true);

                                    }, string.Format("Creado Theme: [Nombre: {0}]", web.Theme.Titulo));

                                    //Set
                                    Apoyo.ExecuteWithTryCatch(() =>
                                    {
                                        currentWeb.SetComposedLookByUrl(web.Theme.Titulo);

                                    }, string.Format("Cambio a Theme: [Nombre: {0}]", web.Theme.Titulo));

                                    //Cambio pagina maestra de sistema
                                    if (!string.IsNullOrEmpty(web.Theme.SystemMasterPage))
                                        Apoyo.ExecuteWithTryCatch(() =>
                                        {
                                            Folder folder = currentWeb.GetCatalog((int)ListTemplateType.MasterPageCatalog).RootFolder;
                                            clientContext.Load(folder);
                                            clientContext.Load(currentWeb, w => w.MasterUrl);
                                            clientContext.ExecuteQueryRetry();

                                            currentWeb.MasterUrl = String.Concat(UrlUtility.EnsureTrailingSlash(folder.ServerRelativeUrl), web.Theme.SystemMasterPage);
                                            currentWeb.Update();
                                            clientContext.ExecuteQueryRetry();
                                        }, string.Format("Cambio a pagina de sistema: [Nombre: {0}]", web.Theme.SystemMasterPage));
                                }

                                #endregion

                                #region Logo

                                if (web.Logo != null && web.Logo.Provision)
                                {
                                    if (!string.IsNullOrWhiteSpace(web.Logo.Url))
                                    {
                                        Apoyo.ExecuteWithTryCatch(() =>
                                        {
                                            //Obtenemos la URL relativa del sitio para apuntar a la Style Library de la Site Collection
                                            clientContext.Load(currentSite, w => w.ServerRelativeUrl);
                                            clientContext.ExecuteQueryRetry();

                                            currentWeb.SiteLogoUrl = String.Concat(currentSite.ServerRelativeUrl, web.Logo.Url);
                                            currentWeb.Update();
                                            currentWeb.Context.ExecuteQuery();
                                        },
                                        string.Format("Actualizando logo: [Url: {0}]", web.Logo.Url));
                                    }
                                }

                                #endregion

                                #region Lists (DEPRECATED)

                                //if (web.Lists.Provision)
                                //    foreach (TenantSiteWebsWebListsList list in web.Lists.List)
                                //    {
                                //        Apoyo.ExecuteWithTryCatch(() =>
                                //        {
                                //            List newList = null;
                                //            string ctAsDefault = string.Empty;

                                //            #region Creación / Obtención de la lista

                                //            //Creación de la lista
                                //            if (!currentWeb.ListExists(list.Name))
                                //                newList = currentWeb.CreateList(
                                //                                    (ListTemplateType)Enum.Parse(typeof(ListTemplateType), list.TemplateType, true),
                                //                                    list.Name,
                                //                                    list.EnableVersioning,
                                //                                    true,
                                //                                    list.UrlPath.Replace("\\", "/"),
                                //                                    list.EnableContentTypes);

                                //            if (newList == null)
                                //                newList = currentWeb.GetListByTitle(list.Name);

                                //            clientContext.Load(newList, w => w.ContentTypes);
                                //            clientContext.ExecuteQueryRetry();

                                //            #endregion

                                //            #region Asociación de Content Types

                                //            if (list.ContentTypes != null && list.ContentTypes.Provision)
                                //                foreach (TenantSiteWebsWebListsListContentTypesContentType contentType in list.ContentTypes.ContentType)
                                //                {
                                //                    //Lo asociamos si no existe
                                //                    if (!newList.ContentTypeExistsByName(contentType.Name))
                                //                    {
                                //                        Apoyo.ExecuteWithTryCatch(() =>
                                //                        {
                                //                            if (contentType.SetAsDefault)
                                //                                ctAsDefault = contentType.Name;

                                //                            newList.AddContentTypeToListByName(contentType.Name, contentType.SetAsDefault, contentType.SearchContentTypeInSiteHierarchy);
                                //                        },
                                //                        string.Format("Asociando contentType: [Nombre: {0}]", contentType.Name));
                                //                    }
                                //                }

                                //            #endregion

                                //            #region Eliminación de Content Types

                                //            if (list.ContentTypes != null && list.ContentTypes.Provision)
                                //                foreach (TenantSiteWebsWebListsListContentTypesRemoveContentType removeContentType in list.ContentTypes.RemoveContentType)
                                //                {
                                //                    //Si existe lo eliminamos
                                //                    if (newList.ContentTypeExistsByName(removeContentType.Name))
                                //                    {
                                //                        Apoyo.ExecuteWithTryCatch(() =>
                                //                        {
                                //                            //Asociamos el Content-Type a la lista
                                //                            newList.RemoveContentTypeByName(removeContentType.Name);
                                //                            clientContext.ExecuteQueryRetry();
                                //                        },
                                //                        string.Format("Eliminando contentType de la lista: [Nombre: {0}]", removeContentType.Name));
                                //                    }
                                //                }

                                //            #endregion

                                //            #region Creación de carpetas

                                //            if (list.EnableFolderCreation != null)
                                //            {
                                //                newList.EnableFolderCreation = list.EnableFolderCreation;
                                //                newList.Update();
                                //                clientContext.ExecuteQueryRetry();
                                //            }

                                //            if (list.Folders != null && list.Folders.Provision)
                                //                foreach (TenantSiteWebsWebListsListFoldersFolder folder in list.Folders.Folder)
                                //                {
                                //                    //if (!newList.RootFolder.FolderExists(folder.Name))
                                //                    //{
                                //                    Apoyo.ExecuteWithTryCatch(() =>
                                //                    {
                                //                        Folder carpeta = null;

                                //                        //Comprobamos si la estructura de carpetas tiene más de un nivel
                                //                        if (folder.Name.Split('/').Length > 1)
                                //                        {
                                //                            carpeta = newList.RootFolder.CreateFolder(folder.Name.Split('/')[0]);

                                //                            for (int i = 1; i < folder.Name.Split('/').Length; i++)
                                //                            {
                                //                                carpeta = carpeta.CreateFolder(folder.Name.Split('/')[i]);
                                //                            }
                                //                        }
                                //                        else if (string.IsNullOrEmpty(folder.Name))
                                //                        {
                                //                            carpeta = newList.RootFolder;
                                //                            clientContext.Load(carpeta);
                                //                            clientContext.ExecuteQueryRetry();
                                //                        }
                                //                        else
                                //                        {
                                //                            carpeta = newList.RootFolder.CreateFolder(folder.Name);
                                //                        }

                                //                        #region Asociación de Content Types a Carpeta

                                //                        if (folder.AssociatedContentTypes != null && folder.AssociatedContentTypes.Provision)
                                //                            if (folder.AssociatedContentTypes.AsignedContentType != null)
                                //                            {
                                //                                List<ContentType> AssociatedContentTypes = new List<ContentType>();

                                //                                #region Obtención tipos de contenido a partir de lista

                                //                                foreach (TenantSiteWebsWebListsListFoldersFolderAssociatedContentTypesAsignedContentType AssociatedContentType
                                //                                    in folder.AssociatedContentTypes.AsignedContentType)
                                //                                {
                                //                                    //Obtengo las copias de los tipos de contenido de la lista a partir de su nombre 
                                //                                    Apoyo.ExecuteWithTryCatch(() =>
                                //                                    {
                                //                                        AssociatedContentTypes.Add(newList.GetContentTypeByName(AssociatedContentType.Name));
                                //                                    });
                                //                                }

                                //                                #endregion

                                //                                #region Ordenación y asociación de los tipos de contenido a la carpeta

                                //                                //Ordenación ascendente o descendente
                                //                                AssociatedContentTypes = (folder.AssociatedContentTypes.OrderBy != null && folder.AssociatedContentTypes.OrderBy.Equals("Descending")) ?
                                //                                                            AssociatedContentTypes.OrderByDescending(x => x.Name).ToList() :
                                //                                                            AssociatedContentTypes.OrderBy(x => x.Name).ToList();

                                //                                #region Si se ha configurado un ct por defecto, lo metemos en primer lugar para que lo establezca como predeterminado. Para los casos de las carpetas raíz como Expendientes en jurídico.

                                //                                if (AssociatedContentTypes.Exists(w => w.Name.Equals(ctAsDefault, StringComparison.InvariantCultureIgnoreCase)))
                                //                                {
                                //                                    ContentType ct = AssociatedContentTypes.Where(w => w.Name.Equals(ctAsDefault, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                                //                                    AssociatedContentTypes.Remove(ct);
                                //                                    AssociatedContentTypes.Insert(0, ct);
                                //                                }

                                //                                #endregion

                                //                                //Asignación de los tipos de contenido a la carpeta
                                //                                Apoyo.ExecuteWithTryCatch(() =>
                                //                                {
                                //                                    carpeta.UniqueContentTypeOrder = AssociatedContentTypes.Select(x => x.Id).ToList();
                                //                                    carpeta.Update();
                                //                                    newList.Update();
                                //                                    clientContext.ExecuteQueryRetry();
                                //                                }, String.Format("Asociando Tipos de contenido - [Carpeta {0} - {1}]", folder.Name, string.Join(", ", AssociatedContentTypes.Select(x => x.Name).ToArray())));

                                //                                #endregion
                                //                            }

                                //                        #endregion

                                //                        #region Asignación de permisos a carpeta

                                //                        if (folder.Permissions != null)
                                //                        {
                                //                            Apoyo.ExecuteWithTryCatch(() =>
                                //                            {
                                //                                //Rompemos herencia si se ha indicado
                                //                                if (folder.Permissions.BreakRoleInheritance)
                                //                                {
                                //                                    carpeta.ListItemAllFields.BreakRoleInheritance(false, true);
                                //                                    clientContext.ExecuteQueryRetry();

                                //                                    #region Eliminamos todas las asignaciones que queden tras romper herencia

                                //                                    clientContext.Load(carpeta.ListItemAllFields.RoleAssignments);
                                //                                    clientContext.ExecuteQueryRetry();

                                //                                    List<RoleAssignment> asignaciones = carpeta.ListItemAllFields.RoleAssignments.ToList();

                                //                                    foreach (RoleAssignment asignacion in asignaciones)
                                //                                    {
                                //                                        asignacion.DeleteObject();
                                //                                        clientContext.ExecuteQueryRetry();
                                //                                    }

                                //                                    #endregion
                                //                                }

                                //                                //Agregamos permisos
                                //                                foreach (TenantSiteWebsWebListsListFoldersFolderPermissionsAddPermission addPermission in folder.Permissions.AddPermission)
                                //                                {
                                //                                    carpeta.ListItemAllFields.AddPermissionLevelToGroup(addPermission.Name, (RoleType)Enum.Parse(typeof(RoleType), addPermission.RoleType, true), addPermission.RemoveExistingPermissionLevels);
                                //                                }
                                //                            },
                                //                            string.Format("Asignando permisos a la carpeta"));
                                //                        }

                                //                        #endregion
                                //                    },
                                //                    string.Format("Creando carpeta: [Nombre: {0}]", folder.Name));
                                //                }
                                //            //}

                                //            #endregion

                                //            #region Creación Vistas

                                //            if (!string.IsNullOrEmpty(list.ViewsSourcePath))
                                //                Apoyo.ExecuteWithTryCatch(() =>
                                //                {
                                //                    currentWeb.CreateViewsFromXMLFile(string.Format("{0}\\{1}\\{2}", site.Url.Replace("/", "\\"), web.Url, list.UrlPath), list.ViewsSourcePath);
                                //                }, string.Format("Creada vista {0}", list.ViewsSourcePath));

                                //            #endregion

                                //            #region Visualización de versiones borrador

                                //            if (list.DraftVisibilityType != null)
                                //            {
                                //                newList.DraftVersionVisibility = (DraftVisibilityType)Enum.Parse(typeof(DraftVisibilityType), list.DraftVisibilityType, true);
                                //                newList.Update();
                                //                clientContext.ExecuteQueryRetry();
                                //            }

                                //            #endregion

                                //            #region Forzar el checkout de los elementos

                                //            if (list.ForceCheckOut)
                                //            {
                                //                newList.ForceCheckout = list.ForceCheckOut;
                                //                newList.Update();
                                //                clientContext.ExecuteQueryRetry();
                                //            }

                                //            #endregion

                                //            #region Límite de versiones principales

                                //            if (list.MajorVersionLimit > 0)
                                //            {
                                //                clientContext.Load(newList, w => w.MajorVersionLimit);
                                //                clientContext.ExecuteQueryRetry();

                                //                newList.MajorVersionLimit = list.MajorVersionLimit;
                                //                newList.Update();
                                //                clientContext.ExecuteQueryRetry();
                                //            }

                                //            #endregion

                                //            #region Límite de versiones secundarias y principales

                                //            if (list.MajorWithMinorVersionsLimit > 0)
                                //            {
                                //                clientContext.Load(newList, w => w.MajorWithMinorVersionsLimit);
                                //                clientContext.ExecuteQueryRetry();

                                //                newList.MajorWithMinorVersionsLimit = list.MajorWithMinorVersionsLimit;
                                //                newList.Update();
                                //                clientContext.ExecuteQueryRetry();
                                //            }

                                //            #endregion
                                //        },
                                //        string.Format("Creando lista: [Nombre: {0}] [Template: {1}]", list.Name, list.TemplateType));
                                //    }

                                #endregion

                                #region Lists

                                if (web.Lists != null && web.Lists.Provision)
                                    foreach (TenantSiteWebsWebListsList list in web.Lists.List)
                                    {
                                        Apoyo.ExecuteWithTryCatch(() =>
                                        {
                                            List newList = null;
                                            string ctAsDefault = string.Empty;

                                            #region Creación / Obtención de la lista

                                            //Creación de la lista
                                            if (!currentWeb.ListExists(list.Name))
                                                newList = currentWeb.CreateList(
                                                                    (ListTemplateType)Enum.Parse(typeof(ListTemplateType), list.TemplateType, true),
                                                                    list.Name,
                                                                    list.EnableVersioning,
                                                                    true,
                                                                    list.UrlPath.Replace("\\", "/"),
                                                                    list.EnableContentTypes);

                                            if (newList == null)
                                                newList = currentWeb.GetListByTitle(list.Name);

                                            clientContext.Load(newList, w => w.ContentTypes);
                                            clientContext.ExecuteQueryRetry();

                                            #endregion

                                            #region Asociación de Content Types

                                            Console.WriteLine();

                                            Apoyo.ExecuteWithTryCatch(() =>
                                            {
                                                List<ContentType> associatedListContentTypes = new List<ContentType>();

                                                if (list.ContentTypes != null && list.ContentTypes.Provision)
                                                {
                                                    clientContext.Load(currentSite.RootWeb.AvailableContentTypes, w => w.Include(a => a.Id, a => a.Name));
                                                    clientContext.ExecuteQueryRetry();

                                                    //Cargamos los tipos de contenido cacheados en una lista para mejorar tiempos
                                                    foreach (TenantSiteWebsWebListsListContentTypesContentType contentType in list.ContentTypes.ContentType)
                                                        associatedListContentTypes.AddRange(currentSite.RootWeb.AvailableContentTypes.Where(w => w.Name == contentType.Name));

                                                    ContentType associatedContentType = null;

                                                    foreach (TenantSiteWebsWebListsListContentTypesContentType contentType in list.ContentTypes.ContentType)
                                                    {
                                                        Apoyo.ExecuteWithTryCatch(() =>
                                                        {
                                                            if (contentType.SetAsDefault)
                                                                ctAsDefault = contentType.Name;

                                                            associatedContentType = associatedListContentTypes.Where(w => w.Name == contentType.Name).FirstOrDefault();

                                                            //newList.AddContentTypeToList(ct, contentType.SetAsDefault);
                                                            associatedContentType = newList.ContentTypes.AddExistingContentType(associatedContentType);
                                                        },
                                                        string.Format("Asociando contentType: [Nombre: {0}]", contentType.Name));

                                                        if (contentType.DocumentTemplatePath != null)
                                                        {
                                                            //Ruta del content type donde subimos la plantilla
                                                            string resourcePath = string.Format("{0}/Forms/{1}", list.UrlPath, contentType.Name);

                                                            //Lectura de la plantilla
                                                            FileInfo file = new FileInfo(contentType.DocumentTemplatePath);

                                                            Apoyo.ExecuteWithTryCatch(() =>
                                                            {
                                                                //Subimos la plantilla al content type
                                                                Folder folder = currentWeb.GetFolderByServerRelativeUrl(resourcePath);
                                                                clientContext.Load(folder);
                                                                clientContext.Load(associatedContentType);
                                                                clientContext.ExecuteQueryRetry();

                                                                Microsoft.SharePoint.Client.File createdFile = folder.UploadFile(file.Name, file.OpenRead(), true);
                                                                clientContext.ExecuteQueryRetry();

                                                                //Actualizamos la ruta de la plantilla en el content type
                                                                associatedContentType.DocumentTemplate = file.Name;
                                                                associatedContentType.Update(false);
                                                                clientContext.ExecuteQueryRetry();
                                                            },
                                                            string.Format("Asociando plantilla al content type: [Plantilla: {0}]", file.Name));
                                                        }
                                                    }
                                                }
                                            },
                                            "Asociación de tipos de contenido");

                                            #endregion

                                            #region Eliminación de Content Types

                                            if (list.ContentTypes != null && list.ContentTypes.Provision)
                                                foreach (TenantSiteWebsWebListsListContentTypesRemoveContentType removeContentType in list.ContentTypes.RemoveContentType)
                                                {
                                                    Apoyo.ExecuteWithTryCatch(() =>
                                                    {
                                                        //Asociamos el Content-Type a la lista
                                                        ContentType contentTypeToRemove = newList.ContentTypes.Where(w => w.Name == removeContentType.Name).FirstOrDefault();
                                                        if (contentTypeToRemove != null)
                                                            contentTypeToRemove.DeleteObject();
                                                    },
                                                    string.Format("Eliminando contentType de la lista: [Nombre: {0}]", removeContentType.Name));
                                                }

                                            #endregion

                                            Apoyo.ExecuteWithTryCatch(() =>
                                            {
                                                clientContext.ExecuteQueryRetry();
                                            }, "ExecuteQuery");

                                            #region Asignación de permisos a la lista

                                            if (list.Permissions != null)
                                            {
                                                Apoyo.ExecuteWithTryCatch(() =>
                                                {
                                                    Console.WriteLine();

                                                    Apoyo.ExecuteWithTryCatch(() =>
                                                    {
                                                        newList.BreakRoleInheritance(list.Permissions.CopyRoleAssignments, list.Permissions.ClearSubscopes);
                                                        clientContext.ExecuteQueryRetry();
                                                    },
                                                    "Rompiendo herencia");

                                                    //Cargamos todos los grupos
                                                    clientContext.Load(currentSite.RootWeb.SiteGroups);
                                                    clientContext.ExecuteQueryRetry();

                                                    #region Agregar permisos

                                                    Apoyo.ExecuteWithTryCatch(() =>
                                                    {
                                                        //Agregamos permisos
                                                        foreach (TenantSiteWebsWebListsListPermissionsAddRoleAssignment addRoleAssignment in list.Permissions.AddRoleAssignment)
                                                        {
                                                            //Obtenemos el grupo
                                                            Principal principal = currentSite.RootWeb.SiteGroups.Where(w => w.Title == addRoleAssignment.Name).FirstOrDefault();

                                                            //Almacenamos el rol del grupo
                                                            RoleDefinition roleDefinition = currentSite.RootWeb.RoleDefinitions.GetByType((RoleType)Enum.Parse(typeof(RoleType), addRoleAssignment.RoleType, true));

                                                            if (addRoleAssignment.RemoveExistingRoleDefinitions)
                                                            {
                                                                newList.RoleAssignments.GetByPrincipal(principal).DeleteObject();
                                                                newList.Context.ExecuteQueryRetry();
                                                            }

                                                            //Creamos una nueva definición de rol
                                                            RoleDefinitionBindingCollection rdc = new RoleDefinitionBindingCollection(clientContext);
                                                            rdc.Add(roleDefinition);
                                                            newList.RoleAssignments.Add(principal, rdc);
                                                        }

                                                        newList.Context.ExecuteQueryRetry();
                                                    },
                                                    "Agregando RoleDefinitions");

                                                    #endregion

                                                    #region Eliminar permisos

                                                    Apoyo.ExecuteWithTryCatch(() =>
                                                    {
                                                        //Agregamos permisos
                                                        foreach (TenantSiteWebsWebListsListPermissionsRemoveRoleAssignment removeRoleAssignment in list.Permissions.RemoveRoleAssignment)
                                                        {
                                                            //Obtenemos el grupo
                                                            Principal principal = currentSite.RootWeb.SiteGroups.Where(w => w.Title == removeRoleAssignment.Name).FirstOrDefault();

                                                            //Cargamos los roleAssignments
                                                            clientContext.Load(newList.RoleAssignments);
                                                            clientContext.ExecuteQueryRetry();

                                                            foreach (RoleAssignment rol in newList.RoleAssignments)
                                                                if (rol.PrincipalId == principal.Id)
                                                                {
                                                                    rol.DeleteObject();
                                                                    newList.Context.ExecuteQueryRetry();
                                                                    break;
                                                                }
                                                        }
                                                    },
                                                    "Eliminando RoleDefinitions");

                                                    #endregion
                                                },
                                                string.Format("Asignando permisos a la lista"));
                                            }

                                            #endregion

                                            #region Creación de carpetas

                                            //Cargamos los contentTypes de la lista para la asignación de tipos de contenido a carpetas
                                            clientContext.Load(newList.ContentTypes, w => w.Include(a => a.Id, a => a.Name));
                                            clientContext.ExecuteQueryRetry();

                                            if (list.Folders != null && list.Folders.Provision)
                                                foreach (TenantSiteWebsWebListsListFoldersFolder folder in list.Folders.Folder)
                                                {
                                                    Console.WriteLine();

                                                    Apoyo.ExecuteWithTryCatch(() =>
                                                    {
                                                        Folder carpeta = null;

                                                        //Comprobamos si la estructura de carpetas tiene más de un nivel
                                                        if (folder.Name.Split('/').Length > 1)
                                                        {
                                                            carpeta = newList.RootFolder.CreateFolder_Hiberus(folder.Name.Split('/')[0]);

                                                            for (int i = 1; i < folder.Name.Split('/').Length; i++)
                                                            {
                                                                carpeta = carpeta.CreateFolder_Hiberus(folder.Name.Split('/')[i]);
                                                            }
                                                        }
                                                        else if (string.IsNullOrEmpty(folder.Name))
                                                        {
                                                            carpeta = newList.RootFolder;
                                                            clientContext.Load(carpeta);
                                                            clientContext.ExecuteQueryRetry();
                                                        }
                                                        else
                                                        {
                                                            carpeta = newList.RootFolder.CreateFolder_Hiberus(folder.Name);
                                                        }

                                                        #region Asociación de Content Types a Carpeta

                                                        if (folder.AssociatedContentTypes != null && folder.AssociatedContentTypes.Provision)
                                                            if (folder.AssociatedContentTypes.AsignedContentType != null)
                                                            {
                                                                List<ContentType> AssociatedContentTypes = new List<ContentType>();

                                                                #region Obtención tipos de contenido a partir de lista

                                                                foreach (TenantSiteWebsWebListsListFoldersFolderAssociatedContentTypesAsignedContentType AssociatedContentType
                                                                    in folder.AssociatedContentTypes.AsignedContentType)
                                                                {
                                                                    //Obtengo las copias de los tipos de contenido de la lista a partir de su nombre 
                                                                    Apoyo.ExecuteWithTryCatch(() =>
                                                                    {
                                                                        AssociatedContentTypes.Add(newList.ContentTypes.Where(w => w.Name == AssociatedContentType.Name).FirstOrDefault());
                                                                    });
                                                                }
                                                                clientContext.ExecuteQueryRetry();

                                                                #endregion

                                                                #region Ordenación y asociación de los tipos de contenido a la carpeta

                                                                //Ordenación ascendente o descendente
                                                                AssociatedContentTypes = (folder.AssociatedContentTypes.OrderBy != null && folder.AssociatedContentTypes.OrderBy.Equals("Descending")) ?
                                                                                            AssociatedContentTypes.OrderByDescending(x => x.Name).ToList() :
                                                                                            AssociatedContentTypes.OrderBy(x => x.Name).ToList();

                                                                #region Si se ha configurado un ct por defecto, lo metemos en primer lugar para que lo establezca como predeterminado. Para los casos de las carpetas raíz como Expendientes en jurídico.

                                                                if (AssociatedContentTypes.Exists(w => w.Name.Equals(ctAsDefault, StringComparison.InvariantCultureIgnoreCase)))
                                                                {
                                                                    ContentType ct = AssociatedContentTypes.Where(w => w.Name.Equals(ctAsDefault, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                                                                    AssociatedContentTypes.Remove(ct);
                                                                    AssociatedContentTypes.Insert(0, ct);
                                                                }

                                                                #endregion

                                                                //Asignación de los tipos de contenido a la carpeta
                                                                Apoyo.ExecuteWithTryCatch(() =>
                                                                {
                                                                    carpeta.UniqueContentTypeOrder = AssociatedContentTypes.Select(x => x.Id).ToList();
                                                                    carpeta.Update();
                                                                    newList.Update();
                                                                    clientContext.ExecuteQueryRetry();
                                                                }, String.Format("Asociando Tipos de contenido - [Carpeta {0} - {1}]", folder.Name, string.Join(", ", AssociatedContentTypes.Select(x => x.Name).ToArray())));

                                                                #endregion
                                                            }

                                                        #endregion

                                                        #region Asignación de permisos a carpeta

                                                        if (folder.Permissions != null)
                                                        {
                                                            Apoyo.ExecuteWithTryCatch(() =>
                                                            {
                                                                Apoyo.ExecuteWithTryCatch(() =>
                                                                {
                                                                    carpeta.ListItemAllFields.BreakRoleInheritance(folder.Permissions.CopyRoleAssignments, folder.Permissions.ClearSubscopes);
                                                                    clientContext.ExecuteQueryRetry();
                                                                },
                                                                "Rompiendo herencia");

                                                                //Cargamos todos los grupos
                                                                clientContext.Load(currentSite.RootWeb.SiteGroups);
                                                                clientContext.ExecuteQueryRetry();

                                                                #region Agregar permisos

                                                                if (folder.Permissions.AddRoleAssignment != null)
                                                                {
                                                                    Apoyo.ExecuteWithTryCatch(() =>
                                                                    {
                                                                        //Agregamos permisos
                                                                        foreach (TenantSiteWebsWebListsListFoldersFolderPermissionsAddRoleAssignment addRoleAssignment in folder.Permissions.AddRoleAssignment)
                                                                        {
                                                                            //Obtenemos el grupo
                                                                            Principal principal = currentSite.RootWeb.SiteGroups.Where(w => w.Title == addRoleAssignment.Name).FirstOrDefault();

                                                                            //Almacenamos el rol del grupo
                                                                            RoleDefinition roleDefinition = currentSite.RootWeb.RoleDefinitions.GetByType((RoleType)Enum.Parse(typeof(RoleType), addRoleAssignment.RoleType, true));

                                                                            if (addRoleAssignment.RemoveExistingRoleDefinitions)
                                                                            {
                                                                                try
                                                                                {
                                                                                    RoleAssignment existingRoleAssignment = carpeta.ListItemAllFields.RoleAssignments.GetByPrincipal(principal);

                                                                                    carpeta.ListItemAllFields.RoleAssignments.GetByPrincipal(principal).DeleteObject();
                                                                                    carpeta.ListItemAllFields.Context.ExecuteQueryRetry();
                                                                                }
                                                                                catch (Exception) { /*No se ha encontrado el grupo a eliminar en los permisos de la carpeta */ }
                                                                            }

                                                                            //Creamos una nueva definición de rol
                                                                            RoleDefinitionBindingCollection rdc = new RoleDefinitionBindingCollection(clientContext);
                                                                            rdc.Add(roleDefinition);
                                                                            carpeta.ListItemAllFields.RoleAssignments.Add(principal, rdc);
                                                                        }

                                                                        carpeta.ListItemAllFields.Context.ExecuteQueryRetry();
                                                                    },
                                                                    "Agregando RoleDefinitions");
                                                                }

                                                                #endregion

                                                                #region Eliminar permisos

                                                                if (folder.Permissions.RemoveRoleAssignment != null)
                                                                {
                                                                    Apoyo.ExecuteWithTryCatch(() =>
                                                                    {
                                                                        //Agregamos permisos
                                                                        foreach (TenantSiteWebsWebListsListFoldersFolderPermissionsRemoveRoleAssignment removeRoleAssignment in folder.Permissions.RemoveRoleAssignment)
                                                                        {
                                                                            //Obtenemos el grupo
                                                                            Principal principal = currentSite.RootWeb.SiteGroups.Where(w => w.Title == removeRoleAssignment.Name).FirstOrDefault();

                                                                            //Cargamos los roleAssignments
                                                                            clientContext.Load(carpeta.ListItemAllFields.RoleAssignments);
                                                                            clientContext.ExecuteQueryRetry();

                                                                            foreach (RoleAssignment rol in carpeta.ListItemAllFields.RoleAssignments)
                                                                                if (rol.PrincipalId == principal.Id)
                                                                                {
                                                                                    rol.DeleteObject();
                                                                                    carpeta.ListItemAllFields.Context.ExecuteQueryRetry();
                                                                                    break;
                                                                                }
                                                                        }
                                                                    },
                                                                    "Eliminando RoleDefinitions");
                                                                }

                                                                #endregion
                                                            },
                                                            string.Format("Asignando permisos a la carpeta"));
                                                        }

                                                        #endregion
                                                    },
                                                    string.Format("Creando carpeta: [Nombre: {0}]", folder.Name));
                                                }

                                            #endregion

                                            #region Creación Vistas

                                            if (!string.IsNullOrEmpty(list.ViewsSourcePath))
                                                Apoyo.ExecuteWithTryCatch(() =>
                                                {
                                                    currentWeb.CreateViewsFromXMLFile(string.Format("{0}\\{1}\\{2}", site.Url.Replace("/", "\\"), web.Url, list.UrlPath), list.ViewsSourcePath);
                                                }, string.Format("Creada vista {0}", list.ViewsSourcePath));

                                            #endregion

                                            #region Configuración: Habilitar la creación de carpetas

                                            if (list.EnableFolderCreation != null)
                                            {
                                                newList.EnableFolderCreation = list.EnableFolderCreation;
                                                newList.Update();
                                            }

                                            #endregion

                                            #region Configuración: Visualización de versiones borrador

                                            if (list.DraftVisibilityType != null)
                                            {
                                                newList.DraftVersionVisibility = (DraftVisibilityType)Enum.Parse(typeof(DraftVisibilityType), list.DraftVisibilityType, true);
                                                newList.Update();
                                            }

                                            #endregion

                                            #region Configuración: Forzar el checkout de los elementos

                                            if (list.ForceCheckOut)
                                            {
                                                newList.ForceCheckout = list.ForceCheckOut;
                                                newList.Update();
                                            }

                                            #endregion

                                            #region Configuración: Habilitar o no versiones secundarias (borradores)

                                            newList.EnableMinorVersions = list.EnableMinorVersions;
                                            newList.Update();

                                            #endregion

                                            #region Configuración: Límite de versiones principales

                                            if (list.MajorVersionLimit > 0)
                                            {
                                                newList.MajorVersionLimit = list.MajorVersionLimit;
                                                newList.Update();
                                            }

                                            #endregion

                                            #region Configuración: Límite de versiones secundarias y principales

                                            if (list.MajorWithMinorVersionsLimit > 0)
                                            {
                                                newList.MajorWithMinorVersionsLimit = list.MajorWithMinorVersionsLimit;
                                                newList.Update();
                                            }

                                            #endregion

                                            //Aplicamos las configuraciones
                                            clientContext.ExecuteQueryRetry();
                                        },
                                        string.Format("Creando lista: [Nombre: {0}] [Template: {1}]", list.Name, list.TemplateType));
                                    }

                                #endregion

                                #region Apps

                                if (web.Apps != null && web.Apps.Provision)
                                {
                                    Guid developmentFeatureId = new Guid("e374875e-06b6-11e0-b0fa-57f5dfd72085");

                                    try
                                    {
                                        //Activamos la característica de desarrollo para poder activar Apps desde CSOM
                                        if (!currentSite.IsFeatureActive(developmentFeatureId))
                                            currentSite.ActivateFeature(developmentFeatureId);

                                        Site siteAppCatalog = o365Tenant.GetSiteByUrl(string.Concat(tenant.Url, web.Apps.AppCatalogUrl));

                                        foreach (TenantSiteWebsWebAppsApp app in web.Apps.App)
                                        {
                                            //Comprobamos si ya existe alguna instancia instalada
                                            ClientObjectList<AppInstance> instances = currentWeb.GetAppInstancesByProductId(new Guid(app.ProductId));
                                            clientContext.Load(instances);
                                            clientContext.ExecuteQueryRetry();

                                            if (instances.Count <= 0)
                                            {
                                                Apoyo.ExecuteWithTryCatch(() =>
                                                {
                                                    Microsoft.SharePoint.Client.File fileApp = siteAppCatalog.RootWeb.GetFileByServerRelativeUrl(
                                                        string.Format("/{0}/{1}/{2}", web.Apps.AppCatalogUrl, web.Apps.AppCatalogListUrl, app.AppFileName));

                                                    clientContext.Load(fileApp);
                                                    clientContext.ExecuteQueryRetry();

                                                    ClientResult<Stream> data = fileApp.OpenBinaryStream();
                                                    clientContext.ExecuteQueryRetry();

                                                    //Instalación de la APP
                                                    AppInstance instance = currentWeb.LoadAndInstallAppInSpecifiedLocale(data.Value, 3082);

                                                    #region Esperamos a que se haya instalado correctamente la APP

                                                    //Obtenemos el ID de instancia
                                                    clientContext.Load(instance, w => w.Id, w => w.Status);
                                                    clientContext.ExecuteQueryRetry();

                                                    Guid instanceID = instance.Id;

                                                    int maxTry = 15;
                                                    int count = 0;
                                                    do
                                                    {
                                                        System.Threading.Thread.Sleep(2000);
                                                        instance = currentWeb.GetAppInstanceById(instanceID);
                                                        clientContext.Load(instance, w => w.Status);
                                                        clientContext.ExecuteQueryRetry();
                                                        count++;
                                                    }
                                                    while (instance != null && instance.Status != AppInstanceStatus.Installed && count < maxTry);

                                                    #endregion
                                                },
                                                string.Format("Instalando App: [ProductID: {0}]", app.ProductId));
                                            }
                                        }
                                    }
                                    finally
                                    {
                                        if (currentSite.IsFeatureActive(developmentFeatureId))
                                            currentSite.DeactivateFeature(developmentFeatureId);
                                    }
                                }

                                #endregion

                                #region Pages

                                if (web.Pages != null && web.Pages.Provision)
                                    foreach (TenantSiteWebsWebPagesPublishingPage page in web.Pages.PublishingPage)
                                    {
                                        Apoyo.ExecuteWithTryCatch(() =>
                                        {
                                            string paginaURL = string.Format("Paginas/{0}.aspx", page.Name);

                                            //Si tiene la opción de borrar, la borramos
                                            if (currentWeb.FileExists(paginaURL) && page.RemoveIfExists)
                                            {
                                                //Obtenemos la URL relativa del sitio
                                                clientContext.Load(currentWeb, w => w.ServerRelativeUrl);
                                                clientContext.ExecuteQueryRetry();

                                                Microsoft.SharePoint.Client.File pagina = currentSite.RootWeb.GetFileByServerRelativeUrl(string.Format("{0}/{1}", currentWeb.ServerRelativeUrl, paginaURL));
                                                pagina.DeleteObject();
                                                currentSite.RootWeb.Update();
                                                clientContext.ExecuteQueryRetry();
                                            }

                                            //Si la página no existe la creamos
                                            if (!currentWeb.FileExists(paginaURL))
                                            {
                                                #region Creación de la página

                                                //Obtenemos la URL relativa del sitio
                                                clientContext.Load(currentSite, w => w.ServerRelativeUrl);
                                                clientContext.ExecuteQueryRetry();

                                                //Obtenemos el PageLayout
                                                Microsoft.SharePoint.Client.File pageFromPageLayout = currentSite.RootWeb.GetFileByServerRelativeUrl(String.Format("{0}_catalogs/masterpage/{1}.aspx",
                                                UrlUtility.EnsureTrailingSlash(currentSite.ServerRelativeUrl),
                                                page.Layout));

                                                ListItem pageLayoutItem = pageFromPageLayout.ListItemAllFields;
                                                clientContext.Load(pageLayoutItem);
                                                clientContext.ExecuteQueryRetry();

                                                //Obtenemos el objeto web de publicación
                                                PublishingWeb publishingWeb = PublishingWeb.GetPublishingWeb(clientContext, currentWeb);
                                                clientContext.Load(publishingWeb);

                                                PublishingPage mPage = publishingWeb.AddPublishingPage(new PublishingPageInformation
                                                {
                                                    Name = string.Format("{0}.aspx", page.Name),
                                                    PageLayoutListItem = pageLayoutItem
                                                });

                                                #endregion

                                                #region Asociación de webparts

                                                //Asociamos los webparts
                                                foreach (TenantSiteWebsWebPagesPublishingPageWebPart webpart in page.WebParts)
                                                {
                                                    Apoyo.ExecuteWithTryCatch(() =>
                                                    {
                                                        //Obtenemos el XML del webpart
                                                        string webPartXml = System.IO.File.ReadAllText(string.Format("{0}\\{1}", webpart.SourcePath, webpart.Name));

                                                        webPartXml = webPartXml.Contains(Apoyo.ResolveSiteCollection) ? webPartXml.Replace(Apoyo.ResolveSiteCollection, currentSite.RootWeb.ServerRelativeUrl) : webPartXml;

                                                        //Creamos la entidad de webpart
                                                        OfficeDevPnP.Core.Entities.WebPartEntity webPartEntity = new OfficeDevPnP.Core.Entities.WebPartEntity()
                                                        {
                                                            WebPartTitle = webpart.Title,
                                                            WebPartXml = webPartXml,
                                                            WebPartZone = webpart.Zone,
                                                            WebPartIndex = webpart.Index
                                                        };

                                                        //Agregamos el webpart
                                                        currentWeb.AddWebPartToWebPartPage(webPartEntity, paginaURL);
                                                    },
                                                    string.Format("Asociando webpart: [Nombre: {0}]", webpart.Name));
                                                }

                                                #endregion

                                                #region Actualización del título de página

                                                //Cargamos la página para establecer el título y hacer el check-in
                                                clientContext.Load(mPage.ListItem, w => w.ParentList);
                                                clientContext.ExecuteQueryRetry();

                                                //Actualización del título de la página
                                                ListItem pageItem = mPage.ListItem;
                                                pageItem["Title"] = page.Title;
                                                pageItem.Update();

                                                #endregion

                                                #region Check-in y Publish

                                                clientContext.Load(pageItem, p => p.File.CheckOutType);
                                                clientContext.ExecuteQueryRetry();

                                                //Check-in
                                                if (pageItem.File.CheckOutType != CheckOutType.None)
                                                    pageItem.File.CheckIn(String.Empty, CheckinType.MajorCheckIn);

                                                //Publicación
                                                pageItem.File.Publish(String.Empty);
                                                if (pageItem.ParentList.EnableModeration)
                                                    pageItem.File.Approve(String.Empty);

                                                clientContext.ExecuteQueryRetry();

                                                #endregion

                                            }
                                        },
                                        string.Format("Creando la página: [Nombre: {0}] [Layout: {1}]", page.Name, page.Layout));
                                    }

                                #endregion

                                #region Navigation

                                if (web.Navigation != null && web.Navigation.Provision)
                                {
                                    #region Global

                                    if (web.Navigation.GlobalNavigation != null)
                                    {
                                        Apoyo.ExecuteWithTryCatch(() =>
                                        {
                                            StandardNavigationSource navigationType = (StandardNavigationSource)Enum.Parse(typeof(StandardNavigationSource), web.Navigation.GlobalNavigation.Type);

                                            switch (navigationType)
                                            {
                                                case StandardNavigationSource.InheritFromParentWeb:

                                                    //Obtenemos la sesión del conjunto de términos de taxonomia y actualizamos la caché
                                                    TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(clientContext);

                                                    //Obtenemos la configuración de navegación de la web y la reestablecemos
                                                    WebNavigationSettings navigationSettings = new WebNavigationSettings(clientContext, currentWeb);//TaxonomyNavigation.GetWebNavigationSettings(clientContext, web);
                                                    clientContext.Load(navigationSettings, w => w.GlobalNavigation);
                                                    clientContext.ExecuteQuery();

                                                    navigationSettings.GlobalNavigation.Source = StandardNavigationSource.InheritFromParentWeb;
                                                    navigationSettings.Update(taxonomySession);
                                                    clientContext.ExecuteQuery();

                                                    break;
                                            }
                                        },
                                        string.Format("Configurando la navegación global"));
                                    }

                                    #endregion

                                    #region Current

                                    if (web.Navigation.CurrentNavigation != null)
                                    {
                                        Apoyo.ExecuteWithTryCatch(() =>
                                        {
                                            StandardNavigationSource navigationType = (StandardNavigationSource)Enum.Parse(typeof(StandardNavigationSource), web.Navigation.CurrentNavigation.Type);

                                            switch (navigationType)
                                            {
                                                case StandardNavigationSource.TaxonomyProvider:

                                                    //Obtenemos la sesión del conjunto de términos de taxonomia y actualizamos la caché
                                                    TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(clientContext);

                                                    //Obtenemos el conjunto de términos de la navegación del XML
                                                    string navigationTermGroupName = web.Navigation.CurrentNavigation.NavigationTermGroupName;
                                                    string navigationTermSetName = web.Navigation.CurrentNavigation.NavigationTermSetName;

                                                    //Obtenemos la configuración de navegación de la web y la reestablecemos
                                                    WebNavigationSettings navigationSettings = new WebNavigationSettings(clientContext, currentWeb);//TaxonomyNavigation.GetWebNavigationSettings(clientContext, web);
                                                    clientContext.Load(navigationSettings, w => w.CurrentNavigation);
                                                    clientContext.ExecuteQuery();

                                                    //Carga de los metadatos
                                                    TermGroup group = clientContext.Site.GetTermGroupByName(navigationTermGroupName);

                                                    clientContext.Load(group.TermStore);
                                                    clientContext.Load(group.TermSets, w => w.Include(t => t.Id, t => t.Name));
                                                    clientContext.ExecuteQuery();

                                                    TermSet termSet = group.TermSets.GetByName(navigationTermSetName);

                                                    clientContext.Load(termSet, w => w.Id);
                                                    clientContext.ExecuteQuery();

                                                    //Actualización de la navegación
                                                    navigationSettings.CurrentNavigation.Source = StandardNavigationSource.TaxonomyProvider;
                                                    navigationSettings.CurrentNavigation.TermStoreId = group.TermStore.Id;
                                                    navigationSettings.CurrentNavigation.TermSetId = termSet.Id;
                                                    navigationSettings.Update(taxonomySession);
                                                    clientContext.ExecuteQuery();

                                                    break;

                                                case StandardNavigationSource.PortalProvider:   //Estructurada

                                                    //Obtenemos la sesión del conjunto de términos de taxonomia y actualizamos la caché
                                                    TaxonomySession taxonomySessionEstructured = TaxonomySession.GetTaxonomySession(clientContext);

                                                    //Obtenemos la configuración de navegación de la web y la reestablecemos
                                                    WebNavigationSettings navigationSettingsEstructurada = new WebNavigationSettings(clientContext, currentWeb);
                                                    clientContext.Load(navigationSettingsEstructurada, w => w.CurrentNavigation);
                                                    clientContext.ExecuteQuery();

                                                    //Actualización de la navegación a navegación estructurada
                                                    navigationSettingsEstructurada.CurrentNavigation.Source = StandardNavigationSource.PortalProvider;
                                                    navigationSettingsEstructurada.Update(taxonomySessionEstructured);
                                                    clientContext.ExecuteQuery();

                                                    //Quitamos la opción de mostrar las páginas y los subsitios de la navegación
                                                    //http://social.technet.microsoft.com/wiki/contents/articles/23512.sharepoint-2010-configure-navigation-settings-in-a-sandbox-solution.aspx
                                                    currentWeb.AllProperties["__CurrentNavigationIncludeTypes"] = "0";
                                                    currentWeb.Update();
                                                    clientContext.ExecuteQueryRetry();

                                                    #region Nodos de la navegacion estructurada

                                                    if (web.Navigation.CurrentNavigation.FirstLevelNode != null && web.Navigation.CurrentNavigation.FirstLevelNode.Count() > 0)
                                                    {
                                                        //Cargo los nodos de la navegacion estructurada
                                                        clientContext.Load(currentWeb, w => w.Navigation.QuickLaunch);
                                                        clientContext.ExecuteQuery();

                                                        //Borro la navegacion anterior
                                                        currentWeb.Navigation.QuickLaunch.ToList().ForEach(x => { x.DeleteObject(); });
                                                        clientContext.ExecuteQuery();

                                                        #region Nodo navegacion estructurada: Primer nivel

                                                        //Recorro los nodos del primer nivel para añadirlos a la navegacion
                                                        foreach (TenantSiteWebsWebNavigationCurrentNavigationFirstLevelNode nodoPrimerNivel in web.Navigation.CurrentNavigation.FirstLevelNode)
                                                        {
                                                            Apoyo.ExecuteWithTryCatch(() =>
                                                            {
                                                                //Para que el nodo se comporte como cabecera, en el XML debe configurarse el nodo como "IsExternal = true" y "Url= ''"

                                                                //Añado el nodo del primer nivel y lo recupero por si es necesario añadir un segundo nivel
                                                                NavigationNode nodoPadre = currentWeb.Navigation.QuickLaunch.Add(new NavigationNodeCreationInformation()
                                                                {
                                                                    Title = nodoPrimerNivel.Title,
                                                                    Url = string.IsNullOrEmpty(nodoPrimerNivel.Url) ? null : nodoPrimerNivel.Url.Contains(Apoyo.ResolveSiteCollection) ? nodoPrimerNivel.Url.Replace(Apoyo.ResolveSiteCollection, currentSite.RootWeb.ServerRelativeUrl) : nodoPrimerNivel.Url,
                                                                    AsLastNode = nodoPrimerNivel.AsLastNode,
                                                                    IsExternal = nodoPrimerNivel.IsExternal,
                                                                });

                                                                #region Nodo navegacion estructurada: Segundo nivel

                                                                if (nodoPrimerNivel.SecondLevelNode != null && nodoPrimerNivel.SecondLevelNode.Count() > 0)
                                                                    //Recorro los nodos del segundo nivel
                                                                    foreach (TenantSiteWebsWebNavigationCurrentNavigationFirstLevelNodeSecondLevelNode nodoSegundoNivel in nodoPrimerNivel.SecondLevelNode)
                                                                    {
                                                                        //Añado el nodo de segundo nivel al padre
                                                                        nodoPadre.Children.Add(new NavigationNodeCreationInformation()
                                                                        {
                                                                            Title = nodoSegundoNivel.Title,
                                                                            Url = nodoSegundoNivel.Url.Contains(Apoyo.ResolveSiteCollection) ? nodoSegundoNivel.Url.Replace(Apoyo.ResolveSiteCollection, currentSite.RootWeb.ServerRelativeUrl) : nodoSegundoNivel.Url,
                                                                            AsLastNode = nodoSegundoNivel.AsLastNode,
                                                                            IsExternal = nodoSegundoNivel.IsExternal
                                                                        });
                                                                    }

                                                                #endregion

                                                            }, string.Format("\nAñadido nodo {0}", nodoPrimerNivel.Title));
                                                        }

                                                        #endregion

                                                        //Actualizo los cambios
                                                        clientContext.ExecuteQueryRetry();
                                                    }

                                                    #endregion

                                                    break;
                                            }
                                        },
                                        string.Format("Configurando la navegación actual"));
                                    }

                                    #endregion
                                }

                                #endregion

                                #region Search

                                if (web.Search != null && web.Search.Provision)
                                {
                                    //Página principal de búsqueda
                                    Apoyo.ExecuteWithTryCatch(() =>
                                    {
                                        //Ruta por defecto del sitio web
                                        String searchDefaultPath = String.Format("{{\"Inherit\":{0},\"ResultsPageAddress\":\"~site/{1}.aspx\",\"ShowNavigation\":{2}}}",
                                                                                                            web.Search.Inherit.ToString().ToLower(),
                                                                                                            web.Search.DefaultResultsPage,
                                                                                                            web.Search.ShowNavigation.ToString().ToLower());
                                        currentWeb.SetPropertyBagValue("SRCH_SB_SET_WEB", searchDefaultPath);

                                    },
                                    String.Format("Cambiando ruta de búsqueda predeterminada: [Nueva pagina: {0}] ]", web.Search.DefaultResultsPage));

                                    //Nodos de navegación de búsqueda
                                    if (web.Search.SearchNodes != null && web.Search.SearchNodes.Provision)
                                    {
                                        currentWeb.DeleteAllNavigationNodes(OfficeDevPnP.Core.Enums.NavigationType.SearchNav);

                                        foreach (TenantSiteWebsWebSearchSearchNodesSearchNode searchNode in web.Search.SearchNodes.SearchNode)
                                            Apoyo.ExecuteWithTryCatch(() =>
                                            {
                                                //Ruta final del nodo de navegación
                                                String targetPath = String.Format("{0}{1}{2}", tenant.Url,
                                                                                    UrlUtility.EnsureTrailingSlash(currentWeb.ServerRelativeUrl.Remove(0, 1).ToString()),
                                                                                    searchNode.TargetPath);

                                                //Añado el nodo a la busqueda del sitio
                                                currentWeb.AddNavigationNode(searchNode.Title,
                                                                            new Uri(targetPath),
                                                                            searchNode.ParentNodeTitle,
                                                                            OfficeDevPnP.Core.Enums.NavigationType.SearchNav);
                                            },
                                            String.Format("Añadido nuevo nodo de búsqueda: [Nuevo nodo: {0}] ]", searchNode.Title));
                                    }
                                }

                                #endregion

                                #region Welcome Page

                                //Aplicamos la página de inicio en caso necesario
                                if (!string.IsNullOrEmpty(web.WelcomePage))
                                {
                                    Apoyo.ExecuteWithTryCatch(() =>
                                    {
                                        currentWeb.SetHomePage(web.WelcomePage);
                                    },
                                    string.Format("Cambiada la pagina principal a {0}", web.WelcomePage));
                                }

                                #endregion
                            }

                        #endregion

                        #region RootWeb Navigation

                        if ((site.Webs.RootWeb != null && site.Webs.RootWeb.Provision) && (site.Webs.RootWeb.Navigation != null && site.Webs.RootWeb.Navigation.Provision))
                        {
                            #region Global

                            if (site.Webs.RootWeb.Navigation.GlobalNavigation != null)
                            {
                                Apoyo.ExecuteWithTryCatch(() =>
                                {
                                    StandardNavigationSource navigationType = (StandardNavigationSource)Enum.Parse(typeof(StandardNavigationSource), site.Webs.RootWeb.Navigation.GlobalNavigation.Type);

                                    switch (navigationType)
                                    {
                                        case StandardNavigationSource.InheritFromParentWeb:

                                            //Obtenemos la sesión del conjunto de términos de taxonomia y actualizamos la caché
                                            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(clientContext);

                                            //Obtenemos la configuración de navegación de la web y la reestablecemos
                                            WebNavigationSettings navigationSettings = new WebNavigationSettings(clientContext, currentSite.RootWeb);//TaxonomyNavigation.GetWebNavigationSettings(clientContext, web);
                                            clientContext.Load(navigationSettings, w => w.GlobalNavigation);
                                            clientContext.ExecuteQuery();

                                            navigationSettings.GlobalNavigation.Source = StandardNavigationSource.InheritFromParentWeb;
                                            navigationSettings.Update(taxonomySession);
                                            clientContext.ExecuteQuery();

                                            break;

                                        case StandardNavigationSource.TaxonomyProvider:

                                            //Obtenemos la sesión del conjunto de términos de taxonomia y actualizamos la caché
                                            TaxonomySession tSession = TaxonomySession.GetTaxonomySession(clientContext);

                                            //Obtenemos el conjunto de términos de la navegación del XML
                                            string navigationTermGroupName = site.Webs.RootWeb.Navigation.GlobalNavigation.NavigationTermGroupName;
                                            string navigationTermSetName = site.Webs.RootWeb.Navigation.GlobalNavigation.NavigationTermSetName;

                                            //Obtenemos la configuración de navegación de la web y la reestablecemos
                                            WebNavigationSettings navSettings = new WebNavigationSettings(clientContext, currentSite.RootWeb);
                                            clientContext.Load(navSettings, w => w.GlobalNavigation);
                                            clientContext.ExecuteQuery();

                                            //Carga de los metadatos
                                            TermGroup group = clientContext.Site.GetTermGroupByName(navigationTermGroupName);

                                            clientContext.Load(group.TermStore);
                                            clientContext.Load(group.TermSets, w => w.Include(t => t.Id, t => t.Name));
                                            clientContext.ExecuteQuery();

                                            TermSet termSet = group.TermSets.GetByName(navigationTermSetName);

                                            clientContext.Load(tSession);
                                            clientContext.Load(termSet, w => w.Id);
                                            clientContext.ExecuteQuery();

                                            //Actualización de la navegación
                                            navSettings.GlobalNavigation.Source = StandardNavigationSource.TaxonomyProvider;
                                            navSettings.GlobalNavigation.TermStoreId = group.TermStore.Id;
                                            navSettings.GlobalNavigation.TermSetId = termSet.Id;
                                            navSettings.Update(tSession);
                                            clientContext.ExecuteQuery();

                                            break;

                                        case StandardNavigationSource.PortalProvider:   //Estructurada

                                            //Obtenemos la sesión del conjunto de términos de taxonomia y actualizamos la caché
                                            TaxonomySession taxonomySessionEstructured = TaxonomySession.GetTaxonomySession(clientContext);

                                            //Obtenemos la configuración de navegación de la web y la reestablecemos
                                            WebNavigationSettings navigationSettingsEstructurada = new WebNavigationSettings(clientContext, currentSite.RootWeb);
                                            clientContext.Load(navigationSettingsEstructurada, w => w.CurrentNavigation);
                                            clientContext.ExecuteQuery();

                                            //Actualización de la navegación a navegación estructurada
                                            navigationSettingsEstructurada.GlobalNavigation.Source = StandardNavigationSource.PortalProvider;
                                            navigationSettingsEstructurada.Update(taxonomySessionEstructured);
                                            clientContext.ExecuteQuery();

                                            //Quitamos la opción de mostrar las páginas y los subsitios de la navegación
                                            //http://social.technet.microsoft.com/wiki/contents/articles/23512.sharepoint-2010-configure-navigation-settings-in-a-sandbox-solution.aspx
                                            currentSite.RootWeb.AllProperties["__GlobalNavigationIncludeTypes"] = "0";
                                            currentSite.RootWeb.Update();
                                            clientContext.ExecuteQueryRetry();

                                            #region Eliminación de los nodos anteriores

                                            //Cargo los nodos de la navegacion estructurada
                                            clientContext.Load(currentSite.RootWeb, w => w.Navigation.TopNavigationBar);
                                            clientContext.ExecuteQuery();

                                            //Borro la navegacion anterior
                                            currentSite.RootWeb.Navigation.TopNavigationBar.ToList().ForEach(x => { x.DeleteObject(); });
                                            clientContext.ExecuteQuery();

                                            #endregion

                                            #region Nodos de la navegacion estructurada

                                            if (site.Webs.RootWeb.Navigation.GlobalNavigation.FirstLevelNode != null && site.Webs.RootWeb.Navigation.GlobalNavigation.FirstLevelNode.Count() > 0)
                                            {
                                                #region Nodo navegacion estructurada: Primer nivel

                                                //Recorro los nodos del primer nivel para añadirlos a la navegacion
                                                foreach (TenantSiteWebsRootWebNavigationGlobalNavigationFirstLevelNode nodoPrimerNivel in site.Webs.RootWeb.Navigation.GlobalNavigation.FirstLevelNode)
                                                {
                                                    Apoyo.ExecuteWithTryCatch(() =>
                                                    {
                                                        //Para que el nodo se comporte como cabecera, en el XML debe configurarse el nodo como "IsExternal = true" y "Url= ''"

                                                        //Añado el nodo del primer nivel y lo recupero por si es necesario añadir un segundo nivel
                                                        NavigationNode nodoPadre = currentSite.RootWeb.Navigation.TopNavigationBar.Add(new NavigationNodeCreationInformation()
                                                        {
                                                            Title = nodoPrimerNivel.Title,
                                                            Url = string.IsNullOrEmpty(nodoPrimerNivel.Url) ? null : nodoPrimerNivel.Url.Contains(Apoyo.ResolveSiteCollection) ? nodoPrimerNivel.Url.Replace(Apoyo.ResolveSiteCollection, currentSite.RootWeb.ServerRelativeUrl) : nodoPrimerNivel.Url,
                                                            AsLastNode = nodoPrimerNivel.AsLastNode,
                                                            IsExternal = nodoPrimerNivel.IsExternal,
                                                        });

                                                        #region Nodo navegacion estructurada: Segundo nivel

                                                        if (nodoPrimerNivel.SecondLevelNode != null && nodoPrimerNivel.SecondLevelNode.Count() > 0)
                                                            //Recorro los nodos del segundo nivel
                                                            foreach (TenantSiteWebsRootWebNavigationGlobalNavigationFirstLevelNodeSecondLevelNode nodoSegundoNivel in nodoPrimerNivel.SecondLevelNode)
                                                            {

                                                                //Añado el nodo de segundo nivel al padre
                                                                nodoPadre.Children.Add(new NavigationNodeCreationInformation()
                                                                {
                                                                    Title = nodoSegundoNivel.Title,
                                                                    Url = nodoSegundoNivel.Url.Contains(Apoyo.ResolveSiteCollection) ? nodoSegundoNivel.Url.Replace(Apoyo.ResolveSiteCollection, currentSite.RootWeb.ServerRelativeUrl) : nodoSegundoNivel.Url,
                                                                    AsLastNode = nodoSegundoNivel.AsLastNode,
                                                                    IsExternal = nodoSegundoNivel.IsExternal
                                                                });
                                                            }

                                                        #endregion

                                                    }, string.Format("\nAñadido nodo {0}", nodoPrimerNivel.Title));
                                                }

                                                #endregion

                                                //Actualizo los cambios
                                                clientContext.ExecuteQueryRetry();
                                            }

                                            #endregion

                                            break;
                                    }
                                },
                                string.Format("Configurando la navegación global"));
                            }

                            #endregion

                            #region Current

                            if (site.Webs.RootWeb.Navigation.CurrentNavigation != null)
                            {
                                Apoyo.ExecuteWithTryCatch(() =>
                                {
                                    StandardNavigationSource navigationType = (StandardNavigationSource)Enum.Parse(typeof(StandardNavigationSource), site.Webs.RootWeb.Navigation.GlobalNavigation.Type);

                                    switch (navigationType)
                                    {
                                        case StandardNavigationSource.TaxonomyProvider:

                                            //Obtenemos la sesión del conjunto de términos de taxonomia y actualizamos la caché
                                            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(clientContext);

                                            //Obtenemos el conjunto de términos de la navegación del XML
                                            string navigationTermGroupName = site.Webs.RootWeb.Navigation.CurrentNavigation.NavigationTermGroupName;
                                            string navigationTermSetName = site.Webs.RootWeb.Navigation.CurrentNavigation.NavigationTermSetName;

                                            //Obtenemos la configuración de navegación de la web y la reestablecemos
                                            WebNavigationSettings navigationSettings = new WebNavigationSettings(clientContext, currentSite.RootWeb);
                                            clientContext.Load(navigationSettings, w => w.CurrentNavigation);
                                            clientContext.ExecuteQuery();

                                            //Carga de los metadatos
                                            TermGroup group = clientContext.Site.GetTermGroupByName(navigationTermGroupName);

                                            clientContext.Load(group.TermStore);
                                            clientContext.Load(group.TermSets, w => w.Include(t => t.Id, t => t.Name));
                                            clientContext.ExecuteQuery();

                                            TermSet termSet = group.TermSets.GetByName(navigationTermSetName);

                                            clientContext.Load(taxonomySession);
                                            clientContext.Load(termSet, w => w.Id);
                                            clientContext.ExecuteQuery();

                                            //Actualización de la navegación
                                            navigationSettings.CurrentNavigation.Source = StandardNavigationSource.TaxonomyProvider;
                                            navigationSettings.CurrentNavigation.TermStoreId = group.TermStore.Id;
                                            navigationSettings.CurrentNavigation.TermSetId = termSet.Id;
                                            navigationSettings.Update(taxonomySession);
                                            clientContext.ExecuteQuery();

                                            break;

                                        case StandardNavigationSource.PortalProvider:   //Estructurada

                                            //Obtenemos la sesión del conjunto de términos de taxonomia y actualizamos la caché
                                            TaxonomySession taxonomySessionEstructured = TaxonomySession.GetTaxonomySession(clientContext);

                                            //Obtenemos la configuración de navegación de la web y la reestablecemos
                                            WebNavigationSettings navigationSettingsEstructurada = new WebNavigationSettings(clientContext, currentSite.RootWeb);
                                            clientContext.Load(navigationSettingsEstructurada, w => w.CurrentNavigation);
                                            clientContext.ExecuteQuery();

                                            //Actualización de la navegación a navegación estructurada
                                            navigationSettingsEstructurada.CurrentNavigation.Source = StandardNavigationSource.PortalProvider;
                                            navigationSettingsEstructurada.Update(taxonomySessionEstructured);
                                            clientContext.ExecuteQuery();

                                            //Quitamos la opción de mostrar las páginas y los subsitios de la navegación
                                            //http://social.technet.microsoft.com/wiki/contents/articles/23512.sharepoint-2010-configure-navigation-settings-in-a-sandbox-solution.aspx
                                            currentSite.RootWeb.AllProperties["__CurrentNavigationIncludeTypes"] = "0";
                                            currentSite.RootWeb.Update();
                                            clientContext.ExecuteQueryRetry();

                                            #region Eliminación de los nodos anteriores

                                            //Cargo los nodos de la navegacion estructurada
                                            clientContext.Load(currentSite.RootWeb, w => w.Navigation.QuickLaunch);
                                            clientContext.ExecuteQuery();

                                            //Borro la navegacion anterior
                                            currentSite.RootWeb.Navigation.QuickLaunch.ToList().ForEach(x => { x.DeleteObject(); });
                                            clientContext.ExecuteQuery();

                                            #endregion

                                            #region Nodos de la navegacion estructurada

                                            if (site.Webs.RootWeb.Navigation.CurrentNavigation.FirstLevelNode != null && site.Webs.RootWeb.Navigation.CurrentNavigation.FirstLevelNode.Count() > 0)
                                            {
                                                #region Nodo navegacion estructurada: Primer nivel

                                                //Recorro los nodos del primer nivel para añadirlos a la navegacion
                                                foreach (TenantSiteWebsRootWebNavigationCurrentNavigationFirstLevelNode nodoPrimerNivel in site.Webs.RootWeb.Navigation.CurrentNavigation.FirstLevelNode)
                                                {
                                                    Apoyo.ExecuteWithTryCatch(() =>
                                                    {
                                                        //Para que el nodo se comporte como cabecera, en el XML debe configurarse el nodo como "IsExternal = true" y "Url= ''"

                                                        //Añado el nodo del primer nivel y lo recupero por si es necesario añadir un segundo nivel
                                                        NavigationNode nodoPadre = currentSite.RootWeb.Navigation.QuickLaunch.Add(new NavigationNodeCreationInformation()
                                                        {
                                                            Title = nodoPrimerNivel.Title,
                                                            Url = string.IsNullOrEmpty(nodoPrimerNivel.Url) ? null : nodoPrimerNivel.Url.Contains(Apoyo.ResolveSiteCollection) ? nodoPrimerNivel.Url.Replace(Apoyo.ResolveSiteCollection, currentSite.RootWeb.ServerRelativeUrl) : nodoPrimerNivel.Url,
                                                            AsLastNode = nodoPrimerNivel.AsLastNode,
                                                            IsExternal = nodoPrimerNivel.IsExternal,
                                                        });

                                                        #region Nodo navegacion estructurada: Segundo nivel

                                                        if (nodoPrimerNivel.SecondLevelNode != null && nodoPrimerNivel.SecondLevelNode.Count() > 0)
                                                            //Recorro los nodos del segundo nivel
                                                            foreach (TenantSiteWebsRootWebNavigationCurrentNavigationFirstLevelNodeSecondLevelNode nodoSegundoNivel in nodoPrimerNivel.SecondLevelNode)
                                                            {
                                                                //Añado el nodo de segundo nivel al padre
                                                                nodoPadre.Children.Add(new NavigationNodeCreationInformation()
                                                                {
                                                                    Title = nodoSegundoNivel.Title,
                                                                    Url = nodoSegundoNivel.Url.Contains(Apoyo.ResolveSiteCollection) ? nodoSegundoNivel.Url.Replace(Apoyo.ResolveSiteCollection, currentSite.RootWeb.ServerRelativeUrl) : nodoSegundoNivel.Url,
                                                                    AsLastNode = nodoSegundoNivel.AsLastNode,
                                                                    IsExternal = nodoSegundoNivel.IsExternal
                                                                });
                                                            }

                                                        #endregion

                                                    }, string.Format("\nAñadido nodo {0}", nodoPrimerNivel.Title));
                                                }

                                                #endregion

                                                //Actualizo los cambios
                                                clientContext.ExecuteQueryRetry();
                                            }

                                            #endregion

                                            break;
                                    }
                                },
                                string.Format("Configurando la navegación actual"));
                            }

                            #endregion
                        }

                        #endregion

                        #region RootWeb Welcome Page

                        //Aplicamos la página de inicio en caso necesario
                        if (site.Webs.RootWeb != null && site.Webs.RootWeb.Provision && !string.IsNullOrEmpty(site.Webs.RootWeb.WelcomePage))
                        {
                            Apoyo.ExecuteWithTryCatch(() =>
                            {
                                currentSite.RootWeb.SetHomePage(site.Webs.RootWeb.WelcomePage);
                            },
                            string.Format("Cambiada la pagina principal a {0}", site.Webs.RootWeb.WelcomePage));
                        }

                        #endregion

                        #region List Items - Aprovisionamiento

                        if (site.ProvisionListItems)
                            if (!string.IsNullOrEmpty(site.ListItemsSourcePath))
                                Apoyo.ExecuteWithTryCatch(() =>
                                {
                                    currentSite.PopulateListItemsFromXMLFile(site.ListItemsSourcePath);
                                }, String.Format("Provisionando listas\n"));


                        #endregion

                        Console.WriteLine("Finalizado: " + DateTime.Now.ToString("yyyy-MM-dd_hh-mm-ss"));
                    }
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(String.Format("\n[Error] - {0}", (ex.InnerException != null) ? ex.InnerException.Message.ToString() : ex.Message.ToString()));
                Console.ResetColor();
            }
            finally
            {
                Console.WriteLine("\n--Fin aprovisionamiento--");
                //Console.ReadLine();
            }            
        }
    }
}