using System;
using System.Xml.Serialization;
using Hiberus.Actualizaciones.Schema;
using System.Security;
using Microsoft.SharePoint.Client;
using System.IO;
using AprovisionamientoO365.Base;
using System.Xml;
using AprovisionamientoO365.Base.Extensions.Fields;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace ActualizacionesO365.Base
{
    /// <summary>
    /// Actualiza la solución en base al esquema 'settings_schema_Actualizacion.xsd'
    /// </summary>
    public class ActualizacionBase
    {
        #region Constantes

        //Atributos Field
        internal const string ATR_ID = "ID";
        internal const string ATR_Name = "Name";
        internal const string ATR_DisplayName = "DisplayName";
        internal const string ATR_Hidden = "Hidden";
        internal const string ATR_Required = "Required";
        internal const string ATR_LCID = "LCID";
        internal const string ATR_Max = "Max";
        internal const string ATR_Min = "Min";
        internal const string ATR_Group = "Group";
        internal const string ATR_ReadOnly = "ReadOnly";
        internal const string ATR_Indexed = "Indexed";
        internal const string ATR_ShowInNewForm = "ShowInNewForm";
        internal const string ATR_ShowInEditForm = "ShowInEditForm";
        internal const string ATR_ShowInViewForm = "ShowInViewForm";
        internal const string ATR_Type = "Type";
        internal const string ATR_ShowField = "ShowField";

        //Atributos Field Custom para los tipos de contenido
        internal const string ATR_HIBERUS_Type = "Hiberus-Type"; //Tipo de field asociado a un tipo de contenido, necesario para saber a qué cambiar la columna / campos únicos a actualizar
        internal const string ATR_HIBERUS_TermGroupName = "Hiberus-TermGroupName";    //Fields de tipo taxonomía asociados a un tipo de contenido
        internal const string ATR_HIBERUS_TermSetName = "Hiberus-TermSetName";        //Fields de tipo taxonomía asociados a un tipo de contenido
        internal const string ATR_HIBERUS_RequiredWebUrl = "Hiberus-RequiredWebUrl";  //Fields de tipo lookup asociados a un tipo de contenido
        internal const string ATR_HIBERUS_RequiredListUrlPath = "Hiberus-RequiredListUrlPath";    //Fiels de tipo lookup asociados a un tipo de contenido

        //Atributos ContentType
        internal const string ATR_Description = "Description";
        internal const string ATR_Inherits = "Inherits";

        //Tipos de columnas
        internal const string TYPE_LookUp = "LookUp";
        internal const string TYPE_LookUpMulti = "LookUpMulti";
        internal const string TYPE_Taxonomy = "TaxonomyFieldType";
        internal const string TYPE_TaxonomyMulti = "TaxonomyFieldTypeMulti";
        internal const string TYPE_Currency = "Currency";
        internal const string TYPE_Number = "Number";
        internal const string TYPE_Calculated = "Calculated";
        internal const string TYPE_Note = "Note";

        #endregion

        public static void Empieza(string settingsXMLName, string nombreFicheroLog = "Hiberus-Actualizacion")
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

                #region Leer fichero XML para Actualizar solucion

                System.IO.StreamReader xml = new System.IO.StreamReader(settingsXMLName);
                if (xml == null)
                    throw new Exception(String.Format("No se ha podido recuperar el fichero '{0}'", settingsXMLName));

                XmlSerializer xSerializer = new XmlSerializer(typeof(Tenant));

                Tenant tenant = (Tenant)xSerializer.Deserialize(xml);

                #endregion

                #region Credenciales

                SecureString password = new SecureString();

                foreach (char c in tenant.Credentials.Password.ToCharArray())
                    password.AppendChar(c);

                SharePointOnlineCredentials credentials = new SharePointOnlineCredentials(tenant.Credentials.Account, password);

                #endregion

                #region Fichero | OutPut

                FileStream filestream = new FileStream(string.Format(@"C:\{0}_{1}.txt", nombreFicheroLog, DateTime.Now.ToString("yyyy-MM-dd_hh-mm-ss")), FileMode.Create);
                var streamwriter = new StreamWriter(filestream);
                streamwriter.AutoFlush = true;
                Console.SetOut(streamwriter);
                Console.SetError(streamwriter);

                Console.WriteLine("Inicio: " + DateTime.Now.ToString("yyyy-MM-dd_hh-mm-ss"));

                #endregion

                using (ClientContext clientContext = new ClientContext(tenant.AdminUrl))
                {
                    //credenciales del usuario introducido
                    clientContext.Credentials = credentials;

                    Microsoft.Online.SharePoint.TenantAdministration.Tenant o365Tenant = new Microsoft.Online.SharePoint.TenantAdministration.Tenant(clientContext);

                    foreach (TenantSite site in tenant.Sites)
                    {
                        #region Obtención Site Collection

                        Site currentSite = Apoyo.ExecuteWithTryCatch(() =>
                        {
                            return o365Tenant.GetSiteByUrl(string.Concat(tenant.Url, site.Url));
                        },
                        string.Format("Obteniendo SiteCollection: [Url: {0}]", site.Url));

                        #endregion

                        #region RootWeb

                        if (site.Webs.RootWeb != null && site.Webs.RootWeb.Update)
                        {
                            Console.WriteLine(string.Empty);
                            Console.WriteLine(">> Actualizando la RootWeb <<");
                            Console.WriteLine(string.Empty);

                            clientContext.Load(currentSite.RootWeb, w => w.ServerRelativeUrl);
                            clientContext.Load(currentSite.RootWeb, w => w.Url);
                            clientContext.ExecuteQueryRetry();

                            #region SiteColumns

                            if (site.Webs.RootWeb.SiteColumns != null && site.Webs.RootWeb.SiteColumns.Update)
                                foreach (TenantSiteWebsRootWebSiteColumnsField field in site.Webs.RootWeb.SiteColumns.Field)
                                    Apoyo.ExecuteWithTryCatch(() =>
                                    {
                                        #region Leemos el XML

                                        string xmlString = System.IO.File.ReadAllText(string.Format("{0}\\{1}", site.Webs.RootWeb.SiteColumns.SourcePath, field.SourceXML));

                                        //Generamos un objeto XMLReader
                                        XmlDocument xmlDocument = new XmlDocument();
                                        xmlDocument.LoadXml(xmlString);

                                        //Obtenemos el Nodo Field
                                        XmlNode fieldNode = xmlDocument.GetElementsByTagName("Field").Item(0);

                                        #endregion

                                        if (fieldNode.Attributes[ATR_Name] == null)
                                            throw new Exception(string.Format("La columna {0} no tiene el atributo 'Name'", field.SourceXML));

                                        #region Actualizaciones - Base

                                        Apoyo.ExecuteWithTryCatch(() =>
                                        {
                                            //Recuperamos la columna
                                            Field fieldToUpdate = currentSite.RootWeb.Fields.GetFieldByName<Field>(fieldNode.Attributes[ATR_Name].Value);

                                            //DisplayName
                                            if (fieldNode.Attributes[ATR_DisplayName] != null && !string.IsNullOrWhiteSpace(fieldNode.Attributes[ATR_DisplayName].Value))
                                                fieldToUpdate.Title = fieldNode.Attributes[ATR_DisplayName].Value;

                                            //Hidden
                                            if (fieldNode.Attributes[ATR_Hidden] != null && !string.IsNullOrWhiteSpace(fieldNode.Attributes[ATR_Hidden].Value))
                                                fieldToUpdate.Hidden = bool.Parse(fieldNode.Attributes[ATR_Hidden].Value);

                                            //Required
                                            if (fieldNode.Attributes[ATR_Required] != null && !string.IsNullOrWhiteSpace(fieldNode.Attributes[ATR_Required].Value))
                                                fieldToUpdate.Required = bool.Parse(fieldNode.Attributes[ATR_Required].Value);

                                            //Group
                                            if (fieldNode.Attributes[ATR_Group] != null && !string.IsNullOrEmpty(fieldNode.Attributes[ATR_Group].Value))
                                                fieldToUpdate.Group = fieldNode.Attributes[ATR_Group].Value;

                                            //ReadOnly
                                            if (fieldNode.Attributes[ATR_ReadOnly] != null && !string.IsNullOrWhiteSpace(fieldNode.Attributes[ATR_ReadOnly].Value))
                                                fieldToUpdate.ReadOnlyField = bool.Parse(fieldNode.Attributes[ATR_ReadOnly].Value);

                                            //Indexed
                                            if (fieldNode.Attributes[ATR_Indexed] != null && !string.IsNullOrWhiteSpace(fieldNode.Attributes[ATR_Indexed].Value))
                                                fieldToUpdate.Indexed = bool.Parse(fieldNode.Attributes[ATR_Indexed].Value);

                                            #region Visibilidad del campo

                                            if (fieldNode.Attributes[ATR_ShowInNewForm] != null ||
                                                    fieldNode.Attributes[ATR_ShowInEditForm] != null ||
                                                    fieldNode.Attributes[ATR_ShowInViewForm] != null)
                                            {
                                                Apoyo.ExecuteWithTryCatch(() =>
                                                {
                                                    //ShowInNewForm
                                                    if (fieldNode.Attributes[ATR_ShowInNewForm] != null)
                                                        fieldToUpdate.SetShowInNewForm(bool.Parse(fieldNode.Attributes[ATR_ShowInNewForm].Value));

                                                    //ShowInEditForm
                                                    if (fieldNode.Attributes[ATR_ShowInEditForm] != null)
                                                        fieldToUpdate.SetShowInEditForm(bool.Parse(fieldNode.Attributes[ATR_ShowInEditForm].Value));

                                                    //ShowInViewForm
                                                    if (fieldNode.Attributes[ATR_ShowInViewForm] != null)
                                                        fieldToUpdate.SetShowInEditForm(bool.Parse(fieldNode.Attributes[ATR_ShowInViewForm].Value));
                                                }, "Actualizada visibilidad del campo");
                                            }

                                            //PushChanges por defecto true
                                            if (field.UpdateAndPushChanges)
                                                fieldToUpdate.UpdateAndPushChanges(true);
                                            else
                                                fieldToUpdate.Update();

                                            #endregion

                                            clientContext.ExecuteQueryRetry();
                                        }, "Actulizados atributos base");

                                        #endregion

                                        #region Actualización específica

                                        //Moneda
                                        if (fieldNode.Attributes[ATR_Type].Value.Equals(TYPE_Currency))
                                            Apoyo.ExecuteWithTryCatch(() =>
                                            {
                                                FieldCurrency fieldCurrencyToUpdate = currentSite.RootWeb.Fields.GetFieldByName<FieldCurrency>(fieldNode.Attributes[ATR_Name].Value);

                                                if (fieldNode.Attributes[ATR_LCID] != null && !string.IsNullOrWhiteSpace(fieldNode.Attributes[ATR_LCID].Value))
                                                    fieldCurrencyToUpdate.CurrencyLocaleId = int.Parse(fieldNode.Attributes[ATR_LCID].Value);

                                                //PushChanges por defecto true
                                                if (field.UpdateAndPushChanges)
                                                    fieldCurrencyToUpdate.UpdateAndPushChanges(true);
                                                else
                                                    fieldCurrencyToUpdate.Update();

                                            }, "Actualizadas referencias de moneda");

                                        //Número
                                        if (fieldNode.Attributes[ATR_Type].Value.Equals(TYPE_Number))
                                            Apoyo.ExecuteWithTryCatch(() =>
                                            {
                                                FieldNumber fieldNumberToUpdate = currentSite.RootWeb.Fields.GetFieldByName<FieldNumber>(fieldNode.Attributes[ATR_Name].Value);

                                                if (fieldNode.Attributes[ATR_Max] != null && !string.IsNullOrWhiteSpace(fieldNode.Attributes[ATR_Max].Value))
                                                    fieldNumberToUpdate.MaximumValue = int.Parse(fieldNode.Attributes[ATR_Max].Value);

                                                if (fieldNode.Attributes[ATR_Min] != null && !string.IsNullOrWhiteSpace(fieldNode.Attributes[ATR_Min].Value))
                                                    fieldNumberToUpdate.MinimumValue = int.Parse(fieldNode.Attributes[ATR_Min].Value);

                                                //PushChanges por defecto true
                                                if (field.UpdateAndPushChanges)
                                                    fieldNumberToUpdate.UpdateAndPushChanges(true);
                                                else
                                                    fieldNumberToUpdate.Update();

                                            }, "Actualizadas referencias de número");

                                        //LookUpField
                                        if (fieldNode.Attributes[ATR_Type].Value.Equals(TYPE_LookUp) ||
                                                    fieldNode.Attributes[ATR_Type].Value.Equals(TYPE_LookUpMulti))
                                            Apoyo.ExecuteWithTryCatch(() =>
                                            {
                                                FieldLookup fieldLookUpToUpdate = currentSite.RootWeb.Fields.GetFieldByName<FieldLookup>(fieldNode.Attributes[ATR_Name].Value);

                                                //Recuperamos el Id de la Web a la que apunta
                                                Web requiredWeb = currentSite.RootWeb.GetWeb(field.RequiredWeb.Url);
                                                requiredWeb.EnsureProperty(w => w.Id);

                                                //Recuperamos el Id de la lista a la que apunta
                                                List lookupList = requiredWeb.GetListByUrl(field.RequiredList.UrlPath);
                                                lookupList.EnsureProperty(w => w.Id);

                                                //Actualizamos el xml directamente, da error si se intenta actualizar con propiedad
                                                fieldLookUpToUpdate.Hiberus_UpdateLookupReferences(requiredWeb.Id, lookupList.Id);

                                            }, "Actualizadas referencias de LookUpField");

                                        //Taxonomía
                                        if (fieldNode.Attributes[ATR_Type].Value.Equals(TYPE_Taxonomy) ||
                                                    fieldNode.Attributes[ATR_Type].Value.Equals(TYPE_TaxonomyMulti))
                                            Apoyo.ExecuteWithTryCatch(() =>
                                            {
                                                //Recuperamos field
                                                TaxonomyField fieldTaxonomyToUpdate = currentSite.RootWeb.Fields.GetFieldByName<TaxonomyField>(fieldNode.Attributes[ATR_Name].Value);

                                                //Obtenemos el almacén de términos del site collection
                                                TermStore termStore = currentSite.GetDefaultSiteCollectionTermStore();

                                                //Obtenemos el grupo de términos asociado a la columna
                                                TermGroup termGroup = termStore.Groups.GetByName(field.TermGroupName);

                                                //Obtenemos el conjunto de términos asociado a la columna
                                                TermSet termSet = termGroup.TermSets.GetByName(field.TermSetName);

                                                clientContext.Load(termStore);
                                                clientContext.Load(termSet);
                                                clientContext.ExecuteQueryRetry();

                                                //Multiples valores
                                                fieldTaxonomyToUpdate.AllowMultipleValues = (fieldNode.Attributes[ATR_Type].Value == TYPE_TaxonomyMulti) ? true : false;

                                                //TermSet ID
                                                fieldTaxonomyToUpdate.TermSetId = termSet.Id;

                                                //PushChanges por defecto true
                                                if (field.UpdateAndPushChanges)
                                                    fieldTaxonomyToUpdate.UpdateAndPushChanges(true);
                                                else
                                                    fieldTaxonomyToUpdate.Update();

                                            }, "Actualizadas referencias de Taxonomía");

                                        #endregion

                                        clientContext.ExecuteQueryRetry();

                                    }, string.Format("Actualizando columna: [SourceXML: {0}]", field.SourceXML));

                            #endregion

                            #region ContentTypes

                            if (site.Webs.RootWeb.ContentTypes != null && site.Webs.RootWeb.ContentTypes.Update)
                            {
                                FieldCollection fields = null;
                                Apoyo.ExecuteWithTryCatch(() =>
                                {
                                    fields = currentSite.RootWeb.Fields;
                                    clientContext.Load(fields, w => w.Include(a => a.Id, a => a.SchemaXml, a => a.StaticName));
                                    clientContext.ExecuteQueryRetry();
                                },
                                string.Format("Obtención de los campos de la web"));

                                foreach (TenantSiteWebsRootWebContentTypesContentType contentType in site.Webs.RootWeb.ContentTypes.ContentType)
                                    Apoyo.ExecuteWithTryCatch(() =>
                                    {
                                        #region Leemos el XML

                                        string xmlString = System.IO.File.ReadAllText(string.Format("{0}\\{1}", site.Webs.RootWeb.ContentTypes.SourcePath, contentType.SourceXML));

                                        //Generamos un objeto XMLReader
                                        XmlDocument xmlDocument = new XmlDocument();
                                        xmlDocument.LoadXml(xmlString);

                                        //Obtenemos el Nodo Field
                                        XmlNode contentTypeNode = xmlDocument.GetElementsByTagName("ContentType").Item(0);
                                        XmlNode fieldRefNode = xmlDocument.GetElementsByTagName("FieldRefs").Item(0);

                                        #endregion

                                        //Recuperamos el tipo de contenido
                                        ContentType contentTypeToUpdate = currentSite.RootWeb.ContentTypes.GetById(contentTypeNode.Attributes[ATR_ID].Value);
                                        
                                        //Cargamos las propiedades a usar
                                        contentTypeToUpdate.EnsureProperties(
                                                c => c.Id, c => c.SchemaXml,
                                                c => c.FieldLinks.Include(ctFl => ctFl.Id),
                                                c => c.Fields.Include(ctF => ctF.Id, ctF => ctF.StaticName, ctF => ctF.FromBaseType ));

                                        #region Actualización base

                                        Apoyo.ExecuteWithTryCatch(() =>
                                        {
                                            //Name
                                            if (contentTypeNode.Attributes[ATR_Name] != null && !string.IsNullOrWhiteSpace(contentTypeNode.Attributes[ATR_Name].Value))
                                                contentTypeToUpdate.Name = contentTypeNode.Attributes[ATR_Name].Value;

                                            //Description
                                            if (contentTypeNode.Attributes[ATR_Description] != null && !string.IsNullOrWhiteSpace(contentTypeNode.Attributes[ATR_Description].Value))
                                                contentTypeToUpdate.Description = contentTypeNode.Attributes[ATR_Description].Value;

                                            //Hidden
                                            if (contentTypeNode.Attributes[ATR_Hidden] != null && !string.IsNullOrWhiteSpace(contentTypeNode.Attributes[ATR_Hidden].Value))
                                                contentTypeToUpdate.Hidden = bool.Parse(contentTypeNode.Attributes[ATR_Hidden].Value);

                                            //Grupo
                                            if (contentTypeNode.Attributes[ATR_Group] != null && !string.IsNullOrWhiteSpace(contentTypeNode.Attributes[ATR_Group].Value))
                                                contentTypeToUpdate.Group = contentTypeNode.Attributes[ATR_Group].Value;

                                            //Añadimos a la cola el update
                                            if (contentType.UpdateChildren)
                                                contentTypeToUpdate.Update(true);
                                            else
                                                contentTypeToUpdate.Update(false);

                                        }, "Actualizados atributos base");
                                      
                                        #endregion

                                        #region Actualización Fields

                                        //Recorremos los fields de los nodos para actualizar u crear
                                        fieldRefNode.ChildNodes.Cast<XmlNode>().ToList().ForEach(fieldNode =>
                                        {
                                            Apoyo.ExecuteWithTryCatch(() => 
                                            {
                                                Field fieldToAddUpdate = currentSite.RootWeb.Fields.GetById(Guid.Parse(fieldNode.Attributes[ATR_ID].Value));
                                                //Cargamos el Id y el Esquema de la columna de sitio de la RootWeb
                                                fieldToAddUpdate.EnsureProperties(f => f.Id, f => f.SchemaXmlWithResourceTokens);

                                                FieldLink fLink = contentTypeToUpdate.FieldLinks.FirstOrDefault(fld => fld.Id == fieldToAddUpdate.Id);
                                                if (fLink == null)  //No existe, creamos    FIELDLINK
                                                {
                                                    XElement fieldElement = XElement.Parse(fieldToAddUpdate.SchemaXmlWithResourceTokens);
                                                    fieldElement.SetAttributeValue("AllowDeletion", "TRUE");    //Valor por defecto del producto
                                                    fieldToAddUpdate.SchemaXml = fieldElement.ToString();
                                                    var fldInfo = new FieldLinkCreationInformation();
                                                    fldInfo.Field = fieldToAddUpdate;   //  <-
                                                    contentTypeToUpdate.FieldLinks.Add(fldInfo);
                                                }
                                                else //Existe, actualizamos FIELD
                                                {
                                                    //Recuperamos el field del tipo de contenido
                                                    Field fieldFromContentTypeToUpdate = contentTypeToUpdate.Fields.GetById(Guid.Parse(fieldNode.Attributes[ATR_ID].Value));

                                                    //Cargamos la columna de sitio del tipo de contenido
                                                    clientContext.Load(fieldFromContentTypeToUpdate);
                                                    clientContext.ExecuteQueryRetry();

                                                    //Actualizamos campos
                                                    #region Actualizaciones - Base

                                                    Apoyo.ExecuteWithTryCatch(() =>
                                                    {
                                                        //DisplayName
                                                        if (fieldNode.Attributes[ATR_DisplayName] != null && !string.IsNullOrWhiteSpace(fieldNode.Attributes[ATR_DisplayName].Value))
                                                            fieldFromContentTypeToUpdate.Title = fieldNode.Attributes[ATR_DisplayName].Value;

                                                        //Hidden
                                                        if (fieldNode.Attributes[ATR_Hidden] != null && !string.IsNullOrWhiteSpace(fieldNode.Attributes[ATR_Hidden].Value))
                                                            fieldFromContentTypeToUpdate.Hidden = bool.Parse(fieldNode.Attributes[ATR_Hidden].Value);

                                                        //Required
                                                        if (fieldNode.Attributes[ATR_Required] != null && !string.IsNullOrWhiteSpace(fieldNode.Attributes[ATR_Required].Value))
                                                            fieldFromContentTypeToUpdate.Required = bool.Parse(fieldNode.Attributes[ATR_Required].Value);

                                                        //Group
                                                        if (fieldNode.Attributes[ATR_Group] != null && !string.IsNullOrEmpty(fieldNode.Attributes[ATR_Group].Value))
                                                            fieldFromContentTypeToUpdate.Group = fieldNode.Attributes[ATR_Group].Value;

                                                        //ReadOnly
                                                        if (fieldNode.Attributes[ATR_ReadOnly] != null && !string.IsNullOrWhiteSpace(fieldNode.Attributes[ATR_ReadOnly].Value))
                                                            fieldFromContentTypeToUpdate.ReadOnlyField = bool.Parse(fieldNode.Attributes[ATR_ReadOnly].Value);

                                                        //Indexed
                                                        if (fieldNode.Attributes[ATR_Indexed] != null && !string.IsNullOrWhiteSpace(fieldNode.Attributes[ATR_Indexed].Value))
                                                            fieldFromContentTypeToUpdate.Indexed = bool.Parse(fieldNode.Attributes[ATR_Indexed].Value);

                                                        #region Visibilidad del campo

                                                        if (fieldNode.Attributes[ATR_ShowInNewForm] != null ||
                                                                fieldNode.Attributes[ATR_ShowInEditForm] != null ||
                                                                fieldNode.Attributes[ATR_ShowInViewForm] != null)
                                                        {
                                                            Apoyo.ExecuteWithTryCatch(() =>
                                                            {
                                                                //ShowInNewForm
                                                                if (fieldNode.Attributes[ATR_ShowInNewForm] != null)
                                                                    fieldFromContentTypeToUpdate.SetShowInNewForm(bool.Parse(fieldNode.Attributes[ATR_ShowInNewForm].Value));

                                                                //ShowInEditForm
                                                                if (fieldNode.Attributes[ATR_ShowInEditForm] != null)
                                                                    fieldFromContentTypeToUpdate.SetShowInEditForm(bool.Parse(fieldNode.Attributes[ATR_ShowInEditForm].Value));

                                                                //ShowInViewForm
                                                                if (fieldNode.Attributes[ATR_ShowInViewForm] != null)
                                                                    fieldFromContentTypeToUpdate.SetShowInEditForm(bool.Parse(fieldNode.Attributes[ATR_ShowInViewForm].Value));
                                                            }, "Actualizada visibilidad del campo");
                                                        }

                                                        //Ponemos el update en cola
                                                        if (contentType.UpdateChildren)
                                                            fieldFromContentTypeToUpdate.UpdateAndPushChanges(true);
                                                        else
                                                            fieldFromContentTypeToUpdate.Update();

                                                        #endregion

                                                        clientContext.ExecuteQueryRetry();
                                                    }, "Actulizados atributos del fieldref base");

                                                    #endregion

                                                    #region Actualización específica

                                                    //Moneda
                                                    if (fieldNode.Attributes[ATR_HIBERUS_Type].Value.Equals(TYPE_Currency))
                                                        Apoyo.ExecuteWithTryCatch(() =>
                                                        {
                                                            FieldCurrency fieldFromContentTypeCurrencyToUpdate = contentTypeToUpdate.Fields.GetFieldByName<FieldCurrency>(fieldNode.Attributes[ATR_Name].Value);

                                                            if (fieldNode.Attributes[ATR_LCID] != null && !string.IsNullOrWhiteSpace(fieldNode.Attributes[ATR_LCID].Value))
                                                                fieldFromContentTypeCurrencyToUpdate.CurrencyLocaleId = int.Parse(fieldNode.Attributes[ATR_LCID].Value);

                                                            //Ponemos el update en cola
                                                            if (contentType.UpdateChildren)
                                                                fieldFromContentTypeCurrencyToUpdate.UpdateAndPushChanges(true);
                                                            else
                                                                fieldFromContentTypeCurrencyToUpdate.Update();

                                                        }, "Actualizadas referencias de moneda");

                                                    //Número
                                                    if (fieldNode.Attributes[ATR_HIBERUS_Type].Value.Equals(TYPE_Number))
                                                        Apoyo.ExecuteWithTryCatch(() =>
                                                        {
                                                            FieldNumber fieldFromContentTypeNumberToUpdate = contentTypeToUpdate.Fields.GetFieldByName<FieldNumber>(fieldNode.Attributes[ATR_Name].Value);

                                                            if (fieldNode.Attributes[ATR_Max] != null && !string.IsNullOrWhiteSpace(fieldNode.Attributes[ATR_Max].Value))
                                                                fieldFromContentTypeNumberToUpdate.MaximumValue = int.Parse(fieldNode.Attributes[ATR_Max].Value);

                                                            if (fieldNode.Attributes[ATR_Min] != null && !string.IsNullOrWhiteSpace(fieldNode.Attributes[ATR_Min].Value))
                                                                fieldFromContentTypeNumberToUpdate.MinimumValue = int.Parse(fieldNode.Attributes[ATR_Min].Value);

                                                            //Ponemos el update en cola
                                                            if (contentType.UpdateChildren)
                                                                fieldFromContentTypeNumberToUpdate.UpdateAndPushChanges(true);
                                                            else
                                                                fieldFromContentTypeNumberToUpdate.Update();

                                                        }, "Actualizadas referencias de número");

                                                    //LookUpField
                                                    if (fieldNode.Attributes[ATR_HIBERUS_Type].Value.Equals(TYPE_LookUp) ||
                                                        fieldNode.Attributes[ATR_HIBERUS_Type].Value.Equals(TYPE_LookUpMulti))
                                                        Apoyo.ExecuteWithTryCatch(() =>
                                                        {
                                                            FieldLookup fieldFromContentTypeLookUpToUpdate = currentSite.RootWeb.Fields.GetFieldByName<FieldLookup>(fieldNode.Attributes[ATR_Name].Value);

                                                            //Recuperamos el Id de la Web a la que apunta
                                                            Web requiredWeb = currentSite.RootWeb.GetWeb(fieldNode.Attributes[ATR_HIBERUS_RequiredWebUrl].Value);
                                                            requiredWeb.EnsureProperty(w => w.Id);

                                                            //Recuperamos el Id de la lista a la que apunta
                                                            List lookupList = requiredWeb.GetListByUrl(fieldNode.Attributes[ATR_HIBERUS_RequiredListUrlPath].Value);
                                                            lookupList.EnsureProperty(w => w.Id);

                                                            //Actualizamos el xml directamente, da error si se intenta actualizar con propiedad
                                                            fieldFromContentTypeLookUpToUpdate.Hiberus_UpdateLookupReferences(requiredWeb.Id, lookupList.Id);

                                                        }, "Actualizadas referencias de LookUpField");

                                                    //Taxonomía
                                                    if (fieldNode.Attributes[ATR_HIBERUS_Type].Value.Equals(TYPE_Taxonomy) ||
                                                        fieldNode.Attributes[ATR_HIBERUS_Type].Value.Equals(TYPE_TaxonomyMulti))
                                                        Apoyo.ExecuteWithTryCatch(() =>
                                                        {
                                                            if (fieldNode.Attributes[ATR_HIBERUS_TermGroupName] == null || fieldNode.Attributes[ATR_HIBERUS_TermSetName] == null)
                                                                throw new ArgumentException("Valores de entrada no válidos");

                                                            //Recuperamos field
                                                            TaxonomyField fieldFromContentTypeTaxonomyToUpdate = contentTypeToUpdate.Fields.GetFieldByName<TaxonomyField>(fieldNode.Attributes[ATR_Name].Value);

                                                            //Obtenemos el almacén de términos del site collection
                                                            TermStore termStore = currentSite.GetDefaultSiteCollectionTermStore();

                                                            //Obtenemos el grupo de términos asociado a la columna
                                                            TermGroup termGroup = termStore.Groups.GetByName(fieldNode.Attributes[ATR_HIBERUS_TermGroupName].Value); //No es un atributo del producto

                                                            //Obtenemos el conjunto de términos asociado a la columna
                                                            TermSet termSet = termGroup.TermSets.GetByName(fieldNode.Attributes[ATR_HIBERUS_TermSetName].Value);  //No es un atributo del pructo

                                                            //cargamos el almacén de términos y el conjunto de términos
                                                            clientContext.Load(termStore);
                                                            clientContext.Load(termSet);
                                                            clientContext.ExecuteQueryRetry();

                                                            //Multiples valores
                                                            if (fieldNode.Attributes[ATR_Type] == null)
                                                                fieldFromContentTypeTaxonomyToUpdate.AllowMultipleValues = false;
                                                            else
                                                                fieldFromContentTypeTaxonomyToUpdate.AllowMultipleValues = (fieldNode.Attributes[ATR_Type].Value == TYPE_TaxonomyMulti) ? true : false;

                                                            //TermSet ID
                                                            fieldFromContentTypeTaxonomyToUpdate.TermSetId = termSet.Id;
                                                        }, "Actualizadas referencias de Taxonomía");

                                                    #endregion

                                                    //Ponemos en cola el update
                                                    if (contentType.UpdateChildren)
                                                        fieldFromContentTypeToUpdate.UpdateAndPushChanges(true);
                                                    else
                                                        fieldFromContentTypeToUpdate.Update();
                                                    clientContext.ExecuteQueryRetry();
                                                }
                                            });
                                        });

                                        #endregion

                                        #region Borrado de fields 

                                        //Obtengo los fields que son diferentes con el padre, es decir, los fields propios del tipo de contenido a actualizar
                                        ContentType parentCTFromContentTypeToUpdate = contentTypeToUpdate.Parent;
                                        parentCTFromContentTypeToUpdate.EnsureProperties(w => w.Fields.Include(a => a.StaticName));

                                        List<string> staticFieldsFromContentTypeToUpdate = contentTypeToUpdate
                                                                                            .Fields
                                                                                            .Select(x => x.StaticName)
                                                                                            .Except(parentCTFromContentTypeToUpdate.Fields.Select(x => x.StaticName))
                                                                                            .ToList();

                                        //Obtenemos la diferencia entre: StaticName de los fields del tipo de contenido / 
                                        //                               StaticName de los fields de XML para borrar las columnas sobrantes
                                        IEnumerable<String> fieldsNamesFromContentTypeToRemove = staticFieldsFromContentTypeToUpdate
                                                                                                            .Except(fieldRefNode.ChildNodes.Cast<XmlNode>().Select(x => x.Attributes[ATR_Name].Value));


                                        //Borramos las columnas que tiene asociadas el tipo de contenido actual que no aplican al XML del tipo de contenido
                                        if (fieldsNamesFromContentTypeToRemove != null && fieldsNamesFromContentTypeToRemove.Count() > 0)
                                            foreach (string fieldNameFromContentTypeToRemove in fieldsNamesFromContentTypeToRemove)
                                                Apoyo.ExecuteWithTryCatch(() =>
                                                {
                                                    //Recuperamos el field del tipo de contenido a borrar
                                                    Field fieldFromContentTypeToRemove = contentTypeToUpdate.Fields.GetFieldByName<Field>(fieldNameFromContentTypeToRemove);

                                                    //Cargamos para que se pueda actualizar
                                                    clientContext.Load(fieldFromContentTypeToRemove);
                                                    clientContext.ExecuteQueryRetry();

                                                    //Borramos
                                                    fieldFromContentTypeToRemove.DeleteObject();

                                                    //Actualizamos tipo de contenido
                                                    if (contentType.UpdateChildren)
                                                        contentTypeToUpdate.Update(true);
                                                    else
                                                        contentTypeToUpdate.Update(false);
                                                    clientContext.ExecuteQueryRetry(); 

                                                }, string.Format("Borrado campo: [{0}]", fieldNameFromContentTypeToRemove));

                                        #endregion

                                        //Actualizamos tipo de contenido
                                        if (contentType.UpdateChildren)
                                            contentTypeToUpdate.Update(true);
                                        else
                                            contentTypeToUpdate.Update(false);
                                        clientContext.ExecuteQueryRetry();

                                    }, string.Format("Actualizando Tipo de contenido: [SourceXML: {0}]", contentType.SourceXML));
                            }

                            #endregion

                            #region Lists

                            if (site.Webs.RootWeb.Lists != null && site.Webs.RootWeb.Lists.Update)
                                foreach (TenantSiteWebsRootWebListsList list in site.Webs.RootWeb.Lists.List)
                                {
                                    Apoyo.ExecuteWithTryCatch(() =>
                                    {

                                        #region Obtención de la lista

                                        List currentList = currentSite.RootWeb.GetListByTitle(list.Name);

                                        clientContext.Load(currentList, w => w.ContentTypes);
                                        clientContext.ExecuteQueryRetry();

                                        #endregion

                                        //currentList.ContentTypes.First().FieldLinks

                                    },
                                    string.Format("Actualizando Lista: [Nombre: {0}] [Template: {1}]", list.Name, list.TemplateType));
                                }

                            #endregion
                        }

                        #endregion

                        #region Webs

                        if (site.Webs.Update && site.Webs.Web != null)

                            foreach (TenantSiteWebsWeb web in site.Webs.Web)
                            {
                                if (!web.Update)
                                    continue;

                                Console.WriteLine(string.Empty);
                                Console.WriteLine(">> Actualizando la Web {0} <<", web.Url);
                                Console.WriteLine(string.Empty);

                                #region Obtención Web

                                Web currentWeb = currentSite.RootWeb.GetWeb(web.Url);

                                clientContext.Load(currentWeb, w => w.ServerRelativeUrl);
                                clientContext.ExecuteQueryRetry();

                                #endregion

                                #region SiteColumns

                                #endregion

                                #region ContentTypes

                                #endregion

                                #region Lists

                                #endregion
                            }

                        #endregion
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
                Console.ReadLine();
            }
        }
    }
}
