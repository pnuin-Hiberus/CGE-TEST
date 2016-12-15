using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using AprovisionamientoO365.Base;
using System.Xml.Linq;

namespace AprovisionamientoO365.Base.Extensions
{
    public static class SiteExtensions
    {
        #region Constantes

        //Custom Action node Constantes
        internal const string NODE_CustomAction = "CustomAction";
        internal const string ATR_Id = "Id";
        internal const string ATR_RegistrationType = "RegistrationType";
        internal const string ATR_RegistrationId = "RegistrationId";
        internal const string ATR_Location = "Location";
        internal const string ATR_Title = "Title";
        internal const string ATR_Url = "Url";
        internal const string ATR_Group = "Group";
        internal const string ATR_Description = "Description";
        internal const string ATR_ImageUrl = "ImageUrl";
        internal const string ATR_Name = "Name";
        internal const string ATR_Remove = "Remove";
        internal const string ATR_Rights = "Rights";
        internal const string ATR_ScriptBlock = "ScriptBlock";
        internal const string ATR_ScriptSrc = "ScriptSrc";
        internal const string ATR_CommandUIExtension = "CommandUIExtension";
        internal const string ATR_Sequence = "Sequence";

        //CommandUIExtension
        internal const string NODE_CommandUIExtension = "CommandUIExtension";

        #endregion

        /// <summary>
        /// Aprovisiona las listas de las webs de una coleccion de sitios
        /// </summary>
        /// <param name="site"></param>
        /// <param name="absolutePathToFile"></param>
        public static void PopulateListItemsFromXMLFile(this Site site, string absolutePathToFile)
        {
            XmlDocument xd = new XmlDocument();
            xd.Load(absolutePathToFile);

            //obtengo las webs
            XmlNodeList webs = xd.SelectNodes("Webs/Web");

            foreach (XmlNode web in webs)
            {
                Web currentWeb = site.OpenWeb(web.Attributes["Url"].Value);

                //obtengo listas
                XmlNodeList listas = web.SelectNodes("Lists/ListInstance");

                foreach (XmlNode lista in listas)
                {
                    var listUrl = lista.Attributes["Url"].Value;

                    List currentList = currentWeb.GetListByUrl(listUrl);
                    currentWeb.Context.Load(currentList);
                    currentWeb.Context.ExecuteQueryRetry();

                    if (currentList == null) continue;

                    XmlNodeList listItems = lista.SelectNodes("Data/Rows/Row");

                    Apoyo.ExecuteWithTryCatch(() =>
                    {
                        #region List items

                        //Obtengo filas para añadir
                        foreach (XmlNode listItem in listItems)
                        {
                            XmlNodeList fieldItems = listItem.SelectNodes("Field");
                            ListItem currentListItem = currentList.AddItem(new ListItemCreationInformation());

                            //Cada campo del listitem
                            foreach (XmlNode fieldItem in fieldItems)
                            {
                                #region normal
                                if (fieldItem.Attributes["Type"] == null || string.IsNullOrEmpty(fieldItem.Attributes["Type"].Value))
                                    currentListItem[fieldItem.Attributes["Name"].Value] = fieldItem.InnerText;
                                #endregion

                                #region Taxonomia

                                else if (fieldItem.Attributes["Type"].Value.Equals("Taxonomy"))
                                {
                                    Field field = currentListItem.ParentList.Fields.GetByInternalNameOrTitle(fieldItem.Attributes["Name"].Value);
                                    TaxonomyField txField = currentWeb.Context.CastTo<TaxonomyField>(field);

                                    TaxonomyFieldValue termValue = new TaxonomyFieldValue();
                                    termValue.Label = fieldItem.InnerText.Split('|')[0];
                                    termValue.TermGuid = fieldItem.InnerText.Split('|')[1];
                                    termValue.WssId = -1;   //Si se pone valor -1 sharepoint resuelve el WssId automáticamente
                                    txField.SetFieldValueByValue(currentListItem, termValue);
                                }

                                #endregion

                                #region LookUp

                                else if (fieldItem.Attributes["Type"].Value.Equals("Lookup"))
                                {
                                    FieldLookupValue lv = new FieldLookupValue();
                                    lv.LookupId = fieldItem.InnerText.ToInt32();
                                    currentListItem[fieldItem.Attributes["Name"].Value] = lv;
                                }

                                #endregion
                            }

                            //Guardo
                            currentListItem.Update();
                            currentWeb.Context.ExecuteQueryRetry();
                        }

                        #endregion

                    }, string.Format("\tWeb [{0}] - Lista [{1}]", web.Attributes["Url"].Value, lista.Attributes["Url"].Value));
                }
            }
        }

        /// <summary>
        /// Adds a custom action to the current site through his XML path file
        /// </summary>
        /// <param name="site">Colección de sitios a añadir el customaction</param>
        /// <param name="absolutePathToFile">path del fichero XML</param>
        public static Boolean AddCustomActionFromXMLFile_Hiberus(this Site site, string absolutePathToFile)
        {
            var xd = XDocument.Load(absolutePathToFile);
            return AddCustomActionFromXML_Hiberus(site, xd);
        }

        /// <summary>
        /// Adds a custom action to the current site through his XML 
        /// </summary>
        /// <param name="site">Colección de sitios a añadir el customaction</param>
        /// <param name="xDocument">Esquema XML</param>
        /// <returns></returns>
        private static Boolean AddCustomActionFromXML_Hiberus(this Site site, XDocument xDocument)
        {
            try
            {
                var ns = xDocument.Root.Name.Namespace;
                var customActionNodes = from cAction in xDocument.Descendants(ns + NODE_CustomAction) select cAction;

                //Recupero el nodo del custom action
                XElement customActionNode = customActionNodes.First();

                if (customActionNode.Attribute(ATR_Id) == null || string.IsNullOrEmpty(customActionNode.Attribute(ATR_Id).Value) ||
                    customActionNode.Attribute(ATR_Title) == null || string.IsNullOrEmpty(customActionNode.Attribute(ATR_Title).Value))
                    throw new ArgumentException("Valores de entrada no informados");

                #region Recuperar si existe / Borrar

                Apoyo.ExecuteWithTryCatch(() => 
                {
                    UserCustomAction userCustomPrevio = site.GetCustomActions().Where(c => c.Title.Equals(customActionNode.Attribute(ATR_Title).Value)).FirstOrDefault();

                    if (userCustomPrevio != null)
                        site.DeleteCustomAction(userCustomPrevio.Id);
                }, string.Format("se ha borrado el custom action anterior"));

                #endregion

                #region Añadir Custom action a la colección de sitios

                //Recupero el nodo de CommandUIExtension
                XElement commandUIExtension = customActionNode.Descendants(ns + NODE_CommandUIExtension).FirstOrDefault();

                //Añado la información del custom action a la clase de PnP para el aprovisionamiento
                OfficeDevPnP.Core.Entities.CustomActionEntity customActionEntity = new OfficeDevPnP.Core.Entities.CustomActionEntity()
                {
                    RegistrationId = customActionNode.Attribute(ATR_RegistrationId) != null ? customActionNode.Attribute(ATR_RegistrationId).Value : string.Empty,
                    RegistrationType = customActionNode.Attribute(ATR_RegistrationType) != null ? (UserCustomActionRegistrationType)Enum.Parse(typeof(UserCustomActionRegistrationType), customActionNode.Attribute(ATR_RegistrationType).Value) : UserCustomActionRegistrationType.None,
                    Sequence = customActionNode.Attribute(ATR_Sequence) != null ? Int32.Parse(customActionNode.Attribute(ATR_Sequence).Value) : 0,
                    Title = customActionNode.Attribute(ATR_Title) != null ? customActionNode.Attribute(ATR_Title).Value : string.Empty,
                    Url = customActionNode.Attribute(ATR_Url) != null ? customActionNode.Attribute(ATR_Url).Value : string.Empty,
                    Description = customActionNode.Attribute(ATR_Description) != null ? customActionNode.Attribute(ATR_Description).Value : string.Empty,
                    Group = customActionNode.Attribute(ATR_Group) != null ? customActionNode.Attribute(ATR_Group).Value : string.Empty,
                    ImageUrl = customActionNode.Attribute(ATR_ImageUrl) != null ? customActionNode.Attribute(ATR_ImageUrl).Value : string.Empty,
                    Location = customActionNode.Attribute(ATR_Location) != null ? customActionNode.Attribute(ATR_Location).Value : string.Empty,
                    Name = customActionNode.Attribute(ATR_Name) != null ? customActionNode.Attribute(ATR_Name).Value : string.Empty,
                    Remove = customActionNode.Attribute(ATR_Remove) != null ? bool.Parse(customActionNode.Attribute(ATR_Remove).Value) : false,
                    ScriptBlock = customActionNode.Attribute(ATR_ScriptBlock) != null ? customActionNode.Attribute(ATR_ScriptBlock).Value : string.Empty,
                    ScriptSrc = customActionNode.Attribute(ATR_ScriptSrc) != null ? customActionNode.Attribute(ATR_ScriptSrc).Value : string.Empty,
                    CommandUIExtension = (commandUIExtension) != null ? commandUIExtension.ToString(SaveOptions.DisableFormatting) : string.Empty
                };

                #endregion

                return site.AddCustomAction(customActionEntity);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }

}
