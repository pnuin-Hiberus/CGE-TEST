using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace AprovisionamientoO365.Base.Extensions
{
    public static class ListExtensions
    {
        /// <summary>
        /// Creates list views based on specific xml structure from file
        /// </summary>
        /// <param name="web"></param>
        /// <param name="listUrl"></param>
        /// <param name="filePath"></param>
        public static void CreateViewsFromXMLFile_Hiberus(this Web web, string listUrl, string filePath)
        {
            if (string.IsNullOrEmpty(listUrl))
                throw new ArgumentNullException("listUrl");

            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentNullException("filePath");

            XmlDocument xd = new XmlDocument();
            xd.Load(filePath);
            CreateViewsFromXML_Hiberus(web, listUrl, xd);
        }

        /// <summary>
        /// Create list views based on xml structure loaded to memory
        /// </summary>
        /// <param name="web"></param>
        /// <param name="listUrl"></param>
        /// <param name="xmlDoc"></param>
        public static void CreateViewsFromXML_Hiberus(this Web web, string listUrl, XmlDocument xmlDoc)
        {
            if (string.IsNullOrEmpty(listUrl))
                throw new ArgumentNullException("listUrl");

            if (xmlDoc == null)
                throw new ArgumentNullException("xmlDoc");

            // Get instances to the list
            List list = web.GetList(listUrl);
            web.Context.Load(list);
            web.Context.ExecuteQueryRetry();

            // Execute the actual xml based creation
            list.CreateViewsFromXML_Hiberus(xmlDoc,web);

        }


        /// <summary>
        /// Actual implementation of the view creation logic based on given xml
        /// </summary>
        /// <param name="list"></param>
        /// <param name="xmlDoc"></param>
        public static View CreateViewsFromXML_Hiberus(this List list, XmlDocument xmlDoc, Web web)
        {
            if (xmlDoc == null)
                throw new ArgumentNullException("xmlDoc");

            // Convert base type to string value used in the xml structure
            string listType = list.BaseType.ToString();
            // Get only relevant list views for matching base list type
            XmlNodeList listViews = xmlDoc.SelectNodes("ListViews/List[@Type='" + listType + "']/View");
            int count = listViews.Count;
            foreach (XmlNode view in listViews)
            {
                string name = view.Attributes["Name"].Value;
                ViewType type = (ViewType)Enum.Parse(typeof(ViewType), view.Attributes["ViewTypeKind"].Value);
                string[] viewFields = view.Attributes["ViewFields"].Value.Split(',');
                uint rowLimit = uint.Parse(view.Attributes["RowLimit"].Value);
                bool defaultView = bool.Parse(view.Attributes["DefaultView"].Value);
                bool paged = bool.Parse(view.Attributes["Paged"].Value);
                string query = view.SelectSingleNode("./ViewQuery").InnerText;
                string jsLink = view.SelectSingleNode("./JSLink") != null ? view.SelectSingleNode("./JSLink").InnerText : String.Empty;

                //Create View
                //list.CreateView_Hiberus(name, type, viewFields, rowLimit, defaultView, query, false, paged);
                list.CreateView_Hiberus(web,name, type, viewFields, rowLimit, defaultView,jsLink , query, false, paged);

                

            }

            return null;
        }

        /// <summary>
        /// Create view to existing list
        /// </summary>
        /// <param name="list"></param>
        /// <param name="viewName"></param>
        /// <param name="viewType"></param>
        /// <param name="viewFields"></param>
        /// <param name="rowLimit"></param>
        /// <param name="setAsDefault"></param>
        /// <param name="query"></param>        
        /// <param name="personal"></param>
        /// <param name="paged"></param>        
        //public static View CreateView_Hiberus(this List list,
        //                                      string viewName,
        //                                      ViewType viewType,
        //                                      string[] viewFields,
        //                                      uint rowLimit,
        //                                      bool setAsDefault,
        //                                      string query = null,
        //                                      bool personal = false,
        //                                      bool paged = false)
        //{
        //    if (string.IsNullOrEmpty(viewName))
        //        throw new ArgumentNullException("viewName");

        //    ViewCreationInformation viewCreationInformation = new ViewCreationInformation();
        //    viewCreationInformation.Title = viewName;
        //    viewCreationInformation.ViewTypeKind = viewType;
        //    viewCreationInformation.RowLimit = rowLimit;
        //    viewCreationInformation.ViewFields = viewFields;
        //    viewCreationInformation.PersonalView = personal;
        //    viewCreationInformation.SetAsDefaultView = setAsDefault;
        //    viewCreationInformation.Paged = paged;
        //    if (!string.IsNullOrEmpty(query))
        //    {
        //        viewCreationInformation.Query = query;
        //    }

        //    View view = list.Views.Add(viewCreationInformation);
        //    list.Context.Load(view);
        //    list.Context.ExecuteQueryRetry();

        //    return view;
        //}

        /// <summary>
        /// Create view to existing list
        /// </summary>
        /// <param name="list"></param>
        /// <param name="viewName"></param>
        /// <param name="viewType"></param>
        /// <param name="viewFields"></param>
        /// <param name="rowLimit"></param>
        /// <param name="setAsDefault"></param>
        /// <param name="query"></param>        
        /// <param name="personal"></param>
        /// <param name="paged"></param>        
        public static View CreateView_Hiberus(this List list,
                                              Web web,
                                              string viewName,
                                              ViewType viewType,
                                              string[] viewFields,
                                              uint rowLimit,
                                              bool setAsDefault,
                                              string jsLink,
                                              string query = null,
                                              bool personal = false,
                                              bool paged = false)
        {
            if (string.IsNullOrEmpty(viewName))
                throw new ArgumentNullException("viewName");

            ViewCreationInformation viewCreationInformation = new ViewCreationInformation();
            viewCreationInformation.Title = viewName;
            viewCreationInformation.ViewTypeKind = viewType;
            viewCreationInformation.RowLimit = rowLimit;
            viewCreationInformation.ViewFields = viewFields;
            viewCreationInformation.PersonalView = personal;
            viewCreationInformation.SetAsDefaultView = setAsDefault;
            viewCreationInformation.Paged = paged;
            if (!string.IsNullOrEmpty(query))
            {
                viewCreationInformation.Query = query;
            }

            View view = list.Views.Add(viewCreationInformation);
            list.Context.Load(view);
            list.Context.ExecuteQueryRetry();

            if (!String.IsNullOrEmpty(jsLink))
            {
                //Hay que abrir un clientcontext contra la coleccion de sitios no vale la tuta de administración
                //porque sino al guardar la webpart da un error
                //FUENTE: http://sharepoint.stackexchange.com/questions/147324/set-authorizationfilter-property-of-a-webpart
                using (ClientContext clientContext = new ClientContext(web.Url))
                {
                    clientContext.Credentials = list.Context.Credentials;

                    Web w = clientContext.Web;
                    clientContext.Load(w);
                    clientContext.ExecuteQuery();

                    Microsoft.SharePoint.Client.File file = w.GetFileByServerRelativeUrl(view.ServerRelativeUrl);
                    LimitedWebPartManager wpm = file.GetLimitedWebPartManager(PersonalizationScope.Shared);
                    clientContext.Load(wpm.WebParts, wps => wps.Include(wp => wp.WebPart.Title));
                    clientContext.ExecuteQuery();

                    //Set the properties for all web parts
                    foreach (WebPartDefinition wpd in wpm.WebParts)
                    {
                        WebPart wp = wpd.WebPart;
                        wp.Properties["JSLink"] = jsLink;
                        wpd.SaveWebPartChanges();
                        clientContext.ExecuteQuery();
                    }
                }
            }



            return view;
        }

        public static void CreateQuestion_Hiberus(string FieldSourcePath, List surveyList)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(FieldSourcePath); ;
            surveyList.Fields.AddFieldAsXml(doc.OuterXml, false, AddFieldOptions.AddToDefaultContentType);
        }

    }
}
