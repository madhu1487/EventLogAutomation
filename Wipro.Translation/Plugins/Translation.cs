using Microsoft.Xrm.Sdk;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.ServiceModel;
using System.Text;
using Wipro.Translation.Plugins.Common;


namespace Wipro.Translation.Plugins
{
    public class Translation : Plugin
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="Translation"/> class.
        /// </summary>
        public Translation() : base(typeof(Translation))
        {
            base.RegisteredEvents.Add(new Tuple<SdkMessageProcessingStepStage, string, string, Action<LocalPluginContext>>(SdkMessageProcessingStepStage.Postoperation, "create", "wiptrans_translationsnippet", new Action<LocalPluginContext>(ExecuteTranslationSnippet)));
            base.RegisteredEvents.Add(new Tuple<SdkMessageProcessingStepStage, string, string, Action<LocalPluginContext>>(SdkMessageProcessingStepStage.Postoperation, "update", "wiptrans_translationsnippet", new Action<LocalPluginContext>(ExecuteTranslationSnippet)));

            //base.RegisteredEvents.Add(new Tuple<SdkMessageProcessingStepStage, string, string, Action<LocalPluginContext>>(SdkMessageProcessingStepStage.Postoperation, "<Plugin Message Name>", "<Entity Logical Name>", new Action<LocalPluginContext>(ExecuteTranslation)));
        }
        #endregion

        #region Protected Subs / Functions
        protected void ExecuteTranslationSnippet(LocalPluginContext localPluginContext)
        {
            Exception exception = null;

            if (localPluginContext == null)
                exception = new ArgumentNullException("localPluginContext");

            // TODO: Implement your custom Plug-in business logic.
            IPluginExecutionContext context = localPluginContext.PluginExecutionContext;
            IOrganizationService service = localPluginContext.OrganizationService;

            //exit if the depth exceeds 2...otherwise infinite loop...and subsequent error.
            if (context.Depth > 2)
                return;

            try
            {
                Entity entity = null;
                EntityReference entityReference = null;

                if (context.InputParameters.Contains("EntityMoniker"))
                    entityReference = (EntityReference)context.InputParameters["EntityMoniker"];
                if (context.InputParameters.Contains("Target"))
                {
                    entity = context.InputParameters["Target"] as Microsoft.Xrm.Sdk.Entity;
                    if (entity == null)
                        entityReference = (EntityReference)context.InputParameters["Target"];
                }
                if (entity != null && entityReference == null)
                    entityReference = entity.ToEntityReference();

                //Null validation
                if (entityReference != null)
                {
                    string traslationText = localPluginContext.GetAttributeValue<string>("wiptrans_translatetext"),
                        translatedText = string.Empty;

                    Int32? languageFrom = localPluginContext.GetAttributeValue<Int32?>("wiptrans_languagefrom"),
                        languageTo = localPluginContext.GetAttributeValue<Int32?>("wiptrans_languageto");

                    if (!string.IsNullOrEmpty(traslationText) && languageFrom != null && languageTo != null)
                    {
                        translatedText = ProcessTranslationProviders(service, languageFrom.Value, languageTo.Value, traslationText, entityReference);
                        if (!string.IsNullOrEmpty(translatedText))
                        {
                            Entity snippet = new Entity(entityReference.LogicalName, entityReference.Id);
                            snippet.Attributes.Add("wiptrans_translatedtext", translatedText);
                            service.Update(snippet);
                        }
                    }
                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                localPluginContext.ThrowException("Fault Exception Occured", ex);
            }
            catch (Exception ex)
            {
                exception = ex;
            }
            finally
            {
                //Throw exception
                if (exception != null)
                    throw new InvalidPluginExecutionException(exception.Message);
            }
        }
        #endregion

        #region private Subs / Functions
        private string ProcessTranslationProviders(IOrganizationService service, Int32 from, Int32 to, string text, EntityReference entityReference)
        {
            List<Int32> languageList = new List<Int32>();
            languageList.Add(from);
            languageList.Add(to);

            EntityCollection translatorProviderCollection = HelperFunctions.RetrieveTranslationProvider(service),
                languageCollection = HelperFunctions.RetrieveLanguage(service, languageList);

            try
            {
                string fromCode = languageCollection.Entities.Where(item => HelperFunctions.GetAttributeValue<Int32>("localeid", item).Equals(from)).Select(item => HelperFunctions.GetAttributeValue<string>("code", item)).FirstOrDefault(),
                        toCode = languageCollection.Entities.Where(item => HelperFunctions.GetAttributeValue<Int32>("localeid", item).Equals(to)).Select(item => HelperFunctions.GetAttributeValue<string>("code", item)).FirstOrDefault(),
                        translatedText = null;

                List<string> translatedTextList = null;

                if (translatorProviderCollection != null && translatorProviderCollection.Entities.Any())
                {
                    // This logic will require significant changes. As for some providers we need to support OAuth based authentication.
                    translatorProviderCollection.Entities.ToList().ForEach(item =>
                    {
                        // These are the parameters the Providers are expecting. Also, there will be additional parameters to improve the accuracy of the Translation.
                        string textKey = HelperFunctions.GetAttributeValue<string>("wiptrans_parametertextkey", item),
                            languageFromKey = HelperFunctions.GetAttributeValue<string>("wiptrans_parametersourcekey", item),
                            languageToKey = HelperFunctions.GetAttributeValue<string>("wiptrans_parameterdestinationkey", item),
                            apiKeyName = HelperFunctions.GetAttributeValue<string>("wiptrans_clientkeyname", item),
                            apiKeyValue = HelperFunctions.GetAttributeValue<string>("wiptrans_clientkey", item),
                            uri = string.Empty;

                        int httpVerb = HelperFunctions.GetAttributeValue<OptionSetValue>("wiptrans_requesttype", item).Value;
                        HttpResponseMessage response = null;

                        // TODO: Add separate check for OAuth Required and a separate complete logic for OAuth enable APIs.
                        // Will move code blocks to different methods to increase readability.
                        // Add unit test to maximize test coverage.
                        switch (httpVerb)
                        {
                            case 1://Get
                                {
                                    if (HelperFunctions.GetAttributeValue<bool>("wiptrans_passkeyasheader", item))
                                    {
                                        // Creating the URI for Get request
                                        uri = string.Format("{0}?{1}={2}&{3}={4}&{5}={6}", HelperFunctions.GetAttributeValue<string>("wiptrans_url", item), textKey, text, languageFromKey, fromCode, languageToKey, toCode);
                                        using (HttpClient aClient = new HttpClient())
                                        using (HttpRequestMessage request = new HttpRequestMessage())
                                        {
                                            request.Method = HttpMethod.Get;
                                            aClient.DefaultRequestHeaders.Add(apiKeyName, apiKeyValue);

                                            response = aClient.GetAsync(new Uri(uri)).Result;
                                        }
                                    }
                                };
                                break;

                            case 2:// Post
                                {

                                    using (HttpClient aClient = new HttpClient())
                                    using (HttpRequestMessage request = new HttpRequestMessage())
                                    {
                                        if (HelperFunctions.GetAttributeValue<bool>("wiptrans_passkeyasheader", item))// If api key to be passed as Header.
                                        {
                                            uri = string.Format("{0}", HelperFunctions.GetAttributeValue<string>("wiptrans_url", item));
                                            request.Headers.TryAddWithoutValidation(apiKeyName, apiKeyValue);
                                        }
                                        else // API key is being passed as parameter.
                                        {
                                            uri = string.Format("{0}?{1}={2}", HelperFunctions.GetAttributeValue<string>("wiptrans_url", item), apiKeyName, apiKeyValue); //"https://translation.googleapis.com/language/translate/v2?key=" + apiKey;
                                        }

                                        request.Method = HttpMethod.Post;
                                        request.RequestUri = new Uri(uri);
                                        request.Headers.TryAddWithoutValidation("Content-Type", "application/json");

                                        //Below section can be done differently as well if requried as it is using Newtonsoft Dll. Not sure if that is supported by Dynamcis.
                                        JObject requestBody = new JObject();
                                        requestBody.Add(translatedText, text);
                                        requestBody.Add(languageFromKey, fromCode);
                                        requestBody.Add(languageToKey, toCode);

                                        request.Content = new StringContent(requestBody.ToString(), Encoding.UTF8, "application/json");

                                        //aClient.Timeout = TimeSpan.FromSeconds(2); // TO add request time out for high volume production environments.
                                        response = aClient.SendAsync(request).Result;
                                    }
                                };
                                break;
                        }

                        // This will throw exception if error occured.
                        response.EnsureSuccessStatusCode();

                        translatedText = response.Content.ReadAsStringAsync().Result;
                        translatedTextList.Add(translatedText);

                        if (!string.IsNullOrEmpty(translatedText))
                        {
                            Entity mapping = new Entity("wiptrans_translationmapper");
                            mapping.Attributes.Add("wiptrans_languagefrom", from);
                            mapping.Attributes.Add("wiptrans_languageto", to);
                            mapping.Attributes.Add("wiptrans_translatetext", text);
                            mapping.Attributes.Add("wiptrans_translatedtext", translatedText);
                            mapping.Attributes.Add("wiptrans_translationprovideridid", item.ToEntityReference());
                            mapping.Attributes.Add("wiptrans_translationsnippetid", entityReference);
                            service.Create(mapping);
                        }

                    });
                }
                else
                {
                    throw new InvalidPluginExecutionException("No translation providers available");
                }

                translatedText = SelectTextFromAvailableTexts(translatedTextList);

                return translatedText;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private string SelectTextFromAvailableTexts(List<string> translatedTextList)
        {
            string finalTranslation = null;
            finalTranslation = translatedTextList.GroupBy(x => x)
                .Where(group => group.Count() > 1)
                .Select(group => group.Key).FirstOrDefault();

            if (finalTranslation == null)
            {
                finalTranslation = translatedTextList.FirstOrDefault(); //ToDo: This section will choose the default translation. For now it will only take the first one as default.
            }

            return finalTranslation;
        }
        #endregion
    }
}
