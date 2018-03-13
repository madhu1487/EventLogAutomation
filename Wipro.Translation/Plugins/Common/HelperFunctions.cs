using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace Wipro.Translation.Plugins.Common
{
    [KnownTypeAttribute(typeof(KeyValuePair<string, object>))]
    public class HelperFunctions
    {
        #region Public Subs / Functions

        public static ManyToManyRelationshipMetadata RetrieveManyToManyRelationshipMetadata(string logicalName, string schemaName, IOrganizationService organizationService)
        {
            ManyToManyRelationshipMetadata manyToManyRelationshipMetadata = null;

            EntityMetadata entityMetadata = RetrieveEntityMetadata(logicalName, EntityFilters.Relationships, organizationService);

            //Null validation
            if (entityMetadata != null)
            {
                //Get many to many relationship metadata
                manyToManyRelationshipMetadata = entityMetadata.ManyToManyRelationships.FirstOrDefault(item => !string.IsNullOrEmpty(item.SchemaName)
                                                                                                    && item.SchemaName.ToLowerInvariant().Equals(schemaName.ToLowerInvariant()));
            }

            return manyToManyRelationshipMetadata;
        }

        public static EntityCollection RetrieveManyToManyRelationship(EntityReference entityReference, string schemaName, IOrganizationService organizationService)
        {
            EntityCollection entityCollection = null;

            //Null validation
            if (entityReference != null)
            {
                //Get many to many metadata from entity metadata where schema name meets criteria
                ManyToManyRelationshipMetadata manyToManyRelationshipMetadata = RetrieveManyToManyRelationshipMetadata(entityReference.LogicalName, schemaName, organizationService);

                entityCollection = RetrieveManyToManyRelationship(entityReference, manyToManyRelationshipMetadata, organizationService);
            }

            return entityCollection;
        }

        public static EntityCollection RetrieveManyToManyRelationship(EntityReference entityRefernce, ManyToManyRelationshipMetadata manyToManyRelationshipMetadata, IOrganizationService organizationService)
        {
            EntityCollection entityCollection = null;

            try
            {
                entityCollection = RetrieveManyToManyRelationship(manyToManyRelationshipMetadata.Entity2LogicalName
                                                                , manyToManyRelationshipMetadata.Entity1IntersectAttribute
                                                                , entityRefernce
                                                                , manyToManyRelationshipMetadata.SchemaName
                                                                , organizationService
                                                                );
            }
            catch
            {
                FilterExpression filterExpression = new FilterExpression();
                filterExpression.AddCondition(new ConditionExpression(manyToManyRelationshipMetadata.Entity2IntersectAttribute, ConditionOperator.Equal, entityRefernce.Id));

                entityCollection = RetrieveManyToManyRelationship(manyToManyRelationshipMetadata.Entity1LogicalName
                                                                , manyToManyRelationshipMetadata.Entity1IntersectAttribute
                                                                , null
                                                                , manyToManyRelationshipMetadata.Entity2LogicalName
                                                                , manyToManyRelationshipMetadata.Entity2IntersectAttribute
                                                                , filterExpression
                                                                , manyToManyRelationshipMetadata.SchemaName
                                                                , organizationService
                                                                );
            }

            return entityCollection;
        }

        public static EntityCollection RetrieveManyToManyRelationship(string currentEntityLogicalName
                                                                    , string otherEntityPrimaryIdAttribute
                                                                    , Microsoft.Xrm.Sdk.EntityReference otherEntityReference
                                                                    , string relationshipSchemaName
                                                                    , IOrganizationService organizationService
                                                                    )
        {
            EntityCollection entityCollection = null;
            Microsoft.Xrm.Sdk.Relationship relationship = new Microsoft.Xrm.Sdk.Relationship(relationshipSchemaName);

            QueryExpression query = new QueryExpression();
            query.EntityName = currentEntityLogicalName;
            query.ColumnSet = new ColumnSet(true);

            RelationshipQueryCollection relatedEntity = new RelationshipQueryCollection();
            relatedEntity.Add(relationship, query);

            RetrieveRequest request = new RetrieveRequest();
            request.RelatedEntitiesQuery = relatedEntity;
            request.ColumnSet = new ColumnSet(otherEntityPrimaryIdAttribute);
            request.Target = otherEntityReference;

            RetrieveResponse response = (RetrieveResponse)organizationService.Execute(request);

            if (((DataCollection<Microsoft.Xrm.Sdk.Relationship, Microsoft.Xrm.Sdk.EntityCollection>)(((RelatedEntityCollection)(response.Entity.RelatedEntities)))).Contains(new Microsoft.Xrm.Sdk.Relationship(relationshipSchemaName))
                && ((DataCollection<Microsoft.Xrm.Sdk.Relationship, Microsoft.Xrm.Sdk.EntityCollection>)(((RelatedEntityCollection)(response.Entity.RelatedEntities))))[new Microsoft.Xrm.Sdk.Relationship(relationshipSchemaName)].Entities.Any())
            {
                DataCollection<Microsoft.Xrm.Sdk.Entity> results = ((DataCollection<Microsoft.Xrm.Sdk.Relationship, Microsoft.Xrm.Sdk.EntityCollection>)(((RelatedEntityCollection)(response.Entity.RelatedEntities))))[new Microsoft.Xrm.Sdk.Relationship(relationshipSchemaName)].Entities;

                if (results != null && results.Any())
                {
                    entityCollection = new EntityCollection();
                    entityCollection.Entities.AddRange(results);
                }
            }

            return entityCollection;
        }

        public static EntityCollection RetrieveManyToManyRelationship(string currentEntityLogicalName
                                                                    , string currentEntityPrimaryIdAttribute
                                                                    , FilterExpression currentEntityFilterExpression
                                                                    , string otherEntityLogicalName
                                                                    , string otherEntityPrimaryIdAttribute
                                                                    , FilterExpression otherEntityFilterExpression
                                                                    , string relationshipSchemaName
                                                                    , IOrganizationService organizationService
                                                                    )
        {
            LinkEntity linkCurrentEntity, linkOtherEntity;
            QueryExpression query = new QueryExpression(currentEntityLogicalName);
            query.ColumnSet = new ColumnSet(true);

            linkCurrentEntity = new LinkEntity(currentEntityLogicalName
                                              , relationshipSchemaName
                                              , currentEntityPrimaryIdAttribute
                                              , currentEntityPrimaryIdAttribute
                                              , JoinOperator.Inner);
            if (currentEntityFilterExpression != null) linkCurrentEntity.LinkCriteria = currentEntityFilterExpression;

            linkOtherEntity = new LinkEntity(relationshipSchemaName
                                            , otherEntityLogicalName
                                            , otherEntityPrimaryIdAttribute
                                            , otherEntityPrimaryIdAttribute
                                            , JoinOperator.Inner);
            if (otherEntityFilterExpression != null) linkOtherEntity.LinkCriteria = otherEntityFilterExpression;

            linkCurrentEntity.LinkEntities.Add(linkOtherEntity);

            query.LinkEntities.Add(linkCurrentEntity);

            EntityCollection entityCollection = organizationService.RetrieveMultiple(query);

            return entityCollection;
        }

        public static Boolean RelationshipExists(Microsoft.Xrm.Sdk.Metadata.ManyToManyRelationshipMetadata manyToManyRelationshipMetadata, EntityReference entity1Reference, EntityReference entity2Reference, IOrganizationService organizationService)
        {
            Boolean relationshipExists = false;

            //Null validation
            if (entity1Reference != null && entity2Reference != null)
            {
                try
                {
                    //Null validation
                    if (manyToManyRelationshipMetadata != null)
                    {
                        //Query N:N relationship where criteria matches
                        QueryExpression query = new QueryExpression(manyToManyRelationshipMetadata.IntersectEntityName)
                        {
                            ColumnSet = new Microsoft.Xrm.Sdk.Query.ColumnSet(false)
                        };
                        query.Criteria.AddCondition(manyToManyRelationshipMetadata.Entity1IntersectAttribute, ConditionOperator.Equal, entity1Reference.Id);
                        query.Criteria.AddCondition(manyToManyRelationshipMetadata.Entity2IntersectAttribute, ConditionOperator.Equal, entity2Reference.Id);

                        //Set realationship exists flag
                        relationshipExists = organizationService.RetrieveMultiple(query).Entities.Any();
                    }
                }
                catch
                {
                    relationshipExists = false;
                }
            }

            return relationshipExists;
        }

        /// <summary>
        /// Serialize Object
        /// </summary>
        /// <param name="serializeableObject">Serializeable Object</param>
        /// <remarks>
        /// </remarks>
        public static string SerializeObject<T>(T serializeableObject)
        {
            string serializedObject = string.Empty;

            if (!object.Equals(serializeableObject, default(T)))
            {
                try
                {
                    System.Runtime.Serialization.Json.DataContractJsonSerializer dataContractJsonSerializer = new System.Runtime.Serialization.Json.DataContractJsonSerializer(typeof(T));
                    System.IO.MemoryStream memoryStream = new System.IO.MemoryStream();

                    dataContractJsonSerializer.WriteObject(memoryStream, serializeableObject);
                    serializedObject = System.Text.Encoding.UTF8.GetString(memoryStream.ToArray());

                    memoryStream.Close();
                }
                catch { serializedObject = string.Empty; }
            }

            return serializedObject;
        }

        /// <summary>
        /// Deserialize Object
        /// </summary>
        /// <param name="json">Json String</param>
        /// <remarks>
        /// </remarks>
        public static T DeserializeObject<T>(string json)
        {
            T rtnObject = default(T);

            try
            {
                using (MemoryStream memoryStream = new MemoryStream(Encoding.Unicode.GetBytes(json)))
                {
                    System.Runtime.Serialization.Json.DataContractJsonSerializer dataContractJsonSerializer = new System.Runtime.Serialization.Json.DataContractJsonSerializer(typeof(T));
                    rtnObject = (T)dataContractJsonSerializer.ReadObject(memoryStream);
                }
            }
            catch { }

            return rtnObject;
        }

        /// <summary>
        /// Get Attribute Value from entity
        /// </summary>
        /// <param name="attributeLogicalName">Attribute Logical Name</param>
        /// <param name="entity">Entity to Retrieve attribute value</param>
        /// <remarks>
        /// Get attribte value from entity parameter
        /// </remarks>
        public static T GetAttributeValue<T>(string attributeLogicalName, Entity entity)
        {
            T rtnValue = default(T);

            if (!string.IsNullOrEmpty(attributeLogicalName))
            {
                attributeLogicalName = attributeLogicalName.ToLowerInvariant();
                try
                {
                    if (entity != null && entity.Contains(attributeLogicalName)) rtnValue = entity.GetAttributeValue<T>(attributeLogicalName);
                }
                catch { rtnValue = default(T); }
            }

            return rtnValue;
        }

        public static string GetSettingValue(string settingName, string defaultValue, IOrganizationService organizationService, EntityReference targetId = null, Boolean siteSetting = false)
        {
            string settingValue = string.Empty;

            if (siteSetting)
            {
                //Query site setting record that meets criteria
                QueryExpression query = new QueryExpression("adx_sitesetting");
                query.ColumnSet = new ColumnSet(new string[] { "adx_value" });
                query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                query.Criteria.AddCondition("adx_name", ConditionOperator.Equal, settingName);

                if (targetId != null) query.Criteria.AddCondition("adx_websiteid", ConditionOperator.Equal, targetId.Id);

                EntityCollection siteSettings = organizationService.RetrieveMultiple(query);

                //Null validation
                if (siteSettings != null && siteSettings.Entities.Any())
                    settingValue = GetAttributeValue<string>("adx_value", siteSettings.Entities.FirstOrDefault());
            }
            else
            {
                //Query site setting record that meets criteria
                QueryExpression query = new QueryExpression("essc_applicationsetting");
                query.ColumnSet = new ColumnSet(new string[] { "essc_value" });
                query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                query.Criteria.AddCondition("essc_name", ConditionOperator.Equal, settingName);

                if (targetId != null) query.Criteria.AddCondition("essc_applicationid", ConditionOperator.Equal, targetId.Id);

                EntityCollection applicationSettings = organizationService.RetrieveMultiple(query);

                //Null validation
                if (applicationSettings != null && applicationSettings.Entities.Any())
                    settingValue = GetAttributeValue<string>("essc_value", applicationSettings.Entities.FirstOrDefault());
            }

            //Set snippet value to default value if empty
            if (string.IsNullOrEmpty(settingValue)) settingValue = defaultValue;

            return settingValue;
        }

        public static WebHeaderCollection RetrieveApplicationWebHeaders(IOrganizationService organizationService)
        {
            WebHeaderCollection webHeaders = null;

            //Null validation
            if (organizationService != null)
            {
                //Retrieve application credentials that meet criteria
                QueryExpression query = new QueryExpression("essc_applicationcredential");
                query.ColumnSet = new ColumnSet(new string[] { "essc_applicationkey" });
                query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active

                FilterExpression filterExpression = new FilterExpression(LogicalOperator.Or);
                filterExpression.AddCondition("essc_name", ConditionOperator.Equal, "CRM Application");
                filterExpression.AddCondition("essc_name", ConditionOperator.Equal, "ADX Application");
                filterExpression.AddCondition("essc_name", ConditionOperator.Equal, "ESSC Application");

                query.Criteria.AddFilter(filterExpression);

                EntityCollection applicationCredentials = organizationService.RetrieveMultiple(query);

                //Null validation
                if (applicationCredentials != null && applicationCredentials.Entities.Any())
                {
                    Entity applicationCredential = applicationCredentials.Entities.FirstOrDefault();

                    webHeaders = new System.Net.WebHeaderCollection();
                    webHeaders.Add("X-HP-ESSC-ApplicationId", applicationCredential.Id.ToString());
                    webHeaders.Add("X-HP-ESSC-ApplicationKey", GetAttributeValue<string>("essc_applicationkey", applicationCredential));
                }
            }

            return webHeaders;
        }

        public static string ExecuteHttpWebRequest(string serialziedObject, string url, WebHeaderCollection headers = null, string method = default(string))
        {
            string httpResponse = string.Empty;

            //String empty validation
            if (!string.IsNullOrEmpty(url))
            {
                try
                {
                    if (headers == null) headers = new WebHeaderCollection();

                    //Default method if equal string empty
                    if (string.IsNullOrEmpty(method)) method = "POST";

                    headers.Add(HttpRequestHeader.Accept, "application/json");
                    headers.Add(HttpRequestHeader.ContentType, "application/json");

                    using (WebClient webClient = new WebClient())
                    {
                        webClient.Encoding = UTF8Encoding.UTF8;
                        webClient.Headers = headers;

                        if (method.ToLowerInvariant().Equals("get")) httpResponse = webClient.DownloadString(url);
                        else httpResponse = webClient.UploadString(url, method, serialziedObject);
                    }
                }
                catch (WebException exception)
                {
                    string response = string.Empty;
                    if (exception.Response != null)
                    {
                        using (StreamReader reader =
                            new StreamReader(exception.Response.GetResponseStream()))
                        {
                            response = reader.ReadToEnd();
                        }
                        exception.Response.Close();
                    }
                    if (exception.Status == WebExceptionStatus.Timeout)
                    {
                        throw new InvalidPluginExecutionException(
                            "The timeout elapsed while attempting to issue the request.", exception);
                    }
                    throw new InvalidPluginExecutionException(String.Format(System.Globalization.CultureInfo.InvariantCulture,
                        "A Web exception occurred while attempting to issue the request. {0}: {1}",
                        exception.Message, response), exception);
                }
            }

            return httpResponse;
        }

        public static OptionSetMetadata GetOptionSetMetadata(string attributeName, string entityLogicaName, IOrganizationService organizationService)
        {
            OptionSetMetadata optionSetMetadata = null;

            if (!string.IsNullOrEmpty(entityLogicaName))
            {
                RetrieveAttributeRequest retrieveAttributeRequest = new RetrieveAttributeRequest
                {
                    EntityLogicalName = entityLogicaName,
                    LogicalName = attributeName,
                    RetrieveAsIfPublished = true
                };
                RetrieveAttributeResponse retrieveAttributeResponse = (RetrieveAttributeResponse)organizationService.Execute(retrieveAttributeRequest);
                AttributeMetadata attributeMetadata = retrieveAttributeResponse.AttributeMetadata;

                if (attributeMetadata is StateAttributeMetadata) { optionSetMetadata = (attributeMetadata as StateAttributeMetadata).OptionSet; }
                else if (attributeMetadata is StatusAttributeMetadata) { optionSetMetadata = (attributeMetadata as StatusAttributeMetadata).OptionSet; }
                else if (attributeMetadata is PicklistAttributeMetadata) { optionSetMetadata = (attributeMetadata as PicklistAttributeMetadata).OptionSet; }
                else if (attributeMetadata is EntityNameAttributeMetadata) { optionSetMetadata = (attributeMetadata as EntityNameAttributeMetadata).OptionSet; }
            }
            else
            {
                RetrieveOptionSetRequest retrieveOptionSetRequest = new RetrieveOptionSetRequest() { Name = attributeName };
                RetrieveOptionSetResponse retrieveOptionSetResponse = (RetrieveOptionSetResponse)organizationService.Execute(retrieveOptionSetRequest);
                if (retrieveOptionSetResponse != null) optionSetMetadata = retrieveOptionSetResponse.OptionSetMetadata as OptionSetMetadata;
            }

            return optionSetMetadata;
        }

        public static string GetOptionSetLabel(OptionMetadata optionMetadata, int? languageCode)
        {
            string label = string.Empty;

            if (optionMetadata != null)
            {
                label = optionMetadata.Label.UserLocalizedLabel.Label;

                if (languageCode != null)
                {
                    LocalizedLabel localizedLabel = optionMetadata.Label.LocalizedLabels.FirstOrDefault(item => item.LanguageCode.Equals(languageCode));
                    if (localizedLabel != null) label = localizedLabel.Label;
                }
            }

            return label;
        }

        public static string GetAttributeFormattedValue(string attributeLogicalName, int? languageCode, Entity entity, AttributeMetadata attributeMetadata, IOrganizationService organizationService)
        {
            string formattedValue = string.Empty;

            //Null validation
            if (attributeMetadata != null)
            {
                object atrributeValue = null;

                switch (attributeMetadata.AttributeType.Value)
                {
                    case AttributeTypeCode.String:
                    case AttributeTypeCode.Memo:
                        formattedValue = GetAttributeValue<string>(attributeLogicalName, entity);
                        break;
                    case AttributeTypeCode.Uniqueidentifier:
                        atrributeValue = GetAttributeValue<Guid?>(attributeLogicalName, entity);
                        if (atrributeValue != null) formattedValue = ((Guid)atrributeValue).ToString();
                        break;
                    case AttributeTypeCode.DateTime:
                        atrributeValue = GetAttributeValue<DateTime?>(attributeLogicalName, entity);
                        if (atrributeValue != null) formattedValue = ((DateTime)atrributeValue).ToOADate().ToString();
                        break;
                    case AttributeTypeCode.Boolean:
                        atrributeValue = GetAttributeValue<Boolean?>(attributeLogicalName, entity);
                        if (atrributeValue != null) formattedValue = atrributeValue.ToString().ToLowerInvariant();
                        break;
                    case AttributeTypeCode.Customer:
                    case AttributeTypeCode.Owner:
                    case AttributeTypeCode.Lookup:
                        atrributeValue = GetAttributeValue<EntityReference>(attributeLogicalName, entity);
                        if (atrributeValue != null)
                        {
                            EntityReference entityReference = (EntityReference)atrributeValue;
                            if (entityReference != null)
                            {
                                formattedValue = entityReference.Name;

                                //Null validation
                                if (string.IsNullOrEmpty(formattedValue))
                                {
                                    //Get entity metadata
                                    EntityMetadata entyMetadata = RetrieveEntityMetadata(entityReference.LogicalName, EntityFilters.Entity, organizationService);

                                    //Null validation
                                    if (entyMetadata != null)
                                    {
                                        //Get entity related to entity
                                        Entity _entity = RetrieveEntity(entityReference, new string[] { entyMetadata.PrimaryNameAttribute }, organizationService);

                                        //Null validation
                                        if (_entity != null)
                                            formattedValue = GetAttributeValue<string>(entyMetadata.PrimaryNameAttribute, _entity);
                                    }
                                }
                            }
                        }
                        break;
                    case AttributeTypeCode.BigInt:
                        atrributeValue = GetAttributeValue<long?>(attributeLogicalName, entity);
                        if (atrributeValue != null) formattedValue = ((long)atrributeValue).ToString(CultureInfo.InvariantCulture);
                        break;
                    case AttributeTypeCode.Integer:
                        atrributeValue = GetAttributeValue<int?>(attributeLogicalName, entity);
                        if (atrributeValue != null) formattedValue = ((int)atrributeValue).ToString(CultureInfo.InvariantCulture);
                        break;
                    case AttributeTypeCode.Picklist:
                    case AttributeTypeCode.State:
                    case AttributeTypeCode.Status:
                        atrributeValue = GetAttributeValue<OptionSetValue>(attributeLogicalName, entity);

                        //Null validation
                        if (atrributeValue != null)
                        {
                            int selectedValue = ((OptionSetValue)atrributeValue).Value;
                            string optionLabel = string.Empty;
                            OptionSetMetadata optionSetMetadata = null;

                            //Check attribute metadata type
                            if (attributeMetadata is PicklistAttributeMetadata) optionSetMetadata = ((PicklistAttributeMetadata)attributeMetadata).OptionSet;
                            else if (attributeMetadata is StateAttributeMetadata) optionSetMetadata = ((StateAttributeMetadata)attributeMetadata).OptionSet;
                            else if (attributeMetadata is StatusAttributeMetadata) optionSetMetadata = ((StatusAttributeMetadata)attributeMetadata).OptionSet;

                            //Null validation
                            if (optionSetMetadata != null)
                            {
                                //Get option metadata from option set metadata by selected value
                                OptionMetadata optionMetadata = optionSetMetadata.Options.FirstOrDefault(item => item.Value.Value.Equals(selectedValue));

                                //Null validation
                                if (optionMetadata != null) formattedValue = GetDisplayName(optionMetadata.Label, languageCode);
                            }
                        }
                        break;
                    case AttributeTypeCode.Money:
                        atrributeValue = GetAttributeValue<Money>(attributeLogicalName, entity);
                        if (atrributeValue != null) formattedValue = ((Money)atrributeValue).Value.ToString(CultureInfo.InvariantCulture);
                        break;
                    case AttributeTypeCode.Decimal:
                        atrributeValue = GetAttributeValue<decimal?>(attributeLogicalName, entity);
                        if (atrributeValue != null) formattedValue = ((decimal)atrributeValue).ToString(CultureInfo.InvariantCulture);
                        break;
                    case AttributeTypeCode.Double:
                        atrributeValue = GetAttributeValue<double?>(attributeLogicalName, entity);
                        if (atrributeValue != null) formattedValue = ((double)atrributeValue).ToString(CultureInfo.InvariantCulture);
                        break;
                }
            }

            return formattedValue;
        }

        public static EntityKeyMetadata RetrieveEntityKeyMetadata(string entityLogicalName, string logicalName, Boolean retrieveAsIfPublished, IOrganizationService organizationService)
        {
            EntityKeyMetadata entityKeyMetadata = null;

            if (!string.IsNullOrEmpty(entityLogicalName) && !string.IsNullOrEmpty(logicalName))
            {
                try
                {
                    RetrieveEntityKeyRequest retrieveEntityKeyRequest = new RetrieveEntityKeyRequest()
                    {
                        EntityLogicalName = entityLogicalName,
                        LogicalName = logicalName,
                        MetadataId = Guid.Empty,
                        RetrieveAsIfPublished = retrieveAsIfPublished,
                    };
                    RetrieveEntityKeyResponse retrieveEntityKeyResponse = (RetrieveEntityKeyResponse)organizationService.Execute(retrieveEntityKeyRequest);

                    //Null validation
                    if (retrieveEntityKeyResponse != null) entityKeyMetadata = retrieveEntityKeyResponse.EntityKeyMetadata;
                }
                catch { }
            }

            return entityKeyMetadata;
        }

        public static EntityCollection RetrieveTranslationProvider(IOrganizationService service)
        {
            EntityCollection collection = null;

            QueryExpression query = new QueryExpression("wiptrans_translationprovider");
            query.ColumnSet = new ColumnSet(true);
            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);//Active
            query.Criteria.AddCondition("wiptrans_clientkeyname", ConditionOperator.NotNull);
            query.Criteria.AddCondition("wiptrans_clientkey", ConditionOperator.NotNull);
            query.Criteria.AddCondition("wiptrans_url", ConditionOperator.NotNull);

            collection = service.RetrieveMultiple(query);

            return collection;
        }

        public static EntityCollection RetrieveLanguage(IOrganizationService service, List<Int32> languages)
        {
            EntityCollection collection = null;
            QueryExpression query = new QueryExpression("languagelocale");
            query.ColumnSet = new ColumnSet(true);

            if (languages != null && languages.Any())
            {
                FilterExpression filter = new FilterExpression(LogicalOperator.Or);
                languages.ForEach(item => { filter.AddCondition("localeid", ConditionOperator.Equal, item); });
                query.Criteria.AddFilter(filter);
            }

            collection = service.RetrieveMultiple(query);

            return collection;
        }
        #endregion

        #region Private Subs / Functions

        /// <summary>
        /// Get entity
        /// </summary>
        /// <param name="organizationService">Organization Service</param>
        /// <param name="entityReference">Entity Reference</param>
        /// <remarks>
        /// Get entity record
        /// </remarks>
        private static Entity RetrieveEntity(Microsoft.Xrm.Sdk.EntityReference entityReference, IOrganizationService organizationService)
        {
            return RetrieveEntity(entityReference, null, organizationService);
        }

        /// <summary>
        /// Get entity
        /// </summary>
        /// <param name="organizationService">Organization Service</param>
        /// <param name="entityReference">Entity Reference</param>
        /// <param name="attributes">Attributes to return</param>
        /// <remarks>
        /// Get entity record
        /// </remarks>
        private static Entity RetrieveEntity(Microsoft.Xrm.Sdk.EntityReference entityReference, string[] attributes, IOrganizationService organizationService)
        {
            Entity entity = null;
            ColumnSet columnSet = new ColumnSet(true);

            if (attributes != null && attributes.Length > 0) columnSet = new ColumnSet(attributes);

            try
            {
                if (entityReference != null && !entityReference.Id.Equals(Guid.Empty))
                {
                    RetrieveResponse retrieveResponse = (RetrieveResponse)organizationService.Execute(new RetrieveRequest()
                    {
                        ColumnSet = columnSet,
                        Target = entityReference
                    });

                    if (retrieveResponse != null && retrieveResponse.Entity != null)
                        entity = retrieveResponse.Entity;
                }
            }
            catch { entity = null; }

            return entity;
        }

        /// <summary>
        /// Get Entity Metadata
        /// </summary>
        /// <param name="entityLogicalName">Entity Logical Name</param>
        /// <param name="entityFilter">Entity Filter</param>
        /// <param name="organizationService">Organization Service</param>
        /// <remarks>
        /// Return Entity Metadata 
        /// </remarks>
        private static EntityMetadata RetrieveEntityMetadata(string entityLogicalName, Microsoft.Xrm.Sdk.Metadata.EntityFilters entityFilter, IOrganizationService organizationService)
        {
            EntityMetadata entityMetadata = null;

            if (organizationService != null)
            {
                try
                {
                    OrganizationRequest request = new OrganizationRequest();
                    request.RequestName = "RetrieveEntity";
                    request["EntityFilters"] = entityFilter;
                    request["LogicalName"] = entityLogicalName;
                    request["MetadataId"] = Guid.Empty;
                    request["RetrieveAsIfPublished"] = false;

                    OrganizationResponse response = organizationService.Execute(request);
                    if (response != null) entityMetadata = (EntityMetadata)response["EntityMetadata"];
                }
                catch { entityMetadata = null; }
            }

            return entityMetadata;
        }

        //get display name for optionset
        private static string GetDisplayName(Label label, int? languageCode)
        {
            string displayName = string.Empty;

            //Null validation
            if (label != null)
            {
                LocalizedLabel localizedLabel = null;

                //Null validation
                if (languageCode != null)
                {
                    //Null validation
                    if (label.LocalizedLabels != null && label.LocalizedLabels.Any())
                        label.LocalizedLabels.FirstOrDefault(item => item.LanguageCode == languageCode); //Get localized label from language code
                }

                displayName = (localizedLabel != null)
                            ? localizedLabel.Label
                            : label.UserLocalizedLabel.Label;
            }

            return displayName;
        }


        #endregion
    }
}
