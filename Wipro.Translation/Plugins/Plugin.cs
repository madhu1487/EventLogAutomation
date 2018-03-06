namespace Wipro.Translation.Plugins
{
    using System;
    using System.Text;
    using System.Collections.ObjectModel;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.ServiceModel;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Messages;
    using Microsoft.Xrm.Sdk.Metadata;
    using Microsoft.Xrm.Sdk.Query;
    using System.Runtime.Serialization;

    /// <summary>
    /// Base class for all Plugins.
    /// </summary>    
    public abstract class Plugin : IPlugin
    {
        /// <summary>
        /// <remarks>Remember to change the "pluginName" to your plugin namespace for error handling naming</remarks>
        /// </summary>
        public class LocalPluginContext
        {
            #region Cosntructors

            private LocalPluginContext()
            {
            }

            internal LocalPluginContext(IServiceProvider serviceProvider)
            {
                if (serviceProvider == null) throw new ArgumentNullException("serviceProvider");

                // Obtain the execution context service from the service provider.
                this.PluginExecutionContext = serviceProvider.GetService(typeof(IPluginExecutionContext)) as IPluginExecutionContext;

                // Obtain the tracing service from the service provider.
                this.TracingService = serviceProvider.GetService(typeof(ITracingService)) as ITracingService;

                // Obtain the Organization Service factory service from the service provider
                IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));

                // Use the factory to generate the Organization Service.
                this.OrganizationService = factory.CreateOrganizationService(this.PluginExecutionContext.UserId);
            }

            #endregion

            #region Private Properties / Variables

            private const string pluginName = "HP.ESSC.Plugins";
            private StringBuilder messages = new StringBuilder();

            #endregion

            #region Protected Properties / Variables

            internal IServiceProvider ServiceProvider
            {
                get;

                private set;
            }

            internal IOrganizationService OrganizationService
            {
                get;

                private set;
            }

            internal IPluginExecutionContext PluginExecutionContext
            {
                get;

                private set;
            }

            internal ITracingService TracingService
            {
                get;

                private set;
            }

            #endregion

            #region Private Subs / Functions

            /// <summary>
            /// Trace Message
            /// </summary>
            /// <param name="message">Message to be traced</param>
            /// <remarks>
            /// Writes message to tracing service
            /// </remarks>
            private void Trace(string message)
            {
                if (string.IsNullOrWhiteSpace(message) || this.TracingService == null) return;
                if (this.PluginExecutionContext == null) this.TracingService.Trace(message);
                else
                {
                    this.TracingService.Trace(
                        "{0}, Correlation Id: {1}, Initiating User: {2}",
                        message,
                        this.PluginExecutionContext.CorrelationId,
                        this.PluginExecutionContext.InitiatingUserId);
                }
            }

            /// <summary>
            /// Get Entity Logical Name
            /// </summary>
            /// <remarks>
            /// Gets the target entity which invoked the plugin from the context input parameter
            /// </remarks>
            private string GetEntityLogicalName()
            {
                string entityLogicalName = string.Empty;

                if (PluginExecutionContext != null)
                {
                    EntityReference entityReference = null;

                    if (PluginExecutionContext.InputParameters.Contains("EntityMoniker")) entityReference = (EntityReference)PluginExecutionContext.InputParameters["EntityMoniker"];
                    if (entityReference == null && PluginExecutionContext.InputParameters.Contains("Target"))
                    {
                        if (PluginExecutionContext.InputParameters["Target"] is Entity) entityReference = ((Entity)PluginExecutionContext.InputParameters["Target"]).ToEntityReference();
                        else if (PluginExecutionContext.InputParameters["Target"] is EntityReference) entityReference = (EntityReference)PluginExecutionContext.InputParameters["Target"];
                    }
                    if (entityReference != null) entityLogicalName = entityReference.LogicalName;
                }

                return entityLogicalName;
            }

            /// <summary>
            /// Get message from inner exception
            /// </summary>
            /// <param name="exception">Exception to extract message from</param>
            /// <param name="includeStackTrace">Include stack trace messages (True/False)</param>
            /// <remarks>
            /// Reads inner exception message and stack trace message and creates a string builder message object
            /// </remarks>
            private StringBuilder GetInnerExceptionMessage(Exception exception, Boolean includeStackTrace)
            {
                return GetInnerExceptionMessage(exception, false, includeStackTrace);
            }

            /// <summary>
            /// Get message from inner exception
            /// </summary>
            /// <param name="exception">Exception to extract message from</param>
            /// <param name="isInnerException">Inner exception? (True/False)</param>
            /// <param name="includeStackTrace">Include stack trace messages (True/False)</param>
            /// <remarks>
            /// Reads inner exception message and stack trace message and creates a string builder message object
            /// </remarks>
            private StringBuilder GetInnerExceptionMessage(Exception exception, Boolean isInnerException, Boolean includeStackTrace)
            {
                StringBuilder message = new StringBuilder();

                if (exception != null && exception.InnerException != null)
                {
                    Exception innerException = exception.InnerException;
                    if (innerException != null)
                    {
                        string errMessage = !string.IsNullOrEmpty(innerException.Message) ? innerException.Message : string.Empty;
                        if (!string.IsNullOrEmpty(errMessage))
                        {
                            if (!isInnerException) { message.AppendLine("Inner Exception Message:"); }
                            message.AppendLine(string.Format("Error Message: {0}", errMessage));
                            if (includeStackTrace) message.AppendLine(string.Format("Stack Trace: {0}", innerException.StackTrace));
                        }

                        if (innerException.InnerException != null)
                        {
                            string msg = GetInnerExceptionMessage(innerException.InnerException, includeStackTrace).ToString();
                            if (!string.IsNullOrEmpty(msg)) { message.AppendLine(msg); }
                        }
                    }
                }

                return message;
            }

            #endregion

            #region Protected Subs / Functions

            /// <summary>
            /// Logs message
            /// </summary>
            /// <param name="message">Message to be logged</param>
            /// <remarks>
            /// Logs custom message to the tracing service
            /// </remarks>
            internal void LogMessage(string message)
            {
                LogMessage(message, false);
            }

            /// <summary>
            /// Logs message
            /// </summary>
            /// <param name="message">Message to be logged</param>
            /// <param name="throwException">Throw exception after message is logged</param>
            /// <remarks>
            /// Logs custom message to the tracing service
            /// </remarks>
            internal void LogMessage(string message, Boolean throwException)
            {
                this.messages.AppendLine(message);

                Trace(message);

                if (throwException) ThrowException(message);
            }

            /// <summary>
            /// Logs error
            /// </summary>
            /// <param name="exception">Exception to be logged</param>
            /// <param name="title">Title giving to the exception</param>
            /// <remarks>
            /// Logs custom message to the tracing service
            /// </remarks>
            internal void LogError(Exception exception, string title)
            {
                LogError(exception, title, false);
            }

            /// <summary>
            /// Logs error
            /// </summary>
            /// <param name="exception">Exception to be logged</param>
            /// <param name="title">Title giving to the exception</param>
            /// <param name="includeStackTrace">Include stack trace?</param>
            /// <remarks>
            /// Logs custom message to the tracing service
            /// </remarks>
            internal void LogError(Exception exception, string title, Boolean includeStackTrace)
            {
                StringBuilder message = GetMessageFromException(exception, includeStackTrace);
                Trace(string.Format("Plug-in error: {0}", title));
                Trace(string.Format("Error details: {0}", message.ToString()));
            }

            /// <summary>
            /// Get message from exception
            /// </summary>
            /// <param name="exception">Exception to extract message from</param>
            /// <param name="title">Title giving to the message</param>
            /// <param name="includeStackTrace">Include stack trace messages (True/False)</param>
            /// <param name="localPluginCotext">Local Plugin Context</param>
            /// <remarks>
            /// Reads exception message and stack trace message and creates a string builder message object
            /// </remarks>
            internal StringBuilder GetMessageFromException(Exception exception, Boolean includeStackTrace)
            {
                StringBuilder message = new StringBuilder();

                if (exception != null)
                {
                    message.AppendLine("An error occurred");
                    message.AppendLine(string.Format("Message: {0}", exception.Message));
                    if (includeStackTrace) message.AppendLine(string.Format("Stack Trace: {0}", exception.StackTrace));

                    string msg = GetInnerExceptionMessage(exception, includeStackTrace).ToString();
                    if (!string.IsNullOrEmpty(msg)) { message.AppendLine(msg); }
                }

                return message;
            }

            /// <summary>
            /// Throws Plugin Exception
            /// </summary>
            /// <param name="message">Message to be thrown</param>
            /// <remarks>
            /// Throws a new instanace of InvalidPluginExecutionException
            /// </remarks>
            internal void ThrowException(string message)
            {
                ThrowException(message, new Exception("Error thrown by application"));
            }

            /// <summary>
            /// Throws Plugin Exception
            /// </summary>
            /// <param name="title">Title giving to the exception</param>
            /// <param name="message">Message to be thrown</param>
            /// <param name="exception">Exception thrown</param>
            /// <remarks>
            /// Throws a new instanace of InvalidPluginExecutionException
            /// </remarks>
            internal void ThrowException(string message, Exception exception, Boolean includeStackTrace = false)
            {
                string messageName = (PluginExecutionContext != null) ? PluginExecutionContext.MessageName : string.Empty;

                StringBuilder messages = GetMessageFromException(exception, includeStackTrace);
                messages.AppendLine("Logged Message");
                messages.AppendLine(this.messages.ToString());

                throw new InvalidPluginExecutionException(string.Format("{0} ({1}) Err Msg: {2}", message, messageName, messages.ToString()), exception);
            }

            #endregion

            #region Public Subs / Functions  

            /// <summary>
            /// Get Attribute Value
            /// </summary>
            /// <param name="attributeLogicalName">Attribute Logical Name</param>
            /// <remarks>
            /// Get attribte value from entity input parameter or from plugin images
            /// </remarks>
            public T GetAttributeValue<T>(string attributeLogicalName)
            {
                T rtnValue = default(T);

                try
                {
                    string entityLogicalName = GetEntityLogicalName();
                    Entity targetEntity = null
                           , targetEntityPrePost = null;

                    if (PluginExecutionContext != null)
                    {
                        //Get entity from input parameter bag
                        targetEntity = (PluginExecutionContext.InputParameters.Contains("Target"))
                                     ? PluginExecutionContext.InputParameters["Target"] as Entity
                                     : null;

                        if (!string.IsNullOrEmpty(entityLogicalName))
                        {
                            //Get Pre/Post entity from Pre/Post entity images
                            if (PluginExecutionContext.PreEntityImages.Contains(entityLogicalName)) { targetEntityPrePost = PluginExecutionContext.PreEntityImages[entityLogicalName]; }
                            if (targetEntityPrePost == null && PluginExecutionContext.PostEntityImages.Contains(entityLogicalName)) { targetEntityPrePost = PluginExecutionContext.PostEntityImages[entityLogicalName]; }
                        }
                    }

                    //Validate is attribute is in target entity. Get attribute value from entity
                    rtnValue = Plugin.GetAttributeValue<T>(attributeLogicalName, targetEntity);

                    //Attribute is not in target entity. Search Pre/Post entity for attribute value
                    if (targetEntityPrePost != null && object.Equals(rtnValue, default(T)))
                        rtnValue = Plugin.GetAttributeValue<T>(attributeLogicalName, targetEntityPrePost);
                }
                catch (Exception ex)
                {
                    //if error return default object
                    rtnValue = default(T);

                    //Log Error
                    LogError(ex, "LocalPluginContext.GetAttributeValue");
                }

                return rtnValue;
            }

            #endregion
        }

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="Plugin"/> class.
        /// </summary>
        /// <param name="childClassName">The <see cref=" cred="Type"/> of the derived class.</param>
        internal Plugin(Type childClassName)
        {
            this.ChildClassName = childClassName.ToString();
        }

        #endregion

        #region Private Properties / Variables

        private Collection<Tuple<SdkMessageProcessingStepStage, string, string, Action<LocalPluginContext>>> registeredEvents;

        #endregion

        #region Protected Properties / Variables

        /// <summary>
        /// Gets the List of events that the plug-in should fire for. Each List
        /// Item is a <see cref="System.Tuple"/> containing the Pipeline Stage, Message and (optionally) the Primary Entity. 
        /// In addition, the fourth parameter provide the delegate to invoke on a matching registration.
        /// </summary>
        protected Collection<Tuple<SdkMessageProcessingStepStage, string, string, Action<LocalPluginContext>>> RegisteredEvents
        {
            get
            {
                if (this.registeredEvents == null)
                {
                    this.registeredEvents = new Collection<Tuple<SdkMessageProcessingStepStage, string, string, Action<LocalPluginContext>>>();
                }

                return this.registeredEvents;
            }
        }

        /// <summary>
        /// Gets or sets the name of the child class.
        /// </summary>
        /// <value>The name of the child class.</value>
        protected string ChildClassName
        {
            get;

            private set;
        }

        protected enum SdkMessageProcessingStepStage
        {
            InitialPreoperation_Forinternaluseonly = 5,
            Prevalidation = 10,
            Preoperation = 20,
            Postoperation = 40,
            Postoperation_Deprecated = 50,
            InternalPostoperationAfterExternalPlugins_Forinternaluseonly = 45,
            InternalPostoperationBeforeExternalPlugins_Forinternaluseonly = 35,
            InternalPreoperationBeforeExternalPlugins_Forinternaluseonly = 15,
            InternalPreoperationAfterExternalPlugins_Forinternaluseonly = 25,
            MainOperation_Forinternaluseonly = 30,
            FinalPostoperation_Forinternaluseonly = 55,
        }

        #endregion

        #region Public Subs / Functions

        /// <summary>
        /// Executes the plug-in.
        /// </summary>
        /// <param name="serviceProvider">The service provider.</param>
        /// <remarks>
        /// For improved performance, Microsoft Dynamics CRM caches plug-in instances. 
        /// The plug-in's Execute method should be written to be stateless as the constructor 
        /// is not called for every invocation of the plug-in. Also, multiple system threads 
        /// could execute the plug-in at the same time. All per invocation state information 
        /// is stored in the context. This means that you should not use global variables in plug-ins.
        /// </remarks>
        public void Execute(IServiceProvider serviceProvider)
        {
            if (serviceProvider == null)
            {
                throw new ArgumentNullException("serviceProvider");
            }

            // Construct the Local plug-in context.
            LocalPluginContext localContext = new LocalPluginContext(serviceProvider);

            localContext.LogMessage(string.Format(CultureInfo.InvariantCulture, "Entered {0}.Execute()", this.ChildClassName));

            try
            {
                // Iterate over all of the expected registered events to ensure that the plugin
                // has been invoked by an expected event
                // For any given plug-in event at an instance in time, we would expect at most 1 result to match.
                Action<LocalPluginContext> entityAction =
                    (from a in this.RegisteredEvents
                     where (a.Item1.GetHashCode() == localContext.PluginExecutionContext.Stage &&
                            a.Item2.ToLowerInvariant() == localContext.PluginExecutionContext.MessageName.ToLowerInvariant() &&
                            (string.IsNullOrEmpty(a.Item3)
                            ? true
                            : a.Item3.ToLowerInvariant() == localContext.PluginExecutionContext.PrimaryEntityName.ToLowerInvariant())
                     )
                     select a.Item4).FirstOrDefault();

                if (entityAction != null)
                {
                    localContext.LogMessage(string.Format(
                        CultureInfo.InvariantCulture,
                        "{0} is firing for Entity: {1}, Message: {2}",
                        this.ChildClassName,
                        localContext.PluginExecutionContext.PrimaryEntityName,
                        localContext.PluginExecutionContext.MessageName));

                    entityAction.Invoke(localContext);

                    // now exit - if the derived plug-in has incorrectly registered overlapping event registrations,
                    // guard against multiple executions.
                    return;
                }
                else localContext.ThrowException(string.Format("Missing Entity Action. Stage: {0} | Message Name: {1} | Primary Entity Name: {2}", localContext.PluginExecutionContext.Stage, localContext.PluginExecutionContext.MessageName, localContext.PluginExecutionContext.PrimaryEntityName));
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                localContext.LogMessage(string.Format(CultureInfo.InvariantCulture, "Exception: {0}", e.ToString()));

                // Handle the exception.
                throw;
            }
            finally
            {
                localContext.LogMessage(string.Format(CultureInfo.InvariantCulture, "Exiting {0}.Execute()", this.ChildClassName));
            }
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

        #endregion

        #region Protected Subs / Functions

        /// <summary>
        /// Update Entity
        /// </summary>
        /// <param name="attritbutes">Attributes To Be Updated</param>
        /// <param name="primaryIdAttribute">Entity Primiary Id Name</param>
        /// <param name="entity">Reference Entity</param>
        /// <param name="localPluginContext">Local Plugin Context</param>
        /// <remarks>
        /// Update entity record
        /// </remarks>
        protected void UpdateEntity(Dictionary<string, object> attritbutes, ref Entity entity, LocalPluginContext localPluginContext)
        {
            if (localPluginContext != null && entity != null)
            {
                if (attritbutes != null && attritbutes.Any())
                {
                    Boolean updating = (localPluginContext.PluginExecutionContext.Stage == SdkMessageProcessingStepStage.Postoperation.GetHashCode())
                            , entityValid = (RetrieveEntity(entity.ToEntityReference(), localPluginContext.OrganizationService) != null);

                    Entity updateEntity = new Entity(entity.LogicalName);

                    if (!entityValid) updating = false;

                    foreach (KeyValuePair<string, object> item in attritbutes)
                    {
                        string attributeName = item.Key;
                        object attribute = item.Value;

                        if (!updating)
                        {
                            if (entity.Contains(attributeName))
                                entity[attributeName] = attribute;
                            else
                                entity.Attributes.Add(attributeName, attribute);
                        }
                        else
                            updateEntity.Attributes.Add(attributeName, attribute);
                    }

                    if (updating)
                    {
                        EntityMetadata entityMetadata = RetrieveEntityMetadata(entity.LogicalName, EntityFilters.Entity, localPluginContext.OrganizationService);

                        if (entityMetadata != null) updateEntity.Attributes.Add(entityMetadata.PrimaryIdAttribute, entity.Id);

                        localPluginContext.OrganizationService.Update(updateEntity);
                    }
                }
            }
        }

        /// <summary>
        /// Get entity
        /// </summary>
        /// <param name="organizationService">Organization Service</param>
        /// <param name="entityReference">Entity Reference</param>
        /// <remarks>
        /// Get entity record
        /// </remarks>
        protected static Entity RetrieveEntity(Microsoft.Xrm.Sdk.EntityReference entityReference, IOrganizationService organizationService)
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
        protected static Entity RetrieveEntity(Microsoft.Xrm.Sdk.EntityReference entityReference, string[] attributes, IOrganizationService organizationService)
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
        /// <param name="organizationService">Organization Service</param>
        /// <remarks>
        /// Return Entity Metadata 
        /// </remarks>
        protected static EntityMetadata RetrieveEntityMetadata(string entityLogicalName, IOrganizationService organizationService)
        {
            return RetrieveEntityMetadata(entityLogicalName
                                    , Microsoft.Xrm.Sdk.Metadata.EntityFilters.Entity | Microsoft.Xrm.Sdk.Metadata.EntityFilters.Attributes | Microsoft.Xrm.Sdk.Metadata.EntityFilters.Relationships
                                    , organizationService);
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
        protected static EntityMetadata RetrieveEntityMetadata(string entityLogicalName, Microsoft.Xrm.Sdk.Metadata.EntityFilters entityFilter, IOrganizationService organizationService)
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
                    request["RetrieveAsIfPublished"] = true;

                    OrganizationResponse response = organizationService.Execute(request);
                    if (response != null) entityMetadata = (EntityMetadata)response["EntityMetadata"];
                }
                catch { entityMetadata = null; }
            }

            return entityMetadata;
        }

        protected T GetInputParameter<T>(string name, IPluginExecutionContext pluginExecutionContext)
        {
            T inputParameter = default(T);

            if (pluginExecutionContext != null)
            {
                if (pluginExecutionContext.InputParameters != null && pluginExecutionContext.InputParameters.Any())
                {
                    inputParameter = pluginExecutionContext.InputParameters.Where(item => item.Key.ToLowerInvariant().Equals(name.ToLowerInvariant()))
                                                                           .Select(item => (T)item.Value)
                                                                           .FirstOrDefault();
                }
            }

            return inputParameter;
        }

        #endregion
    }
}
