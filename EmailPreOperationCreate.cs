/*================================================================================================================================

  This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.  

  THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
  INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  

  We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object 
  code form of the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks to market Your software 
  product in which the Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in which the 
  Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims 
  or lawsuits, including attorneys’ fees, that arise or result from the use or distribution of the Sample Code.

 =================================================================================================================================*/
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;

namespace IncomingInReplyToEmailQuarantine
{
    /// <summary>
    /// This plug-in evaluates incoming messages to determine if they are forwards or replies of emails that currently reside in Dynamics 365. If this email message being evaluated for automatic promotion into Dynamics 365 has had the recipients modifed to no longer include the origin email's "from" value, we will fire logic to remove the regarding value that inreplyto correlation will automatically append to the created email.
    /// </summary>
    /// <remarks>
    /// Register this plug-in on the Create message, email entity, and pre-operation.
    /// </remarks>

    public sealed class EmailPreOperationCreate : IPlugin
    {
        /// <summary>
        /// Execute method that is required by the IPlugin interface.
        /// </summary>
        /// <param name="serviceProvider">The service provider from which you can obtain the
        /// tracing service, plug-in execution context, organization service, and more.</param>
        public void Execute(IServiceProvider serviceProvider)
        {
            //Extract the tracing service for use in debugging sandboxed plug-ins.
            ITracingService tracingService =
                (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            // Obtain the organization service reference.
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.
                Entity entity = (Entity)context.InputParameters["Target"];

                // Verify that the target entity represents an email.
                // If not, this plug-in was not registered correctly.
                if (entity.LogicalName != "email")
                {
                    tracingService.Trace("Exiting plugin: Deployment error - Check IncomingInReplyToEmailQuarantine.EmailPreOperationCreate plugin registration.");
                    return;
                }
                try
                {
                    tracingService.Trace("Begining: IncomingInReplyToEmailQuarantine.EmailPreOperationCreate.Execute");
                    bool direction = false;
                    int emailstatus = 0;

                    // Check if Target entity is not null
                    Entity targetEmail = entity;
                    if (targetEmail == null)
                    {
                        tracingService.Trace("Exiting plugin: Target entity is empty.");
                        return;
                    }
                    // Check email direction
                    direction = ((bool)targetEmail["directioncode"]);
                    if (direction != false)
                    {
                        tracingService.Trace("Exiting plugin: Record is not an incoming email.");
                        return;
                    }
                    // Check if Email is inreplyto
                    if (!targetEmail.Contains("inreplyto"))
                    {
                        tracingService.Trace("Exiting plugin: Email does not have a valid inreplyto value.");
                        return;
                    }
                    // Check if Email is Received within D365
                    emailstatus = ((OptionSetValue)targetEmail["statuscode"]).Value;
                    tracingService.Trace("Email Status Code: " + emailstatus.ToString());
                    if (emailstatus != 4)
                    {
                        tracingService.Trace("Exiting plugin: Record does not have a status value of Received.");
                        return;
                    }
                    //Get email with the same messageid as the the inreplyto value (this is the email being replied to or forwarded)
                    tracingService.Trace("Getting email with the same messageid as the the inreplyto value.");
                    Entity originEmail = getEmailWithSameMessageidAsInreplyto(targetEmail["inreplyto"].ToString(), tracingService, service);
                    if (originEmail == null)
                    {
                        tracingService.Trace("Exiting plugin: Origin email with the same messageid as the the inreplyto value on evaluated message was not found.");
                        return;
                    }
                    //Check if the origin email has a Regarding value (if the origin email has no regarding value, inreplyto correlation will not set a value upon promotion)
                    if (!originEmail.Contains("regardingobjectid"))
                    {
                        tracingService.Trace("Exiting plugin: Regarding of origin email is null.");
                        return;
                    }
                    // Check if the From email address on the origin email exist in any of the target email activity parties.
                    tracingService.Trace("Checking if the From email address on the origin email exists within any of the target email activity parties.");
                    if (originEmail.Contains("from"))
                    {
                        EntityCollection from = (EntityCollection)originEmail["from"];
                        string originEmailFrom = from[0]["addressused"].ToString().ToLower();
                        tracingService.Trace("Origin email From value {0}", originEmailFrom);
                        if (targetEmail.Contains("to"))
                        {
                            tracingService.Trace("Checking target email To activity parties.");
                            if (isTargetFromOnOriginActivityParty(originEmailFrom, (EntityCollection)targetEmail["to"], tracingService))
                            {
                                tracingService.Trace("Exiting plugin: The From value of the origin email was in the TO of the reply or forward.");
                                return;
                            }
                        }
                        if (targetEmail.Contains("cc"))
                        {
                            tracingService.Trace("Checking target email CC activity parties.");
                            if (isTargetFromOnOriginActivityParty(originEmailFrom, (EntityCollection)targetEmail["cc"], tracingService))
                            {
                                tracingService.Trace("Exiting plugin: The From value of the origin email was in the CC of the reply or forward.");
                                return;
                            }
                        }
                        if (targetEmail.Contains("bcc"))
                        {
                            tracingService.Trace("Checking target email BCC activity parties.");
                            if (isTargetFromOnOriginActivityParty(originEmailFrom, (EntityCollection)targetEmail["bcc"], tracingService))
                            {
                                tracingService.Trace("Exiting plugin: The From value of the origin email was in the BCC of the reply or forward.");
                                return;
                            }
                        }
                        // Backup Origin Regarding
                        EntityReference targetRegardingObject = (EntityReference)targetEmail["regardingobjectid"];
                        // TODO: Example of incident EntityReference
                        // Add ***_originregarding to Email entity and rename below in code
                        // Check if Regarding is a Case
                        if (targetRegardingObject.LogicalName == "incident")
                        {
                            // store removed regardingobject values as Entity Reference
                            tracingService.Trace("Regarding Object is a Case");
                            targetEmail["pfe_originregarding"] = targetRegardingObject;
                            tracingService.Trace("Origin Regarding object backed up: {0}", targetRegardingObject.Id.ToString());
                        }
                        // TODO: Example of generic string to capture related entity
                        // Add ***_quarantinedregardingobjectid and ***_quarantinedregardingobjectlogicalname to Email entity and rename below in code
                        // store removed regardingobject values as string
                        targetEmail.Attributes.Add("pfe_quarantinedregardingobjectid", targetRegardingObject.Id.ToString());
                        targetEmail.Attributes.Add("pfe_quarantinedregardingobjectlogicalname", targetRegardingObject.LogicalName);
                        // Remove existing Regarding object value
                        targetEmail["regardingobjectid"] = null;
                        tracingService.Trace("Current Regarding object been removed, New reference is null.");
                        tracingService.Trace("Email quarantined {0}", targetEmail.Id.ToString());
                    }
                }
                catch (Exception ex)
                {
                    tracingService.Trace(ex.Message);
                }
                finally
                {
                    tracingService.Trace("Completed: IncomingInReplyToEmailQuarantine.EmailPreOperationCreate.Execute");
                }
            }
        }
        /// <summary>
        /// Get emails with the same messageid as the the inreplyto value
        /// </summary>
        /// <param name="inreplyto"></param>
        /// <param name="tracingService"></param>
        /// <param name="svc"></param>
        /// <returns></returns>
        private Entity getEmailWithSameMessageidAsInreplyto(string inreplyto, ITracingService tracingService, IOrganizationService svc)
        {
            tracingService.Trace("inreplyto: {0}", inreplyto);

            if (inreplyto == null)
            {
                return null;
            }
            QueryExpression query = new QueryExpression
            {
                EntityName = "email",
                ColumnSet = new ColumnSet("activityid", "messageid", "regardingobjectid", "from"),
                NoLock = true,
                Criteria = {
                        Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "messageid",
                                Operator = ConditionOperator.Equal,
                                Values = { inreplyto }
                            }
                        }
                    },
            };
            EntityCollection results = svc.RetrieveMultiple(query);
            return results?.Entities?.FirstOrDefault();
        }
        /// <summary>
        /// Check Activity Party collection for email address
        /// </summary>
        /// <param name="originEmailFrom"></param>
        /// <param name="activityparties"></param>
        /// <returns></returns>
        private static bool isTargetFromOnOriginActivityParty(string originEmailFrom, EntityCollection activityparties, ITracingService tracingService)
        {
            foreach (Entity activityparty in activityparties.Entities)
            {
                tracingService.Trace("Orgin Email addressused: {0}", activityparty["addressused"].ToString().ToLower());
                if (originEmailFrom == activityparty["addressused"].ToString().ToLower())
                {
                    return true;
                }
            }
            return false;
        }
    }
}