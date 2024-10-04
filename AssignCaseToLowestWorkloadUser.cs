using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Web.UI.WebControls;

namespace AssignCaseToLowestWorkloadUserPlugin
{
    public class AssignCaseToLowestWorkloadUser : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Get the context, tracing service, and organization service from the service provider
            //This line retrieves the execution context of the plugin.The IPluginExecutionContext contains information about the event that triggered the plugin
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            //This line retrieves a tracing service.
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            //This line gets the IOrganizationServiceFactory, which is responsible for creating instances of IOrganizationService. This service allows interaction with Dynamics 365 data such as retrieving, creating, updating, or deleting records.
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            //This line creates an instance of the IOrganizationService for the user who triggered the plugin 
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("AssignCaseToUserWithFewestCases Plugin execution started.");

            // Check if the Target parameter exists and is of type Entity
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                Entity caseEntity = (Entity)context.InputParameters["Target"];
                tracingService.Trace("Target entity is: {0}", caseEntity.LogicalName);

                // Ensure the entity is a case (incident)
                if (caseEntity.LogicalName != "incident")
                {
                    tracingService.Trace("Entity is not a case. Exiting plugin.");
                    return;
                }

                try
                {
                    tracingService.Trace("Retrieving all active users with the specified criteria.");

                    // FetchXML to retrieve users based on provided criteria (including positionid)
                    string fetchXml = @"
                <fetch version='1.0' output-format='xml-platform' mapping='logical' savedqueryid='da675923-da9d-4702-aad6-3207283a5ace' no-lock='false' distinct='true'>
                    <entity name='systemuser'>
                        <attribute name='entityimage_url'/>
                        <attribute name='parentsystemuserid'/>
                        <order attribute='fullname' descending='false'/>
                        <attribute name='title'/>
                        <attribute name='address1_telephone1'/>
                        <attribute name='businessunitid'/>
                        <attribute name='fullname'/>
                        <attribute name='systemuserid'/>
                        <attribute name='positionid'/>
                        <filter type='and'>
                            <condition attribute='isdisabled' operator='eq' value='0'/>
                            <condition attribute='accessmode' operator='eq' value='0'/>
                            <condition attribute='positionid' operator='eq' value='{4f0b53fa-cf74-ef11-ac20-6045bdc5d905}' uiname='Associate' uitype='position'/>
                        </filter>
                    </entity>
                </fetch>";

                    // Retrieve the users based on the FetchXML
                    EntityCollection users = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    tracingService.Trace("Total active users retrieved: {0}", users.Entities.Count);

                    if (users.Entities.Count > 0)
                    {
                        Entity leastBusyUser = null;
                        int minCaseCount = int.MaxValue;

                        // Loop through each user and count the number of cases assigned
                        foreach (Entity user in users.Entities)
                        {
                            Guid userId = user.Id;
                            tracingService.Trace("Checking cases for user: {0}", user["fullname"]);

                            // Query to retrieve active cases for each user
                            QueryExpression caseQuery = new QueryExpression("incident")
                            {
                                ColumnSet = new ColumnSet("incidentid"),
                                Criteria = new FilterExpression
                                {
                                    Conditions =
                                {
                                    new ConditionExpression("ownerid", ConditionOperator.Equal, userId),
                                    new ConditionExpression("statecode", ConditionOperator.Equal, 0) // Active cases
                                }
                                }
                            };

                            EntityCollection userCases = service.RetrieveMultiple(caseQuery);
                            int userCaseCount = userCases.Entities.Count;

                            tracingService.Trace("User {0} has {1} active cases.", user["fullname"], userCaseCount);

                            // Check if the current user has fewer cases
                            if (userCaseCount < minCaseCount)
                            {
                                leastBusyUser = user;
                                minCaseCount = userCaseCount;
                            }
                        }

                        if (leastBusyUser != null)
                        {
                            tracingService.Trace("Assigning case to user: {0} who has {1} cases.", leastBusyUser["fullname"], minCaseCount);

                            // Update the case with the new owner
                            Entity updatedCaseEntity = new Entity(caseEntity.LogicalName, caseEntity.Id);
                            updatedCaseEntity["ownerid"] = new EntityReference("systemuser", leastBusyUser.Id);
                            service.Update(updatedCaseEntity);
                        }
                        else
                        {
                            tracingService.Trace("No suitable user found to assign the case.");
                        }
                    }
                    else
                    {
                        tracingService.Trace("No active users found based on the FetchXML criteria.");
                    }
                }
                catch (Exception ex)
                {
                    tracingService.Trace("Error in AssignCaseToUserWithFewestCases: {0}", ex.ToString());
                    throw new InvalidPluginExecutionException("An error occurred in the AssignCaseToUserWithFewestCases plugin.", ex);
                }
            }
            else
            {
                tracingService.Trace("Target is not an entity.");
            }

            tracingService.Trace("AssignCaseToUserWithFewestCases Plugin execution ended.");
        }
    }
}