using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Task3_Scenario2_SingleActiveCasePlugin
{

    public class Plugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            try
            {
                // Verify that the context is for the Create message and the primary entity is 'incident'
                if (context.MessageName != "Create" || context.PrimaryEntityName != "incident")
                {
                    tracingService.Trace("The plugin cannot work with this context.");
                    return;
                }

                // Get the target entity from the input parameters
                if (!(context.InputParameters["Target"] is Entity incident))
                {
                    tracingService.Trace("The target entity is not available.");
                    return;
                }

                // Verify that the 'customerid' attribute is present
                if (!incident.Attributes.TryGetValue("customerid", out object customerObj) || !(customerObj is EntityReference customerRef))
                {
                    tracingService.Trace("The customer is not available.");
                    return;
                }

                // Obtain the organization service reference
                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                // Query for existing active cases for the same customer
                QueryExpression query = new QueryExpression("incident")
                {
                    ColumnSet = new ColumnSet("incidentid"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression("customerid", ConditionOperator.Equal, customerRef.Id),
                            new ConditionExpression("statecode", ConditionOperator.Equal, 0) // Active state
                        }
                    }
                };

                tracingService.Trace("Retrieving all active cases for customer {0}.", customerRef.Id);
                EntityCollection existingCases = service.RetrieveMultiple(query);

                // If there is at least one active case, prevent the creation of a new case
                if (existingCases.Entities.Count > 0)
                {
                    throw new InvalidPluginExecutionException("An active case already exists for this customer. Please resolve the active case before creating a new one.");
                }
                tracingService.Trace("No active cases found for the customer. Creating case.");
            }
            catch (InvalidPluginExecutionException ex)
            {
                tracingService.Trace(ex.Message);
                throw;
            }
            catch (Exception ex)
            {
                tracingService.Trace("An error occurred in the Case creation plugin: {0}", ex.ToString());
                throw new InvalidPluginExecutionException($"An error occurred in the Case creation plugin. {ex.Message}", ex);
            }
        }
    }
}
