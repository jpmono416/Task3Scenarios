using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Task3_Scenario2_SingleActiveCasePlugin
{
    /// <summary>
    /// A plugin that checks whether an active case exists for a given customer
    /// before creating a new case in Dynamics 365.
    /// </summary>
    public class SingleActiveCasePerAccountPlugin : IPlugin
    {
        /// <summary>
        /// Main entry point for the plugin. Checks if the target entity is valid,
        /// attempts to retrieve the customer, and verifies if an active case is
        /// already open for that customer.
        /// </summary>
        /// <param name="serviceProvider">The service provider from the platform.</param>
        /// <remarks>
        /// Throws an InvalidPluginExecutionException if an active case is found or if
        /// any unexpected error occurs during execution.
        /// </remarks>
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            try
            {
                if (!ValidateEntities(context, out Entity incident, out EntityReference customerRef))
                {
                    tracingService.Trace("Please ensure your context, case and contact are present");
                    return;
                }

                IOrganizationService service = CreateOrgService(serviceProvider, context);
                tracingService.Trace("Retrieving all active cases for customer {0}.", customerRef.Id);

                if (HasActiveCaseForCustomer(service, customerRef.Id))
                {
                    throw new InvalidPluginExecutionException(
                        "An active case already exists for this customer. " +
                        "Please resolve the active case before creating a new one."
                    );
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
                throw new InvalidPluginExecutionException(
                    $"An error occurred in the Case creation plugin. {ex.Message}",
                    ex
                );
            }
        }

        /// <summary>
        /// Validates the plugin context, retrieves the incident entity, and attempts to get the customer reference.
        /// </summary>
        /// <param name="context"></param>
        /// <param name="incident"></param>
        /// <param name="customerRef"></param>
        /// <returns>Whether both operations were successful</returns>
        private bool ValidateEntities(
            IPluginExecutionContext context, 
            out Entity incident, 
            out EntityReference customerRef)
        {
            incident = null;
            customerRef = null;

            return IsCreateOperationValid(context, out incident) && TryGetCustomer(incident, out customerRef);
        }

        /// <summary>
        /// Validates if the plugin context is a Create operation targeting the incident entity.
        /// </summary>
        /// <param name="context">Plugin execution context from the platform.</param>
        /// <param name="incident">The incident entity if valid; otherwise null.</param>
        /// <returns>True if valid, otherwise false.</returns>
        private bool IsCreateOperationValid(IPluginExecutionContext context, out Entity incident)
        {
            incident = null;
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity target &&
                target.LogicalName == "incident")
            {
                incident = target;
                return true;
            }
            return false;
        }

        /// <summary>
        /// Attempts to retrieve the customer EntityReference from a given incident record.
        /// </summary>
        /// <param name="incident">The incident record to inspect.</param>
        /// <param name="customerRef">The Customer EntityReference if found; otherwise null.</param>
        /// <returns>True if the customer was retrieved successfully, otherwise false.</returns>
        private bool TryGetCustomer(Entity incident, out EntityReference customerRef)
        {
            customerRef = null;
            if (incident.Attributes.TryGetValue("customerid", out object customerObj) &&
                customerObj is EntityReference reference)
            {
                customerRef = reference;
                return true;
            }
            return false;
        }

        /// <summary>
        /// Creates an IOrganizationService for the current user.
        /// </summary>
        /// <param name="serviceProvider">The service provider from the platform.</param>
        /// <param name="context">The plugin execution context.</param>
        /// <returns>An IOrganizationService instance.</returns>
        private IOrganizationService CreateOrgService(IServiceProvider serviceProvider, IPluginExecutionContext context)
        {
            var factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            return factory.CreateOrganizationService(context.UserId);
        }

        /// <summary>
        /// Checks whether there are any active cases for the specified customer.
        /// </summary>
        /// <param name="service">An instance of IOrganizationService for data query.</param>
        /// <param name="tracingService">Tracing service for logging.</param>
        /// <param name="customerId">ID of the customer to query for active cases.</param>
        /// <returns>True if an active case exists, otherwise false.</returns>
        private bool HasActiveCaseForCustomer(IOrganizationService service, Guid customerId)
        {
            EntityCollection existingCases = service.RetrieveMultiple(CreateActiveCaseQuery(customerId));
            return existingCases.Entities.Count > 0;
        }

        /// <summary>
        /// Builds a query for retrieving active cases (statecode = 0) for a specific customer.
        /// </summary>
        /// <param name="customerId">ID of the customer to query for active cases.</param>
        /// <returns>A QueryExpression configured to find active cases for the given customer.</returns>
        private QueryExpression CreateActiveCaseQuery(Guid customerId)
        {
            return new QueryExpression("incident")
            {
                ColumnSet = new ColumnSet("incidentid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("customerid", ConditionOperator.Equal, customerId),
                        new ConditionExpression("statecode", ConditionOperator.Equal, 0) // Active
                    }
                }
            };
        }
    }
}
