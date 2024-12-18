using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace CAF_Update_Company_and_Group
{
    public class UpdateCompanyGroup : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            tracingService.Trace("Code started");

            if (context.MessageName.ToLower() != "create" && context.MessageName.ToLower() != "update")
                return;
            tracingService.Trace("Message is create or update");

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity target)
            {
                tracingService.Trace("Target is entity");
                string logicalName = target.LogicalName;
                if (logicalName != "alletech_caf")
                    return;
                tracingService.Trace("Entity logical name is CAF");

                if (target.Attributes.Contains("alletech_name") && target["alletech_name"] != null)
                {
                    EntityReference companyLookup = target.GetAttributeValue<EntityReference>("alletech_name");

                    if (companyLookup != null && string.IsNullOrEmpty(companyLookup.Name))
                    {
                        tracingService.Trace("Company lookup name is empty, attempting to retrieve from CRM.");
                        Entity companyEntity = service.Retrieve(companyLookup.LogicalName, companyLookup.Id, new ColumnSet("name"));
                        if (companyEntity != null && companyEntity.Contains("name"))
                        {
                            companyLookup.Name = companyEntity["name"].ToString();
                            tracingService.Trace("Company lookup name retrieved from CRM: " + companyLookup.Name);
                        }

                    }
                    else
                    {
                        tracingService.Trace("Company lookup name already available: " + companyLookup.Name);
                    }

                    if (companyLookup != null && !string.IsNullOrEmpty(companyLookup.Name))
                    {
                        tracingService.Trace("Company lookup is not null and name is: " + companyLookup.Name);

                        QueryExpression query = new QueryExpression("spectra_company")
                        {
                            ColumnSet = new ColumnSet("spectra_name", "spectra_groupid", "spectra_unifyparentorgid"),
                            Criteria = new FilterExpression()
                        };
                        query.Criteria.AddCondition("spectra_name", ConditionOperator.Equal, companyLookup.Name);
                        query.AddOrder("createdon", OrderType.Descending);

                        EntityCollection companyResults = service.RetrieveMultiple(query);

                        if (companyResults.Entities.Count > 0)
                        {
                            tracingService.Trace("Matching spectra_company record(s) found.");
                            Entity company = companyResults.Entities[0];

                            if (company.Contains("spectra_unifyparentorgid") && company["spectra_unifyparentorgid"] != null)
                            {
                                tracingService.Trace("spectra_unifyparentorgid is not null");

                                Entity updateEntity = new Entity("alletech_caf");
                                updateEntity.Id = target.Id;
                                updateEntity["spectra_company"] = new EntityReference("spectra_company", company.Id);

                                if (company.Contains("spectra_groupid") && company["spectra_groupid"] != null)
                                {
                                    updateEntity["spectra_group"] = new EntityReference("spectra_group", company.GetAttributeValue<EntityReference>("spectra_groupid").Id);
                                }

                                service.Update(updateEntity);
                                tracingService.Trace("Updated CAF record with spectra_company and spectra_group.");
                            }
                            else
                            {
                                tracingService.Trace("spectra_unifyparentorgid is null, so no updates will be made.");
                            }
                        }
                        else
                        {
                            //throw new InvalidPluginExecutionException("No matching spectra_company record found");
                            tracingService.Trace("No matching spectra_company record found.");
                            SetNewCompanyAsDefault(service, tracingService, target.Id);
                        }
                    }
                    else
                    {
                        tracingService.Trace("Company lookup is null or name is empty, setting default company.");
                        SetNewCompanyAsDefault(service, tracingService, target.Id);
                    }
                }
                else
                {
                    tracingService.Trace("object reference issue");
                    SetNewCompanyAsDefault(service, tracingService, target.Id);
                }
            }
        }

        private void SetNewCompanyAsDefault(IOrganizationService service, ITracingService tracingService, Guid entityId)
        {
            tracingService.Trace("Inside SetNewCompanyAsDefault method");

            QueryExpression queryNewCompany = new QueryExpression("spectra_company");
            queryNewCompany.ColumnSet.AddColumns("spectra_name", "spectra_groupid");
            queryNewCompany.Criteria.AddCondition("spectra_name", ConditionOperator.Equal, "New Company");

            EntityCollection newCompanyResults = service.RetrieveMultiple(queryNewCompany);

            if (newCompanyResults.Entities.Count > 0)
            {
                tracingService.Trace("Default 'New Company' record found.");
                Entity newCompany = newCompanyResults.Entities[0];

                Entity updateEntity = new Entity("alletech_caf");
                updateEntity.Id = entityId;
                updateEntity["spectra_company"] = new EntityReference("spectra_company", newCompany.Id);

                if (newCompany.Contains("spectra_groupid") && newCompany["spectra_groupid"] != null)
                {
                    updateEntity["spectra_group"] = new EntityReference("spectra_group", newCompany.GetAttributeValue<EntityReference>("spectra_groupid").Id);
                }

                service.Update(updateEntity);
                tracingService.Trace("Updated CAF record with default New Company and spectra_group.");
               // throw new InvalidPluginExecutionException("set new company sucessfully...");
            }
            else
            {
                tracingService.Trace("No 'New Company' found to set as default.");
            }
        }
    }
}
