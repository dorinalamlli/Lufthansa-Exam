using Lufthansa.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Data;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Lufthansa.Plugins
{
    // plugin on update of Start Date and End Date at Agreement (post operation, async)
    // step on create of Agreement (pre operation, sync)
    public class Agreement : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            bool tcValue = false;
            if (context.MessageName.ToLower() == "create")
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    Entity target = (Entity)context.InputParameters["Target"];
                    int agreementType = target.Contains("lhind_agreementtype") ? target.GetAttributeValue<OptionSetValue>("lhind_agreementtype").Value : -1;
                    EntityReference accountER = target.Contains("lhind_relatedaccount") ? target.GetAttributeValue<EntityReference>("lhind_relatedaccount") : null;

                    if (accountER != null)
                    {
                        if ((agreementType == (int)Enums.AgreementType.Onboarding) || (agreementType == (int)Enums.AgreementType.NDA))
                        {
                            string fetchXML = $@"
                            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='lhind_agreement'>
                                    <attribute name='lhind_agreementid' />
                                    <attribute name='lhind_name' />
                                    <attribute name='createdon' />
                                    <order attribute='lhind_name' descending='false' />
                                    <filter type='and'>
                                        <condition attribute='lhind_relatedaccount' operator='eq' value='{accountER.Id}' />
                                        <condition attribute='lhind_agreementtype' operator='eq' value='{agreementType}' />
                                    </filter>
                                </entity>
                            </fetch>";

                            FetchExpression fetchExpression = new FetchExpression(fetchXML);

                            EntityCollection results = service.RetrieveMultiple(fetchExpression);

                            if (results.Entities.Count > 0)
                            {
                                throw new InvalidPluginExecutionException("You can't create record.");
                            }
                        }
                    }

                    ManageLCOnAgreementCreation(target, service);
                }
            }
            else if (context.MessageName.ToLower() == "update")
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    Entity target = (Entity)context.InputParameters["Target"];
                    Entity postImage = (Entity)context.PostEntityImages["lhind_agreement"];

                    ManageLCOnAgreementUpdate(target, postImage, service);
                }
            }

        }

        public void ManageLCOnAgreementCreation(Entity target, IOrganizationService service)
        {
            bool tcValue = false;
            if
            (
                (target.Contains("lhind_agreementstartdate") && (target.GetAttributeValue<DateTime>("lhind_agreementstartdate") != DateTime.MinValue))
                &&
                ((target.Contains("lhind_agreementenddate")) && (target.GetAttributeValue<DateTime>("lhind_agreementenddate") != DateTime.MinValue))
            )
            {
                tcValue = true;
            }

            EntityReference accountER = target.Contains("lhind_relatedaccount") ? target.GetAttributeValue<EntityReference>("lhind_relatedaccount") : null;
            if (accountER != null)
            {
                string fetchXML = $@"
                        <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='opportunity'>
                                <attribute name='opportunityid'/>
                                <filter type='and'>
                                    <condition attribute='parentaccountid' operator='eq' value='{accountER.Id}'/>
                                </filter>
                            </entity>
                        </fetch>";

                FetchExpression fetchExpression = new FetchExpression(fetchXML);

                EntityCollection results = service.RetrieveMultiple(fetchExpression);

                foreach (Entity opportunity in results.Entities)
                {
                    Entity opportunityToUpdate = new Entity(opportunity.LogicalName, opportunity.Id);
                    opportunityToUpdate["lhind_tcs"] = tcValue;
                    service.Update(opportunityToUpdate);
                }
            }
        }

        public void ManageLCOnAgreementUpdate(Entity target, Entity postImage, IOrganizationService service)
        {
            bool tcValue = false;

            if (
                ((postImage.Contains("lhind_agreementtype")) && (postImage.GetAttributeValue<OptionSetValue>("lhind_agreementtype").Value == (int)Enums.AgreementType.Onboarding))
                &&
                ((target.Contains("lhind_agreementstartdate")) || (target.Contains("lhind_agreementenddate")))
               )
            {
                if (
                    (
                     (target.Contains("lhind_agreementstartdate") && (target.GetAttributeValue<DateTime>("lhind_agreementstartdate") != DateTime.MinValue))
                     ||
                     (postImage.Contains("lhind_agreementstartdate") && (postImage.GetAttributeValue<DateTime>("lhind_agreementstartdate") != DateTime.MinValue))
                    )
                    &&
                    (
                     ((target.Contains("lhind_agreementenddate")) && (target.GetAttributeValue<DateTime>("lhind_agreementenddate") != DateTime.MinValue))
                     ||
                     ((postImage.Contains("lhind_agreementenddate")) && (postImage.GetAttributeValue<DateTime>("lhind_agreementenddate") != DateTime.MinValue))
                    )
                   )
                {
                    tcValue = true;
                }

                EntityReference accountER = postImage.Contains("lhind_relatedaccount") ? postImage.GetAttributeValue<EntityReference>("lhind_relatedaccount") : null;
                if (accountER != null)
                {
                    string fetchXML = $@"
                            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='opportunity'>
                                    <attribute name='opportunityid'/>
                                    <filter type='and'>
                                        <condition attribute='parentaccountid' operator='eq' value='{accountER.Id}'/>
                                    </filter>
                                </entity>
                            </fetch>";

                    FetchExpression fetchExpression = new FetchExpression(fetchXML);

                    EntityCollection results = service.RetrieveMultiple(fetchExpression);

                    foreach (Entity opportunity in results.Entities)
                    {
                        Entity opportunityToUpdate = new Entity(opportunity.LogicalName, opportunity.Id);
                        opportunityToUpdate["lhind_tcs"] = tcValue;
                        service.Update(opportunityToUpdate);
                    }
                }
            }
        }

    }
}
