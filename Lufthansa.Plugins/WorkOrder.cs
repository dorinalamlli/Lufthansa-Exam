using Lufthansa.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Text;

namespace Lufthansa.Plugins
{
    // plugin on create of Work Order (pre operation, sync)
    public class WorkOrder : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            if (context.MessageName.ToLower() == "create")
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    Entity target = (Entity)context.InputParameters["Target"];

                    if ((target.Contains("lhind_scheduledon")) && (target.Contains("lhind_assignedagent")))
                    {
                        int scheduledOn = target.GetAttributeValue<OptionSetValue>("lhind_scheduledon").Value;
                        EntityReference assignedAgentER = target.GetAttributeValue<EntityReference>("lhind_assignedagent");
                        Entity assignedAgent = service.Retrieve(assignedAgentER.LogicalName, assignedAgentER.Id, new ColumnSet("lhind_isscheduledmonday", "lhind_isscheduledtuesday", "lhind_isscheduledwednesday", "lhind_isscheduledthurday", "lhind_isscheduledfriday", "lhind_isscheduledsaturday", "lhind_isscheduledsunday"));

                        bool isCreationAllowed = true;

                        switch (scheduledOn)
                        {
                            case (int)Enums.ScheduledOn.Monday:
                                isCreationAllowed = assignedAgent.Contains("lhind_isscheduledmonday") ? assignedAgent.GetAttributeValue<bool>("lhind_isscheduledmonday") : false;
                                break;
                            case (int)Enums.ScheduledOn.Tuesday:
                                isCreationAllowed = assignedAgent.Contains("lhind_isscheduledtuesday") ? assignedAgent.GetAttributeValue<bool>("lhind_isscheduledtuesday") : false;
                                break;
                            case (int)Enums.ScheduledOn.Wednesday:
                                isCreationAllowed = assignedAgent.Contains("lhind_isscheduledwednesday") ? assignedAgent.GetAttributeValue<bool>("lhind_isscheduledwednesday") : false;
                                break;
                            case (int)Enums.ScheduledOn.Thursday:
                                isCreationAllowed = assignedAgent.Contains("lhind_isscheduledthurday") ? assignedAgent.GetAttributeValue<bool>("lhind_isscheduledthurday") : false;
                                break;
                            case (int)Enums.ScheduledOn.Friday:
                                isCreationAllowed = assignedAgent.Contains("lhind_isscheduledfriday") ? assignedAgent.GetAttributeValue<bool>("lhind_isscheduledfriday") : false;
                                break;
                            case (int)Enums.ScheduledOn.Saturday:
                                isCreationAllowed = assignedAgent.Contains("lhind_isscheduledsaturday") ? assignedAgent.GetAttributeValue<bool>("lhind_isscheduledsaturday") : false;
                                break;
                            case (int)Enums.ScheduledOn.Sunday:
                                isCreationAllowed = assignedAgent.Contains("lhind_isscheduledsunday") ? assignedAgent.GetAttributeValue<bool>("lhind_isscheduledsunday") : false;
                                break;
                            default:
                                break;
                        }

                        if (!isCreationAllowed)
                        {
                            throw new InvalidPluginExecutionException("You can not create this Work Order!");
                        }
                    }
                }
            }
        }
    }
}
