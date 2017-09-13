namespace CnCrm.WfActivities {
  using System;
  using System.Activities;
  using System.ServiceModel;
  using CnCrm.WfActivities.HelperCode;
  using Microsoft.Xrm.Sdk;
  using Microsoft.Xrm.Sdk.Workflow;
  using Microsoft.Xrm.Sdk.Query;
  using Microsoft.Crm.Sdk.Messages;
  public class SetBizProcessFlowV2 : CodeActivity {
    #region Input/Output Variables

    /// <summary>
    /// Gets or sets the Business Process Flow.
    /// </summary>
    [RequiredArgument]
    [Input("Business Process Flow")]
    [ReferenceTarget("workflow")]
    public InArgument<EntityReference> BpfEntityReference { get; set; }

    /// <summary>
    /// Gets or sets the Business Process Flow Stage.
    /// </summary>
    [RequiredArgument]
    [Input("Business Process Flow Stage Name")]
    public InArgument<string> BpfStageName { get; set; }

    #endregion

    #region Constant Variables

    private const string CHILD_CLASS_NAME = "SetBizProcessFlowV2";
    private const string ACTIVE_STAGE_ID = "activestageid";

    #endregion

    protected override void Execute(CodeActivityContext executionContext) {
      //*** Create the tracing service
      ITracingService tracingService = executionContext.GetExtension<ITracingService>();

      if (tracingService == null) {
        throw new InvalidPluginExecutionException("Failed to retrieve the tracing service.");
      }
      //*** Create the context
      IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();

      if (context == null) {
        throw new InvalidPluginExecutionException("Failed to retrieve the workflow context.");
      }

      tracingService.Trace("{0}.Execute(): ActivityInstanceId: {1}; WorkflowInstanceId: {2}; CorrelationId: {3}; InitiatingUserId: {4} -- Entering", CHILD_CLASS_NAME, executionContext.ActivityInstanceId, executionContext.WorkflowInstanceId, context.CorrelationId, context.InitiatingUserId);

      IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
      IOrganizationService serviceProxy = serviceFactory.CreateOrganizationService(context.UserId);

      if (context.InputParameters.Contains(Common.Target) && context.InputParameters[Common.Target] is Entity) {
        try {
          //*** Grab the Target Entity
          var theEntity = (Entity)context.InputParameters[Common.Target];

          tracingService.Trace("Active Stage Name: {0}", theEntity.EntityState);
          //-------------------------------------------------------------------------------------------------------------
          var processInstancesRequest = new RetrieveProcessInstancesRequest {
            EntityId = theEntity.Id,
            EntityLogicalName = theEntity.LogicalName
          };

          var processInstancesResponse = (RetrieveProcessInstancesResponse)serviceProxy.Execute(processInstancesRequest);
          var processCount = processInstancesResponse.Processes.Entities.Count;

          if (processCount > 0) {
            tracingService.Trace("{0}: Count of Process Instances concurrently associated with the Entity record: {1}", CHILD_CLASS_NAME, processCount);
            tracingService.Trace("{0}: BPF Definition Name currently set for the Entity record: {1}, Id: {2}", CHILD_CLASS_NAME, processInstancesResponse.Processes.Entities[0].Attributes[CrmEarlyBound.Workflow.Fields.Name], processInstancesResponse.Processes.Entities[0].Id.ToString());

            var bpfEntityRef = this.BpfEntityReference.Get<EntityReference>(executionContext);
            var colSet = new ColumnSet();
            colSet.AddColumn(CrmEarlyBound.Workflow.Fields.UniqueName);
            var bpfEntity = serviceProxy.Retrieve(bpfEntityRef.LogicalName, bpfEntityRef.Id, colSet);

            tracingService.Trace("{0}: Switching to BPF Unique Name: {1}, Id: {2}", CHILD_CLASS_NAME, bpfEntity.Attributes[CrmEarlyBound.Workflow.Fields.UniqueName].ToString(), bpfEntity.Id.ToString());

            var bpfStageName = this.BpfStageName.Get<string>(executionContext).Trim();
            var qe = new QueryExpression {
              EntityName = CrmEarlyBound.Workflow.EntityLogicalName,
              ColumnSet = new ColumnSet(new string[] { CrmEarlyBound.Workflow.Fields.Name }),
              Criteria = new FilterExpression {
                Conditions = {
                  new ConditionExpression {
                    AttributeName = CrmEarlyBound.Workflow.Fields.UniqueName, Operator = ConditionOperator.Equal, Values = { bpfEntity.Attributes[CrmEarlyBound.Workflow.Fields.UniqueName] } //new_bpf_472aceaabf7c4f1db4d13ac3c7076c65
                  }
                }
              },
              NoLock = true,
              Distinct = false
            };

            #region Convert Query Expression to FetchXML

            var conversionRequest = new QueryExpressionToFetchXmlRequest {
              Query = qe
            };
            var conversionResponse = (QueryExpressionToFetchXmlResponse)serviceProxy.Execute(conversionRequest);
            var fetchXml = conversionResponse.FetchXml;

            tracingService.Trace("{0}: [{1}], Message: {2}", CHILD_CLASS_NAME, fetchXml, context.MessageName);

            #endregion Convert the query expression to FetchXML.

            tracingService.Trace("{0}: Built BPF Query, Now Executing...", CHILD_CLASS_NAME);

            var entColByQuery = serviceProxy.RetrieveMultiple(qe).Entities; //// Execute Query with Filter Expressions
            //-------------------------------------------------------------------------------------------------------------
            if (entColByQuery != null && entColByQuery.Count > 0) { //// Search and handle related entities
              tracingService.Trace("{0}: Found matching Business Process Flows...", CHILD_CLASS_NAME);

              var bpfId = new Guid();
              var bpfEntityName = String.Empty;

              foreach (var entity in entColByQuery) { //// Loop related entities and retrieve Workflow Names
                bpfId = entity.Id;
                bpfEntityName = entity.GetAttributeValue<string>(CrmEarlyBound.Workflow.Fields.Name);
                break;
              }

              if (bpfId != Guid.Empty) {
                tracingService.Trace("{0}: Successfully retrieved the Business Process Flow that we'll be switching to: {1}, Id: {2}", CHILD_CLASS_NAME, bpfEntityName, bpfId.ToString());

                System.Threading.Thread.Sleep(2000); // Wait for 2 seconds before switching the process
                //*** Set to the new or same Business BpfEntityName Flow
                var setProcReq = new SetProcessRequest {
                  Target = new EntityReference(theEntity.LogicalName, theEntity.Id),
                  NewProcess = new EntityReference(CrmEarlyBound.Workflow.EntityLogicalName, bpfId)
                };

                tracingService.Trace("{0}: ***Ready To Update - Business Process Flow", CHILD_CLASS_NAME);
                var setProcResp = (SetProcessResponse)serviceProxy.Execute(setProcReq);
                tracingService.Trace("{0}: ***Updated", CHILD_CLASS_NAME);
              }
            } else {
              tracingService.Trace("{0}: No Business Process Flows were found with Unique Name: {1}", CHILD_CLASS_NAME, bpfEntity.Attributes[CrmEarlyBound.Workflow.Fields.UniqueName].ToString());
            }
            //-------------------------------------------------------------------------------------------------------------
            //*** Verify if the Process Instance was switched successfully for the Entity record
            processInstancesRequest = new RetrieveProcessInstancesRequest {
              EntityId = theEntity.Id,
              EntityLogicalName = theEntity.LogicalName
            };

            processInstancesResponse = (RetrieveProcessInstancesResponse)serviceProxy.Execute(processInstancesRequest);
            processCount = processInstancesResponse.Processes.Entities.Count;

            if (processCount > 0) {
              var activeProcessInstance = processInstancesResponse.Processes.Entities[0]; //*** First Entity record is the Active Process Instance
              var activeProcessInstanceId = activeProcessInstance.Id; //*** Active Process Instance Id to be used later for retrieval of the active path of the process instance

              tracingService.Trace("{0}: Successfully Switched to '{1}' BPF for the Entity Record.", CHILD_CLASS_NAME, activeProcessInstance.Attributes[CrmEarlyBound.Workflow.Fields.Name]);
              tracingService.Trace("{0}: Count of process instances concurrently associated with the entity record: {1}.", CHILD_CLASS_NAME, processCount);
              var message = "All process instances associated with the entity record:";

              for (var i = 0; i < processCount; i++) {
                message = message + " " + processInstancesResponse.Processes.Entities[i].Attributes[CrmEarlyBound.Workflow.Fields.Name] + ",";
              }

              tracingService.Trace("{0}: {1}", CHILD_CLASS_NAME, message.TrimEnd(message[message.Length - 1]));

              //*** Retrieve the Active Stage ID of the Active Process Instance
              var activeStageId = new Guid(activeProcessInstance.Attributes[CrmEarlyBound.ProcessStage.Fields.ProcessStageId].ToString());
              var activeStagePosition = 0;
              var newStageId = new Guid();
              var newStagePosition = 0;

              //*** Retrieve the BPF Stages in the active path of the Active Process Instance
              var activePathRequest = new RetrieveActivePathRequest {
                ProcessInstanceId = activeProcessInstanceId
              };
              var activePathResponse = (RetrieveActivePathResponse)serviceProxy.Execute(activePathRequest);

              tracingService.Trace("{0}: Retrieved the BPF Stages in the Active Path of the Process Instance:", CHILD_CLASS_NAME);

              for (var i = 0; i < activePathResponse.ProcessStages.Entities.Count; i++) {
                var curStageName = activePathResponse.ProcessStages.Entities[i].Attributes[CrmEarlyBound.ProcessStage.Fields.StageName].ToString();

                tracingService.Trace("{0}: Looping Through Stage #{1}: {2} (StageId: {3}, IndexId: {4})", CHILD_CLASS_NAME, i + 1, curStageName, activePathResponse.ProcessStages.Entities[i].Attributes[CrmEarlyBound.ProcessStage.Fields.ProcessStageId], i);
                //*** Retrieve the Active Stage Name and Stage Position based on a successful match of the activeStageId
                if (activePathResponse.ProcessStages.Entities[i].Attributes[CrmEarlyBound.ProcessStage.Fields.ProcessStageId].Equals(activeStageId)) {
                  activeStagePosition = i;
                  tracingService.Trace("{0}: Concerning the Process Instance -- Initial Active Stage Name: {1} (StageId: {2})", CHILD_CLASS_NAME, curStageName, activeStageId);
                }
                //*** Retrieve the New Stage Id, Stage Name, and Stage Position based on a successful match of the stagename
                if (curStageName.Equals(bpfStageName, StringComparison.InvariantCultureIgnoreCase)) {
                  newStageId = new Guid(activePathResponse.ProcessStages.Entities[i].Attributes[CrmEarlyBound.ProcessStage.Fields.ProcessStageId].ToString());
                  newStagePosition = i;
                  tracingService.Trace("{0}: Concerning the Process Instance -- Desired New Stage Name: {1} (StageId: {2})", CHILD_CLASS_NAME, curStageName, newStageId);
                }
              }
              //-------------------------------------------------------------------------------------------------------------
              //***Update the Business Process Flow Instance record to the desired Active Stage
              Entity retrievedProcessInstance;
              ColumnSet columnSet;
              var stageShift = newStagePosition - activeStagePosition;

              if (stageShift > 0) {
                tracingService.Trace("{0}: Number of Stages Shifting Forward: {1}", CHILD_CLASS_NAME, stageShift);
                //*** Stages only move in 1 direction --> Forward
                for (var i = activeStagePosition; i <= newStagePosition; i++) {
                  System.Threading.Thread.Sleep(1000);
                  //*** Retrieve the Stage Id of the next stage that you want to set as active
                  var newStageName = activePathResponse.ProcessStages.Entities[i].Attributes[CrmEarlyBound.ProcessStage.Fields.StageName].ToString();
                  newStageId = new Guid(activePathResponse.ProcessStages.Entities[i].Attributes[CrmEarlyBound.ProcessStage.Fields.ProcessStageId].ToString());

                  tracingService.Trace("{0}: Setting To Stage #{1}: {2} (StageId: {3}, IndexId: {4})", CHILD_CLASS_NAME, i + 1, newStageName, newStageId, i);
                  //*** Retrieve the BpfEntityName Instance record to update its Active Stage
                  columnSet = new ColumnSet();
                  columnSet.AddColumn(ACTIVE_STAGE_ID);
                  retrievedProcessInstance = serviceProxy.Retrieve(bpfEntity.Attributes[CrmEarlyBound.Workflow.Fields.UniqueName].ToString(), activeProcessInstanceId, columnSet);
                  //*** Set the next Stage as the Active Stage
                  retrievedProcessInstance[ACTIVE_STAGE_ID] = new EntityReference(CrmEarlyBound.ProcessStage.EntityLogicalName, newStageId); //(ProcessStage.EntityLogicalName, activeStageId);

                  try {
                    tracingService.Trace("{0}: ***Ready To Update -- BPF Stage", CHILD_CLASS_NAME);
                    serviceProxy.Update(retrievedProcessInstance);
                    tracingService.Trace("{0}: ***Updated", CHILD_CLASS_NAME);
                  } catch (FaultException<OrganizationServiceFault> ex) { //*** Determine BPF Stage Requirements
                    foreach (var stageAttribute in activePathResponse.ProcessStages.Entities[i].Attributes) {
                      if (stageAttribute.Key.Equals("clientdata")) {
                        tracingService.Trace("{0}: Attribute Key: {1}, Value: {2}", CHILD_CLASS_NAME, stageAttribute.Key, stageAttribute.Value.ToString());
                        break;
                      }
                    }

                    tracingService.Trace(FullStackTraceException.Create(ex).ToString());
                    throw;
                  }
                }
              } else {
                tracingService.Trace("{0}: Number of Stages Shifting Backwards: {1}", CHILD_CLASS_NAME, stageShift);
              }
              //-------------------------------------------------------------------------------------------------------------
              //***Retrieve the Business Process Flow Instance record again to verify its Active Stage information
              columnSet = new ColumnSet();
              columnSet.AddColumn(ACTIVE_STAGE_ID);
              retrievedProcessInstance = serviceProxy.Retrieve(bpfEntity.Attributes[CrmEarlyBound.Workflow.Fields.UniqueName].ToString(), activeProcessInstanceId, columnSet);

              var activeStageEntityRef = retrievedProcessInstance[ACTIVE_STAGE_ID] as EntityReference;

              if (activeStageEntityRef != null) {
                if (activeStageEntityRef.Id.Equals(newStageId)) {
                  tracingService.Trace("{0}: Concerning the Process Instance -- Modified -- Active Stage Name: {1} (StageId: {2})", CHILD_CLASS_NAME, activeStageEntityRef.Name, activeStageEntityRef.Id);
                }
              }
            } else {
              tracingService.Trace("{0}:The RetrieveProcessInstancesRequest object returned 0", CHILD_CLASS_NAME);
            }
          }
        } catch (FaultException<OrganizationServiceFault> ex) {
          tracingService.Trace("{0}: Fault Exception: An Error Occurred During Workflow Activity Execution", CHILD_CLASS_NAME);
          tracingService.Trace("{0}: Fault Timestamp: {1}", CHILD_CLASS_NAME, ex.Detail.Timestamp);
          tracingService.Trace("{0}: Fault Code: {1}", CHILD_CLASS_NAME, ex.Detail.ErrorCode);
          tracingService.Trace("{0}: Fault Message: {1}", CHILD_CLASS_NAME, ex.Detail.Message);
          ////localContext.Trace("{0}: Fault Trace: {1}", this.ChildClassName, ex.Detail.TraceText);
          tracingService.Trace("{0}: Fault Inner Exception: {1}", CHILD_CLASS_NAME, null == ex.Detail.InnerFault ? "No Inner Fault" : "Has Inner Fault");
          //*** Display the details of the inner exception.
          if (ex.InnerException != null) {
            Exception innerEx = ex;
            var i = 0;
            while (innerEx.InnerException != null) {
              innerEx = innerEx.InnerException;
              tracingService.Trace("{0}: Inner Exception: {1}, Message: {2};", CHILD_CLASS_NAME, i++, innerEx.Message);
            }
          }

          throw new InvalidPluginExecutionException(OperationStatus.Failed, ex.Detail.ErrorCode, ex.Message);
        } catch (Exception ex) {
          tracingService.Trace("{0}: Exception: An Error Occurred During Workflow Activity Execution", CHILD_CLASS_NAME);
          tracingService.Trace("{0}: Exception Message: {1}", CHILD_CLASS_NAME, ex.Message);
          //*** Display the details of the inner exception.
          if (ex.InnerException != null) {
            Exception innerEx = ex;
            var i = 0;
            while (innerEx.InnerException != null) {
              innerEx = innerEx.InnerException;
              tracingService.Trace("{0}: Inner Exception: {1}, Message: {2};", CHILD_CLASS_NAME, i++, innerEx.Message);
            }
          }

          throw new InvalidPluginExecutionException(OperationStatus.Failed, ex.HResult, ex.Message);
        } finally {
          tracingService.Trace("{0}.Execute(): ActivityInstanceId: {1}; WorkflowInstanceId: {2}; CorrelationId: {3} -- Exiting", CHILD_CLASS_NAME, executionContext.ActivityInstanceId, executionContext.WorkflowInstanceId, context.CorrelationId);
          // Uncomment to force plugin failure for Debugging
          //--> throw new InvalidPluginExecutionException(String.Format("{0}.Execute(): Plug-in Warning: Manually forcing exception for logging purposes.", CHILD_CLASS_NAME));
        }
      }
    }
  }
}